# ============================================================
#  core.py - the "engine" of the Travel Organizer.
#  Used by app.py (the website) and check_email.py (the
#  background email checker that GitHub runs on a schedule).
#  You normally never need to edit this file.
# ============================================================
import base64
import email
import hashlib
import html as html_lib
import imaplib
import io
import json
import os
import re
from datetime import date, datetime, time, timedelta, timezone
from email import policy
from email.utils import getaddresses
from functools import lru_cache

import docx
from anthropic import Anthropic
from supabase import create_client

# ---------- Settings ----------
MODEL = "claude-sonnet-4-5"       # reads documents and chats
FAST_MODEL = "claude-haiku-4-5"   # cheaper model used to find map coordinates
FIELDS = ["type", "title", "start_date", "start_time", "end_date", "end_time",
          "location", "confirmation_number", "provider", "details"]
TYPE_ICONS = {"flight": "✈️", "hotel": "🏨", "car": "🚗", "train": "🚆",
              "bus": "🚌", "activity": "🎟️", "restaurant": "🍽️", "cruise": "🚢",
              "other": "📌"}
UNSORTED_CODE = "unsorted"
MAX_TEXT = 60000          # max characters of text sent to Claude per block
IMAP_HOST = "imap.gmail.com"
MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
          "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


def setting(name, default=""):
    """Read a secret: first from environment variables (GitHub Actions),
    then from Streamlit secrets (the website)."""
    value = os.environ.get(name)
    if value:
        return value
    try:
        import streamlit as st
        return str(st.secrets.get(name, default) or default)
    except Exception:
        return default


@lru_cache(maxsize=1)
def claude():
    return Anthropic(api_key=setting("ANTHROPIC_API_KEY"))


@lru_cache(maxsize=1)
def db():
    return create_client(setting("SUPABASE_URL"), setting("SUPABASE_KEY"))


# ============================================================
#  READING FILES AND EMAILS
# ============================================================
def html_to_text(raw_html):
    raw_html = re.sub(r"(?is)<(script|style).*?</\1>", " ", raw_html)
    raw_html = re.sub(r"(?i)<br\s*/?>|</p>|</div>|</tr>|</li>|</h\d>", "\n", raw_html)
    text = re.sub(r"<[^>]+>", " ", raw_html)
    text = html_lib.unescape(text)
    text = re.sub(r"[ \t]+", " ", text)
    return re.sub(r"\n\s*\n+", "\n\n", text)


def text_block(text):
    text = (text or "").strip()[:MAX_TEXT]
    return {"type": "text", "text": text or "(empty)"}


def pdf_block(data):
    """Claude can read PDFs directly (even scanned ones)."""
    return {"type": "document",
            "source": {"type": "base64", "media_type": "application/pdf",
                       "data": base64.standard_b64encode(data).decode()}}


def docx_to_text(data):
    d = docx.Document(io.BytesIO(data))
    lines = [p.text for p in d.paragraphs]
    for table in d.tables:
        for row in table.rows:
            lines.append(" | ".join(cell.text for cell in row.cells))
    return "\n".join(lines)


def email_to_blocks(msg):
    """Turn an email (including its attachments) into content Claude can read."""
    header = (f"Email subject: {msg.get('subject', '')}\n"
              f"From: {msg.get('from', '')}\nDate: {msg.get('date', '')}\n\n")
    body = ""
    part = msg.get_body(preferencelist=("html", "plain"))
    if part is not None:
        try:
            body = part.get_content()
        except Exception:
            body = ""
        if part.get_content_type() == "text/html":
            body = html_to_text(body)
    blocks = [text_block(header + body)]

    for att in msg.iter_attachments():
        ctype = att.get_content_type()
        fname = (att.get_filename() or "").lower()
        if ctype == "message/rfc822":                 # email forwarded "as attachment"
            inner = att.get_content()
            if hasattr(inner, "get_body"):
                blocks += email_to_blocks(inner)
            continue
        payload = att.get_payload(decode=True)
        if not payload:
            continue
        if ctype == "application/pdf" or fname.endswith(".pdf"):
            blocks.append(pdf_block(payload))
        elif fname.endswith(".docx"):
            blocks.append(text_block(docx_to_text(payload)))
        elif fname.endswith(".txt"):
            blocks.append(text_block(payload.decode("utf-8", errors="ignore")))
    return blocks


def file_to_blocks(name, data):
    """Turn an uploaded file into content Claude can read."""
    lower = name.lower()
    if lower.endswith(".pdf"):
        return [pdf_block(data)]
    if lower.endswith(".docx"):
        return [text_block(docx_to_text(data))]
    if lower.endswith(".eml"):
        return email_to_blocks(email.message_from_bytes(data, policy=policy.default))
    return [text_block(data.decode("utf-8", errors="ignore"))]


# ============================================================
#  CLAUDE HELPERS
# ============================================================
def _ask(model, content, max_tokens):
    response = claude().messages.create(
        model=model, max_tokens=max_tokens,
        messages=[{"role": "user", "content": content}])
    return "".join(b.text for b in response.content if b.type == "text")


def _json_between(text, open_ch, close_ch, default):
    start, end = text.find(open_ch), text.rfind(close_ch)
    if start == -1 or end == -1 or end < start:
        return default
    try:
        return json.loads(text[start:end + 1])
    except json.JSONDecodeError:
        return default


EXTRACT_PROMPT = """You are a travel assistant. Read the reservation document(s) above
and extract EVERY separate reservation (each flight leg, hotel stay, car rental,
train, activity, restaurant booking, etc.). The content may be a forwarded email;
ignore forwarding headers and signatures.

Return ONLY a JSON array (no explanation, no markdown). Each item must be:
{
 "type": "flight | hotel | car | train | bus | activity | restaurant | cruise | other",
 "title": "short description, e.g. 'Flight AI 101 Delhi → London' or 'Hilton Paris Opera'",
 "start_date": "YYYY-MM-DD",
 "start_time": "HH:MM (24h, local time) or null",
 "end_date": "YYYY-MM-DD or null",
 "end_time": "HH:MM (24h, local time) or null",
 "location": "address / city, or for transport: 'From City (CODE) → To City (CODE)'",
 "confirmation_number": "booking reference or null",
 "provider": "airline / hotel / company name",
 "details": "other useful info: terminal, seat, passengers, room type, price, check-in rules, phone"
}

Rules:
- A round-trip flight = 2 items; connecting flights = 1 item per leg.
- Flights/trains: start = departure, end = arrival (include end_date if arrival is the next day).
- Hotel: start = check-in, end = check-out.
- If the year is missing, assume the next upcoming date after today (TODAY).
- If nothing is found, return [].
"""


def extract_reservations(blocks):
    prompt = EXTRACT_PROMPT.replace("TODAY", date.today().isoformat())
    text = _ask(MODEL, blocks + [{"type": "text", "text": prompt}], 8000)
    items = _json_between(text, "[", "]", [])
    return [i for i in items if isinstance(i, dict)] if isinstance(items, list) else []


GEO_PROMPT = """Give the approximate latitude and longitude for each numbered place below.
If a place is a route like "Delhi (DEL) → London (LHR)", use the STARTING place.
Return ONLY a JSON array with exactly {n} items, in the same order.
Each item must be [latitude, longitude], or null if you cannot identify the place.

{places}"""


def geocode_many(places):
    """Ask Claude for map coordinates. Returns {place_text: (lat, lon)}."""
    unique = list(dict.fromkeys(p.strip() for p in places if p and p.strip()))
    result = {}
    for i in range(0, len(unique), 30):
        chunk = unique[i:i + 30]
        listing = "\n".join(f"{n + 1}. {p}" for n, p in enumerate(chunk))
        text = _ask(FAST_MODEL, GEO_PROMPT.format(n=len(chunk), places=listing), 1500)
        coords = _json_between(text, "[", "]", [])
        if not isinstance(coords, list) or len(coords) != len(chunk):
            continue
        for place, c in zip(chunk, coords):
            try:
                lat, lon = float(c[0]), float(c[1])
            except (TypeError, ValueError, IndexError):
                continue
            if -90 <= lat <= 90 and -180 <= lon <= 180:
                result[place] = (lat, lon)
    return result


# ============================================================
#  DATABASE: TRIPS
# ============================================================
def load_trips():
    return db().table("trips").select("*").order("id").execute().data or []


def _slug(name):
    return re.sub(r"[^a-z0-9]", "", name.lower())[:20] or "trip"


def create_trip(name, code=None):
    existing = {t["email_code"] for t in load_trips()}
    base = code or _slug(name)
    code, n = base, 2
    while code in existing or (code == UNSORTED_CODE and base != UNSORTED_CODE):
        code, n = f"{base}{n}", n + 1
    return db().table("trips").insert({"name": name, "email_code": code}).execute().data[0]


def get_or_create_unsorted():
    rows = db().table("trips").select("*").eq("email_code", UNSORTED_CODE).execute().data
    return rows[0] if rows else create_trip("📥 Unsorted emails", UNSORTED_CODE)


def delete_trip(trip_id):
    db().table("trips").delete().eq("id", trip_id).execute()


# ============================================================
#  DATABASE: RESERVATIONS
# ============================================================
def clean(value):
    if value is None or (isinstance(value, float) and value != value):
        return None
    s = str(value).strip()
    return None if s.lower() in ("", "null", "none", "nan", "nat") else s


def normalize_date(value):
    if not value:
        return None
    try:
        return datetime.strptime(str(value).strip()[:10], "%Y-%m-%d").date().isoformat()
    except ValueError:
        raise ValueError(f"'{value}' is not a valid date. Please use YYYY-MM-DD, e.g. 2025-07-14")


def save_reservations(rows, source, trip_id):
    records = []
    for row in rows:
        rec = {f: clean(row.get(f)) for f in FIELDS}
        if not (rec["title"] or rec["start_date"]):
            continue
        rec["type"] = (rec["type"] or "other").lower()
        if rec["type"] not in TYPE_ICONS:
            rec["type"] = "other"
        rec["start_date"] = normalize_date(rec["start_date"])
        rec["end_date"] = normalize_date(rec["end_date"])
        rec.update(source_file=source, trip_id=trip_id, latitude=None, longitude=None)
        records.append(rec)
    if not records:
        return 0
    try:                                   # map coordinates (never blocks saving)
        coords = geocode_many([r["location"] for r in records])
        for r in records:
            c = coords.get((r["location"] or "").strip())
            if c:
                r["latitude"], r["longitude"] = c
    except Exception:
        pass
    db().table("reservations").insert(records).execute()
    return len(records)


def load_reservations(trip_id):
    rows = (db().table("reservations").select("*").eq("trip_id", trip_id)
            .execute().data or [])
    rows.sort(key=lambda r: (r.get("start_date") or "9999-12-31", r.get("start_time") or ""))
    return rows


def delete_reservation(res_id):
    db().table("reservations").delete().eq("id", res_id).execute()


def move_reservation(res_id, new_trip_id):
    db().table("reservations").update({"trip_id": new_trip_id}).eq("id", res_id).execute()


def update_missing_coordinates(trip_id):
    rows = (db().table("reservations").select("id,location").eq("trip_id", trip_id)
            .is_("latitude", "null").execute().data or [])
    rows = [r for r in rows if r.get("location")]
    coords = geocode_many([r["location"] for r in rows])
    updated = 0
    for r in rows:
        c = coords.get(r["location"].strip())
        if c:
            db().table("reservations").update(
                {"latitude": c[0], "longitude": c[1]}).eq("id", r["id"]).execute()
            updated += 1
    return updated


# ============================================================
#  CALENDAR EXPORT (.ics)
# ============================================================
def _ics_text(s):
    return (str(s).replace("\\", "\\\\").replace(";", "\\;").replace(",", "\\,")
            .replace("\r\n", "\\n").replace("\n", "\\n"))


def _fold(line):
    """Calendar files require long lines to be split (max 75 bytes)."""
    out, cur = [], ""
    for ch in line:
        if len((cur + ch).encode("utf-8")) > 74:
            out.append(cur)
            cur = " " + ch
        else:
            cur += ch
    out.append(cur)
    return "\r\n".join(out)


def _parse_time(s):
    if not s:
        return None
    m = re.search(r"(\d{1,2})[:.h](\d{2})", str(s))
    if not m:
        return None
    h, mi = int(m.group(1)), int(m.group(2))
    low = str(s).lower()
    if "pm" in low and h < 12:
        h += 12
    if "am" in low and h == 12:
        h = 0
    return time(h, mi) if h < 24 and mi < 60 else None


def _parse_date(s):
    try:
        return date.fromisoformat(str(s)[:10])
    except (TypeError, ValueError):
        return None


def build_ics(rows, calendar_name):
    stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    lines = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//Travel Organizer//EN",
             "CALSCALE:GREGORIAN", "METHOD:PUBLISH",
             "X-WR-CALNAME:" + _ics_text(calendar_name)]
    for r in rows:
        start_day = _parse_date(r.get("start_date"))
        if not start_day:
            continue
        end_day = _parse_date(r.get("end_date")) or start_day
        if end_day < start_day:
            end_day = start_day
        t1, t2 = _parse_time(r.get("start_time")), _parse_time(r.get("end_time"))
        if t1:                                         # event with a time
            start = datetime.combine(start_day, t1)
            if t2:
                end = datetime.combine(end_day, t2)
            elif end_day > start_day:
                end = datetime.combine(end_day, time(12, 0))
            else:
                end = start + timedelta(hours=1)
            if end <= start:
                end = start + timedelta(hours=1)
            when = [f"DTSTART:{start:%Y%m%dT%H%M%S}", f"DTEND:{end:%Y%m%dT%H%M%S}"]
        else:                                          # all-day event
            when = [f"DTSTART;VALUE=DATE:{start_day:%Y%m%d}",
                    f"DTEND;VALUE=DATE:{end_day + timedelta(days=1):%Y%m%d}"]

        icon = TYPE_ICONS.get((r.get("type") or "other").lower(), "📌")
        desc = "\n".join(x for x in [
            f"Provider: {r['provider']}" if r.get("provider") else "",
            f"Confirmation: {r['confirmation_number']}" if r.get("confirmation_number") else "",
            r.get("details") or "",
            "(Times are local times as shown on the booking.)",
        ] if x)
        lines += ["BEGIN:VEVENT",
                  f"UID:reservation-{r.get('id')}@travel-organizer",
                  f"DTSTAMP:{stamp}", *when,
                  "SUMMARY:" + _ics_text(f"{icon} {r.get('title') or 'Reservation'}")]
        if r.get("location"):
            lines.append("LOCATION:" + _ics_text(r["location"]))
        lines += ["DESCRIPTION:" + _ics_text(desc), "END:VEVENT"]
    lines.append("END:VCALENDAR")
    return "\r\n".join(_fold(line) for line in lines) + "\r\n"


# ============================================================
#  EMAIL INBOX (Gmail)
# ============================================================
def trip_address(code):
    """yourinbox@gmail.com + trip code -> yourinbox+code@gmail.com"""
    address = setting("EMAIL_ADDRESS")
    if "@" not in address:
        return ""
    user, domain = address.split("@", 1)
    return f"{user}+{code}@{domain}"


def _trip_code_from(msg):
    found = []
    for h in ("To", "Cc", "Delivered-To", "X-Original-To", "X-Forwarded-To", "Envelope-To"):
        found += [str(v) for v in (msg.get_all(h) or [])]
    for _, addr in getaddresses(found):
        m = re.search(r"\+([^@+]+)@", addr.lower())
        if m:
            return m.group(1)
    return None


def _claim(message_id):
    """Mark an email as 'being processed'. Returns False if already handled."""
    try:
        db().table("processed_emails").insert(
            {"message_id": message_id, "status": "processing"}).execute()
        return True
    except Exception as e:
        if "duplicate" in str(e).lower() or "23505" in str(e):
            return False
        raise


def recent_emails(limit=20):
    return (db().table("processed_emails").select("*")
            .order("processed_at", desc=True).limit(limit).execute().data or [])


def check_inbox(max_emails=5, days=14):
    """Read new emails from the trip inbox and save their reservations."""
    address = setting("EMAIL_ADDRESS")
    password = setting("EMAIL_APP_PASSWORD").replace(" ", "")
    if not address or not password:
        return {"enabled": False, "processed": []}

    trips_by_code = {t["email_code"]: t for t in load_trips()}
    results = []
    imap = imaplib.IMAP4_SSL(IMAP_HOST)
    try:
        imap.login(address, password)
        imap.select("INBOX", readonly=True)       # read-only: we never change your emails

        since = date.today() - timedelta(days=days)
        since_str = f"{since.day:02d}-{MONTHS[since.month - 1]}-{since.year}"
        _, data = imap.search(None, "SINCE", since_str)
        nums = data[0].split() if data and data[0] else []
        if not nums:
            return {"enabled": True, "processed": []}

        # 1. Get the unique ID of every recent email
        candidates = []
        _, fetched = imap.fetch(",".join(n.decode() for n in nums), "(BODY.PEEK[HEADER])")
        for part in fetched:
            if not isinstance(part, tuple):
                continue
            num = part[0].split()[0].decode()
            hdr = email.message_from_bytes(part[1], policy=policy.default)
            mid = str(hdr.get("Message-ID") or "").strip()
            if not mid:
                mid = "noid-" + hashlib.sha256(part[1]).hexdigest()[:32]
            candidates.append((num, mid))

        # 2. Skip ones we've already processed
        cutoff = (datetime.now(timezone.utc) - timedelta(days=days + 7)).isoformat()
        known = {r["message_id"] for r in (
            db().table("processed_emails").select("message_id")
            .gte("processed_at", cutoff).execute().data or [])}
        new = [(n, m) for n, m in candidates if m not in known][:max_emails]

        # 3. Process each new email
        for num, mid in new:
            if not _claim(mid):
                continue
            subject, trip_id, trip_name = "(no subject)", None, ""
            try:
                _, full = imap.fetch(num, "(BODY.PEEK[])")
                raw = next(p[1] for p in full if isinstance(p, tuple))
                msg = email.message_from_bytes(raw, policy=policy.default)
                subject = str(msg.get("Subject") or "(no subject)")
                items = extract_reservations(email_to_blocks(msg))
                if items:
                    code = _trip_code_from(msg)
                    trip = trips_by_code.get(code) if code else None
                    if trip is None:
                        trip = get_or_create_unsorted()
                        trips_by_code[trip["email_code"]] = trip
                    trip_id, trip_name = trip["id"], trip["name"]
                    n = save_reservations(items, f"email: {subject}"[:200], trip_id)
                    status = f"saved {n} reservation(s)"
                else:
                    status = "no reservations found"
            except Exception as e:
                status = f"error: {e}"[:500]
            db().table("processed_emails").update(
                {"status": status, "subject": subject[:300], "trip_id": trip_id}
            ).eq("message_id", mid).execute()
            results.append({"subject": subject, "trip": trip_name, "status": status})
    finally:
        try:
            imap.logout()
        except Exception:
            pass
    return {"enabled": True, "processed": results}
