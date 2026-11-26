# ============================================================
#  core.py - the "engine" of the Travel Organizer (version 3.2).
#  Used by app.py (the website) and check_email.py.
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
import urllib.parse
import urllib.request
from datetime import date, datetime, time, timedelta, timezone
from email import policy
from email.utils import getaddresses
from functools import lru_cache
from zoneinfo import ZoneInfo

import docx
from anthropic import Anthropic
from supabase import create_client

# ---------- Settings ----------
MODEL = "claude-sonnet-4-5"       # preferred model for reading, chat and the assistant
FAST_MODEL = "claude-haiku-4-5"   # preferred cheaper model (map, optional assistant)
FIELDS = ["type", "title", "start_date", "start_time", "end_date", "end_time",
          "location", "confirmation_number", "provider", "details"]
TYPE_ICONS = {"flight": "✈️", "hotel": "🏨", "car": "🚗", "train": "🚆",
              "bus": "🚌", "activity": "🎟️", "restaurant": "🍽️", "cruise": "🚢",
              "other": "📌"}
CONFIRMED, PENDING = "confirmed", "pending"
UNSORTED_CODE = "unsorted"
MAX_TEXT = 60000
IMAP_HOST = "imap.gmail.com"
MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
          "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


def setting(name, default=""):
    """Read a secret from environment variables (GitHub) or Streamlit secrets."""
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


@lru_cache(maxsize=4)
def resolve_model(preferred, family):
    """Use the preferred model if available, else the newest of the same family."""
    try:
        ids = [m.id for m in claude().models.list(limit=100)]
    except Exception:
        return preferred
    if preferred in ids:
        return preferred
    for model_id in ids:
        if family in model_id:
            return model_id
    return ids[0] if ids else preferred


def main_model():
    return resolve_model(MODEL, "sonnet")


def fast_model():
    return resolve_model(FAST_MODEL, "haiku")


# ============================================================
#  DATES AND TIMES
# ============================================================
def local_tz():
    try:
        return ZoneInfo(setting("APP_TIMEZONE", "UTC").strip() or "UTC")
    except Exception:
        return timezone.utc


def now_local():
    """Current date/time in the app's time zone (APP_TIMEZONE secret)."""
    return datetime.now(local_tz()).replace(tzinfo=None)


def today_local():
    return now_local().date()


def parse_date(s):
    try:
        return date.fromisoformat(str(s)[:10])
    except (TypeError, ValueError):
        return None


def parse_time(s):
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
        if ctype == "message/rfc822":
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
def ask(content, max_tokens=3000, model=None):
    response = claude().messages.create(
        model=model or main_model(), max_tokens=max_tokens,
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
 "details": "other useful info: terminal, seat, passengers, room type, TOTAL PRICE WITH CURRENCY, check-in rules, phone"
}

Rules:
- A round-trip flight = 2 items; connecting flights = 1 item per leg.
- Flights/trains: start = departure, end = arrival (include end_date if arrival is the next day).
- Hotel: start = check-in, end = check-out.
- If the year is missing, assume the next upcoming date after today (TODAY).
- If nothing is found, return [].
"""


def extract_reservations(blocks):
    prompt = EXTRACT_PROMPT.replace("TODAY", today_local().isoformat())
    text = ask(blocks + [{"type": "text", "text": prompt}], 8000)
    items = _json_between(text, "[", "]", [])
    return [i for i in items if isinstance(i, dict)] if isinstance(items, list) else []


GEO_PROMPT = """Give the approximate latitude and longitude for each numbered place below.
If a place is a route like "Delhi (DEL) → London (LHR)", use the STARTING place.
Return ONLY a JSON array with exactly {n} items, in the same order.
Each item must be [latitude, longitude], or null if you cannot identify the place.

{places}"""


def geocode_many(places):
    unique = list(dict.fromkeys(p.strip() for p in places if p and p.strip()))
    result = {}
    for i in range(0, len(unique), 30):
        chunk = unique[i:i + 30]
        listing = "\n".join(f"{n + 1}. {p}" for n, p in enumerate(chunk))
        text = ask(GEO_PROMPT.format(n=len(chunk), places=listing), 1500, fast_model())
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
#  DATABASE: RESERVATIONS (with review + duplicate blocking)
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


def normalize_time(value):
    if not value:
        return None
    t = parse_time(value)
    if t is None:
        raise ValueError(f"'{value}' is not a valid time. Please use HH:MM, e.g. 14:30")
    return t.strftime("%H:%M")


def _prepare_record(row, strict=True):
    """Clean one reservation. strict=False (emails): bad dates/times become notes."""
    rec = {f: clean(row.get(f)) for f in FIELDS}
    if not (rec["title"] or rec["start_date"]):
        return None
    rec["type"] = (rec["type"] or "other").lower()
    if rec["type"] not in TYPE_ICONS:
        rec["type"] = "other"
    notes = []
    for field, fix in (("start_date", normalize_date), ("end_date", normalize_date),
                       ("start_time", normalize_time), ("end_time", normalize_time)):
        try:
            rec[field] = fix(rec[field])
        except ValueError:
            if strict:
                raise
            notes.append(f"{field.replace('_', ' ')} as written: {rec[field]}")
            rec[field] = None
    if notes:
        rec["details"] = "; ".join([x for x in [rec["details"]] if x] + notes)
    return rec


def _norm(s):
    return re.sub(r"[^a-z0-9]", "", str(s or "").lower())


def _same_booking(a, b):
    """Decide whether two reservations describe the same booking."""
    if (a.get("type") or "other") != (b.get("type") or "other"):
        return False
    if str(a.get("start_date") or "")[:10] != str(b.get("start_date") or "")[:10]:
        return False
    ca, cb = _norm(a.get("confirmation_number")), _norm(b.get("confirmation_number"))
    if ca and cb:
        if ca != cb:
            return False
        ta, tb = a.get("start_time"), b.get("start_time")
        if ta and tb and ta != tb:            # same booking ref, other time: maybe another leg
            return _norm(a.get("title")) == _norm(b.get("title"))
        return True
    title = _norm(a.get("title"))
    return bool(title) and title == _norm(b.get("title"))


def _differs(new, old):
    return new is not None and str(new) != str(old if old is not None else "")


def save_reservations(rows, source, trip_id, status=CONFIRMED, strict=True):
    """Save reservations, updating duplicates instead of copying them.
    Returns {'added': n, 'updated': n, 'skipped': n}."""
    existing = db().table("reservations").select("*").eq("trip_id", trip_id).execute().data or []
    added = updated = skipped = 0
    for row in rows:
        rec = _prepare_record(row, strict)
        if rec is None:
            continue
        match = next((e for e in existing if _same_booking(rec, e)), None)
        if match:
            changes = {f: rec[f] for f in FIELDS if _differs(rec[f], match.get(f))}
            if not changes:
                skipped += 1
                continue
            if "location" in changes:
                changes.update(latitude=None, longitude=None)
            if status == PENDING:
                changes["status"] = PENDING      # changed by email: review it again
            changes["source_file"] = source
            db().table("reservations").update(changes).eq("id", match["id"]).execute()
            match.update(changes)
            updated += 1
            continue
        rec.update(source_file=source, trip_id=trip_id, status=status,
                   latitude=None, longitude=None)
        inserted = db().table("reservations").insert(rec).execute().data
        existing.append(inserted[0] if inserted else rec)
        added += 1
    if added or updated:
        try:
            update_missing_coordinates(trip_id)
        except Exception:
            pass
    return {"added": added, "updated": updated, "skipped": skipped}


def load_reservations(trip_id):
    rows = (db().table("reservations").select("*").eq("trip_id", trip_id)
            .execute().data or [])
    rows.sort(key=lambda r: (r.get("start_date") or "9999-12-31", r.get("start_time") or ""))
    return rows


def update_reservation(res_id, row, approve=False):
    rec = _prepare_record(row, strict=True)
    if rec is None:
        raise ValueError("Please enter at least a title or a start date.")
    old = db().table("reservations").select("location,trip_id").eq("id", res_id).execute().data
    if old and (old[0].get("location") or "") != (rec["location"] or ""):
        rec.update(latitude=None, longitude=None)
    if approve:
        rec["status"] = CONFIRMED
    db().table("reservations").update(rec).eq("id", res_id).execute()
    if old:
        try:
            update_missing_coordinates(old[0]["trip_id"])
        except Exception:
            pass


def approve_reservation(res_id):
    db().table("reservations").update({"status": CONFIRMED}).eq("id", res_id).execute()


def approve_all(trip_id):
    (db().table("reservations").update({"status": CONFIRMED})
     .eq("trip_id", trip_id).eq("status", PENDING).execute())


def delete_reservation(res_id):
    db().table("reservations").delete().eq("id", res_id).execute()


def move_reservation(res_id, new_trip_id):
    db().table("reservations").update({"trip_id": new_trip_id}).eq("id", res_id).execute()


def update_missing_coordinates(trip_id):
    rows = (db().table("reservations").select("id,location").eq("trip_id", trip_id)
            .is_("latitude", "null").execute().data or [])
    rows = [r for r in rows if r.get("location")]
    if not rows:
        return 0
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
    out, cur = [], ""
    for ch in line:
        if len((cur + ch).encode("utf-8")) > 74:
            out.append(cur)
            cur = " " + ch
        else:
            cur += ch
    out.append(cur)
    return "\r\n".join(out)


def build_ics(rows, calendar_name):
    stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    lines = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//Travel Organizer//EN",
             "CALSCALE:GREGORIAN", "METHOD:PUBLISH",
             "X-WR-CALNAME:" + _ics_text(calendar_name)]
    for r in rows:
        start_day = parse_date(r.get("start_date"))
        if not start_day:
            continue
        end_day = parse_date(r.get("end_date")) or start_day
        if end_day < start_day:
            end_day = start_day
        t1, t2 = parse_time(r.get("start_time")), parse_time(r.get("end_time"))
        if t1:
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
        else:
            when = [f"DTSTART;VALUE=DATE:{start_day:%Y%m%d}",
                    f"DTEND;VALUE=DATE:{end_day + timedelta(days=1):%Y%m%d}"]
        icon = TYPE_ICONS.get((r.get("type") or "other").lower(), "📌")
        desc = "\n".join(x for x in [
            f"Provider: {r['provider']}" if r.get("provider") else "",
            f"Confirmation: {r['confirmation_number']}" if r.get("confirmation_number") else "",
            r.get("details") or "",
            "(Times are local times as shown on the booking.)",
        ] if x)
        lines += ["BEGIN:VEVENT", f"UID:reservation-{r.get('id')}@travel-organizer",
                  f"DTSTAMP:{stamp}", *when,
                  "SUMMARY:" + _ics_text(f"{icon} {r.get('title') or 'Reservation'}")]
        if r.get("location"):
            lines.append("LOCATION:" + _ics_text(r["location"]))
        lines += ["DESCRIPTION:" + _ics_text(desc), "END:VEVENT"]
    lines.append("END:VCALENDAR")
    return "\r\n".join(_fold(line) for line in lines) + "\r\n"


# ============================================================
#  WEATHER (free Open-Meteo service, no account needed)
# ============================================================
def weather_icon(code):
    try:
        c = int(code)
    except (TypeError, ValueError):
        return "🌡️"
    if c == 0:
        return "☀️"
    if c in (1, 2):
        return "🌤️"
    if c == 3:
        return "☁️"
    if c in (45, 48):
        return "🌫️"
    if 51 <= c <= 57:
        return "🌦️"
    if 61 <= c <= 67 or 80 <= c <= 82:
        return "🌧️"
    if 71 <= c <= 77 or c in (85, 86):
        return "❄️"
    if c >= 95:
        return "⛈️"
    return "🌡️"


def fetch_forecast(lat, lon):
    """16-day forecast. Returns {'YYYY-MM-DD': {'code', 'tmax', 'tmin', 'rain'}}."""
    params = urllib.parse.urlencode({
        "latitude": lat, "longitude": lon, "timezone": "auto", "forecast_days": 16,
        "daily": "weather_code,temperature_2m_max,temperature_2m_min,"
                 "precipitation_probability_max"})
    req = urllib.request.Request("https://api.open-meteo.com/v1/forecast?" + params,
                                 headers={"User-Agent": "travel-organizer"})
    with urllib.request.urlopen(req, timeout=10) as resp:
        daily = json.loads(resp.read().decode()).get("daily", {})

    def pick(key, i):
        values = daily.get(key) or []
        return values[i] if i < len(values) else None

    return {day: {"code": pick("weather_code", i), "tmax": pick("temperature_2m_max", i),
                  "tmin": pick("temperature_2m_min", i),
                  "rain": pick("precipitation_probability_max", i)}
            for i, day in enumerate(daily.get("time", []))}


def day_places(rows):
    """For each day: where you'll be (hotel that night, else first booking that day)."""
    with_coords = [r for r in rows if r.get("latitude") is not None
                   and r.get("longitude") is not None and parse_date(r.get("start_date"))]
    stays = [r for r in with_coords if (r.get("type") or "") in ("hotel", "cruise")]
    result = {}
    for day in sorted({str(r["start_date"])[:10] for r in rows if parse_date(r.get("start_date"))}):
        d = parse_date(day)
        pick = None
        for h in stays:
            s = parse_date(h["start_date"])
            e = parse_date(h.get("end_date")) or s
            if s <= d < e or s == d:
                pick = h
                break
        if pick is None:
            pick = next((r for r in with_coords if str(r["start_date"])[:10] == day), None)
        if pick:
            label = (pick.get("location") or "").split("→")[0].split("->")[0].strip()[:40]
            result[day] = (float(pick["latitude"]), float(pick["longitude"]), label)
    return result


# ============================================================
#  TRIP ASSISTANT (Claude tools)
# ============================================================
def trip_facts(rows):
    """Checks the app can do reliably by itself (given to Claude as facts)."""
    dated = [r for r in rows if parse_date(r.get("start_date"))]
    if not dated:
        return {}
    first = min(parse_date(r["start_date"]) for r in dated)
    last = max(parse_date(r.get("end_date")) or parse_date(r["start_date"]) for r in dated)
    covered = set()
    for r in dated:
        if (r.get("type") or "") in ("hotel", "cruise", "flight", "train", "bus"):
            s = parse_date(r["start_date"])
            e = parse_date(r.get("end_date")) or s
            d = s
            while d < e and (d - s).days < 366:
                covered.add(d)
                d += timedelta(days=1)
    missing, d = [], first
    while d < last and (d - first).days < 366:
        if d not in covered:
            missing.append(d.isoformat())
        d += timedelta(days=1)
    return {
        "trip_start": first.isoformat(), "trip_end": last.isoformat(),
        "nights_without_accommodation": missing,
        "bookings_without_date": [r.get("title") for r in rows if not parse_date(r.get("start_date"))],
        "bookings_without_confirmation_number": [
            r.get("title") for r in dated if not r.get("confirmation_number")
            and r.get("type") in ("flight", "hotel", "train", "car", "cruise")],
    }


def trip_places(rows):
    """Places the traveller stays at or visits, with the dates they are there."""
    places = {}
    for r in rows:
        loc = (r.get("location") or "").strip()
        kind = (r.get("type") or "other").lower()
        d = parse_date(r.get("start_date"))
        if not loc or kind == "car":
            continue
        if kind in ("flight", "train", "bus", "cruise"):
            parts = re.split(r"→|->", loc)
            if len(parts) < 2:
                continue
            loc = parts[-1].strip()                  # the destination
            d = parse_date(r.get("end_date")) or d
        label = loc[:80]
        key = _norm(label)
        if not key:
            continue
        p = places.setdefault(key, {"label": label, "first": d, "last": d})
        end = parse_date(r.get("end_date")) if kind == "hotel" else None
        for x in (d, end):
            if x:
                p["first"] = min(p["first"] or x, x)
                p["last"] = max(p["last"] or x, x)
    return sorted(places.values(), key=lambda p: (p["first"] or date.max, p["label"]))


INSIGHT_PROMPTS = {
    "packing": """Create a practical packing list for this trip. Consider the destinations,
the dates/season and likely weather, the length of the trip and the activities in the bookings.
Travellers: {travellers}. Extra notes: {notes}.
Group it into sections with headings (Documents & money, Clothing, Toiletries, Electronics,
Health, Activity-specific, Kids - only if relevant). Use '- [ ] item' lines with quantities
where useful. Finish with 3-5 short tips specific to this trip. Respond in Markdown only.""",

    "gaps": """Go through the schedule day by day and find free time of roughly 3 hours or more
during the daytime (about 08:00-22:00). Allow for travel to/from airports and stations,
check-in/check-out, and rest after long or overnight journeys.
For each free slot give: the day and approximate time window, where the traveller will be
(based on their hotel or latest booking), and 3 suggestions nearby - what it is, why it's good,
rough travel time and rough cost. Say when something should be booked in advance, and remind
them to check opening hours. Traveller interests: {notes}.
If the schedule is fully booked, say so. Markdown, with a heading per day.""",

    "health": """Do a careful health check of this itinerary.
Facts calculated by the app (reliable): {facts}
Check for: tight connections (under ~1 h domestic / ~2 h international between flights, or from
arrival to the next departure), overlapping bookings, hotel check-out after departure or check-in
long after arrival, nights without accommodation, bookings missing dates/times/confirmation
numbers, missing transfers (e.g. airport to hotel) and anything else unusual.
Then give travel-document reminders for travellers holding a passport from: {passport}.
Cover passport validity (many countries require 6 months), visas/e-visas and transit visas.
If the passport country is not specified, give general advice and say rules depend on nationality.
Always tell them to confirm entry rules on the official government website, as rules change.
Markdown with these sections: '🔴 Problems', '🟡 Worth checking', '🛂 Documents & entry rules',
'🟢 Looks good'. Be specific and mention dates and booking titles.""",

    "phrases": """Work out which countries this trip visits and their main local languages.
Skip languages the traveller already speaks: {languages}.
For each language give a Markdown table of about 20 useful phrases with the columns:
English | Local language (in local script) | How to say it (simple English pronunciation).
Include greetings, please/thank you, sorry/excuse me, do you speak English, numbers 1-10 (one row),
how much is this, the bill please, where is the toilet, help, I am allergic to..., directions,
and phrases related to these bookings (e.g. hotel check-in, taxi to the airport).
Add 3 quick etiquette tips per country. If the traveller already speaks the languages of
every destination, say so and give local etiquette tips instead.""",

    "cost": """Find the price of each reservation, if it is mentioned in any field (especially 'details').
Return ONLY a JSON array with one item per reservation that has a price:
{"id": reservation id, "item": "short name", "category": "flight|hotel|car|train|bus|activity|restaurant|cruise|other",
 "amount": number (total paid for that booking, no currency symbol), "currency": "3-letter code e.g. INR, USD, EUR",
 "note": "e.g. 'includes taxes', 'deposit only', 'per night x 3'"}
If a price is per night or per person and the total is clear, give the total.
Never guess prices that are not mentioned. Return [] if no prices are found.""",

    "activities": """Suggest things to do for this traveller.
Area to focus on: {place}
(If it says 'all places', cover every city or area the traveller stays in or visits,
spread fairly across them.)
Traveller interests: {notes}. Who is travelling: {travellers}. Budget level: {budget}.
Give about {count} suggestions. Mix famous highlights with local favourites, and include at
least one food experience per area. Do not suggest things that are already booked.
Choose a sensible suggested_date when the traveller is actually in that area (use the
reservations), preferring days that are not already busy. Consider the season, the weather
and the days places are usually closed.
Return ONLY a JSON array (no markdown, no explanation). Each item:
{"place": "city or area", "name": "name of the place or experience",
 "category": "sightseeing|food|nature|culture|shopping|nightlife|family|adventure|relaxation",
 "description": "1-2 sentences", "why": "why it suits this traveller and trip",
 "address": "address or area, good enough to search on a map",
 "suggested_date": "YYYY-MM-DD or null", "best_time": "e.g. 'early morning', 'sunset'",
 "duration_hours": number, "approx_cost": "e.g. 'Free' or 'about EUR 20 per person'",
 "book_ahead": true or false, "tip": "one practical tip"}""",
}


def trip_context(trip_name, rows):
    slim = [{"id": r.get("id"), **{f: r.get(f) for f in FIELDS}} for r in rows]
    return (f"Trip name: {trip_name}\nToday's date: {today_local().isoformat()}\n"
            "Times are local times at each place.\n"
            f"Reservations (JSON):\n{json.dumps(slim, default=str, ensure_ascii=False)}\n\n")


def run_insight(kind, trip_name, rows, **inputs):
    values = {k: (str(v).strip() if v else "") or "not specified" for k, v in inputs.items()}
    if kind == "health":
        values["facts"] = json.dumps(trip_facts(rows), default=str)
    prompt = INSIGHT_PROMPTS[kind]
    for key in ("travellers", "notes", "passport", "languages", "facts",
                "place", "budget", "count"):
        prompt = prompt.replace("{" + key + "}", values.get(key, "not specified"))
    prompt = trip_context(trip_name, rows) + prompt
    # Cheaper model if the ASSISTANT_MODEL secret is "haiku".
    # The health check always uses the main model (visa advice needs extra care).
    cheap = setting("ASSISTANT_MODEL", "sonnet").strip().lower() == "haiku"
    model = fast_model() if cheap and kind != "health" else main_model()
    if kind in ("cost", "activities"):
        limit = 8000 if kind == "activities" else 3000
        items = _json_between(ask(prompt, limit, model), "[", "]", [])
        items = [i for i in items if isinstance(i, dict)] if isinstance(items, list) else []
        return json.dumps(items, ensure_ascii=False)
    return ask(prompt, 6000, model)


def load_insights(trip_id):
    rows = db().table("trip_insights").select("*").eq("trip_id", trip_id).execute().data or []
    return {r["kind"]: r for r in rows}


def save_insight(trip_id, kind, content):
    db().table("trip_insights").upsert(
        {"trip_id": trip_id, "kind": kind, "content": content,
         "updated_at": datetime.now(timezone.utc).isoformat()},
        on_conflict="trip_id,kind").execute()


# ============================================================
#  PRINTABLE PDF ITINERARY
# ============================================================
_LATIN_SWAPS = {"→": "->", "–": "-", "—": "-", "₹": "Rs.", "€": "EUR ", "’": "'",
                "‘": "'", "“": '"', "”": '"', "•": "-", "…": "..."}


def _pdf_fonts():
    for folder in ("/usr/share/fonts/truetype/dejavu", "/usr/share/fonts/dejavu",
                   "/usr/share/fonts/TTF"):
        regular = os.path.join(folder, "DejaVuSans.ttf")
        bold = os.path.join(folder, "DejaVuSans-Bold.ttf")
        if os.path.exists(regular) and os.path.exists(bold):
            return regular, bold
    return None


def build_pdf(trip_name, rows):
    from fpdf import FPDF
    from fpdf.enums import XPos, YPos

    pdf = FPDF(format="A4")
    pdf.set_auto_page_break(auto=True, margin=15)
    fonts = _pdf_fonts()
    if fonts:
        pdf.add_font("DejaVu", "", fonts[0])
        pdf.add_font("DejaVu", "B", fonts[1])
        family = "DejaVu"
    else:
        family = "Helvetica"

    def safe(s):
        s = str(s or "")
        if fonts:   # remove emojis (fonts can't draw them)
            return "".join(ch for ch in s if ord(ch) <= 0xFFFF
                           and not 0xFE00 <= ord(ch) <= 0xFE0F and ord(ch) != 0x200D)
        for a, b in _LATIN_SWAPS.items():
            s = s.replace(a, b)
        return s.encode("latin-1", "replace").decode("latin-1")

    def write(text, size=10, bold=False, indent=0, height=5.5, fill=False):
        pdf.set_font(family, "B" if bold else "", size)
        pdf.set_x(pdf.l_margin + indent)
        pdf.multi_cell(0, height, safe(text), fill=fill,
                       new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.add_page()
    pdf.set_fill_color(225, 236, 250)
    write(trip_name, size=18, bold=True, height=9)
    dated = [r for r in rows if parse_date(r.get("start_date"))]
    if dated:
        first = min(parse_date(r["start_date"]) for r in dated)
        last = max(parse_date(r.get("end_date")) or parse_date(r["start_date"]) for r in dated)
        write(f"{first:%d %b %Y} - {last:%d %b %Y}", size=11)
    write(f"Printed {today_local():%d %b %Y}. Times are local times as shown on each booking.",
          size=8.5)
    pdf.ln(3)

    groups = {}
    for r in rows:
        groups.setdefault(str(r.get("start_date") or "")[:10], []).append(r)
    for day in sorted(groups, key=lambda d: (d == "", d)):
        pdf.ln(2)
        d = parse_date(day)
        write(d.strftime("%A, %d %B %Y") if d else "No date", size=12.5, bold=True,
              height=8, fill=True)
        pdf.ln(1)
        for r in groups[day]:
            kind = (r.get("type") or "other").upper()
            write(f"{r.get('start_time') or '--:--'}   {kind}   {r.get('title') or 'Reservation'}",
                  size=11, bold=True, height=6.5)
            lines = []
            if r.get("end_date") or r.get("end_time"):
                lines.append(f"Until: {r.get('end_date') or ''} {r.get('end_time') or ''}".strip())
            if r.get("location"):
                lines.append(f"Where: {r['location']}")
            if r.get("provider"):
                lines.append(f"Provider: {r['provider']}")
            if r.get("confirmation_number"):
                lines.append(f"Confirmation: {r['confirmation_number']}")
            if r.get("details"):
                det = str(r["details"])
                lines.append("Notes: " + (det[:700] + "..." if len(det) > 700 else det))
            for text in lines:
                write(text, size=9.5, indent=6, height=5)
            pdf.ln(2)
    return bytes(pdf.output())


# ============================================================
#  BACKUP (Excel file with all data)
# ============================================================
def _all_rows(table):
    out, start = [], 0
    while True:
        chunk = (db().table(table).select("*").order("id")
                 .range(start, start + 999).execute().data or [])
        out += chunk
        if len(chunk) < 1000:
            return out
        start += 1000


def export_backup_excel():
    import pandas as pd
    trips = _all_rows("trips")
    names = {t["id"]: t["name"] for t in trips}
    reservations = _all_rows("reservations")
    for r in reservations:
        r["trip_name"] = names.get(r.get("trip_id"), "")
    insights = _all_rows("trip_insights")
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        pd.DataFrame(reservations).to_excel(xw, sheet_name="reservations", index=False)
        pd.DataFrame(trips).to_excel(xw, sheet_name="trips", index=False)
        pd.DataFrame(insights).to_excel(xw, sheet_name="assistant_results", index=False)
    return buf.getvalue()


# ============================================================
#  EMAIL INBOX (Gmail)
# ============================================================
def trip_address(code):
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


def check_inbox(max_emails=5, days=60):
    """Read new emails; their reservations are saved as 'pending' for review."""
    address = setting("EMAIL_ADDRESS")
    password = setting("EMAIL_APP_PASSWORD").replace(" ", "")
    if not address or not password:
        return {"enabled": False, "processed": []}

    trips_by_code = {t["email_code"]: t for t in load_trips()}
    results = []
    imap = imaplib.IMAP4_SSL(IMAP_HOST)
    try:
        imap.login(address, password)
        imap.select("INBOX", readonly=True)
        since = date.today() - timedelta(days=days)
        since_str = f"{since.day:02d}-{MONTHS[since.month - 1]}-{since.year}"
        _, data = imap.search(None, "SINCE", since_str)
        nums = data[0].split() if data and data[0] else []
        if not nums:
            return {"enabled": True, "processed": []}

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

        cutoff = (datetime.now(timezone.utc) - timedelta(days=days + 7)).isoformat()
        known = {r["message_id"] for r in (
            db().table("processed_emails").select("message_id")
            .gte("processed_at", cutoff).execute().data or [])}
        new = [(n, m) for n, m in candidates if m not in known][:max_emails]

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
                    res = save_reservations(items, f"email: {subject}"[:200], trip_id,
                                            status=PENDING, strict=False)
                    status = (f"saved for review: {res['added']} new, {res['updated']} updated, "
                              f"{res['skipped']} duplicate(s) skipped")
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
