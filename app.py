# ============================================================
#  TRAVEL ORGANIZER v3.3 - the website (what people see).
#  The "engine" is in core.py
# ============================================================
import json
from datetime import datetime, time, timedelta, timezone
from urllib.parse import quote_plus

import pandas as pd
import pydeck as pdk
import streamlit as st

import core

st.set_page_config(page_title="Travel Organizer", page_icon="✈️", layout="wide")

CALENDAR_HELP = """
**Google Calendar (on a computer):**
1. Tip: first create a new calendar for this trip (left side → **Other calendars → + → Create new calendar**).
2. Click ⚙️ **Settings → Import & export → Import**.
3. Select the downloaded `.ics` file, choose your trip calendar, and click **Import**.

**iPhone / iPad:** open the downloaded file and tap **Add All**.

**Outlook:** **Calendar → Add calendar → Upload from file**.

**Changed your reservations?** Delete the trip calendar, recreate it and import the new file.
"""


# ---------- Optional password protection ----------
def password_gate():
    required = st.secrets.get("APP_PASSWORD", "")
    if not required or st.session_state.get("auth_ok"):
        return True
    st.title("✈️ Travel Organizer")
    entered = st.text_input("Enter the app password", type="password")
    if entered == required:
        st.session_state.auth_ok = True
        st.rerun()
    elif entered:
        st.error("Wrong password")
    return False


if not password_gate():
    st.stop()


# ---------- Cached helpers ----------
@st.cache_data(ttl=300, show_spinner=False)
def auto_check_inbox():
    try:
        return core.check_inbox(max_emails=5, days=60)
    except Exception as e:
        return {"enabled": True, "processed": [], "error": str(e)}


@st.cache_data(ttl=3 * 3600, show_spinner=False)
def forecast_for(lat, lon):
    try:
        return core.fetch_forecast(lat, lon)
    except Exception:
        return {}


with st.spinner("📧 Checking the trip inbox for new emails…"):
    inbox_result = auto_check_inbox()


# ---------- Small display helpers ----------
def icon_for(r):
    return core.TYPE_ICONS.get((r.get("type") or "other").lower(), "📌")


def summary_lines(r):
    info = []
    if r.get("end_date") or r.get("end_time"):
        info.append(f"**Until:** {r.get('end_date') or ''} {r.get('end_time') or ''}")
    if r.get("location"):
        info.append(f"**📍 Where:** {r['location']}")
    if r.get("provider"):
        info.append(f"**🏢 Provider:** {r['provider']}")
    if r.get("confirmation_number"):
        info.append(f"**🔖 Confirmation:** `{r['confirmation_number']}`")
    return info


def edit_form(r, approve=False):
    """Form to edit one reservation."""
    rid = r["id"]
    types = list(core.TYPE_ICONS)
    current = (r.get("type") or "other").lower()
    with st.form(f"edit_form_{rid}"):
        c1, c2 = st.columns([1, 3])
        e_type = c1.selectbox("Type", types, key=f"e_type_{rid}",
                              index=types.index(current) if current in types else types.index("other"))
        e_title = c2.text_input("Title", value=r.get("title") or "", key=f"e_title_{rid}")
        c3, c4, c5, c6 = st.columns(4)
        e_sd = c3.date_input("Start date", value=core.parse_date(r.get("start_date")),
                             format="YYYY-MM-DD", key=f"e_sd_{rid}")
        e_st = c4.text_input("Start time (HH:MM)", value=r.get("start_time") or "", key=f"e_st_{rid}")
        e_ed = c5.date_input("End date", value=core.parse_date(r.get("end_date")),
                             format="YYYY-MM-DD", key=f"e_ed_{rid}")
        e_et = c6.text_input("End time (HH:MM)", value=r.get("end_time") or "", key=f"e_et_{rid}")
        c7, c8, c9 = st.columns(3)
        e_loc = c7.text_input("Location", value=r.get("location") or "", key=f"e_loc_{rid}")
        e_prov = c8.text_input("Provider", value=r.get("provider") or "", key=f"e_prov_{rid}")
        e_conf = c9.text_input("Confirmation #", value=r.get("confirmation_number") or "",
                               key=f"e_conf_{rid}")
        e_det = st.text_area("Details", value=r.get("details") or "", key=f"e_det_{rid}")
        label = "💾 Save & approve" if approve else "💾 Save changes"
        if st.form_submit_button(label, type="primary"):
            ok = False
            try:
                core.update_reservation(rid, {
                    "type": e_type, "title": e_title,
                    "start_date": e_sd.isoformat() if e_sd else None, "start_time": e_st,
                    "end_date": e_ed.isoformat() if e_ed else None, "end_time": e_et,
                    "location": e_loc, "provider": e_prov,
                    "confirmation_number": e_conf, "details": e_det}, approve=approve)
                ok = True
            except Exception as e:
                st.error(f"Could not save: {e}")
            if ok:
                st.rerun()


def show_countdown(rows):
    now = core.now_local()
    today = now.date()
    dated = [r for r in rows if core.parse_date(r.get("start_date"))]
    if not dated:
        return
    first = min(core.parse_date(r["start_date"]) for r in dated)
    last = max(core.parse_date(r.get("end_date")) or core.parse_date(r["start_date"]) for r in dated)
    c1, c2 = st.columns(2)
    if today < first:
        n = (first - today).days
        c1.metric("🛫 Trip starts in", f"{n} day" + ("" if n == 1 else "s"))
    elif today <= last:
        c1.metric("🌍 Trip in progress", f"Day {(today - first).days + 1} of {(last - first).days + 1}")
    else:
        c1.metric("✅ Trip finished", f"{(today - last).days} day(s) ago")

    upcoming = []
    for r in dated:
        d = core.parse_date(r["start_date"])
        t = core.parse_time(r.get("start_time"))
        if t:
            when = datetime.combine(d, t)
            if when > now:
                upcoming.append((when, True, r))
        elif d >= today:
            upcoming.append((datetime.combine(d, time(23, 59)), False, r))
    if upcoming:
        when, timed, r = min(upcoming, key=lambda x: x[0])
        if timed:
            mins = int((when - now).total_seconds() // 60)
            days, rest = divmod(mins, 1440)
            hours, minutes = divmod(rest, 60)
            left = f"in {days} d {hours} h" if days else f"in {hours} h {minutes} min"
        else:
            left = "today" if when.date() == today else f"on {when:%a %d %b}"
        c2.markdown(f"**⏭️ Next up:** {icon_for(r)} {r.get('title') or 'Reservation'}  \n{left}")
    st.caption(f"Countdown uses the {core.local_tz()} time zone. Booking times are local to "
               "each place, so it's approximate when you are in another time zone.")


def weather_line(day, places, today):
    p = places.get(day)
    d = core.parse_date(day)
    if not p or not d or not (today <= d <= today + timedelta(days=15)):
        return None
    fc = forecast_for(round(p[0], 2), round(p[1], 2)).get(day)
    if not fc or fc.get("tmax") is None or fc.get("tmin") is None:
        return None
    rain = f" · 💧 {fc['rain']}% chance of rain" if fc.get("rain") is not None else ""
    return (f"{core.weather_icon(fc.get('code'))} {round(fc['tmax'])}° / "
            f"{round(fc['tmin'])}°C{rain} · {p[2]}")


# ============================================================
#  SIDEBAR: trips + backup
# ============================================================
trips = core.load_trips()
trip_names = {t["id"]: t["name"] for t in trips}
trip = None

with st.sidebar:
    st.header("🧳 Trips")
    if trips:
        ids = [t["id"] for t in trips]
        qp = st.query_params.get("trip", "")
        start_index = ids.index(int(qp)) if qp.isdigit() and int(qp) in ids else 0
        trip_id = st.selectbox("Current trip", ids, index=start_index,
                               format_func=lambda i: trip_names[i])
        st.query_params["trip"] = str(trip_id)
        trip = next(t for t in trips if t["id"] == trip_id)
        st.caption("🔗 The link in your address bar now opens this trip directly.")

    with st.expander("➕ Create a new trip", expanded=not trips):
        with st.form("new_trip", clear_on_submit=True):
            new_name = st.text_input("Trip name", placeholder="e.g. Paris Summer 2025")
            if st.form_submit_button("Create trip"):
                if new_name.strip():
                    new_trip = core.create_trip(new_name.strip())
                    st.query_params["trip"] = str(new_trip["id"])
                    st.rerun()
                else:
                    st.error("Please type a name.")

    if trip:
        with st.expander("🗑️ Delete this trip"):
            st.warning("This permanently deletes the trip and ALL its reservations.")
            sure = st.checkbox(f"Yes, delete '{trip['name']}'")
            if st.button("Delete trip", disabled=not sure):
                core.delete_trip(trip["id"])
                st.query_params.clear()
                st.rerun()

    st.divider()
    with st.expander("💾 Backup all data"):
        st.caption("Downloads every trip and reservation as an Excel file. "
                   "Do this before each trip and keep it somewhere safe (e.g. Google Drive).")
        if st.button("Prepare backup file"):
            with st.spinner("Collecting your data…"):
                st.session_state.backup = core.export_backup_excel()
        if "backup" in st.session_state:
            st.download_button(
                "⬇️ Download backup (.xlsx)", st.session_state.backup,
                file_name=f"travel-backup-{core.today_local().isoformat()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if not trip:
    st.title("✈️ Travel Organizer")
    st.info("Create your first trip using the sidebar on the left "
            "(on a phone, tap the **>** arrow at the top-left).")
    st.stop()

trip_id = trip["id"]
all_rows = core.load_reservations(trip_id)
pending = [r for r in all_rows if r.get("status") == core.PENDING]
rows = [r for r in all_rows if r.get("status") != core.PENDING]
other_trips = [t for t in trips if t["id"] != trip_id]

st.title(f"✈️ {trip['name']}")
itin_label = f"📅 Itinerary ({len(pending)} to review)" if pending else "📅 Itinerary"
tab_itin, tab_add, tab_map, tab_chat, tab_ai, tab_email = st.tabs(
    [itin_label, "➕ Add Reservation", "🗺️ Map", "💬 Chat with Claude",
     "🤖 Trip Assistant", "📧 Email Inbox"])


# ============================================================
#  TAB 1: ITINERARY
# ============================================================
with tab_itin:
    # ----- Reservations waiting for review -----
    if pending:
        st.warning(f"📥 **{len(pending)} reservation(s) from emails need your review.** "
                   "They are not in the itinerary, calendar, map or PDF until you approve them.")
        if st.button("✅ Approve all", key="approve_all"):
            core.approve_all(trip_id)
            st.rerun()
        for r in pending:
            with st.container(border=True):
                st.markdown(f"**{icon_for(r)} {r.get('title') or 'Reservation'}** · "
                            f"{r.get('start_date') or 'no date'} {r.get('start_time') or ''}")
                info = summary_lines(r)
                if info:
                    st.markdown("  \n".join(info))
                if r.get("source_file"):
                    st.caption(f"Source: {r['source_file']}")
                b1, b2, _ = st.columns([1, 1, 4])
                if b1.button("✅ Approve", key=f"ap_{r['id']}"):
                    core.approve_reservation(r["id"])
                    st.rerun()
                if b2.button("🗑️ Delete", key=f"pdel_{r['id']}"):
                    core.delete_reservation(r["id"])
                    st.rerun()
                with st.expander("✏️ Edit before approving"):
                    if r.get("details"):
                        st.caption(r["details"])
                    edit_form(r, approve=True)
        st.divider()

    # ----- Countdown + downloads -----
    if rows:
        show_countdown(rows)
    top1, top2, top3 = st.columns([2, 1, 1])
    hide_past = top1.checkbox("Hide past days")
    if rows:
        top2.download_button("📆 Calendar (.ics)", data=core.build_ics(rows, trip["name"]),
                             file_name=f"{trip['email_code']}.ics", mime="text/calendar")
        try:
            pdf_bytes = core.build_pdf(trip["name"], rows)
            top3.download_button("🖨️ Printable PDF", data=pdf_bytes,
                                 file_name=f"{trip['email_code']}-itinerary.pdf",
                                 mime="application/pdf")
        except Exception as e:
            top3.caption(f"PDF not available: {e}")
        with st.expander("ℹ️ How to add this trip to Google / Apple / Outlook calendar"):
            st.markdown(CALENDAR_HELP)
    else:
        st.info("No confirmed reservations in this trip yet. Use the **➕ Add Reservation** "
                "or **📧 Email Inbox** tab.")

    # ----- Day-by-day list -----
    today = core.today_local()
    today_str = today.isoformat()
    places = core.day_places(rows)
    groups = {}
    for r in rows:
        day = str(r.get("start_date") or "Undated")[:10]
        if hide_past and day != "Undated" and day < today_str:
            continue
        groups.setdefault(day, []).append(r)

    far_days = False
    for day in sorted(groups, key=lambda d: (d == "Undated", d)):
        if day == "Undated":
            st.subheader("📌 No date")
        else:
            st.subheader("🗓️ " + core.parse_date(day).strftime("%A, %d %B %Y"))
            w = weather_line(day, places, today)
            if w:
                st.caption(w)
            elif core.parse_date(day) > today + timedelta(days=15):
                far_days = True

        for r in groups[day]:
            with st.container(border=True):
                c1, c2 = st.columns([10, 1])
                time_part = f"{r['start_time']} · " if r.get("start_time") else ""
                c1.markdown(f"### {icon_for(r)} {time_part}{r.get('title') or 'Reservation'}")
                info = summary_lines(r)
                if info:
                    c1.markdown("  \n".join(info))
                if c2.button("🗑️", key=f"del_{r['id']}", help="Delete this reservation"):
                    core.delete_reservation(r["id"])
                    st.rerun()
                with st.expander("✏️ Edit · more details · move"):
                    if r.get("details"):
                        st.write(r["details"])
                    if r.get("source_file"):
                        st.caption(f"Source: {r['source_file']}")
                    edit_form(r)
                    if other_trips:
                        target = st.selectbox("Move to trip", [t["id"] for t in other_trips],
                                              format_func=lambda i: trip_names[i],
                                              key=f"mvsel_{r['id']}")
                        if st.button("➡️ Move", key=f"mv_{r['id']}"):
                            core.move_reservation(r["id"], target)
                            st.rerun()
    if far_days:
        st.caption("☀️ Weather forecasts appear about 2 weeks before each day.")


# ============================================================
#  TAB 2: ADD RESERVATIONS
# ============================================================
def show_save_result(res):
    st.success(f"Saved: {res['added']} new, {res['updated']} updated, "
               f"{res['skipped']} duplicate(s) skipped. See the 📅 Itinerary tab.")


with tab_add:
    st.info(f"New reservations will be saved to the trip: **{trip['name']}**. "
            "Duplicates (same confirmation number and day) are updated instead of copied.")
    st.header("1️⃣ Upload documents")
    st.caption("Supported: PDF, Word (.docx), text (.txt), email (.eml).")
    files = st.file_uploader("Choose one or more files", type=["pdf", "docx", "txt", "eml"],
                             accept_multiple_files=True)
    st.header("…or paste text (e.g. a confirmation email)")
    pasted = st.text_area("Paste reservation text here", height=150)

    if st.button("🤖 Read with Claude", type="primary"):
        if not files and not pasted.strip():
            st.warning("Please upload a file or paste some text first.")
        else:
            found, sources, meter = [], [], []
            with st.spinner("Claude is reading your documents…"):
                for f in files or []:
                    try:
                        found += core.extract_reservations(
                            core.file_to_blocks(f.name, f.getvalue()), meter=meter)
                        sources.append(f.name)
                    except Exception as e:
                        st.error(f"Could not read {f.name}: {e}")
                if pasted.strip():
                    try:
                        found += core.extract_reservations([core.text_block(pasted)], meter=meter)
                        sources.append("pasted text")
                    except Exception as e:
                        st.error(f"Could not read pasted text: {e}")
            if found:
                df = pd.DataFrame(found).reindex(columns=core.FIELDS).fillna("").astype(str)
                st.session_state.review_table = df
                st.session_state.review_source = ", ".join(sources)
                st.session_state.review_trip = trip_id
                st.session_state.review_cost = core.meter_summary(meter)
            else:
                st.warning("Claude didn't find any reservations in that content.")
                if meter:
                    st.caption("💰 " + core.meter_summary(meter))

    if "review_table" in st.session_state:
        target_trip = st.session_state.get("review_trip", trip_id)
        st.subheader(f"✅ Check the details, then save to **{trip_names.get(target_trip, trip['name'])}**")
        if st.session_state.get("review_cost"):
            st.caption("💰 Reading cost: " + st.session_state.review_cost)
        st.caption("Click any cell to edit it. Dates must look like 2025-07-14, times like 14:30.")
        edited = st.data_editor(st.session_state.review_table, num_rows="dynamic", key="editor")
        col_a, col_b = st.columns(2)
        if col_a.button("💾 Save these reservations", type="primary"):
            try:
                with st.spinner("Saving and finding map locations…"):
                    res = core.save_reservations(edited.to_dict("records"),
                                                 st.session_state.review_source, target_trip)
                del st.session_state["review_table"]
                st.session_state.pop("review_cost", None)
                show_save_result(res)
            except Exception as e:
                st.error(f"Save failed: {e}")
        if col_b.button("✖️ Discard"):
            del st.session_state["review_table"]
            st.session_state.pop("review_cost", None)
            st.rerun()

    st.divider()
    st.header("2️⃣ Or add one manually")
    with st.form("manual_form", clear_on_submit=True):
        c1, c2 = st.columns(2)
        m_type = c1.selectbox("Type", list(core.TYPE_ICONS.keys()))
        m_title = c2.text_input("Title *", placeholder="e.g. Flight BA 117 London → New York")
        c3, c4, c5, c6 = st.columns(4)
        m_sd = c3.date_input("Start date *", value=None)
        m_st = c4.time_input("Start time", value=None)
        m_ed = c5.date_input("End date", value=None)
        m_et = c6.time_input("End time", value=None)
        c7, c8, c9 = st.columns(3)
        m_loc = c7.text_input("Location", placeholder="e.g. Eiffel Tower, Paris")
        m_prov = c8.text_input("Provider")
        m_conf = c9.text_input("Confirmation #")
        m_det = st.text_area("Details (tip: include the price, e.g. 'Total INR 12,500')")
        if st.form_submit_button("➕ Add reservation"):
            if not m_title or not m_sd:
                st.error("Title and start date are required.")
            else:
                try:
                    res = core.save_reservations([{
                        "type": m_type, "title": m_title, "start_date": m_sd.isoformat(),
                        "start_time": m_st.strftime("%H:%M") if m_st else None,
                        "end_date": m_ed.isoformat() if m_ed else None,
                        "end_time": m_et.strftime("%H:%M") if m_et else None,
                        "location": m_loc, "provider": m_prov,
                        "confirmation_number": m_conf, "details": m_det,
                    }], "manual entry", trip_id)
                    show_save_result(res)
                except Exception as e:
                    st.error(f"Save failed: {e}")


# ============================================================
#  TAB 3: MAP
# ============================================================
def pick_zoom(spread):
    for limit, zoom in [(100, 1), (40, 2), (15, 3), (5, 5), (1, 7)]:
        if spread > limit:
            return zoom
    return 10


with tab_map:
    missing = [r for r in rows if r.get("location") and r.get("latitude") is None]
    if missing:
        st.info(f"{len(missing)} reservation(s) don't have map coordinates yet.")
        if st.button("📍 Find map locations"):
            with st.spinner("Asking Claude where these places are…"):
                try:
                    n = core.update_missing_coordinates(trip_id)
                    st.success(f"Found {n} location(s).")
                except Exception as e:
                    st.error(f"Could not find locations: {e}")
            st.rerun()

    points = [r for r in rows if r.get("latitude") is not None and r.get("longitude") is not None]
    if not points:
        st.info("Nothing to show on the map yet. Add reservations with a location.")
    else:
        df = pd.DataFrame([{
            "lat": float(r["latitude"]), "lon": float(r["longitude"]),
            "label": f"{icon_for(r)} {r.get('title') or ''}",
            "when": f"{r.get('start_date') or ''} {r.get('start_time') or ''}".strip(),
            "where": r.get("location") or "",
        } for r in points])
        layers = []
        if len(df) > 1:
            segments = pd.DataFrame([{
                "start": [df.lon[i], df.lat[i]], "end": [df.lon[i + 1], df.lat[i + 1]]}
                for i in range(len(df) - 1)])
            layers.append(pdk.Layer("LineLayer", data=segments,
                                    get_source_position="start", get_target_position="end",
                                    get_color=[0, 120, 255, 160], get_width=3))
        layers.append(pdk.Layer("ScatterplotLayer", data=df, get_position="[lon, lat]",
                                get_fill_color=[230, 57, 70, 220], get_radius=30000,
                                radius_min_pixels=7, radius_max_pixels=25, pickable=True))
        spread = max(df.lat.max() - df.lat.min(), df.lon.max() - df.lon.min())
        view = pdk.ViewState(latitude=float(df.lat.mean()), longitude=float(df.lon.mean()),
                             zoom=pick_zoom(spread))
        st.pydeck_chart(pdk.Deck(map_style=None, initial_view_state=view, layers=layers,
                                 tooltip={"text": "{label}\n{when}\n{where}"}))
        st.caption("🔴 Places   🔵 Order of travel. Tap a dot for details. Positions are approximate.")
        st.dataframe(df[["when", "label", "where"]].rename(
            columns={"when": "When", "label": "What", "where": "Where"}), hide_index=True)


# ============================================================
#  TAB 4: CHAT
# ============================================================
with tab_chat:
    st.caption("Claude can see all confirmed reservations in this trip. Try: 'Summarize my trip', "
               "'What do I have on Friday?'")
    chats = st.session_state.setdefault("chats", {})
    history = chats.setdefault(trip_id, [])
    if st.button("🧹 Clear chat"):
        chats[trip_id] = []
        st.rerun()
    for m in history:
        with st.chat_message(m["role"]):
            st.markdown(m["content"])
    question = st.chat_input("Ask Claude about this trip…")
    if question:
        history.append({"role": "user", "content": question})
        with st.chat_message("user"):
            st.markdown(question)
        system_prompt = (
            "You are a friendly, helpful travel assistant inside a travel organizer app.\n"
            + core.trip_context(trip["name"], rows)
            + "Use them to answer accurately. Point out conflicts, tight connections or missing "
              "hotel nights if you notice them. You may also give general travel advice.")
        with st.chat_message("assistant"):
            try:
                with core.claude().messages.stream(model=core.main_model(), max_tokens=2000,
                                                   system=system_prompt,
                                                   messages=history) as stream:
                    reply = st.write_stream(stream.text_stream)
                    final = stream.get_final_message()
                history.append({"role": "assistant", "content": reply})
                chat_meter = []
                core.record_usage(chat_meter, final.model, final.usage)
                st.caption("💰 " + core.meter_summary(chat_meter))
            except Exception as e:
                history.pop()
                st.error(f"Claude error: {e}")


# ============================================================
#  TAB 5: TRIP ASSISTANT
# ============================================================
TOOLS = {
    "🗓️ Fill the gaps": ("gaps", "Finds free time in your schedule and suggests things to do nearby."),
    "🎯 Things to do": ("activities", "All activity ideas around a place in your trip. "
                                      "Add any of them to your itinerary with one click."),
    "🧳 Packing list": ("packing", "A packing list based on your destinations, dates and activities."),
    "⚠️ Health check": ("health", "Flags tight connections, missing hotel nights, timing problems "
                                  "and passport/visa reminders."),
    "💰 Cost summary": ("cost", "Totals the prices found in your bookings, by category."),
    "🌍 Local phrases": ("phrases", "Useful phrases for each country you visit, with pronunciation."),
}
THINGS_TO_DO = "🎯 Things to do"
ACTIVITY_ICONS = {"sightseeing": "🏛️", "food": "🍜", "nature": "🌳", "culture": "🎭",
                  "shopping": "🛍️", "nightlife": "🌃", "family": "👨‍👩‍👧", "adventure": "🧗",
                  "relaxation": "🧘"}
ALL_PLACES = "🌍 All places in my trip"
OTHER_PLACE = "✏️ Type another place…"


def place_label(p):
    f, l = p["first"], p["last"]
    if f and l and f != l:
        return f"{p['label']} ({f:%d %b} – {l:%d %b})"
    if f:
        return f"{p['label']} ({f:%d %b})"
    return p["label"]


def open_things_to_do(label):
    """Button callback: switch to 'Things to do' with a place pre-selected."""
    st.session_state.tool = THINGS_TO_DO
    st.session_state.in_place = label


def show_activities(items, store_key):
    if not items:
        st.info("No suggestions came back. Try Regenerate, or choose a specific place.")
        return
    cats = sorted({str(i.get("category") or "other").lower() for i in items})
    picked = st.multiselect("Show categories", cats, default=cats, key=f"cats_{store_key}")
    for n, a in enumerate(items):
        cat = str(a.get("category") or "other").lower()
        if cat not in picked:
            continue
        name = str(a.get("name") or "Activity")
        sdate = core.parse_date(a.get("suggested_date"))
        with st.container(border=True):
            st.markdown(f"#### {ACTIVITY_ICONS.get(cat, '📌')} {name}")
            facts = [a.get("place"), cat.title(),
                     f"📅 Suggested: {sdate:%a %d %b}" if sdate else None,
                     a.get("best_time"),
                     f"about {a['duration_hours']} h" if a.get("duration_hours") else None,
                     a.get("approx_cost"),
                     "🎟️ Book ahead" if a.get("book_ahead") is True else None]
            st.caption(" · ".join(str(x) for x in facts if x))
            if a.get("description"):
                st.write(a["description"])
            if a.get("why"):
                st.markdown(f"*Why: {a['why']}*")
            if a.get("tip"):
                st.markdown(f"💡 **Tip:** {a['tip']}")
            query = quote_plus(f"{name}, {a.get('address') or a.get('place') or ''}")
            st.link_button("🗺️ Open in Google Maps",
                           f"https://www.google.com/maps/search/?api=1&query={query}")
            with st.expander("➕ Add to my itinerary"):
                with st.form(f"add_act_{store_key}_{n}"):
                    c1, c2 = st.columns(2)
                    a_date = c1.date_input("Date", value=sdate, format="YYYY-MM-DD",
                                           key=f"ad_date_{store_key}_{n}")
                    a_time = c2.time_input("Start time (optional)", value=None,
                                           key=f"ad_time_{store_key}_{n}")
                    if st.form_submit_button("➕ Add", type="primary"):
                        if not a_date:
                            st.error("Please pick a date.")
                        else:
                            end_date = end_time = None
                            try:
                                hours = float(a.get("duration_hours") or 0)
                            except (TypeError, ValueError):
                                hours = 0
                            if a_time and hours > 0:
                                end = datetime.combine(a_date, a_time) + timedelta(hours=hours)
                                end_time = end.strftime("%H:%M")
                                if end.date() != a_date:
                                    end_date = end.date().isoformat()
                            details = " ".join(str(x) for x in [
                                a.get("description"),
                                f"Cost: {a['approx_cost']}." if a.get("approx_cost") else "",
                                f"Best time: {a['best_time']}." if a.get("best_time") else "",
                                "Book ahead!" if a.get("book_ahead") is True else "",
                                f"Tip: {a['tip']}" if a.get("tip") else "",
                                "(Suggested by Claude - check opening hours.)"] if x)
                            try:
                                res = core.save_reservations([{
                                    "type": "restaurant" if cat == "food" else "activity",
                                    "title": name, "start_date": a_date.isoformat(),
                                    "start_time": a_time.strftime("%H:%M") if a_time else None,
                                    "end_date": end_date, "end_time": end_time,
                                    "location": a.get("address") or a.get("place"),
                                    "provider": None, "confirmation_number": None,
                                    "details": details}], "Trip Assistant suggestion", trip_id)
                                if res["added"]:
                                    st.success("Added! Refresh to see it in the 📅 Itinerary tab.")
                                elif res["skipped"]:
                                    st.info("This is already in your itinerary.")
                                else:
                                    st.success("Updated the existing entry.")
                            except Exception as e:
                                st.error(f"Could not add: {e}")


with tab_ai:
    if not rows:
        st.info("Add (and approve) some reservations first.")
    else:
        choice = st.radio("Choose a tool", list(TOOLS), horizontal=True, key="tool")
        kind, description = TOOLS[choice]
        st.caption(description + " Results are saved, so you only pay (a few cents) "
                                 "when you click Generate.")
        trip_place_list = core.trip_places(rows)

        inputs, store_key, ready = {}, kind, True
        if kind == "activities":
            choices = {place_label(p): p["label"] for p in trip_place_list}
            options = [ALL_PLACES] + list(choices) + [OTHER_PLACE]
            if st.session_state.get("in_place") not in options:
                st.session_state.in_place = ALL_PLACES
            sel = st.selectbox("Where?", options, key="in_place")
            if sel == OTHER_PLACE:
                place = st.text_input("Type a place", key="in_place_other",
                                      placeholder="e.g. Montmartre, Paris").strip()
                ready = bool(place)
            elif sel == ALL_PLACES:
                place = "all places"
            else:
                place = choices[sel]
            c1, c2 = st.columns(2)
            inputs["notes"] = c1.text_input("Your interests", key="in_act_int",
                                            placeholder="e.g. food, history, hidden gems")
            inputs["travellers"] = c2.text_input("Who is travelling?", key="in_act_trav",
                                                 placeholder="e.g. 2 adults and a 6-year-old")
            c3, c4 = st.columns(2)
            inputs["budget"] = c3.select_slider("Budget", ["Low", "Medium", "High"],
                                                value="Medium", key="in_budget")
            inputs["count"] = c4.slider("How many ideas", 5, 20, 10, key="in_count",
                                        help="Fewer ideas = lower cost.")
            inputs["place"] = place or "all places"
            store_key = "activities:" + (place or "all places").lower()[:80]
        elif kind == "packing":
            inputs["travellers"] = st.text_input("Who is travelling?", key="in_trav",
                                                 placeholder="e.g. 2 adults and a 6-year-old")
            inputs["notes"] = st.text_input("Anything else?", key="in_pack_notes",
                                            placeholder="e.g. trekking day, formal dinner, carry-on only")
        elif kind == "gaps":
            inputs["notes"] = st.text_input("Your interests", key="in_interests",
                                            placeholder="e.g. food, museums, beaches, kid-friendly")
        elif kind == "health":
            facts = core.trip_facts(rows)
            nights = facts.get("nights_without_accommodation") or []
            if nights:
                st.warning("🏨 Nights with no hotel booked (free check, no Claude): "
                           + ", ".join(nights))
            elif facts:
                st.success("🏨 Every night of the trip has accommodation or overnight travel.")
            inputs["passport"] = st.text_input("Passport country of the travellers", key="in_pass",
                                               placeholder="e.g. India")
        elif kind == "phrases":
            inputs["languages"] = st.text_input("Languages you already speak", key="in_lang",
                                                value="English")

        saved = core.load_insights(trip_id).get(store_key)
        if st.button("🔄 Regenerate" if saved else "✨ Generate", type="primary",
                     disabled=not ready):
            with st.spinner("Claude is working on it… (about 20–60 seconds)"):
                meter = []
                try:
                    content = core.run_insight(kind, trip["name"], rows, meter=meter, **inputs)
                    core.save_insight(trip_id, store_key, content)
                    saved = {"content": content,
                             "updated_at": datetime.now(timezone.utc).isoformat()}
                except Exception as e:
                    st.error(f"Claude error: {e}")
                if meter:
                    st.session_state.setdefault("last_cost", {})[store_key] = \
                        core.meter_summary(meter)

        last_cost = st.session_state.get("last_cost", {}).get(store_key)
        if last_cost:
            st.info("💰 Last generation: " + last_cost)

        if saved:
            st.caption(f"Generated {(saved.get('updated_at') or '')[:16].replace('T', ' ')} UTC. "
                       "Changed your bookings since then? Click Regenerate.")
            if kind == "activities":
                try:
                    items = json.loads(saved["content"] or "[]")
                except json.JSONDecodeError:
                    items = []
                show_activities(items, store_key)
                st.caption("⚠️ Suggestions are from Claude's general knowledge. Check opening "
                           "hours and prices before you go.")
            elif kind == "cost":
                try:
                    items = json.loads(saved["content"] or "[]")
                except json.JSONDecodeError:
                    items = []
                df = pd.DataFrame(items).reindex(
                    columns=["id", "item", "category", "amount", "currency", "note"])
                df["amount"] = pd.to_numeric(df["amount"], errors="coerce")
                df = df.dropna(subset=["amount"])
                if df.empty:
                    st.info("No prices found. Add prices to the reservation details "
                            "(e.g. 'Total INR 12,500') and regenerate.")
                else:
                    df["currency"] = df["currency"].fillna("?").astype(str).str.upper()
                    df["category"] = df["category"].fillna("other")
                    totals = df.groupby("currency")["amount"].sum()
                    cols = st.columns(min(len(totals), 4))
                    for i, (cur, total) in enumerate(totals.items()):
                        cols[i % len(cols)].metric(f"Total ({cur})", f"{total:,.2f}")
                    st.subheader("By category")
                    st.dataframe(df.pivot_table(index="category", columns="currency",
                                                values="amount", aggfunc="sum", fill_value=0))
                    st.subheader("Each booking")
                    st.dataframe(df[["item", "category", "amount", "currency", "note"]],
                                 hide_index=True)
                    unpriced = len(rows) - df["id"].nunique()
                    if unpriced > 0:
                        st.caption(f"{unpriced} reservation(s) had no price in their details. "
                                   "Different currencies are not converted.")
            else:
                st.markdown(saved["content"])
                st.download_button("⬇️ Download as text file", saved["content"],
                                   file_name=f"{trip['email_code']}-{kind}.txt",
                                   mime="text/plain")
            if kind == "health":
                st.caption("⚠️ Always confirm visa and entry rules on the official "
                           "government website. They change often.")

        # ----- From "Fill the gaps": jump to the full list for an area -----
        if kind == "gaps":
            st.divider()
            st.markdown("#### 🎯 Want to see all suggestions for an area?")
            st.caption("Opens the 🎯 Things to do tool with that place selected. "
                       "Then click ✨ Generate.")
            buttons = [(place_label(p), f"🎯 See all things to do in {p['label'][:40]}")
                       for p in trip_place_list]
            buttons.append((ALL_PLACES, "🌍 See all things to do for the whole trip"))
            cols = st.columns(min(len(buttons), 3))
            for i, (label, text) in enumerate(buttons):
                cols[i % len(cols)].button(text, key=f"gap_more_{i}",
                                           on_click=open_things_to_do, args=(label,),
                                           use_container_width=True)


# ============================================================
#  TAB 6: EMAIL INBOX
# ============================================================
with tab_email:
    address = core.trip_address(trip["email_code"])
    if not address:
        st.warning("Email forwarding isn't set up yet. Add EMAIL_ADDRESS and "
                   "EMAIL_APP_PASSWORD to the app's Secrets.")
    else:
        st.markdown(f"Forward booking confirmations for **{trip['name']}** to:")
        st.code(address, language=None)
        st.caption("Emailed reservations appear at the top of the 📅 Itinerary tab for review. "
                   "Emails to the plain inbox address (without +code) go to '📥 Unsorted emails'. "
                   "The inbox is checked when someone opens the app (at most every 5 minutes).")
        if inbox_result.get("error"):
            st.error(f"Couldn't check the inbox: {inbox_result['error']}")
        if st.button("🔄 Check for new emails now", type="primary"):
            with st.spinner("Reading new emails… (about 10–30 seconds per email)"):
                try:
                    res = core.check_inbox(max_emails=10, days=60)
                    if res["processed"]:
                        for item in res["processed"]:
                            st.write(f"• **{item['subject']}** → {item['trip'] or '—'}: "
                                     f"{item['status']}")
                        st.info("Refresh the page to see them in the 📅 Itinerary tab.")
                    else:
                        st.info("No new emails.")
                except Exception as e:
                    st.error(f"Couldn't check the inbox: {e}")
        st.subheader("Recently processed emails")
        log = core.recent_emails(20)
        if log:
            st.dataframe(pd.DataFrame([{
                "When (UTC)": (e.get("processed_at") or "")[:16].replace("T", " "),
                "Subject": e.get("subject") or "",
                "Trip": trip_names.get(e.get("trip_id"), "—"),
                "Result": e.get("status") or "",
            } for e in log]), hide_index=True)
        else:
            st.caption("No emails processed yet.")
