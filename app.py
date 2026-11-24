# ============================================================
#  TRAVEL ORGANIZER - the website (what people see).
#  The "engine" (reading files, emails, calendar, map) is in core.py
# ============================================================
import json
from datetime import date

import pandas as pd
import pydeck as pdk
import streamlit as st

import core

st.set_page_config(page_title="Travel Organizer", page_icon="✈️", layout="wide")

CALENDAR_HELP = """
**Google Calendar (on a computer):**
1. Tip: first create a new calendar for this trip (left side → **Other calendars → + → Create new calendar**). This makes it easy to remove later.
2. Click ⚙️ **Settings → Import & export → Import**.
3. Select the downloaded `.ics` file, choose your trip calendar, and click **Import**.

**iPhone / iPad:** open the downloaded file and tap **Add All**.

**Outlook:** go to **Calendar → Add calendar → Upload from file**, then choose the `.ics` file.

**Changed your reservations?** Delete the trip calendar, recreate it, and import the new file. This avoids duplicates.
Times appear exactly as on your bookings (local time at each place).
"""


# ---------- Optional password protection ----------
def password_gate():
    required = st.secrets.get("APP_PASSWORD", "")
    if not required:
        return True
    if st.session_state.get("auth_ok"):
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


# ---------- Automatic inbox check (at most once every 5 minutes) ----------
@st.cache_data(ttl=300, show_spinner=False)
def auto_check_inbox():
    try:
        return core.check_inbox(max_emails=5)
    except Exception as e:
        return {"enabled": True, "processed": [], "error": str(e)}


with st.spinner("📧 Checking the trip inbox for new emails…"):
    inbox_result = auto_check_inbox()


# ============================================================
#  SIDEBAR: choose / create / delete trips
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
        st.caption("🔗 Tip: the link in your browser's address bar now opens this trip directly.")

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

if not trip:
    st.title("✈️ Travel Organizer")
    st.info("Create your first trip using the sidebar on the left "
            "(on a phone, tap the **>** arrow at the top-left).")
    st.stop()

trip_id = trip["id"]
rows = core.load_reservations(trip_id)
other_trips = [t for t in trips if t["id"] != trip_id]

st.title(f"✈️ {trip['name']}")
tab_itin, tab_add, tab_map, tab_chat, tab_email = st.tabs(
    ["📅 Itinerary", "➕ Add Reservation", "🗺️ Map", "💬 Chat with Claude", "📧 Email Inbox"])


# ============================================================
#  TAB 1: ITINERARY + CALENDAR EXPORT
# ============================================================
with tab_itin:
    top1, top2 = st.columns([3, 2])
    hide_past = top1.checkbox("Hide past days")
    if rows:
        top2.download_button("📆 Download calendar file (.ics)",
                             data=core.build_ics(rows, trip["name"]),
                             file_name=f"{trip['email_code']}.ics",
                             mime="text/calendar", use_container_width=True)
        with st.expander("ℹ️ How to add this trip to Google / Apple / Outlook calendar"):
            st.markdown(CALENDAR_HELP)
    else:
        st.info("No reservations in this trip yet. Use the **➕ Add Reservation** "
                "or **📧 Email Inbox** tab.")

    today_str = date.today().isoformat()
    groups = {}
    for r in rows:
        day = r.get("start_date") or "Undated"
        if hide_past and day != "Undated" and day < today_str:
            continue
        groups.setdefault(day, []).append(r)

    for day in sorted(groups, key=lambda d: (d == "Undated", d)):
        heading = ("📌 No date" if day == "Undated"
                   else "🗓️ " + date.fromisoformat(day).strftime("%A, %d %B %Y"))
        st.subheader(heading)

        for r in groups[day]:
            icon = core.TYPE_ICONS.get((r.get("type") or "other").lower(), "📌")
            with st.container(border=True):
                c1, c2 = st.columns([10, 1])
                time_part = f"{r['start_time']} · " if r.get("start_time") else ""
                c1.markdown(f"### {icon} {time_part}{r.get('title') or 'Reservation'}")

                info = []
                if r.get("end_date") or r.get("end_time"):
                    info.append(f"**Until:** {r.get('end_date') or ''} {r.get('end_time') or ''}")
                if r.get("location"):
                    info.append(f"**📍 Where:** {r['location']}")
                if r.get("provider"):
                    info.append(f"**🏢 Provider:** {r['provider']}")
                if r.get("confirmation_number"):
                    info.append(f"**🔖 Confirmation:** `{r['confirmation_number']}`")
                if info:
                    c1.markdown("  \n".join(info))

                with c1.expander("More details / move to another trip"):
                    if r.get("details"):
                        st.write(r["details"])
                    if r.get("source_file"):
                        st.caption(f"Source: {r['source_file']}")
                    if other_trips:
                        target = st.selectbox("Move to trip", [t["id"] for t in other_trips],
                                              format_func=lambda i: trip_names[i],
                                              key=f"mvsel_{r['id']}")
                        if st.button("➡️ Move", key=f"mv_{r['id']}"):
                            core.move_reservation(r["id"], target)
                            st.rerun()

                if c2.button("🗑️", key=f"del_{r['id']}", help="Delete this reservation"):
                    core.delete_reservation(r["id"])
                    st.rerun()


# ============================================================
#  TAB 2: ADD RESERVATIONS
# ============================================================
with tab_add:
    st.info(f"New reservations will be saved to the trip: **{trip['name']}**")
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
            found, sources = [], []
            with st.spinner("Claude is reading your documents…"):
                for f in files or []:
                    try:
                        found += core.extract_reservations(core.file_to_blocks(f.name, f.getvalue()))
                        sources.append(f.name)
                    except Exception as e:
                        st.error(f"Could not read {f.name}: {e}")
                if pasted.strip():
                    try:
                        found += core.extract_reservations([core.text_block(pasted)])
                        sources.append("pasted text")
                    except Exception as e:
                        st.error(f"Could not read pasted text: {e}")
            if found:
                df = pd.DataFrame(found).reindex(columns=core.FIELDS).fillna("").astype(str)
                st.session_state.pending = df
                st.session_state.pending_source = ", ".join(sources)
                st.session_state.pending_trip = trip_id
            else:
                st.warning("Claude didn't find any reservations in that content.")

    if "pending" in st.session_state:
        target_trip = st.session_state.get("pending_trip", trip_id)
        st.subheader(f"✅ Check the details, then save to **{trip_names.get(target_trip, trip['name'])}**")
        st.caption("Click any cell to edit it. Dates must look like 2025-07-14.")
        edited = st.data_editor(st.session_state.pending, num_rows="dynamic",
                                use_container_width=True, key="editor")
        col_a, col_b = st.columns(2)
        if col_a.button("💾 Save these reservations", type="primary"):
            try:
                with st.spinner("Saving and finding map locations…"):
                    n = core.save_reservations(edited.to_dict("records"),
                                               st.session_state.pending_source, target_trip)
                del st.session_state["pending"]
                st.success(f"Saved {n} reservation(s)! See them in the 📅 Itinerary tab.")
            except Exception as e:
                st.error(f"Save failed: {e}")
        if col_b.button("✖️ Discard"):
            del st.session_state["pending"]
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
        m_det = st.text_area("Details")
        if st.form_submit_button("➕ Add reservation"):
            if not m_title or not m_sd:
                st.error("Title and start date are required.")
            else:
                core.save_reservations([{
                    "type": m_type, "title": m_title,
                    "start_date": m_sd.isoformat(),
                    "start_time": m_st.strftime("%H:%M") if m_st else None,
                    "end_date": m_ed.isoformat() if m_ed else None,
                    "end_time": m_et.strftime("%H:%M") if m_et else None,
                    "location": m_loc, "provider": m_prov,
                    "confirmation_number": m_conf, "details": m_det,
                }], "manual entry", trip_id)
                st.success("Added! See it in the 📅 Itinerary tab.")


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
            "label": f"{core.TYPE_ICONS.get(r.get('type') or 'other', '📌')} {r.get('title') or ''}",
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
        st.caption("🔴 Places   🔵 Order of travel. Hover over or tap a dot for details. "
                   "Map positions are approximate.")
        st.dataframe(df[["when", "label", "where"]].rename(
            columns={"when": "When", "label": "What", "where": "Where"}),
            hide_index=True, use_container_width=True)


# ============================================================
#  TAB 4: CHAT WITH CLAUDE (separate chat per trip)
# ============================================================
with tab_chat:
    st.caption("Claude can see all reservations in this trip. Try: 'Summarize my trip', "
               "'What do I have on Friday?', 'Any nights without a hotel?'")
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
        slim = [{f: r.get(f) for f in core.FIELDS} for r in core.load_reservations(trip_id)]
        system_prompt = (
            "You are a friendly, helpful travel assistant inside a travel organizer app. "
            f"Today's date is {date.today().isoformat()}. "
            f"The user is looking at the trip '{trip['name']}'. Its reservations (JSON):\n"
            f"{json.dumps(slim, default=str)}\n\n"
            "Use them to answer accurately. Point out conflicts, tight connections or missing "
            "hotel nights if you notice them. You may also give general travel advice.")
        with st.chat_message("assistant"):
            try:
                with core.claude().messages.stream(model=core.main.model(), max_tokens=2000,
                                                   system=system_prompt,
                                                   messages=history) as stream:
                    reply = st.write_stream(stream.text_stream)
                history.append({"role": "assistant", "content": reply})
            except Exception as e:
                history.pop()
                st.error(f"Claude error: {e}")


# ============================================================
#  TAB 5: EMAIL INBOX
# ============================================================
with tab_email:
    address = core.trip_address(trip["email_code"])
    if not address:
        st.warning("Email forwarding isn't set up yet. Add EMAIL_ADDRESS and "
                   "EMAIL_APP_PASSWORD to the app's Secrets (see the guide).")
    else:
        st.markdown(f"Forward booking confirmations for **{trip['name']}** to:")
        st.code(address, language=None)
        st.caption("Emails sent to the plain inbox address (without +code) go to the "
                   "'📥 Unsorted emails' trip, and you can move them from there. The inbox is "
                   "checked automatically every few minutes while the app is in use.")

        if inbox_result.get("error"):
            st.error(f"Couldn't check the inbox: {inbox_result['error']}")

        if st.button("🔄 Check for new emails now", type="primary"):
            with st.spinner("Reading new emails… (about 10–30 seconds per email)"):
                try:
                    res = core.check_inbox(max_emails=10)
                    if res["processed"]:
                        for item in res["processed"]:
                            st.write(f"• **{item['subject']}** → {item['trip'] or '—'}: "
                                     f"{item['status']}")
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
            } for e in log]), hide_index=True, use_container_width=True)
        else:
            st.caption("No emails processed yet.")
