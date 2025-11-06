import streamlit as st
import pandas as pd
from icalendar import Calendar
from datetime import datetime, timedelta
import io
import re

st.set_page_config(page_title="Google Calendar Event Extractor", page_icon="📅", layout="centered")

st.title("📅 Google Calendar Accepted Events Extractor")

st.markdown("""
Upload your **Google Calendar `.ics` file**, and this app will:
- Extract all events you **accepted (Yes)** this week  
- Exclude **all-day events** by default  
- Show their **date**, **time**, and **accepted attendees**  
- Let you **download the results as CSV or Excel**
""")

# --- File upload ---
ics_file = st.file_uploader("📤 Upload your Google Calendar .ics file", type=["ics"])

# --- Try to extract email from filename ---
user_email = None
if ics_file is not None:
    # remove trailing .ics extension (case-insensitive) to avoid capturing it as part of the email
    filename_no_ext = re.sub(r'(?i)\.ics$', '', ics_file.name)
    # now extract a normal-looking email (requires a dot in the domain)
    match = re.search(r'([\w\.-]+@[\w\.-]+\.\w+)', filename_no_ext)
    if match:
        user_email = match.group(1)

# Show extracted email and allow edit
user_email = st.text_input("✉️ Your email (used to check your RSVP)", value=user_email or "")

# --- Date range: Default to current week ---
today = datetime.now()
start_of_week = today - timedelta(days=today.weekday())
end_of_week = start_of_week + timedelta(days=6)

start_date = st.sidebar.date_input("Start date", start_of_week.date())
end_date = st.sidebar.date_input("End date", end_of_week.date())

if ics_file and user_email:
    cal = Calendar.from_ical(ics_file.read())
    events_data = []

    for component in cal.walk():
        if component.name == "VEVENT":
            summary = str(component.get('summary', ''))
            start = component.get('dtstart').dt
            end = component.get('dtend').dt

            # Skip all-day events (datetime.date instead of datetime.datetime)
            if not isinstance(start, datetime) or not isinstance(end, datetime):
                continue

            # Skip events outside selected date range
            if not (start_date <= start.date() <= end_date):
                continue

            attendees = component.get('attendee')
            attendees_list = []
            accepted_by_user = False

            if attendees:
                if not isinstance(attendees, list):
                    attendees = [attendees]

                for attendee in attendees:
                    params = attendee.params
                    email = str(attendee).replace("mailto:", "").lower()
                    partstat = params.get('PARTSTAT', '').upper()

                    if partstat == "ACCEPTED":
                        attendees_list.append(email)

                    if user_email.lower() in email and partstat == "ACCEPTED":
                        accepted_by_user = True

            # Duration in hours
            duration = round((end - start).total_seconds() / 3600, 2)

            if accepted_by_user:
                events_data.append({
                    "Event": summary,
                    "Date": start.strftime("%Y-%m-%d"),
                    "Time": f"{start.strftime('%H:%M')} - {end.strftime('%H:%M')}",
                    "Duration (hrs)": duration,
                    "Accepted Attendees": ", ".join(attendees_list)
                })

    if events_data:
        df = pd.DataFrame(events_data)
        st.success(f"✅ Found {len(df)} accepted events between {start_date} and {end_date}.")
        st.dataframe(df, use_container_width=True)

        # Download as CSV
        csv = df.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="💾 Download CSV",
            data=csv,
            file_name="accepted_events.csv",
            mime="text/csv"
        )

        # Download as Excel
        towrite = io.BytesIO()
        with pd.ExcelWriter(towrite, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Accepted Events")
        towrite.seek(0)
        st.download_button(
            label="📘 Download Excel",
            data=towrite,
            file_name="accepted_events.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("No accepted events found for your email within this week.")
else:
    st.info("⬆️ Please upload your .ics file (named as your email) and confirm your email.")
