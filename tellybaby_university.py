
# tellybaby_offline_fallback.py
import streamlit as st
import google.generativeai as genai
from datetime import datetime, timedelta
import pandas as pd
import io
import os

st.set_page_config(page_title="TellyBaby (Offline-safe) 🤖🎓", page_icon="🎓")

# ---------- CONFIG ----------
API_KEY = "AIzaSyBbClhmx7SxERZOIVCV0JtS-bnINlRlLOo"
genai.configure(api_key=API_KEY)

if API_KEY:
    try:
        genai.configure(api_key="AIzaSyBbClhmx7SxERZOIVCV0JtS-bnINlRlLOo")
        model = genai.GenerativeModel("models/gemini-2.5-flash")
        GEMINI_AVAILABLE = True
    except Exception:
        # If configuration fails, we'll still continue in offline mode.
        GEMINI_AVAILABLE = False
        model = None
else:
    # No API key provided — offline mode only
    GEMINI_AVAILABLE = False
    model = None

#Default UNIVERSITY DATA (can be overwritten by uploads)
default_timetable = {
    "monday": [
        "10:00-11:50 AM: R Programming (VEDANTA - VEDTF 305 LH 3) - Gaurav Dhiman",
        "2:00-2:50 PM: Machine Learning (VISVESVARAYA - VERASF 201) - Anushikha Singh",
        "4:00-5:50 PM: Artificial Intelligence Lab (VEDANTA - VED5F 508) - Vishakha Arya"
    ],
    "tuesday": [
        "10:00-10:50 AM: Computer Networks (VISVESVARAYA - VERA4F 406) - Shahid Ul Haq",
        "11:00-11:50 AM: Artificial Intelligence (VEDANTA - VEDTF 329 LH 6) - Vishakha Arya",
        "2:00-2:50 PM: Machine Learning (VISVESVARAYA - VERASF 201) - Anushikha Singh",
        "3:00-3:50 PM: R Programming (VISVESVARAYA - VERASF 201) - Gaurav Dhiman",
        "4:00-5:50 PM: Advanced Java Programming Lab (VEDANTA - VEDTF 329 LH 6) - Piyush Anand"
    ],
    "wednesday": [
        "10:00-11:50 AM: Machine Learning Lab (VEDANTA - VEDTF 305 LH 3) - Anushikha Singh",
        "2:00-2:50 PM: Computer Networks (VISVESVARAYA - VERA4F 402) - Shahid Ul Haq",
        "3:00-3:50 PM: Advanced Java Programming (VISVESVARAYA - VERASF 201) - Piyush Anand"
    ],
    "thursday": [
        "11:00-11:50 AM: Artificial Intelligence (VEDANTA - VED4F 405 LH7) - Vishakha Arya",
        "2:00-2:50 PM: Aptitude and Soft Skills (VISVAKARMA - VISKSF WL206) - Shaifali Streeting",
        "3:00-3:50 PM: Advanced Java Programming (VISVESVARAYA - VERASF 201) - Piyush Anand",
        "4:00-5:50 PM: Computer Networks Lab (VEDANTA - VED5F 508) - Shahid Ul Haq"
    ],
    "friday": [
        "9:00-9:50 AM: Aptitude and Soft Skills (VEDANTA - VEDSF 227 LH05 Room) - Vivek Dheeman",
        "10:00-10:50 AM: Artificial Intelligence (VEDANTA - VED4F 402 LH1 Room) - Vishakha Arya",
        "11:00-11:50 AM: Computer Networks (VISVESVARAYA - VERA4F 406) - Shahid Ul Haq",
        "12:00-12:50 PM: R Programming (VISVESVARAYA - VERASF 201) - Gaurav Dhiman",
        "2:00-2:50 PM: Advanced Java Programming (VISVESVARAYA - VERASF 201) - Piyush Anand"
    ]
}

default_attendance = {
    "aptitude and soft skills": 50.00,
    "environmental risk assessment and disaster management": 0.00,
    "advanced java programming": 72.97,
    "computer networks": 56.76,
    "artificial intelligence": 61.11,
    "r programming": 55.56,
    "machine learning": 67.86
}

default_exam_schedule = {
    "artificial intelligence": {"date": "2025-11-10", "marks": 88, "total": 100},
    "computer networks": {"date": "2025-11-12", "marks": 75, "total": 100},
    "r programming": {"date": "2025-11-15", "marks": 82, "total": 100},
    "machine learning": {"date": "2025-11-18", "marks": 79, "total": 100},
    "advanced java programming": {"date": "2025-11-21", "marks": 85, "total": 100},
}

#Session state initialization
if "timetable" not in st.session_state:
    st.session_state.timetable = default_timetable.copy()
if "attendance" not in st.session_state:
    st.session_state.attendance = default_attendance.copy()
if "exam_schedule" not in st.session_state:
    st.session_state.exam_schedule = default_exam_schedule.copy()
if "chat_history" not in st.session_state:
    st.session_state.chat_history = []
if "reminders" not in st.session_state:
    # reminder: dicts with keys: text, date (YYYY-MM-DD), time (HH:MM), done(bool)
    st.session_state.reminders = []

# ---------- Helpers ----------
def parse_uploaded_timetable(file_bytes, filename):
    """Parses timetable in the format with columns like Day, Time, Course Code, Course Title, etc."""
    try:
        if filename.endswith((".xls", ".xlsx")):
            df = pd.read_excel(io.BytesIO(file_bytes), engine="openpyxl")
        else:
            df = pd.read_csv(io.BytesIO(file_bytes))
    except Exception as e:
        st.error(f"Could not parse timetable file: {e}")
        return None

    # Normalize column names
    df.columns = [c.strip().lower() for c in df.columns]

    required_cols = {"day", "time", "course title"}
    if required_cols.issubset(set(df.columns)):
        out = {}
        for _, row in df.iterrows():
            day = str(row["day"]).strip().lower()
            time = str(row["time"]).strip()
            course = str(row["course title"]).strip()
            code = str(row.get("course code", "")).strip()
            faculty = str(row.get("couse instructore", "")).strip()
            room = str(row.get("resource name", "")).strip()

            # Build formatted string
            formatted = f"{time}: {course}"
            if code:
                formatted += f" ({code})"
            if room:
                formatted += f" - {room}"
            if faculty:
                formatted += f" - {faculty}"

            out.setdefault(day, []).append(formatted)
        return out
    else:
        st.warning("Timetable file doesn't contain 'Day', 'Time', and 'Course Title' columns.")
        return None



def parse_uploaded_attendance(file_bytes, filename):
    try:
        if filename.endswith((".xls", ".xlsx")):
            try:
                df = pd.read_excel(io.BytesIO(file_bytes), engine="openpyxl")
            except Exception:
                df = pd.read_excel(io.BytesIO(file_bytes), engine="xlrd")
        else:
            df = pd.read_csv(io.BytesIO(file_bytes))
    except Exception as e:
        st.error(f"Could not parse attendance file: {e}")
        return None

    if {"Subject", "Attendance"}.issubset(set(df.columns)):
        out = {}
        for _, row in df.iterrows():
            subj = str(row["Subject"]).strip().lower()
            try:
                val = float(row["Attendance"])
            except:
                try:
                    val = float(str(row["Attendance"]).replace("%", ""))
                except:
                    val = 0.0
            out[subj] = val
        return out
    else:
        st.warning("Attendance file doesn't have Subject/Attendance columns. Trying best-effort parse.")
        out = {}
        for _, row in df.iterrows():
            if len(row) >= 2:
                subj = str(row.iloc[0]).strip().lower()
                try:
                    val = float(row.iloc[1])
                except:
                    val = 0.0
                out[subj] = val
        return out if out else None


def parse_uploaded_exams(file_bytes, filename):
    try:
        if filename.endswith((".xls", ".xlsx")):
            try:
                df = pd.read_excel(io.BytesIO(file_bytes), engine="openpyxl")
            except Exception:
                df = pd.read_excel(io.BytesIO(file_bytes), engine="xlrd")
        else:
            df = pd.read_csv(io.BytesIO(file_bytes))
    except Exception as e:
        st.error(f"Could not parse exams file: {e}")
        return None

    if {"Subject", "Date"}.issubset(set(df.columns)):
        out = {}
        for _, row in df.iterrows():
            subj = str(row["Subject"]).strip().lower()
            date = str(row["Date"]).strip()
            marks = row.get("Marks", None)
            total = row.get("Total", None)
            try:
                marks = int(marks) if pd.notna(marks) else None
            except:
                marks = None
            try:
                total = int(total) if pd.notna(total) else None
            except:
                total = None
            out[subj] = {"date": date, "marks": marks, "total": total}
        return out
    else:
        st.warning("Exams file doesn't have Subject/Date columns. Trying best-effort parse.")
        return None

def get_day_timetable(day):
    day = day.lower()
    return "\n".join(st.session_state.timetable.get(day, [])) or "No classes scheduled for that day."

def get_next_class():
    now = datetime.now()
    day = now.strftime("%A").lower()
    today_list = st.session_state.timetable.get(day, [])
    if not today_list:
        return "You have no classes today 🎉"
    # try to parse start times (very permissive)
    for cls in today_list:
        # find something like "10:00" or "10-11"
        token = cls.split()[0]  # usually time-range at start
        try:
            if "-" in token:
                start = token.split("-")[0]
            else:
                start = token
            # strip AM/PM if present at end of string:
            start = start.replace("AM", "").replace("PM", "").strip()
            dt = datetime.strptime(start, "%H:%M")
            # compare hours/minutes
            if (dt.hour, dt.minute) > (now.hour, now.minute):
                return f"Your next class is: {cls}"
        except:
            # try %I:%M %p inside entire cls
            try:
                # look for AM/PM
                if "AM" in cls or "PM" in cls:
                    # naive: find first token with colon
                    parts = [p for p in cls.split() if ":" in p]
                    if parts:
                        tkn = parts[0]
                        # handle range like 10:00-11:50
                        if "-" in tkn:
                            tstart = tkn.split("-")[0]
                        else:
                            tstart = tkn
                        # if AM/PM in whole string, default to AM
                        if "AM" in cls:
                            tstart = tstart + " AM"
                        elif "PM" in cls:
                            tstart = tstart + " PM"
                        dt = datetime.strptime(tstart, "%I:%M %p")
                        if (dt.hour, dt.minute) > (now.hour, now.minute):
                            return f"Your next class is: {cls}"
            except:
                continue
    return "All your classes for today are done ✅"

def get_attendance(subject):
    subject = subject.lower()
    if subject in st.session_state.attendance:
        return f"Your attendance in {subject.title()} is {st.session_state.attendance[subject]}%."
    return "I couldn’t find that subject. Please check the name again."

def get_exam_info(subject=None):
    if subject:
        subject = subject.lower()
        if subject in st.session_state.exam_schedule:
            info = st.session_state.exam_schedule[subject]
            marks = info.get("marks")
            total = info.get("total")
            marks_str = f" Marks: {marks}/{total}." if marks is not None and total is not None else ""
            return f"{subject.title()} exam is on {info['date']}.{marks_str}"
        else:
            return "Exam not found for that subject."
    else:
        exams = []
        for sub, info in st.session_state.exam_schedule.items():
            marks = info.get("marks")
            total = info.get("total")
            marks_str = f" ({marks}/{total})" if marks is not None and total is not None else ""
            exams.append(f"{sub.title()} - {info['date']}{marks_str}")
        return "\n".join(exams) or "No exams scheduled."

def get_upcoming_exam():
    today = datetime.now().date()
    upcoming = None
    for sub, info in st.session_state.exam_schedule.items():
        try:
            exam_date = datetime.strptime(info['date'], "%Y-%m-%d").date()
        except:
            # skip malformed
            continue
        if exam_date >= today:
            # pick earliest one
            if upcoming is None or exam_date < upcoming[1]:
                upcoming = (sub, exam_date)
    if upcoming:
        days_left = (upcoming[1] - today).days
        return f"Your next exam is {upcoming[0].title()} on {upcoming[1]} (in {days_left} days) 📚"
    else:
        return "No upcoming exams 🎉"

def ask_gemini(query, max_tokens=512):
    """Return text on success, None on failure."""
    if not GEMINI_AVAILABLE or model is None:
        return None
    try:
        response = model.generate_content(query)
        # defensive access
        if response and getattr(response, "candidates", None):
            cand = response.candidates[0]
            if getattr(cand, "content", None) and getattr(cand.content, "parts", None):
                return cand.content.parts[0].text
        return None
    except Exception as e:
        # Log error to the app for debugging, but do not crash
        st.session_state.chat_history.append(("System", f"Gemini error: {str(e)}"))
        return None

# ---------- Sidebar: upload + reminders ----------
with st.sidebar:
    st.header("Dashboard")
    st.write(f"📅 Date: {datetime.now().strftime('%A, %B %d, %Y')}")
    st.write(f"🕒 Next Class: {get_next_class()}")
    overall_att = sum(st.session_state.attendance.values()) / max(1, len(st.session_state.attendance))
    st.write(f"🧮 Overall Attendance: {overall_att:.2f}%")
    st.write(get_upcoming_exam())
    st.divider()

    st.subheader("Upload / Update Data")
    tt_file = st.file_uploader("Upload Timetable (CSV/XLSX)", type=["csv", "xls", "xlsx"], key="tt_upload")
    if tt_file is not None:
        parsed = parse_uploaded_timetable(tt_file.read(), tt_file.name)
        if parsed:
            st.session_state.timetable = parsed
            st.success("Timetable uploaded and applied.")
    at_file = st.file_uploader("Upload Attendance (CSV/XLSX)", type=["csv", "xls", "xlsx"], key="att_upload")
    if at_file is not None:
        parsed = parse_uploaded_attendance(at_file.read(), at_file.name)
        if parsed:
            st.session_state.attendance = parsed
            st.success("Attendance uploaded and applied.")
    ex_file = st.file_uploader("Upload Exams/Marks (CSV/XLSX)", type=["csv", "xls", "xlsx"], key="ex_upload")
    if ex_file is not None:
        parsed = parse_uploaded_exams(ex_file.read(), ex_file.name)
        if parsed:
            st.session_state.exam_schedule = parsed
            st.success("Exam schedule uploaded and applied.")

    st.divider()
    st.subheader("Reminders")
    with st.form("add_reminder", clear_on_submit=True):
        r_text = st.text_input("Reminder text (e.g., Study AI)")
        r_date = st.date_input("Date")
        r_time = st.time_input("Time")
        submitted = st.form_submit_button("Add reminder")
        if submitted:
            st.session_state.reminders.append({
                "text": r_text,
                "date": r_date.strftime("%Y-%m-%d"),
                "time": r_time.strftime("%H:%M"),
                "done": False
            })
            st.success("Reminder added.")

    if st.session_state.reminders:
        st.write("Upcoming reminders:")
        for idx, rem in enumerate(st.session_state.reminders):
            status = "✅" if rem["done"] else "⏰"
            st.write(f"{idx+1}. {status} {rem['date']} {rem['time']} — {rem['text']}")
            if st.button(f"Mark done {idx+1}", key=f"done_{idx}"):
                st.session_state.reminders[idx]["done"] = True
                st.experimental_rerun()
        if st.button("Clear all reminders"):
            st.session_state.reminders = []
            st.experimental_rerun()

# ---------- Main UI ----------
st.title("🤖 TellyBaby (Offline-safe) - University Assistant 🎓")
if not GEMINI_AVAILABLE:
    st.warning("Gemini API not configured or not available. The app will work in offline mode for university queries.")

st.write("Ask me about your timetable, attendance, exams, or set reminders. For general questions I'll try Gemini when available.")

# quick buttons for common actions
col1, col2, col3 = st.columns(3)
if col1.button("Show today's timetable"):
    today = datetime.now().strftime("%A").lower()
    st.session_state.chat_history.append(("You", "Show today's timetable"))
    st.session_state.chat_history.append(("TellyBaby", get_day_timetable(today)))
if col2.button("What's my next class?"):
    st.session_state.chat_history.append(("You", "What's my next class?"))
    st.session_state.chat_history.append(("TellyBaby", get_next_class()))
if col3.button("Upcoming exam"):
    st.session_state.chat_history.append(("You", "Show upcoming exam"))
    st.session_state.chat_history.append(("TellyBaby", get_upcoming_exam()))

# chat input
user_input = st.text_input("You:", "")

if user_input:
    user_input_lower = user_input.lower().strip()

    if "bye" in user_input_lower:
        reply = "Bye-bye! 💖 Take care!"
    elif "timetable" in user_input_lower or "class" in user_input_lower:
        # If user asks about specific day e.g., "timetable for tuesday" or "what is my time table for tomorrow?"
        if "tomorrow" in user_input_lower:
            tomorrow = (datetime.now() + timedelta(days=1)).strftime("%A").lower()
            reply = f"Timetable for tomorrow ({tomorrow.title()}):\n" + get_day_timetable(tomorrow)
        else:
            found_day = None
            for d in st.session_state.timetable.keys():
                if d in user_input_lower:
                    found_day = d
                    break
            reply = get_day_timetable(found_day or datetime.now().strftime("%A").lower())
    elif "next class" in user_input_lower or ("next" in user_input_lower and "class" in user_input_lower):
        reply = get_next_class()
    elif "attendance" in user_input_lower:
        # try to find subject in query
        found = False
        for sub in st.session_state.attendance.keys():
            if sub in user_input_lower:
                reply = get_attendance(sub)
                found = True
                break
        if not found:
            # maybe user asked "what is my attendance" -> show summary table
            if "what" in user_input_lower or "show" in user_input_lower:
                df = pd.DataFrame([
                    {"Subject": k.title(), "Attendance (%)": v}
                    for k, v in st.session_state.attendance.items()
                ])
                st.session_state.chat_history.append(("You", user_input))
                st.session_state.chat_history.append(("TellyBaby", "Showing attendance table below."))
                st.dataframe(df)
                reply = "(attendance table displayed)"
            else:
                reply = "Please specify the subject name to check attendance (e.g., 'attendance in machine learning')."
    elif "exam" in user_input_lower or "marks" in user_input_lower:
        found = False
        for sub in st.session_state.exam_schedule.keys():
            if sub in user_input_lower:
                reply = get_exam_info(sub)
                found = True
                break
        if not found:
            reply = get_exam_info()
    elif "remind me" in user_input_lower or user_input_lower.startswith("remind"):
        # quick natural parse: "remind me to study AI tomorrow at 7pm"
        # We'll implement a very basic parser for common case: "remind me to <text> on YYYY-MM-DD at HH:MM"
        reply = "To add reminders use the sidebar form. Example: 'Remind me to study AI on 2025-10-30 at 19:00'"
    elif "date" in user_input_lower or "day" in user_input_lower:
        reply = f"Today is {datetime.now().strftime('%A, %B %d, %Y')}."
    else:
        # fallback: try Gemini if available
        gemini_answer = ask_gemini(user_input)
        if gemini_answer is not None:
            reply = gemini_answer
        else:
            # offline fallback message
            reply = ("I can't reach Gemini right now. I can still answer university queries "
                     "like timetable, attendance, exams, or set reminders. Try asking those, "
                     "or upload your files in the sidebar to update data.")

    st.session_state.chat_history.append(("You", user_input))
    st.session_state.chat_history.append(("TellyBaby", reply))

# display chat
for sender, msg in st.session_state.chat_history:
    if sender == "You":
        st.markdown(f"**🧑‍🎓 You:** {msg}")
    elif sender == "TellyBaby":
        st.markdown(f"**🤖 TellyBaby:** {msg}")
    else:
        st.caption(f"**{sender}:** {msg}")

# end

