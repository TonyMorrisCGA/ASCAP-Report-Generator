import streamlit as st
import pandas as pd
import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from datetime import datetime, timedelta
import re

st.set_page_config(page_title="CGA ASCAP Playlist Converter", layout="wide")
st.title("🎸 Classical Guitar Alive! — ASCAP Radio Log Generator")
st.markdown("Automatically scrapes PRX → generates the official ASCAP Radio Log format.")

# === PRX Scraper (unchanged - works well) ===
url = st.text_input("PRX Episode URL", placeholder="https://exchange.prx.org/pieces/626837-...")

if st.button("🔄 Fetch & Parse from PRX URL", type="primary"):
    if not url:
        st.error("Please enter a PRX URL.")
    else:
        try:
            fetch_url = url if "?" in url else url + "?m=false"
            response = requests.get(fetch_url, timeout=15)
            soup = BeautifulSoup(response.text, "html.parser")

            page_title = soup.find("h1").get_text(strip=True) if soup.find("h1") else ""
            episode_title = page_title or "Classical Guitar Alive! Episode"

            air_date = ""
            total_duration = ""
            for text in soup.stripped_strings:
                if ("July" in text or "2026" in text) and not air_date:
                    air_date = text.strip()
                if re.search(r"\d{2}:\d{2}", text) and not total_duration:
                    total_duration = text.strip()

            table = None
            for t in soup.find_all("table"):
                headers = [th.get_text(strip=True).lower() for th in t.find_all("th")]
                if "title" in headers and "artist" in headers and "length" in headers:
                    table = t
                    break

            if table:
                rows = table.find_all("tr")
                data = []
                for i, row in enumerate(rows[1:], start=1):
                    cells = [td.get_text(strip=True) for td in row.find_all("td")]
                    if len(cells) >= 6:
                        title = cells[0]
                        artist = cells[1]
                        album = cells[2]
                        label = cells[3]
                        year = cells[4]
                        length = cells[5]

                        composer = title.split(":")[0].strip() if ":" in title else title
                        recording = f"{album} ({label} {year})".strip(" ()")

                        data.append({
                            "Track #": i,
                            "Title": title,
                            "Composer(s)": composer,
                            "Performer(s)": artist,
                            "Duration (MM:SS)": length,
                            "Recording/Album": recording,
                            "Notes": ""
                        })

                if data:
                    st.session_state["playlist_df"] = pd.DataFrame(data)
                    st.session_state["episode_title"] = episode_title
                    if air_date: st.session_state["air_date"] = air_date
                    if total_duration: st.session_state["total_duration"] = total_duration
                    st.success(f"✅ Scraped {len(data)} tracks successfully!")
                    st.rerun()
        except Exception as e:
            st.error(f"Error: {str(e)}")

# === Editable Table ===
st.divider()
st.subheader("📋 Edit Playlist (optional)")

if "playlist_df" not in st.session_state:
    st.session_state["playlist_df"] = pd.DataFrame(columns=[
        "Track #", "Title", "Composer(s)", "Performer(s)", "Duration (MM:SS)", "Recording/Album", "Notes"
    ])

edited_df = st.data_editor(st.session_state["playlist_df"], num_rows="dynamic", use_container_width=True)
st.session_state["playlist_df"] = edited_df

# Metadata
col1, col2 = st.columns(2)
with col1:
    program_title = st.text_input("Program", "Classical Guitar Alive!")
    episode_title = st.text_input("Episode Title", st.session_state.get("episode_title", ""))
with col2:
    air_date = st.text_input("Air Date (stations will edit)", st.session_state.get("air_date", "July 20, 2026"))
    total_duration = st.text_input("Total Duration", st.session_state.get("total_duration", "58:57"))

# === Generate ASCAP Radio Log (NEW FORMAT) ===
if st.button("📥 Generate ASCAP Radio Log Excel", type="primary"):
    if edited_df.empty or len(edited_df) == 0:
        st.error("Please add tracks to the table first.")
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "ASCAP Radio Log Template"

        # Exact ASCAP headers
        headers = [
            "Station Call Letters",
            "Date of Play",
            "Time of Play",
            "Song/Composition Title",
            "Song/Composition Artist",
            "Song/Composition Composer",
            "Duration"
        ]
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="4472C4")
            cell.alignment = Alignment(horizontal="center", wrap_text=True)

        # Calculate cumulative start times starting at 8:01 AM
        start_time = datetime(1900, 1, 1, 8, 1)
        current_time = start_time

        for idx, row in enumerate(edited_df.itertuples(index=False), start=2):
            duration_str = str(row[4]) if pd.notna(row[4]) else "0:00"
            try:
                mins, secs = map(int, duration_str.split(":"))
                duration_td = timedelta(minutes=mins, seconds=secs)
            except:
                duration_td = timedelta(minutes=0)

            # Write row
            ws.cell(row=idx, column=1, value="EXPL-FM")
            ws.cell(row=idx, column=2, value=air_date)
            ws.cell(row=idx, column=3, value=current_time.strftime("%I:%M %p"))
            ws.cell(row=idx, column=4, value=row[1])           # Title
            ws.cell(row=idx, column=5, value=row[3])           # Performer
            ws.cell(row=idx, column=6, value=row[2])           # Composer
            ws.cell(row=idx, column=7, value=duration_str)

            # Move to next start time
            current_time += duration_td

        # Auto column widths
        for column in ws.columns:
            max_length = max(len(str(cell.value or "")) for cell in column)
            ws.column_dimensions[column[0].column_letter].width = min(max_length + 3, 70)

        # Filename
        safe_name = re.sub(r'[^a-zA-Z0-9_-]', '_', episode_title)[:50]
        filename = f"CGA_{safe_name}_ASCAP_Radio_Log.xlsx"
        wb.save(filename)

        with open(filename, "rb") as f:
            st.download_button(
                label="⬇️ Download ASCAP Radio Log Excel",
                data=f,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        st.success("✅ ASCAP Radio Log generated! Stations just need to change EXPL-FM and the Date column.")
