import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows
import re

st.set_page_config(page_title="CGA ASCAP Playlist Converter", layout="wide")
st.title("🎸 Classical Guitar Alive! — ASCAP Playlist to Excel")
st.markdown("**How to use:** Paste the playlist from the PRX page → click **Parse** → edit the table if needed → click **Generate Excel**")

# === Input Section ===
col1, col2 = st.columns(2)
with col1:
    program_title = st.text_input("Program Title", "Classical Guitar Alive!")
with col2:
    episode_title = st.text_input("Episode Title", "26-29 Composers Looking Back in Time")

col3, col4 = st.columns(2)
with col3:
    air_date = st.text_input("Air Date", "July 20, 2026")
with col4:
    total_duration = st.text_input("Total Duration", "58:57")

playlist_text = st.text_area(
    "Paste the full playlist section from the PRX page here",
    height=250,
    placeholder="Example:\n1. Title of piece\n   Performer: ...\n   Duration: 12:34\n   Album: ..."
)

if st.button("🔄 Parse Playlist Text into Table"):
    lines = [line.strip() for line in playlist_text.split('\n') if line.strip()]
    parsed_data = []
    for line in lines:
        # Simple parsing - you can improve this later
        parsed_data.append({
            "Track #": "",
            "Title": line,
            "Composer(s)": "",
            "Performer(s)": "",
            "Duration (MM:SS)": "",
            "Recording/Album": "",
            "Notes": ""
        })
    
    if parsed_data:
        st.session_state["playlist_df"] = pd.DataFrame(parsed_data)
        st.success(f"✅ Parsed {len(parsed_data)} lines into the table below. Edit as needed!")
    else:
        st.warning("No lines found to parse.")

# === Editable Table ===
st.subheader("📋 Playlist Tracks (edit here)")

if "playlist_df" not in st.session_state:
    st.session_state["playlist_df"] = pd.DataFrame(columns=[
        "Track #", "Title", "Composer(s)", "Performer(s)", 
        "Duration (MM:SS)", "Recording/Album", "Notes"
    ])

edited_df = st.data_editor(
    st.session_state["playlist_df"],
    num_rows="dynamic",
    use_container_width=True,
    key="playlist_editor"
)

# Save edits back to session state
st.session_state["playlist_df"] = edited_df

# === Generate Button ===
if st.button("📥 Generate & Download ASCAP Excel", type="primary"):
    if edited_df.empty or len(edited_df) == 0:
        st.error("Please add at least one track to the table above (click the + button).")
    else:
        wb = Workbook()
        
        # Sheet 1: Program Info
        ws_info = wb.active
        ws_info.title = "Program Info"
        ws_info['A1'] = "ASCAP Music Performance Report"
        ws_info['A1'].font = Font(bold=True, size=14)
        
        info_data = [
            ("Program", program_title),
            ("Episode", episode_title),
            ("Air Date", air_date),
            ("Total Duration", total_duration)
        ]
        for i, (label, value) in enumerate(info_data, start=3):
            ws_info[f'A{i}'] = label
            ws_info[f'B{i}'] = value
            ws_info[f'A{i}'].font = Font(bold=True)
        
        # Sheet 2: Playlist
        ws_playlist = wb.create_sheet("Playlist")
        headers = ["Track #", "Title", "Composer(s)", "Performer(s)", 
                   "Duration (MM:SS)", "Recording/Album", "Notes"]
        
        for col, header in enumerate(headers, 1):
            cell = ws_playlist.cell(row=1, column=col, value=header)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="4472C4")
            cell.alignment = Alignment(horizontal="center", wrap_text=True)
        
        for r_idx, row in enumerate(edited_df.itertuples(index=False), start=2):
            for c_idx, value in enumerate(row, start=1):
                ws_playlist.cell(row=r_idx, column=c_idx, value=value)
        
        # Auto-adjust column widths
        for column in ws_playlist.columns:
            max_length = max(len(str(cell.value or "")) for cell in column)
            ws_playlist.column_dimensions[column[0].column_letter].width = min(max_length + 3, 70)
        
        # Filename
        safe_episode = re.sub(r'[^a-zA-Z0-9_-]', '_', episode_title)[:50]
        filename = f"CGA_{safe_episode}_ASCAP.xlsx"
        
        # Save and offer download
        wb.save(filename)
        with open(filename, "rb") as f:
            st.download_button(
                label="⬇️ Download ASCAP Excel File",
                data=f,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        st.success("✅ Excel file generated successfully!")
