import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows
import requests
from bs4 import BeautifulSoup
import re

st.set_page_config(page_title="CGA ASCAP Playlist Converter", layout="wide")
st.title("🎸 Classical Guitar Alive! ASCAP Playlist Converter")
st.markdown("Paste a PRX episode URL or the playlist text below. Get an ASCAP-ready Excel file instantly.")

# Input options
input_type = st.radio("Input Method", ["PRX Episode URL (auto-scrape)", "Paste Playlist Text (manual)"])

playlist_data = []

if input_type == "PRX Episode URL (auto-scrape)":
    url = st.text_input("PRX Piece URL (e.g. https://exchange.prx.org/pieces/623971)")
    if st.button("Fetch & Convert"):
        if url:
            try:
                response = requests.get(url)
                soup = BeautifulSoup(response.text, 'html.parser')
                # Basic extraction - adjust selectors if PRX changes structure
                title = soup.find('h1').get_text(strip=True) if soup.find('h1') else "Unknown Episode"
                # For full robustness, you can enhance parsing or fall back to manual
                st.info(f"Fetched: {title}. For best results, use manual paste below if needed.")
                # Placeholder - in practice, enhance with more specific parsing
                st.warning("Auto-scrape is basic. Paste the full playlist text for accuracy.")
            except:
                st.error("Could not fetch URL. Try manual paste.")
else:
    episode_info = st.text_area("Episode Info (optional but recommended)", 
                                "Program: Classical Guitar Alive!\nEpisode: [Title]\nDate: [Date]\nTotal Duration: 58:57")
    playlist_text = st.text_area("Paste the full playlist section from the PRX page", height=300)

    if st.button("Convert to ASCAP Excel"):
        # Simple parsing example - stations can paste numbered list or structured text
        lines = playlist_text.strip().split('\n')
        playlist_data = []
        for i, line in enumerate(lines):
            if re.match(r'^\d+|\bIntro\b', line.strip()):
                # Basic splitting - improve as needed or let user edit below
                playlist_data.append({"Track #": i+1, "Title": line.strip(), "Composer(s)": "", 
                                      "Performer(s)": "", "Duration (MM:SS)": "", 
                                      "Recording/Album": "", "Notes": ""})

        st.success(f"Parsed {len(playlist_data)} tracks.")

# Manual editing table (always available for refinement)
if playlist_data or 'playlist_data' in locals():
    st.subheader("Edit Playlist Data")
    df = pd.DataFrame(playlist_data) if playlist_data else pd.DataFrame(columns=["Track #", "Title", "Composer(s)", "Performer(s)", "Duration (MM:SS)", "Recording/Album", "Notes"])
    edited_df = st.data_editor(df, num_rows="dynamic")

    program_title = st.text_input("Program Title", "Classical Guitar Alive!")
    episode_title = st.text_input("Episode Title", "26-27 Music from the Summer of 1739")
    air_date = st.text_input("Air Date", "July 6, 2026")

    if st.button("Generate & Download ASCAP Excel"):
        # Create Excel (same professional format as before)
        wb = Workbook()
        ws_info = wb.active
        ws_info.title = "Program Info"
        ws_info['A1'] = "ASCAP Music Performance Report"
        ws_info['A1'].font = Font(bold=True, size=14)
        
        info = {
            "Program": program_title,
            "Episode": episode_title,
            "Air Date": air_date,
            "Total Duration": "58:57"  # Update as needed
        }
        row = 3
        for key, val in info.items():
            ws_info[f'A{row}'] = key
            ws_info[f'B{row}'] = val
            ws_info[f'A{row}'].font = Font(bold=True)
            row += 1

        ws_playlist = wb.create_sheet("Playlist")
        headers = ["Track #", "Title", "Composer(s)", "Performer(s)", "Duration (MM:SS)", "Recording/Album", "Notes"]
        for col, h in enumerate(headers, 1):
            cell = ws_playlist.cell(1, col, h)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", start_color="4472C4")
            cell.alignment = Alignment(horizontal="center")

        for r_idx, row_data in enumerate(dataframe_to_rows(edited_df, index=False, header=False), 2):
            for c_idx, value in enumerate(row_data, 1):
                ws_playlist.cell(r_idx, c_idx, value)

        # Auto column widths
        for column in ws_playlist.columns:
            max_length = max(len(str(cell.value or "")) for cell in column)
            ws_playlist.column_dimensions[column[0].column_letter].width = min(max_length + 2, 60)

        filename = f"CGA_{episode_title.replace(' ', '_')}_ASCAP.xlsx"
        wb.save(filename)
        
        with open(filename, "rb") as f:
            st.download_button("📥 Download ASCAP Excel File", f, file_name=filename, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        
        st.success("✅ File ready!")
