import requests
import json
import tkinter as tk
from datetime import datetime, timedelta, timezone
from tkinter import messagebox
import pandas as pd
import re
from datetime import datetime

def extract_and_format_schedule(excel_file_path, output_file_path):
    # Load the Excel file
    excel_data = pd.read_excel(excel_file_path)

    # Prepare the output text
    output_text = ""

    # Get the current year
    current_year = datetime.now().year

    # Iterate over each row (performance)
    for index, row in excel_data.iterrows():
        # Extract performance details from the first cell
        details = row.iloc[0].split('\n')
        title = re.sub(r'\d+월 \d+일\(.+?\) ', '', details[0]).strip()

        # Extracting location, date and time, and duration
        location_pattern = re.search(r'(대공연장|소공연장)', details[1])
        location = location_pattern.group(0).strip() if location_pattern else "정보 없음"

        date_pattern = re.search(r'(\d+월 \d+일)', details[0])
        date = date_pattern.group(0) if date_pattern else "정보 없음"

        time_pattern = re.search(r'공연시간 : ([\d:]+)', details[1])
        time = time_pattern.group(1).strip() if time_pattern else "정보 없음"

        # Format the date with the current year
        date_with_year = f"{current_year}년 {date} {time}"

        duration_pattern = re.search(r'러닝타임 (\d+분)', details[1])
        duration = duration_pattern.group(1).strip() if duration_pattern else "정보 없음"

        intermission_pattern = re.search(r'인터미션 (\d+분)', details[1])
        intermission = intermission_pattern.group(1).strip() if intermission_pattern else None

        # Format the text with or without intermission
        if intermission:
            formatted_duration = f"{duration}(인터미션: {intermission})예정"
        else:
            formatted_duration = f"{duration}(인터미션 없음)예정"

        # Extract names of available individuals
        names = row[1:].dropna().tolist()
        formatted_names = '\t'.join(names)

        # Format the text
        output_text += f"공연명: {title}\n공연장소: {location}\n공연시간: {date_with_year}\n러닝타임: {formatted_duration}\n{formatted_names}\n\n"

    # Save the formatted data to the output file
    with open(output_file_path, 'w', encoding='utf-8') as file:
        file.write(output_text.strip())

# Example usage
extract_and_format_schedule('스케줄.xlsm', 'yedang.txt')
