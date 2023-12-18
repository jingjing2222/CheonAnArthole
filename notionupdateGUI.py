import requests
import json
import tkinter as tk
from datetime import datetime, timedelta, timezone
from tkinter import messagebox
import pandas as pd
import re
from datetime import datetime

# Your Notion API Key and Database ID
NOTION_API_KEY = 'secret_rvjpkymAcp6B11Vgs2cvsDs6lAyDkMUGvtuoZlV472d'
DATABASE_ID = 'b2131b5f967c41a4bf8820e056a8c77b'

# Notion API Headers
headers = {
    "Authorization": "Bearer " + NOTION_API_KEY,
    "Notion-Version": "2022-02-22",
    "Content-Type": "application/json"
}

# Korea Standard Time (UTC+9)
KST = timezone(timedelta(hours=9))

# Function to add a row to Notion
def add_row_to_notion(data):
    create_row_url = f"https://api.notion.com/v1/pages"

    payload = {
        'parent': {'database_id': DATABASE_ID},
        'properties': data['properties']
    }

    response = requests.post(
        create_row_url,
        headers=headers,
        data=json.dumps(payload)
    )

    if response.status_code == 200:
        print("Row added successfully!")
    else:
        print("Failed to add row:", response.content)

# Function to process and update performances from a text file
def get_performances():
    with open('yedang.txt', 'r', encoding='utf-8') as file:
        performances = file.read().split('\n\n')
        for performance in performances:
            check = 0
            data = performance.split('\n')
            try:
                performance_name = data[0].split(': ')[1]
                performance_place = data[1].split(': ')[1]
                performance_time = data[2].split(': ')[1]
                print(performance_time)
            except IndexError:
                print(f"Error processing data: {data}")
                continue

            if '미정' in performance_time:
                continue
            else:
                try:
                    performance_time_start = datetime.strptime(performance_time.strip(), "%Y년 %m월 %d일 %H:%M")
                    performance_time_start = performance_time_start.replace(tzinfo=KST)
                except ValueError:
                    performance_time_start = datetime.strptime(performance_time.strip(), "%Y년 %m월 %d일")
                    performance_time_start = performance_time_start.replace(tzinfo=KST)

            performance_duration = data[3].split(': ')[1]
            performance_duration = performance_duration.split('(')[0].strip()

            staff_names = data[4].split('\t')
            staff_list = [{'name': staff_name.strip()} for staff_name in staff_names]

            parsed_data = {
                'parent': {'database_id': DATABASE_ID},
                'properties': {
                    '공연명': {
                        'title': [{'text': {'content': performance_name}}]
                    },
                    '공연장분류': {
                        'select': {'name': performance_place}
                    },
                    '근무일시': {
                        'date': {'start': performance_time_start.isoformat()}
                    },
                    '러닝타임': {
                        'rich_text': [{'text': {'content': performance_duration}}]
                    },
                    '근무인원': {
                        'multi_select': staff_list
                    }
                }
            }

            add_row_to_notion(parsed_data)
    messagebox.showinfo("Success", "업데이트 완료.")

# Function to quit the program
def quit_program():
    root.destroy()

# Function to extract and format schedule from an Excel file
def extract_and_format_schedule(excel_file_path, output_file_path):
    excel_data = pd.read_excel(excel_file_path)
    output_text = ""
    current_year = datetime.now().year

    for index, row in excel_data.iterrows():
        try:
            details = row.iloc[0].split('\n')
            title = re.sub(r'\d+월 \d+일\(.+?\) ', '', details[0]).strip()

            location_pattern = re.search(r'(대공연장|소공연장)', details[1])
            location = location_pattern.group(0).strip() if location_pattern else "정보 없음"

            date_pattern = re.search(r'(\d+월 \d+일)', details[0])
            date = date_pattern.group(0) if date_pattern else "정보 없음"

            # 여러 시간이 포함된 경우에 대한 처리
            time_pattern = re.search(r'공연시간 : ([\d:, ]+)', details[1])
            if time_pattern:
                times = time_pattern.group(1).strip()
                times_list = [time.strip() for time in times.split(',')]
                # 시간이 여러 개일 경우에만 제목에 추가
                if len(times_list) > 1:
                    title += " (" + ', '.join(times_list) + ")"
                first_time = times_list[0]
            else:
                first_time = "정보 없음"

            date_with_year = f"{current_year}년 {date} {first_time}"

            duration_pattern = re.search(r'러닝타임 (\d+분)', details[1])
            duration = duration_pattern.group(1).strip() if duration_pattern else "정보 없음"

            intermission_pattern = re.search(r'인터미션 (\d+분)', details[1])
            intermission = intermission_pattern.group(1).strip() if intermission_pattern else None

            if intermission:
                formatted_duration = f"{duration}(인터미션: {intermission})예정"
            else:
                formatted_duration = f"{duration}(인터미션 없음)예정"

            names = row[1:].dropna().tolist()
            formatted_names = '\t'.join(names)

            output_text += f"공연명: {title}\n공연장소: {location}\n공연시간: {date_with_year}\n러닝타임: {formatted_duration}\n{formatted_names}\n\n"
        except Exception as e:
            print(f"엑셀 행 {index+1} 처리 중 오류 발생: {e}")
            continue

    with open(output_file_path, 'w', encoding='utf-8') as file:
        file.write(output_text.strip())

# Create the main GUI window
root = tk.Tk()
root.title("Notion Update")
root.geometry("300x200")

frame = tk.Frame(root)
frame.place(relx=0.5, rely=0.5, anchor='center')

date_label = tk.Label(frame, text="스케줄.xlsm은 잘 조절하셨나요?\n1열부터 공연으로 시작합니다.\n공연 정보와 근무 인원만 필요로 합니다.")
date_label.pack()

button_frame = tk.Frame(frame)
button_frame.pack(pady=10)

excel_button = tk.Button(button_frame, text="Excel-text 실행", command=lambda: extract_and_format_schedule('스케줄.xlsm', 'yedang.txt'))
excel_button.pack(side="left", padx=10)

fetch_button = tk.Button(button_frame, text="노션 업데이트", command=get_performances)
fetch_button.pack(side="left", padx=10)

quit_button = tk.Button(button_frame, text="종료", command=quit_program)
quit_button.pack(side="right", padx=10)

root.mainloop()