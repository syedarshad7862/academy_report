from openpyxl import load_workbook
from datetime import datetime
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH

# Create a new Document
document = Document()

# Output folder for saving docx files
output_folder = r'C:\Users\ThinkPad\Desktop\python project\\'
# output_folder = r'C:\Users\ThinkPad\Desktop\academy report\\'

wb = load_workbook(filename = r'C:\Users\ThinkPad\Desktop\academy report\Master Sheet - Student Performance Report.xlsx', data_only=True)

sheet_names = ["Attendance report", "Rating"]

skip_rows = {
    'Attendance report': 5,  # Skip 5 rows for 'Sheet1'
    'Rating': 2   # Skip 2 rows for 'Sheet2'
}

start_cells = {
      'Attendance report': 1,
      "Rating": 2
      }

# dictionary to store row data
sheets_data = {}
for sheet_name in sheet_names:
      sheet = wb[sheet_name]
      print(f"Processing sheet: {sheet_name}")
      rows_to_skips = skip_rows.get(sheet_name, 0)
      start_column = start_cells.get(sheet_name,0)
      sheet_data = []
      for i, row in enumerate(sheet.rows, start=1):
            if i <= rows_to_skips:
                  continue
            row_data = []
            for j,cell in enumerate(row[start_column:], start=start_column):
                  if i == rows_to_skips + 1 and isinstance(cell.value, datetime):
                      row_data.append(cell.value.strftime("%d-%m-%Y"))
                  elif isinstance(cell.value, float) and cell.value <= 1.0:
                        row_data.append(f"{cell.value * 100:.2f}%")
                  else:
                        row_data.append(cell.value)
                        
            if row_data:
                  sheet_data.append(row_data)
                  # print(f"Row {i}: {row_data}")
            sheets_data[sheet_name] = sheet_data
            
            # row_data = [cell.value for cell in row[start_column:]]
            # print("after printing row")
            # sheet_data.append(row_data) 
            # print(f"Row {i} data: {[cell.value for cell in row]}")

# print(f" it is dictionary len: {len(sheets_data)}it list len {len(sheet_data)}")

# print(len(sheets_data[sheet_names[0]])) # it has 18 elements
attendance_sheet_data = sheets_data[sheet_names[0]]
# print(attendance_sheet_data[2:])
# print(attendance_sheet_data[0][1], attendance_sheet_data[0][12])
# session_dates = attendance_sheet_data[0][1:]
# for sheet_name, data in sheets_data.items():
#       print(f"\n data from {sheet_name}:")
#       for row in data:
#             print(row)

# Loop through each student row and generate a report
session_dates = attendance_sheet_data[0][1:]  # Skip the header
valid_dates = [date for date in session_dates if date is not None]
start_date = valid_dates[0]
end_date = valid_dates[-1]
# print(f"data of arr : {attendance_sheet_data[5:]}")
# Loop through each student row and generate a report
for student_row in attendance_sheet_data[5:]:
    print(student_row)
    student_name = student_row[0]
    attended_sessions = student_row[14]  # Attended sessions count
    attendance_percentage = student_row[15]  # Attendance percentage
    print(f"attended: {attended_sessions} and {attendance_percentage}")
    # Create a new Document
    document = Document()

    title1 = document.add_paragraph('Daari Academy – Your Path to Success')
    title1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub_title = document.add_paragraph("Student Performance Report")
    sub_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    course_title = document.add_paragraph("Course name: CMA US - Part 2")
    course_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    report_name = document.add_paragraph("Report Name: Weekly Report")
    report_name.alignment = WD_ALIGN_PARAGRAPH.CENTER

    valid_dates = [date for date in session_dates if date is not None]
    start_date = valid_dates[0] if valid_dates else "Unknown"
    end_date = valid_dates[-1] if valid_dates else "Unknown"
    print(start_date,end_date)
    report_duration = document.add_paragraph()
    # report_duration.add_run(f"Report Duration: From {attendance_sheet_data[0][1]} to {attendance_sheet_data[0][12]}")
    report_duration.add_run(f"Report Duration: From {start_date} to {end_date}")
    report_duration.alignment = WD_ALIGN_PARAGRAPH.CENTER

    paragraph = document.add_paragraph()
    paragraph.add_run(f"Student Name: {student_name}").bold = True

    # Attendance report table
    table = document.add_table(rows=1, cols=3)

    # Add header row
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Total Live Session Scheduled'
    hdr_cells[1].text = 'Attended'
    hdr_cells[2].text = 'Percentage'
    
    # Add attendance data
    row_cells = table.add_row().cells
    row_cells[0].text = str(len(valid_dates))  # Total scheduled sessions
    row_cells[1].text = str(attended_sessions)
    row_cells[2].text = str(attendance_percentage)
    
#     Rating section
#     document.add_paragraph(f"Rating: {rating}")
#     document.add_paragraph(f"Action needed: {action_needed}")
    
        # Save the document for the student
        
    file_name = f"{student_name.replace(' ', '_')}.docx"
    output_path = f"{output_folder}{file_name}"
    document.save(output_path)
    print(f"Document saved for {student_name}: {output_path}")

