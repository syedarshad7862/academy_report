from openpyxl import load_workbook
from datetime import datetime
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import RGBColor
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

# for docx to pdf
from docx2pdf import convert

# Create a new Document
document = Document()

# Output folder for saving docx files
output_folder = r'C:\Users\ThinkPad\Desktop\python project\\'
# output_folder = r'C:\Users\ThinkPad\Desktop\academy report\\'

wb = load_workbook(filename = r'C:\Users\ThinkPad\Desktop\academy report\Master Sheet - Student Performance Report.xlsx', data_only=True)


# it is for rating sheet
rating_sheet = wb["Rating"]

skip_row_of_rating = 3
start_cell_of_rating = 2

rating_sheet_data = {}
sheet_data_attendance = []
sheet_data_assignment = []
sheet_data_mock = []
# rating_names = ["Attendance report"]
for i, row in enumerate(rating_sheet.rows, start=1):
      if i <= skip_row_of_rating:
            continue
      print(f"row {i} row_data {row[start_cell_of_rating].value}")
      row_data = []
      for j, cell in enumerate(row[start_cell_of_rating:], start=start_cell_of_rating):
            print(cell.value)
            row_data.append(cell.value)
            # rating_sheet_data["data"] = row_data
      if row_data:
            if len(sheet_data_attendance) < 3:
                 sheet_data_attendance.append(row_data)
            elif len(sheet_data_assignment) < 3:
                 sheet_data_assignment.append(row_data)
            else:
                 sheet_data_mock.append(row_data)           
rating_sheet_data["attendance report"] = sheet_data_attendance
rating_sheet_data["assignment report"] = sheet_data_assignment
rating_sheet_data["mock report"] = sheet_data_mock

         
     
# print(f" it only attendance report{rating_sheet_data} and {len(rating_sheet_data)}")
# print(f" it only attendance report{rating_sheet_data['Attendance report']}")


attendance_rating_criteria = rating_sheet_data["attendance report"]
assignment_rating_criteria = rating_sheet_data["assignment report"]
mock_test_rating_criteria = rating_sheet_data["mock report"]
# print(f"rating_criteria {rating_criteria}")

# function for percentage of attendance rating
def get_rating(p):
    for criteria in attendance_rating_criteria:
        attendance_report, rating, lower_bound, upper_bound, action, empty, empty2,areas_of_improvement, final_score_basis,overall_feedback= criteria
        # print(rating)
        if  lower_bound <= p <= upper_bound:
            return attendance_report, rating, action,areas_of_improvement, final_score_basis
    return "unknown", "no action"

# function for percentage of assignment rating
def get_assignment_rating(p):
    for criteria in assignment_rating_criteria: #i have change here
        assignment_report, rating, lower_bound, upper_bound, action, deadline, empty, areas_of_improvement, final_score_basis, empty2 = criteria
        # print(rating)
        if  lower_bound <= p <= upper_bound:
            return assignment_report, rating, action, deadline,areas_of_improvement,final_score_basis
    return "unknown", "no action", "no deadline"

# function for percentage of mock test rating
def get_mock_test_rating(p):
    for criteria in mock_test_rating_criteria:
        mock_test_report, rating, lower_bound, upper_bound, action, deadline, mock_test_feedback, areas_of_improvement, final_score_basis, empty = criteria
        # print(rating)
        if  lower_bound <= p <= upper_bound:
            return mock_test_report,rating, action, deadline, mock_test_feedback,areas_of_improvement,final_score_basis
    return "unknown", "no action", "no deadline"
# function for feedback of mock test
def get_feed_back(p):
    for criteria in mock_test_rating_criteria:
        _, rating, lower_bound, upper_bound, action, deadline, mock_test_feedback,areas_of_improvement, final_score_basis, empty = criteria
        # print(rating)
        if  lower_bound <= p <= upper_bound:
            return mock_test_feedback
    return "no feedback"

# function for overall report and final socre
def get_final_score(p):
    for criteria in attendance_rating_criteria:
        attendance_report, rating, lower_bound, upper_bound, action, empty, empty2,areas_of_improvement, final_score_basis, overall_feedback= criteria
        # print(rating)
        if  lower_bound <= p <= upper_bound:
            return overall_feedback
    return "unknown", "no action"

# it for looping multiples sheets
sheet_names = ["Attendance report", "Assignment report", "Mock Test Report", "Mock Test Meta"]

skip_rows = {
    'Attendance report': 5, 
    'Assignment report': 6,
    'Mock Test Report': 6,
    'Mock Test Meta': 3
      
}

start_cells = {
      'Attendance report': 1,
      'Assignment report': 2,
      'Mock Test Report': 1,
      'Mock Test Meta': 2
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

# print(f" it is dictionary len: {len(sheets_data)} it list len {len(sheet_data)}")
# print(f" it is dictionary data : {sheets_data}")

# print(len(sheets_data[sheet_names[0]])) # it has 18 elements
attendance_sheet_data = sheets_data[sheet_names[0]]
assignment_sheet_data = sheets_data[sheet_names[1]]
mock_test_sheet_data = sheets_data[sheet_names[2]]
mock_test_meta_sheet_data = sheets_data[sheet_names[3]]

# make dictionary of mock_test_meta for grades
grade_action_mapping = {row[0] : {"action": row[1], "score": row[2]} for row in mock_test_meta_sheet_data if row[0] is not None }
print(grade_action_mapping)
# print(assignment_sheet_data)
# loop for assignments
# for i,std_assignment in enumerate(assignment_sheet_data[5:], start=1):
#     submitted = std_assignment[6]
#     assignment_percentage = (submitted/assignment_due) * 100 # assignment percentage
#     print(f"row {i} student data {std_assignment}, {submitted}, {assignment_percentage}")
#     if i == 13:
#          break
# print(attendance_sheet_data[2:])
# print(attendance_sheet_data[0][1], attendance_sheet_data[0][12])
# session_dates = attendance_sheet_data[0][1:]
# for sheet_name, data in sheets_data.items():
#       print(f"\n data from {sheet_name}:")
#       for row in data:
#             print(row)

session_dates = attendance_sheet_data[0][1:]  # Skip the header
valid_dates = [date for date in session_dates if date is not None]
start_date = valid_dates[0]
end_date = valid_dates[-1]
# calculates the sessions
total_sessions = attendance_sheet_data[3][14]
# calculates the assignemts
assignment_due = assignment_sheet_data[3][6]
# mock test due
mock_test_due = mock_test_sheet_data[3][6]
# print(f"data of arr : {attendance_sheet_data[3][14]}")

            #function for assignment names table start here
def parse_assignment_report(data):
    # Extract assignment names
    assignments = data[1][1:6]  # Get the columns for assignment names
    
    # Extract student data
    student_data = []
    for row in data[5:]:
        if row[0]:  # Check if the first cell (Student Name) exists
            student = {
                "name": row[0].strip(),
                "assignments": row[1:6],  # Assignment submission status
                "total_due": row[6],
                "total_submitted": row[7],
            }
            student_data.append(student)
    
    return assignments, student_data  

def get_pending_assignments(assignments, student):
    pending = []
    for assignment, submitted in zip(assignments, student["assignments"]):
        if submitted == 0:  # Check for unsubmitted (0)
            pending.append({
                "name": assignment,
                # "rating": student["rating"],
                # "action_needed": student["action_needed"],
            })
            
    return pending
# function for mock test grades
def parse_mock_test_report(data):
    # Extract assignment names
    mock_test_names = data[1][1:6]  # Get the columns for mock test names
    
    # Extract student data
    student_data = []
    for row in data[5:]:
        if row[0]:  # Check if the first cell (Student Name) exists
            student = {
                "name": row[0].strip(),
                "mock_test_grade": row[1:6],  # mock test submission status
                "total_due": row[6],
                "total_submitted": row[7],
            }
            student_data.append(student)
    
    return mock_test_names, student_data

def get_moct_test(mock_test_names, student):
    pending = []
    for mock_test, submitted in zip(mock_test_names, student["mock_test_grade"]):
        if submitted == 0 or submitted == "A" or submitted == "B" or submitted == "C" or submitted == "D" or submitted == "E":  # Check for unsubmitted (0)
            pending.append({
                "name": mock_test,
                "grade": student["mock_test_grade"],
                # "action_needed": student["action_needed"],
            })
            
    return pending
# Loop through each student row and generate a report
for i,student_row in enumerate(attendance_sheet_data[5:]):
    print(student_row)
    student_name = student_row[0]
    attended_sessions = student_row[14]  # Attended sessions count
    print(attended_sessions)
#     attendance_percentage = student_row[15]  # Attendance percentage
    attendance_percentage = (attended_sessions/total_sessions) * 100 # Attendance percentage
#     print(f"attended: {attended_sessions} and {attendance_percentage * 100:.2f}%")
    print(f"attended: {attended_sessions} and {attendance_percentage}")
     
    # calling function for attendance rating
    attendance_report,rating, action_needed,areas_of_improvement, final_score_basis = get_rating(attendance_percentage)
    print(rating,action_needed)
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
    line = report_duration._element
    p_pr = line.get_or_add_pPr()
    p_borders = parse_xml(
        r'<w:pBdr {}>'
        r'  <w:bottom w:val="single" w:sz="12" w:space="1" w:color="000000"/>'
        r'</w:pBdr>'.format(nsdecls('w'))
    )
    p_pr.append(p_borders)
    paragraph = document.add_paragraph(f"Student Name: ")
    paragraph.add_run(student_name).font.color.rgb = RGBColor(0, 0, 255)  # RGB for blue
    line2 = paragraph._element
    p_pr = line.get_or_add_pPr()
    p_borders = parse_xml(
        r'<w:pBdr {}>'
        r'  <w:bottom w:val="single" w:sz="12" w:space="1" w:color="000000"/>'
        r'</w:pBdr>'.format(nsdecls('w'))
    )
    p_pr.append(p_borders)
    
    

    attendance_title = document.add_paragraph()
    attendance_title.add_run("A. Attendance Report")
    attendance_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # Attendance report table
    table = document.add_table(rows=1, cols=3)

    # Add header row
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Total Live Session Scheduled'
    hdr_cells[1].text = 'Attended'
    hdr_cells[2].text = 'Percentage'
     
    # table.style = "Grid Table 6 Colorfull - Accent 5"
    table.style = "Light Grid"
    # convert the 1 to 100%
    percentage = f"{attendance_percentage:.2f}%"
    print(percentage)
    # Add attendance data
    row_cells = table.add_row().cells
    row_cells[0].text = str(total_sessions)  # Total scheduled sessions
    row_cells[1].text = str(attended_sessions)
    row_cells[2].text = str(percentage)
    print(percentage)
    #Rating section
    paragraph2 = document.add_paragraph(f"Rating: ")
    paragraph2.add_run(rating).font.color.rgb = RGBColor(0, 0, 255)
    
    paragraph3 = document.add_paragraph(f"Action needed:")
    paragraph3.add_run(action_needed).font.color.rgb = RGBColor(0,0,255)
    
      #Assignment report start here
    attendance_title = document.add_paragraph()
    attendance_title.add_run("B. Assignment Report")
    attendance_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
#     for i,std_assignment in enumerate(assignment_sheet_data[5:], start=1):
#     for i,std_assignment in enumerate(assignment_sheet_data[5:], start=1):
    std_assignment = assignment_sheet_data[i+5]
    submitted = std_assignment[6]
    assignment_percentage = (submitted/assignment_due) * 100 # assignment percentage
    print(f"row {i} student data {std_assignment}, {submitted}, {assignment_percentage}")
      
    
        # Assignment report table
    table = document.add_table(rows=1, cols=3)
    table.style = "Light Grid"
    # Add header row
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Assignment Due'
    hdr_cells[1].text = 'Submitted'
    hdr_cells[2].text = 'Percentage'
    
    # convert the 1 to 100%
    converted_assignment_percentage = f"{assignment_percentage:.2f}%"
#     print(assignment_percentage)
    # Add assignment data
    row_cells = table.add_row().cells
    row_cells[0].text = str(assignment_due)  # Total assignment
    row_cells[1].text = str(submitted)
    row_cells[2].text = str(converted_assignment_percentage)
    
    # calling function for assignment rating
    assignment_report,assignment_rating, assignment_action_needed, deadline,areas_of_improvement_in_assignment,assignment_score = get_assignment_rating(assignment_percentage)
    print(assignment_rating,assignment_action_needed)
    #Rating section
    asm_rating = document.add_paragraph(f"Rating: ")
    asm_rating.add_run(assignment_rating).font.color.rgb = RGBColor(0, 0, 255)  # RGB for blue
    act_needed = document.add_paragraph(f"Action needed: ")
    act_needed.add_run(assignment_action_needed).font.color.rgb = RGBColor(0, 0, 255)  # RGB for blue
    
                #assignment names table start here
    # Parse the data
    assignments, student_data = parse_assignment_report(sheets_data['Assignment report'])
    # print(assignments)
    # print(student_data)
    test_data = student_data[i]
    
    # Select a student (e.g., "JISNA JOSEPH")
    search_student_name = student_name # student_name
    # selected_student = next((s for s in student_data if s["name"] == search_student_name), None)
    # selected_student = next((s for s in student_data[0 + i] if s == search_student_name), None)
    selected_student = test_data["name"]

    if selected_student:
        pending_assignments = get_pending_assignments(assignments, test_data)
        print(f"Pending assignments for {search_student_name}:")
                    # Add header row
            # assignment detail table
        table3 = document.add_table(rows=1, cols=3)
        hdr_cells = table3.rows[0].cells
        hdr_cells[0].text = 'Assignment/s are to be submitted.'
        hdr_cells[1].text = 'Submission Rating'
        hdr_cells[2].text = 'Deadline'
        table3.style = "Light Grid"
        for n, assignment in enumerate(pending_assignments, start=1) :
            # print(f"- {assignment['name']} | Rating: {assignment['rating']} | Action: {assignment['action_needed']}")
            print(f"- {assignment['name']}, Submission Rating : {assignment_rating}, Deadline: {deadline}")


            
            row_cells = table3.add_row().cells
            row_cells[0].text = str(assignment["name"])  
            row_cells[1].text = str(assignment_rating)
            row_cells[2].text = str(deadline)
            
        
    else:
        print(f"No data found for student: {search_student_name}")
    
    
    #Mock test report start here
    attendance_title = document.add_paragraph()
    attendance_title.add_run("C. Mock Test Report")
    attendance_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    std_mock_test = mock_test_sheet_data[i+5]
    mock_test_attended = std_mock_test[6]
    mock_test_percentage = (mock_test_attended/mock_test_due) * 100 # mock test percentage
    print(f"row {i} student data {std_mock_test}, {mock_test_attended}, {mock_test_percentage}")
    
        # mock test report table
    table3 = document.add_table(rows=1, cols=3)
    table3.style = "Light Grid"
    
        # Add header row
    hdr_cells = table3.rows[0].cells
    hdr_cells[0].text = 'Mock Test Conducted'
    hdr_cells[1].text = 'Mock Test Attended'
    hdr_cells[2].text = 'Percentage'
    
    # convert the 1 to 100%
    converted_moct_test_percentage = f"{mock_test_percentage:.2f}%"
    # Add mock test data
    row_cells = table3.add_row().cells
    row_cells[0].text = str(mock_test_due)  # Total mock test
    row_cells[1].text = str(mock_test_attended)
    row_cells[2].text = str(converted_moct_test_percentage)

    # calling function for mock test rating
    mock_test_report,mock_test_rating, mock_test_action_needed, mock_deadline, mock_test_feedback,areas_of_improvement_in_mock,mock_test_score = get_mock_test_rating(mock_test_percentage)
    print(mock_test_rating,mock_test_action_needed, mock_test_feedback)
    #Rating section
    mock_rating = document.add_paragraph(f"Rating: ")
    mock_rating.add_run(f"{mock_test_rating}").font.color.rgb = RGBColor(0,0,255)
    mock_action = document.add_paragraph(f"Action needed: ")
    mock_action.add_run(f"{mock_test_action_needed}").font.color.rgb = RGBColor(0,0,255)  
    
    # mock test grade table start here
    mock_test_names, student_of_mock_test = parse_mock_test_report(sheets_data['Mock Test Report'])
    mock_data = student_of_mock_test[i]
    search_student_name = student_name # student_name
    mock_selected_student = mock_data["name"]
    result = []
    if mock_selected_student:
        mock_names = get_moct_test(mock_test_names, mock_data)
        grades = mock_data["mock_test_grade"]
        print(f"mock test  for {search_student_name}:")
        student_result = {"name": student_name, "grades": []}
        for mock_test, grade in zip(mock_names, grades):
            if grade in grade_action_mapping:
                print(grade)
                action_details = grade_action_mapping[grade]
                student_result["grades"].append({
                 "mock_test_name": mock_test,   
                "grade": grade,
                "action": action_details["action"],
                "score": action_details["score"]
            })
            else:
                 student_result["grades"].append({
                "grade": grade,
                "action": "No action available",
                "score": None
                })
        result.append(student_result)     
                 
                  
        # for g, mock_test in enumerate(mock_names):
        #     # print(f"- {assignment['name']} | Rating: {assignment['rating']} | Action: {assignment['action_needed']}")
        #     print(f"- {mock_test['name']} grade : {mock_test['grade'][g]}")
        
    else:
        print(f"No data found for student: {search_student_name}")
    total_score = 0    
    table4 = document.add_table(rows=1, cols=3)
    table4.style = "Light Grid"
        # Add header row
    hdr_cells = table4.rows[0].cells
    hdr_cells[0].text = 'Mock Test Name'
    hdr_cells[1].text = 'Grade'
    hdr_cells[2].text = 'Future Action'
    for student in result:
        print(f"Student Name: {student['name']}")
        for grade_detail in student["grades"]:
            total_score += grade_detail['score']
            # Add mock test data
            row_cells = table4.add_row().cells
            row_cells[0].text = str(grade_detail['mock_test_name']['name'])  # Total mock test
            row_cells[1].text = str(grade_detail['grade'])
            row_cells[2].text = str(grade_detail['action'])
            # print(f"mock test name: {grade_detail['mock_test_name']['name']}  Grade: {grade_detail['grade']} | Action: {grade_detail['action']} | Score: {grade_detail['score']} | total_score = {total_score}")
        print("-" * 50) 
    print(f"{(total_score/50)*100}")  
    mock_percentage = (total_score/50)*100
    Overall_Score_rating = document.add_paragraph("Overall Score rating: ")
    Overall_Score_rating.add_run(f" Based on the Grade & assigned score (A-10, B-8, C-6, D-4, E-2) = {mock_percentage}%").font.color.rgb = RGBColor(0, 0, 255)
    feed_back = get_feed_back(mock_percentage)
    print(feed_back)   
    mock_test_feedback = document.add_paragraph(f"FeedBack:")
    mock_test_feedback.add_run(feed_back).font.color.rgb = RGBColor(0, 0, 255)
    
    overall_report = document.add_paragraph("Overall Report")
    overall_report.alignment = WD_ALIGN_PARAGRAPH.CENTER
     #Overall Report
    table5 = document.add_table(rows=1, cols=3)
    table5.style = "Light Grid"
        # Add header row
    hdr_cells = table5.rows[0].cells
    hdr_cells[0].text = 'Title'
    hdr_cells[1].text = 'Ratin/ Overall Grade'
    hdr_cells[2].text = 'Area of Improvement'

    row1 = table5.add_row().cells
    row1[0].text = str(attendance_report)
    row1[1].text = str(rating)
    row1[2].text = str(areas_of_improvement)
    row2 = table5.add_row().cells
    row2[0].text = str(assignment_report)
    row2[1].text = str(assignment_rating)
    row2[2].text = str(areas_of_improvement_in_assignment)
    row3 = table5.add_row().cells
    row3[0].text = str(mock_test_report)
    row3[1].text = str(mock_test_rating)
    row3[2].text = str(areas_of_improvement_in_mock)
    final_score = final_score_basis + mock_test_score + assignment_score
    # print(final_score/30*100)
    final_score_percentage = (final_score/30)*100
    overall_feedback = get_final_score(final_score_percentage)
    # print(overall_feedback)
    final_score = document.add_paragraph("Final Score: ")
    final_score.add_run(f"{final_score_percentage}%").font.color.rgb = RGBColor(0,0,225)
    feedback = document.add_paragraph("Feedback: ")
    feedback.add_run(overall_feedback).font.color.rgb = RGBColor(0,0,255)
    # Save the document for the student        
    file_name = f"{student_name.replace(' ', '_')}.docx"
    output_path = f"{output_folder}{file_name}"
    document.save(output_path)
    # convert
    print(f"Document saved for {student_name}: {output_path}")

