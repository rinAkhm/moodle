import json
import requests
import xlsxwriter
from requests.adapters import HTTPAdapter
import os
from dotenv import load_dotenv



load_dotenv()
token = os.getenv('token')
domain = os.getenv('domain')

#filter items
columfilterstring = "semester"

adapter = HTTPAdapter(max_retries=5)
session = requests.Session()
session.mount(domain, adapter)

#get courses 
url = ('%s/webservice/rest/server.php?moodlewsrestformat=json&wsfunction=core_course_get_courses&wstoken=%s')%(domain,token)
courses = session.get(url).json()
categories = [512, 521, 518, 517, 515, 511, 509, 507, 505]
print(f'all courses in moodle: {len(courses)}')

course_list = [course for course in courses for id_course in categories if id_course == course["id"]]
print(f'Courses found: {len(course_list)}')

allstudents = set()
pivot = {}
for i, course in enumerate(course_list):
    cid = course['id']
    url = f'{domain}/webservice/rest/server.php?moodlewsrestformat=json&wsfunction=core_enrol_get_enrolled_users&wstoken={token}&courseid={cid}'
    students = session.get(url).json()
    for student in students:
        for role in student['roles']:
            if role['shortname']=='student':
                allstudents.add((student["lastname"], student["firstname"], student["email"], student["id"]))
                if student["id"] not in pivot:
                    pivot[student["id"]] = {}
    print(f"\tstudents are ready ({len(students)})")

# get all grades
for i, course in enumerate(course_list):
    cid = course['id']
    print(f"Processing course {course['shortname']} / id={cid}, {i + 1} of {len(course_list)}")
    url = f'{domain}/webservice/rest/server.php?moodlewsrestformat=json&wsfunction=gradereport_user_get_grade_items&wstoken={token}&courseid={cid}'
    try:
        report = session.get(url, timeout=30).json()
    except IOError as e:
        print("\tEXCEPTION: ", e.strerror)
        continue

    if 'usergrades' not in report:
        print('\tfailed to process')
        continue
    gradeitems = report['usergrades']
    hideGrade = False
    for user in gradeitems:
        userid = user['userid']
        grades = user['gradeitems']
        for grade in grades:
            if columfilterstring.lower() in (grade['itemname'] or "").lower():
                # if grade['gradeishidden']: continue
                if not hideGrade: print(f"\tFound column: {grade['itemname']}")
                hideGrade = True
                if not userid in pivot: pivot[userid] = {}
                pivot[userid][cid] = grade['gradeformatted'] if 'gradeformatted' in grade else '??'
    print('\tGrades are ready')
    if not hideGrade: print("\t!!! No column found !!!")

print("Students in total: ", len(allstudents))

#Prepare workbook

allstudents = set()
ab = list('ABCDEFGHIJKLMNOPQRSTUVWXYZ')
columns = list(ab)
for p in ab:
    columns += [p + let for let in ab]
print(len(columns))

######################################################

workbook = xlsxwriter.Workbook("Midsemester.xlsx")
# The workbook object is then used to add new
# worksheet via the add_worksheet() method.
worksheet = workbook.add_worksheet()
worksheet.write('A1', 'last name')
worksheet.write('B1', 'first name')
worksheet.write('C1', 'email')


for i, course in enumerate(course_list):
    worksheet.write(columns[i + 3] + '1', course['fullname'])


studs = sorted(list(allstudents))
for i, student in enumerate(studs):
    worksheet.write(columns[0] + str(i + 2), student[0])
    worksheet.write(columns[1] + str(i + 2), student[1])
    worksheet.write(columns[2] + str(i + 2), student[2])
    
    for j, course in enumerate(course_list):
        grade = pivot[student[3]][course['id']] if course['id'] in pivot[student[3]] else ''
        worksheet.write(columns[j + 3] + str(i + 2), grade)

##################################################
workbook.close()