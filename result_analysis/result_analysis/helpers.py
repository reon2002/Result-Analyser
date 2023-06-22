import os
import tabula
import pandas as pd
from django.shortcuts import render,redirect
# from .helpers import handle_uploaded_file,convert_to_excel,extract_data
from django.core.files.storage import FileSystemStorage
from django.http import FileResponse,HttpResponseNotFound
from django.shortcuts import HttpResponse
import tabula
import pandas as pd
import os
from django.conf import settings
import csv
from openpyxl import Workbook
import time
import win32com.client as win32
import pythoncom
import matplotlib.pyplot as plt
from openpyxl.drawing.image import Image

def handle_uploaded_file(file):
    # Specify the upload directory
    upload_dir = 'media/uploads/'
    # Create the directory if it doesn't exist
    os.makedirs(upload_dir, exist_ok=True)
    # Save the uploaded file to the upload directory
    file_path = os.path.join(upload_dir, file.name)
    with open(file_path, 'wb') as destination:
        for chunk in file.chunks():
            destination.write(chunk)
    return file_path

def convert_to_excel(file_path):
    # Specify the output directory for the converted Excel file
    output_dir = 'media/converted/'
    # Create the directory if it doesn't exist
    os.makedirs(output_dir, exist_ok=True)
    # Convert the PDF to Excel using Tabula
    output_path = os.path.join(output_dir, 'output.xlsx')
    tabula.convert_into(file_path, output_path, output_format='xlsx', pages='all')
    return output_path

def extract_data(excel_path):
    # Read the Excel file into a pandas DataFrame
    df = pd.read_excel(excel_path)
    # Extract the data from the DataFrame
    data = df.values.tolist()  # Convert DataFrame to a list of lists
    return data

# def perform_analysis(data):
#     # Perform additional analysis on the extracted data
#     # ...
#     # Return the analyzed data
#     return analyzed_data

def create_empty_csv(directory, filename):
    # Create an empty CSV file in the specified directory
    df = pd.DataFrame()
    file_path = os.path.join(directory, filename)
    df.to_csv(file_path, index=False)
    return file_path

def analyzeS1(uploaded_file):
    
    # Get the current directory
    current_dir = os.getcwd()

    # Define the path to the media directory
    media_dir = os.path.join(current_dir, 'media')

    # Define the path to the uploads directory inside the media directory
    uploads_dir = os.path.join(media_dir, 'uploads')

    # Create the uploads directory if it doesn't exist
    if not os.path.exists(uploads_dir):
        os.makedirs(uploads_dir)

    # Initialize csv_path variable
    csv_path = None

    try:
        file = tabula.read_pdf(uploaded_file, pages="all")
        tabula.convert_into(uploaded_file, os.path.join(uploads_dir, "output2.csv"), output_format="csv", pages="all")
        csv_path = os.path.join(uploads_dir, "output2.csv")
    except Exception as e:
        print("Error:", e)


    file = file

    li=[]
    # Open the CSV file
    with open(csv_path, 'r') as file:
        # Create a CSV reader object
        reader = csv.reader(file)

        # Iterate over each row in the CSV file
        for row in reader:
            # Access the values in each row
            # Example: Print the values in each row
            for i in row:
                if i.startswith('TRV20IT'):
                    li.append((row))
                elif i.startswith('LTRV20IT'):
                    li.append((row))
                elif i.startswith('PKD20IT'):
                    li.append((row))
                else:
                    continue
    time.sleep(1)
    output_file='output.csv'
    create_empty_csv(uploads_dir,output_file)
    length=len(li)
    csv1_path = os.path.join(uploads_dir, "output.csv")
    with open(csv1_path, 'w', newline='') as file1:
        writer = csv.writer(file1)
        for i in range(length):
            writer.writerow(li[i])


    # Function to find index of a substring from a list
    def findsubstring(s1):
         with open("output.csv", 'r') as file:
            reader=csv.reader(file)
            for s in reader:
                r=s[1].split(",")
                l2=list(r)
                for j in l2:
                 if s1 in j:
                    return(l2.index(j))
                    break
            else:
                return(-1)
                
    def print_string_between_parentheses(input_string):
        opening_index = input_string.find('(')
        closing_index = input_string.find(')')
        
        if opening_index != -1 and closing_index != -1:
            substring = input_string[opening_index + 1 : closing_index]
            return substring

    def add_list(st):
        with open(csv1_path, 'r') as file:
            reader = csv.reader(file)
            li = []
            for row in reader:
                r1=row[1].split(",")
                if findsubstring(st) != -1:
                    index = findsubstring(st)

                    li.append(print_string_between_parentheses(r1[index]))
                else:
                    li.append("Absent")
            return li

    def add_regID(filename):
        with open(filename, 'r') as file:
            reader = csv.reader(file)
            li = []
            listTRV = []
            for row in reader:
                if 'TRV' in row[0]:
                    listTRV.append(row[0])
                elif 'PKD' in row[0]:
                    listTRV.append(row[0])
            return listTRV


    grade_points = {'S': 10, 'A+': 9.0, 'A': 8.5, 'B+': 8.0, 'B': 7.5, 'C+': 7.0, 'C': 6.5, 'D': 6.0, 'P': 5.5, 'F': 0, 'FE': 0, 'Ab': 0}
    credits = {'MAT101':4,'PHT100':4,'EST110':3,'EST120':4,'HUN101':0,'PHL120':1,'ESL120':1}
    r1=[]
   
    time.sleep(1)
    listTRV = add_regID(csv1_path)
    l=list(credits.keys())
    listc1=add_list(l[0])
    listc2=add_list(l[1])
    listc3=add_list(l[2])
    listc4=add_list(l[3])
    listc5=add_list(l[4])
    listc6=add_list(l[5])
    listc7=add_list(l[6])
    
                           

    listnames=[listc1,listc2,listc3,listc4,listc5,listc6,listc7]
    # Calculate SGPA
                    
    sgpa=0
    grade_counts = {course_code: {'S': 0, 'A+': 0, 'A': 0, 'B+': 0, 'B': 0, 'C+': 0, 'C': 0, 'D': 0, 'P': 0, 'F': 0, 'FE': 0, 'Absent': 0} for course_code in credits}
    SGPA=[]
    grade_points = {'S': 10, 'A+': 9.0, 'A': 8.5, 'B+': 8.0, 'B': 7.5, 'C+': 7.0, 'C': 6.5, 'D': 6.0, 'P': 5.5, 'F': 0, 'FE': 0, 'Absent': 0}
    for i in range(length):
        has_failing_grade = False
        total_credit_points = 0
        total_credits = 0
        j=0
        for course_code in credits:
            # grade = eval("list" + course_code)[i]
        
            li=listnames[j]
            # print(li)
            grade=li[i]
            # print(grade,end=" ")
            if grade in grade_points:
                if grade in ['F', 'FE', 'Absent']:
                    grade_counts[course_code][grade] += 1
                    has_failing_grade = True
                else:
                    grade_counts[course_code][grade] += 1
                total_credit_points += credits[course_code] * grade_points[grade]
                total_credits += credits[course_code]
                j=j+1
        if has_failing_grade:
            sgpa = 0
        else:
            sgpa = total_credit_points / total_credits
        SGPA.append(sgpa)

    def find_top_students(filename, sgpa_list, reg_id_list):
        student_data = zip(reg_id_list, sgpa_list)
        sorted_students = sorted(student_data, key=lambda x: x[1], reverse=True)
        top_students = sorted_students[:3]
        return top_students
        

    # function to check if a particular combination of stu and sem already exists in the table
    # def check_combination_exists(stu_id, sem):
    #     combination_exists = Student.objects.filter(stu_id=stu_id, sem=sem).exists()
    #     return combination_exists
    # # to insert new rows to the table
    # def insert_student(stu_id, sem, sgpa):
    #                 student = Student(stu_id=stu_id, sem=sem, sgpa=sgpa)
    #                 student.save()
    # Write the result to a new CSV file
    time.sleep(1)
    output_file='output4.csv'
    create_empty_csv(uploads_dir,output_file)
    csv4_path = os.path.join(uploads_dir, "output4.csv")
    with open(csv4_path, 'w', newline='') as file1:
        writer = csv.writer(file1)
        writer.writerow(['RegID', l[0], l[1],l[2],l[3],l[4],l[5],l[6],'SGPA'])
        for i in range(length):
            # students = Student.objects.all()
            # if check_combination_exists(listTRV[i],"S1"):
            #    for student in students: 
            #     new_id=listTRV[i]
            #     new_sem="S1"
            #     new_sgpa=SGPA[i]
            #     student.stu_id = new_id
            #     student.sem = new_sem
            #     student.sgpa = new_sgpa
            #    Student.objects.bulk_update(students, ['stu_id', 'sem', 'sgpa'])
            # else:
            #     insert_student(listTRV[i],"S1",SGPA[i])



            #save data add id,add sem,add sgpa
            #get all sgpas of listTRV[i]
            #cgpa calculation
            
            writer.writerow([listTRV[i],add_list(l[0])[i],add_list(l[1])[i],add_list(l[2])[i],add_list(l[3])[i],add_list(l[4])[i],
                            add_list(l[5])[i],add_list(l[6])[i],SGPA[i]])#write cgpa

# save in this loop^^(database)

    top_students=find_top_students(csv1_path, SGPA, listTRV)
    # Convert the CSV to Excel format
    wb = Workbook()
    ws = wb.active

    time.sleep(1)
    with open(csv4_path, 'r') as file2:
        reader = csv.reader(file2)
        for row in reader:
            ws.append(row)
    ws.append([])  # Add an empty row for separation
    ws.append(["Top Three Students"])
    ws.append(["RegID", "SGPA"])
    for student in top_students:
        ws.append(student)
    # Add grade counts per subject to the Excel file
    ws.append([])  # Add an empty row for separation
    ws.append(["Grade Counts per Subject"])
    for course_code, grade_count in grade_counts.items():
        ws.append([])  # Add an empty row for separation
        ws.append(["Subject: " + course_code])
        ws.append(["Grade", "Count"])
        for grade, count in grade_count.items():
            ws.append([grade, count])
    ws.append([])  # Add an empty row for separation
    ws.append(["Supplementary Count for Each Course"])
    ws.append(['Course', 'Absent', 'F', 'FE'])
    # for course_code in dict:
    #     ws.append([course_code, (grade_count.get(course_code, 0)), count_F.get(course_code, 0), count_FE.get(course_code, 0)])
    for course_code,grades in grade_counts.items():
        fe_count=grades.get('FE',0)
        f_count=grades.get('F',0)
        ab_count=grades.get('Absent',0)
        ws.append([course_code,ab_count, f_count, fe_count])
       

    dict_passfail={course_code: {'passcount': 0, 'failcount': 0} for course_code in credits}
    for i in range(length):
        j=0
        for course_code in credits:
            # grade = eval("list" + course_code)[i]
        
            li=listnames[j]
            # print(li)
            grade=li[i]
            # print(grade,end=" ")
            if grade in grade_points:
                if grade in ['F', 'FE', 'Absent']:
                    dict_passfail[course_code]['failcount'] += 1
                else:
                    dict_passfail[course_code]['passcount'] += 1
                j=j+1

    dict_passpercent={course_code: 0 for course_code in credits}
    percent=0
    list_codes=list(credits.keys())
    j=0
    for i in dict_passfail.values():
            course_code=list_codes[j]
            # print(i['passcount'])
            percent=i['passcount']*100/(i['passcount']+i['failcount'])
            dict_passpercent[course_code]=percent
            j=j+1
            
            # print(percent)

    # print(dict_passfail)
    ws.append([])  # Add an empty row for separation
    ws.append(["Pass Percentage for Each Course"])
    ws.append(['Course','Percent'])
    # for course_code in dict:
    #     ws.append([course_code, (grade_count.get(course_code, 0)), count_F.get(course_code, 0), count_FE.get(course_code, 0)])
    for course_code,percent in dict_passpercent.items():
            ws.append([course_code,percent])
    plt.bar(list(dict_passpercent.keys()), list(dict_passpercent.values()))
    plt.xlabel('Course')
    plt.ylabel('PassPercent')
    plt.title('Bar Chart')
    plt.savefig('chart.png')
    img = Image('chart.png')
    img.width = 400
    img.height = 300
    ws.add_image(img, 'E213')
    # wb.save('output4.xlsx')


    return wb
    # output_path = os.path.join("media", "converted", "output4.xlsx")
    # # workbook = Workbook()
    # wb.save(output_path)

    # return output_path
