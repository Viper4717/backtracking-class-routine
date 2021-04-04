import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile

# taking the excel file as input
file_fullpath = 'input.xlsx'
excel_input = pd.ExcelFile(file_fullpath)
sheet_to_df_map = {}
for sheet_name in excel_input.sheet_names:
    sheet_to_df_map[sheet_name] = excel_input.parse(sheet_name)

# 1.5 = 1 class
# 0.75 = 0.5 class
# More than 1.5 = 2 class
# lab(1.5 credit) = 1 3-hour class
# year _ theory/lab _ uniqueID _ totalGroup _ sectionID _ classID

# function to encode time slots
def time_parse(time):
    encoded_time_list = []
    lunch_split = time.split(";")
    for slot in lunch_split:
        start_end_split = slot.split("-")
        start_end_list = []
        for hour in start_end_split:
            hour_min_split = hour.split(":")
            # print(hour_min_split[0]+":"+hour_min_split[1])
            # print(hour_min_split[1][-2:])
            if(hour_min_split[1][-2:] == "am" or hour_min_split[1][-2:] == "AM"):
                hour_to_min = (int(hour_min_split[0])*60) + int(hour_min_split[1][:2])
            else:
                hour_to_min = ((int(hour_min_split[0])+12)*60) + int(hour_min_split[1][:2])
            start_end_list.append(hour_to_min)
        for i in range(start_end_list[0], start_end_list[1], 30):
            encoded_time_list.append(i)
    str_encoded_time_list = map(str, encoded_time_list)
    return list(str_encoded_time_list)

# function to encode the course codes
def code_parse(code, code_len):
    encoded_code = ""
    for i in range(code_len):
        if(i == 1):
            continue
        encoded_code += code[i]
    return encoded_code

# storing the free time values for every teacher
teacher_to_time_map = {}
for index, row in sheet_to_df_map["SampleInputWithSolution"].iterrows():
    if(isinstance(row[0], str)):
        date_index = 0
        for time in row[1:]:
            teacher_name = row[0]
            if(type(time) != float):
                encoded_times = time_parse(time)
                encoded_days_times = [str(date_index) + etime for etime in encoded_times]
                if(teacher_name in teacher_to_time_map):
                    teacher_to_time_map[teacher_name].extend(encoded_days_times)
                else:
                    teacher_to_time_map[teacher_name] = encoded_days_times
            date_index+=1

# print(teacher_to_time_map["ST"])
# print(teacher_to_time_map["SP"])

# creating the time domains for the courses and encoding the course codes
total_courses = 0
course_variable = {}
course_variable_time_domain = {}
course_to_teacher_map = {}
teacher_to_course_map = {}
course_to_credit_map = {}
section_course_map = {}
lab_course_set = set()
assigned_courses = sheet_to_df_map["AssignedCourses"]
for index, row in assigned_courses.iterrows():
    teacher_name = row[0]
    teacher_to_course_map[teacher_name] = []
    for name in row[1:]:
        course_list = sheet_to_df_map["UndergradCurriculumOptional"]
        if(type(name) != float):
            # course_name = None
            # code = None
            if(("Section" or "section") in name):
                course_name = name[:len(name)-10]
                cr = course_list.loc[course_list["Course"] == course_name]["Credit"]
                code = course_name[-4:]
            else:
                cr = course_list.loc[course_list["Course"] == name]["Credit"]
                code = name[-4:]
            credit = cr.to_list()
            code_len = len(code)
            encoded_code = code_parse(code, code_len)
            # list_for_set = None
            # course_teacher_df_to_list = None
            if(encoded_code[1] == "1" or (int(encoded_code[1])>4)): # means the course is a lab course
                if(name not in lab_course_set):
                    # finding out the course teachers for this specific lab course
                    course_teacher_df = assigned_courses.loc[(assigned_courses["Course1"] == name) | (assigned_courses["Course2"] == name) |
                    (assigned_courses["Course3"] == name) | (assigned_courses["Course4"] == name) | (assigned_courses["Course5"] == name)]["Teacher"]
                    lab_course_set.add(name)
                    course_teacher_df_to_list = course_teacher_df.to_list()
                    # course_to_teacher_map[encoded_code] = course_teacher_df_to_list
                    # if course teacher is more than 1, finding the common time slots
                    if(len(course_teacher_df_to_list) > 1):
                        for i in range(len(course_teacher_df_to_list)):
                            if(i>0):
                                intersection = list_for_set.intersection(teacher_to_time_map[course_teacher_df_to_list[i]])
                                list_for_set = intersection
                            else:
                                list_for_set = set(teacher_to_time_map[course_teacher_df_to_list[i]])
                        list_for_set = list(list_for_set)
                    else:
                        list_for_set = teacher_to_time_map[course_teacher_df_to_list[0]]
            if(("Section" or "section") in name): # finding out the highest section value for that lab course
                if(course_name not in lab_course_set):
                    highest_section_value = 0
                    for i in range(1, 6):
                        column_name = "Course" + str(i)
                        course_name_with_section = assigned_courses[assigned_courses[column_name].str.contains(course_name, na=False)][column_name]
                        if(not course_name_with_section.empty):
                            course_name_with_section_list = course_name_with_section.to_list()
                            for cname in course_name_with_section_list:
                                highest_section_value = max(highest_section_value, int(cname[-1:]))
                    lab_course_set.add(course_name)
                    section_course_map[course_name] = highest_section_value
                    encoded_code += str(highest_section_value) + name[-1:] + '0'
                else:
                    encoded_code += str(section_course_map[course_name]) + name[-1:] + '0'
            else:
                encoded_code += "100"
            # course_variable[encoded_code] = None
            if(encoded_code[1] == "1" or (int(encoded_code[1])>4)):
                course_variable_time_domain[encoded_code] = list_for_set
                course_to_teacher_map[encoded_code] = course_teacher_df_to_list
                # print(name + " " + encoded_code)
                # print(course_teacher_df_to_list)
            else:
                course_variable_time_domain[encoded_code] = teacher_to_time_map[teacher_name]
                course_to_teacher_map[encoded_code] = [teacher_name]
                # print(name + " " + encoded_code)
                # print(teacher_name)
            # if credit > 1.5, means there has to be 2 classes
            teacher_to_course_map[teacher_name].append(encoded_code)
            course_to_credit_map[encoded_code] = credit[0]
            total_courses += 1
            if(credit[0]>1.5):
                encoded_code_2 = encoded_code[:-1] + "1"
                # course_variable[encoded_code_2] = None
                if(encoded_code[1] == "1" or (int(encoded_code[1])>4)):
                    course_variable_time_domain[encoded_code_2] = list_for_set
                    course_to_teacher_map[encoded_code_2] = course_teacher_df_to_list
                else:
                    course_variable_time_domain[encoded_code_2] = teacher_to_time_map[teacher_name]
                    course_to_teacher_map[encoded_code_2] = [teacher_name]
                teacher_to_course_map[teacher_name].append(encoded_code_2)
                course_to_credit_map[encoded_code_2] = credit[0]
                total_courses += 1

def prune_data(current_course, booked_time_list):
    year = current_course[0]
    theo_or_lab = current_course[1]
    c_id = current_course[2]
    total_sec = current_course[3]
    sec_id = current_course[4]
    class_id = current_course[5]
    prunemap = {}
    # first remove the current course from dictionary
    prunemap[current_course] = course_variable_time_domain[current_course]
    course_variable_time_domain.pop(current_course)
    # finding out the teachers associated with this course
    teachers = course_to_teacher_map[current_course]
    # finding out all the courses these teachers are associated with
    for t in range(len(teachers)):
        if(t>0):
            courses_to_prune.extend(teacher_to_course_map[teachers[t]])
        else:
            courses_to_prune = teacher_to_course_map[teachers[t]]
    # pruning the common teacher courses
    for crs in courses_to_prune:
        removed_time_list = []
        for item in booked_time_list:
            if(item in course_variable_time_domain[crs]):
                course_variable_time_domain[crs].remove(item)
                removed_time_list.append(item)
        prunemap[crs] = removed_time_list
    # finding all the courses for current course year
    same_year_courses = [key for key, val in course_variable_time_domain.items() if year == key[0]]
    for crs in same_year_courses:
        if(year == "4"):
            continue
            # genjam
        else:
            if((int(theo_or_lab) + int(crs[1])) > 1):
                continue
                # genjam 2
            else:
                removed_time_list = []
                for item in booked_time_list:
                    if(item in course_variable_time_domain[crs]):
                        course_variable_time_domain[crs].remove(item)
                        removed_time_list.append(item)
                prunemap[crs] = removed_time_list
    return prunemap

# main backtracking function
def backtrack():
    if(len(course_variable) == total_courses):
        print("Routine created")
        return
    # finding out the minimum length of lists of time in the domain
    min_len_dom = min([len(course_variable_time_domain[crs]) for crs in course_variable_time_domain])
    for key, val in course_variable_time_domain.items():
        if len(val) == min_len_dom:
            current_course = key
            break
    # testing for every valid time in course-time domain
    for dtime in course_variable_time_domain[current_course]:
        course_credit = course_to_credit_map[current_course]
        day = dtime[0]
        time = int(dtime[1:])
        if(current_course[1] == "1"):
            end_time = time + 180
        else:
            if(course_credit<3.0):
                end_time = time + 60
            else:
                end_time = time + 90
        booked_time_list = []
        book = True
        for gap in range(time, end_time, 30):
            booked_time_list.append(day+str(gap))
        for btime in booked_time_list:
            if(btime not in course_variable_time_domain[current_course]):
                book = False
                break
        if(book):
            prunemap = prune_data(current_course, booked_time_list)
            backtrack()
        # reverse pruning

# def backtrack(data, considerNumber):
#     if(variable.isFull())
#         saveSolutions = variable.courses
#         return
    
#     course = // choose argmin(Domain) among Course -> Domain map.
    
#     if(courseDomain[course].isEmpty()):
#         if(considerNumber == 0)
#             return
#         considerNumber--

#     for i in courseDomain[course]:
#       variable.course.append(i)
#       variable.count++
#       list = [i, i+1, i+2]// list depends on number of slots booked.
#       pruneMap = pruneDomainForList (list)
#       backtrack(data)
#       fillDomainForList (pruneMap)