import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile

# taking the excel file as input
file_fullpath = 'input.xlsx'
excel_input = pd.ExcelFile(file_fullpath)
sheet_to_df_map = {}
for sheet_name in excel_input.sheet_names:
    sheet_to_df_map[sheet_name] = excel_input.parse(sheet_name)

# function to encode time slots
def time_parse(time):
    # print(time)
    encoded_time_list = []
    lunch_split = time.split(";")
    # print(lunch_split)
    for slot in lunch_split:
        start_end_split = slot.split("-")
        # print(start_end_split)
        start_end_list = []
        for hour in start_end_split:
            hour_min_split = hour.split(":")
            # print(hour_min_split)
            if(hour_min_split[1][-2:] == "am" or hour_min_split[1][-2:] == "AM"):
                hour_to_min = (int(hour_min_split[0])*60) + int(hour_min_split[1][:2])
            else:
                if(hour_min_split[0] == "12"):
                    hour_to_min = (int(hour_min_split[0])*60) + int(hour_min_split[1][:2])
                else:
                    hour_to_min = ((int(hour_min_split[0])+12)*60) + int(hour_min_split[1][:2])
            start_end_list.append(hour_to_min)
        # print("start time " + str(start_end_list[0]))
        # print("end time " + str(start_end_list[1]))
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
for index, row in sheet_to_df_map["ValidTimeSlots"].iterrows():
    if(isinstance(row[0], str)):
        date_index = 0
        for time in row[1:]:
            # print("date is " + str(date_index))
            teacher_name = row[0]
            if(type(time) != float):
                encoded_times = time_parse(time)
                encoded_days_times = [str(date_index) + etime for etime in encoded_times]
                if(teacher_name in teacher_to_time_map):
                    teacher_to_time_map[teacher_name].extend(encoded_days_times)
                else:
                    teacher_to_time_map[teacher_name] = encoded_days_times
            date_index+=1

# creating the time domains for the courses and encoding the course codes
total_classes = 0
course_variable = {}
course_variable_time_domain = {}
course_to_teacher_map = {}
teacher_to_course_map = {}
course_to_credit_map = {}
section_course_map = {}
course_set = set()
assigned_courses = sheet_to_df_map["AssignedCourses"]
for index, row in assigned_courses.iterrows():
    teacher_name = row[0]
    teacher_to_course_map[teacher_name] = []
    for name in row[1:]:
        course_list = sheet_to_df_map["UndergradCurriculumOptional"]
        if(type(name) != float):
            if(name not in course_set):
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
                if(encoded_code[1] == "1" or (int(encoded_code[1])>4)): # means the course is a lab course
                    # finding out the course teachers for this specific lab course
                    course_teacher_df = assigned_courses.loc[(assigned_courses["Course1"] == name) | (assigned_courses["Course2"] == name) |
                    (assigned_courses["Course3"] == name) | (assigned_courses["Course4"] == name) | (assigned_courses["Course5"] == name)]["Teacher"]
                    course_teacher_df_to_list = course_teacher_df.to_list()
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
                        list_for_set = teacher_to_time_map[course_teacher_df_to_list[0]].copy()
                if(("Section" or "section") in name): # finding out the highest section value for that lab course
                    if(course_name not in section_course_map):
                        highest_section_value = 0
                        for i in range(1, 5):
                            column_name = "Course" + str(i)
                            course_name_with_section = assigned_courses[assigned_courses[column_name].str.contains(course_name, na=False)][column_name]
                            if(not course_name_with_section.empty):
                                course_name_with_section_list = course_name_with_section.to_list()
                                for cname in course_name_with_section_list:
                                    highest_section_value = max(highest_section_value, int(cname[-1:]))
                        section_course_map[course_name] = highest_section_value
                    encoded_code += str(section_course_map[course_name]) + name[-1:] + '0'
                else:
                    encoded_code += "100"
                if(encoded_code[1] == "1" or (int(encoded_code[1])>4)):
                    list_for_set.sort()
                    course_variable_time_domain[encoded_code] = list_for_set.copy()
                    course_to_teacher_map[encoded_code] = course_teacher_df_to_list.copy()
                else:
                    teacher_to_time_map[teacher_name].sort()
                    course_variable_time_domain[encoded_code] = teacher_to_time_map[teacher_name].copy()
                    course_to_teacher_map[encoded_code] = [teacher_name]
                # if credit > 1.5, means there has to be 2 classes
                teacher_to_course_map[teacher_name].append(encoded_code)
                course_to_credit_map[encoded_code] = credit[0]
                total_classes += 1
                if(credit[0]>1.5):
                    encoded_code_2 = encoded_code[:-1] + "1"
                    if(encoded_code[1] == "1" or (int(encoded_code[1])>4)):
                        course_variable_time_domain[encoded_code_2] = list_for_set.copy()
                        course_to_teacher_map[encoded_code_2] = course_teacher_df_to_list.copy()
                    else:
                        course_variable_time_domain[encoded_code_2] = teacher_to_time_map[teacher_name].copy()
                        course_to_teacher_map[encoded_code_2] = [teacher_name]
                    teacher_to_course_map[teacher_name].append(encoded_code_2)
                    course_to_credit_map[encoded_code_2] = credit[0]
                    total_classes += 1
                course_set.add(name)

# function to remove time slots from course time domains
def remove_data(crs, booked_time_list):
    removed_time_list = []
    for item in booked_time_list:
        if(item in course_variable_time_domain[crs]):
            course_variable_time_domain[crs].remove(item)
            removed_time_list.append(item)
    return removed_time_list

# function to add time slots back to the course time domains
def fill_data(prunemap):
    for key, val in prunemap.items():
        if key in course_variable_time_domain:
            course_variable_time_domain[key].extend(prunemap[key].copy())
        else:
            course_variable_time_domain[key] = prunemap[key].copy()

# function to prune course time domains
def prune_data(current_course, booked_time_list):
    year = current_course[0]
    theo_or_lab = current_course[1]
    c_id = current_course[2]
    total_sec = current_course[3]
    sec_id = current_course[4]
    class_id = current_course[5]
    local_prunemap = {}
    # first remove the current course from dictionary
    local_prunemap[current_course] = course_variable_time_domain[current_course].copy()
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
        if(crs in course_variable_time_domain):
            local_prunemap[crs] = remove_data(crs, booked_time_list)
    # finding all the courses for current course year
    same_year_courses = [key for key, val in course_variable_time_domain.items() if year == key[0]]
    for crs in same_year_courses:
        if(crs not in local_prunemap):
            if(theo_or_lab == "0"): # mandatory theory. No other classes can run in parallel
                local_prunemap[crs] = remove_data(crs, booked_time_list)
            elif(theo_or_lab == "1"): # mandatory lab
                if(crs[1] == "0" or int(crs[1]) > 1): # mandatory theory or optional lab/theory can't run in parallel
                    local_prunemap[crs] = remove_data(crs, booked_time_list)
                elif(total_sec == "1"):
                    if(course_to_credit_map[current_course] == 0.75): # checking if the current course has 0.75 credit
                        if(course_to_credit_map[crs] != 0.75 and crs[3] != "2"): # all courses except 0.75 credit labs and mandatory labs with 2 sections can't run in parallel
                            local_prunemap[crs] = remove_data(crs, booked_time_list)
                    else: # 1.5 credit and only 1 lab section, other labs can't run in parallel
                        local_prunemap[crs] = remove_data(crs, booked_time_list)
                elif(total_sec != crs[3] or sec_id == crs[4]): # unequal total sections or same lab sections, can't run in parallel
                    local_prunemap[crs] = remove_data(crs, booked_time_list)
            else: # optional theory/labs for 4th year
                if(int(crs[1]) < 2): # mandatory theory/labs can't run in parallel
                    local_prunemap[crs] = remove_data(crs, booked_time_list)
                else:
                    if(int(c_id)%2 != int(crs[2])%2): # option 1 and option 2 lab/theory can't run in parallel
                        local_prunemap[crs] = remove_data(crs, booked_time_list)
                    elif(abs(int(theo_or_lab)-int(crs[1])) == 3): # corresponding opt. lab list for opt. theory classes
                        if(c_id == crs[2]): # corresponding lab for theory can't run in parallel
                            local_prunemap[crs] = remove_data(crs, booked_time_list)
    return local_prunemap

result_list = []
total = 10000
n = 0
# main backtracking function
def backtrack(n):
    if(len(course_variable) == total_classes):
        result_list.append(dict(course_variable))
        print(course_variable)
        n+=1
        return n
    second_class_flag = False
    # finding out the minimum length of lists of time in the domain
    res = sorted(course_variable_time_domain, key = lambda key: len(course_variable_time_domain[key]))
    # min_val = min([len(course_variable_time_domain[crs]) for crs in course_variable_time_domain])
    # min_len_dom = [key for key, val in test_dict.items() if len(val) == min_val]
    # for key, val in course_variable_time_domain.items():
    for i in range(len(res)):
        # if len(val) == min_len_dom:
        current_course = res[i]
        if(current_course[-1] == "1"):
            first_class_code = current_course[:-1] + "0"
            if(first_class_code in course_variable):
                first_class_day = course_variable[first_class_code][0]
                second_class_flag = True
            else:
                continue
        break
    if(not course_variable_time_domain[current_course]):
        return n
    course_credit = course_to_credit_map[current_course]
    loop_list = course_variable_time_domain[current_course].copy()
    # testing for every valid time in course-time domain
    for dtime in loop_list:
        # print(n)
        # print(total)
        if(n>=total):
            break
        day = dtime[0]
        if(second_class_flag):
            if(day <= first_class_day):
                continue
        time = int(dtime[1:])
        if(current_course[1] == "1" or int(current_course[1])>4):
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
        if(not book):
            continue
        course_variable[current_course] = dtime
        prunemap = prune_data(current_course, booked_time_list)
        n = backtrack(n)
        course_variable.pop(current_course)
        fill_data(prunemap)
    return n

p = backtrack(n)
print(len(result_list))
