# import xlsxwriter module
import xlsxwriter

workbook = xlsxwriter.Workbook('Example2.xlsx')
worksheet = workbook.add_worksheet()

# Start from the first cell.
# Rows and columns are zero indexed.
row = 0
column = 0

# # iterating through content list
# for item in content:
#     # write operation perform
#     worksheet.write(row, column, item)
#
#     # incrementing the value of row by one
#     # with each iterations.
#     row += 1


filename = 'ENROLLMENT.csv'

import pandas as pd
import numpy as np

df = pd.read_csv(filename, sep=",")

df.head()

courses = df["COURSECODE"].unique()
courses = np.sort(courses)

course_matrix = {}
for i in range(len(courses)):
    worksheet.write(i + 1, 0, courses[i])
    worksheet.write(0, i + 1, courses[i])

    print(i, "-----------------------------------")
    row_course = courses[i]
    filtered_row_course = df.query('COURSECODE == "' + row_course + '"')
    row_users = filtered_row_course['USERID'].unique()
    for j in range(len(courses)):
        col_course = courses[j]
        dict_key = row_course + "_" + col_course
        if dict_key in course_matrix:
            worksheet.write(i + 1, j + 1, course_matrix[dict_key])
            continue
        filtered_col_course = df.query('COURSECODE == "' + col_course + '"')
        col_users = filtered_col_course['USERID'].unique()
        common_elements = len([x for x in row_users if x in col_users])
        course_matrix[row_course + "_" + col_course] = common_elements
        course_matrix[col_course + "_" + row_course] = common_elements
        worksheet.write(i + 1, j + 1, common_elements)

    print(row_course, " işlendi")

workbook.close()
