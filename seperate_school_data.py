import os
import pandas as pd
import numpy as np


def student_total_sector_to_school(student_total, school_name, path):
    for school in school_name[:-1]:
        this_school = student_total.groupby("学校").get_group(school)
        school_path = path + "{}/".format(school)
        if not os.path.exists(school_path):
            os.makedirs(school_path)
        excel_name = "{}-学生信息.xlsx".format(school)
        with pd.ExcelWriter(school_path + excel_name, engine='xlsxwriter') as writer:
            this_school.to_excel(writer, index=False)


def subject_total_sector_to_school(subject_total, school_name, path, subject_name):
    for school in school_name[:-1]:
        this_school = subject_total.groupby('学校').get_group(school)
        school_path = path + "{}/".format(school)
        subject_total_path = school_path + "小题分/"
        if not os.path.exists(subject_total_path):
            os.makedirs(subject_total_path)
        excel_name = "小题分({}).xls".format(subject_name)
        with pd.ExcelWriter(subject_total_path + excel_name) as writer:
            this_school.to_excel(writer, index=False)