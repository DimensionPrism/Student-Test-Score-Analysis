import os
import pandas as pd
import numpy as np
import sector_analysis
import subject_anaylsis
import seperate_school_data
import school_subject_analysis
import school_sector_analysis

# 读取小题分
ITEM_SCORE = pd.read_excel("./考试数据/item_score.xlsx")

# 读取各科总分和小题满分
CHIN_TOTAL = pd.read_excel("./考试数据/小题分/小题分(语文).xls")
CHIN_PROB_TOTAL = pd.read_excel("./考试数据/小题满分.xlsx", sheet_name="语文")

MATH_TOTAL = pd.read_excel("./考试数据/小题分/小题分(数学).xls")
MATH_PROB_TOTAL = pd.read_excel("./考试数据/小题满分.xlsx", sheet_name="数学")

ENGL_TOTAL = pd.read_excel("./考试数据/小题分/小题分(英语).xls")
ENGL_PROB_TOTAL = pd.read_excel("./考试数据/小题满分.xlsx", sheet_name="英语")

HIST_TOTAL = pd.read_excel("./考试数据/小题分/小题分(历史).xls")
HIST_PROB_TOTAL = pd.read_excel("./考试数据/小题满分.xlsx", sheet_name="历史")

PHYS_TOTAL = pd.read_excel("./考试数据/小题分/小题分(物理).xls")
PHYS_PROB_TOTAL = pd.read_excel("./考试数据/小题满分.xlsx", sheet_name="物理")

CHEM_TOTAL = pd.read_excel("./考试数据/小题分/小题分(化学).xls")
CHEM_PROB_TOTAL = pd.read_excel("./考试数据/小题满分.xlsx", sheet_name="化学")

POLI_TOTAL = pd.read_excel("./考试数据/小题分/小题分(道德与法治).xls")
POLI_PROB_TOTAL = pd.read_excel("./考试数据/小题满分.xlsx", sheet_name="道德与法治")

# 读取学生总分
SCHOOL_NAME = []
STUDENT_TOTAL = pd.read_excel("./考试数据/报名信息-学生成绩.xlsx")
[SCHOOL_NAME.append(x) for x in STUDENT_TOTAL['学校'] if x not in SCHOOL_NAME]
SCHOOL_NAME.append("全体")

SUBJECT_LIST = ["语文", "数学", "英语", "历史", "物理", "化学", "道德与法治", "总分"]
SUBJECT_FULL_SCORE_LIST = [120, 100, 100, 100, 100, 100, 100, 720]

SUBJECT_TOTAL_LIST = [
    CHIN_TOTAL,
    MATH_TOTAL,
    ENGL_TOTAL,
    HIST_TOTAL,
    PHYS_TOTAL,
    CHEM_TOTAL,
    POLI_TOTAL,
]

SUBJECT_PROB_TOTAL = [
    CHIN_PROB_TOTAL,
    MATH_PROB_TOTAL,
    ENGL_PROB_TOTAL,
    HIST_PROB_TOTAL,
    PHYS_PROB_TOTAL,
    CHEM_PROB_TOTAL,
    POLI_PROB_TOTAL,
]

SUBJECT_TOTAL_SCORE = pd.concat([
    STUDENT_TOTAL[['学校', '分数']],
    STUDENT_TOTAL['分数.1'],
    STUDENT_TOTAL['分数.2'],
    STUDENT_TOTAL['分数.3'],
    STUDENT_TOTAL['分数.4'],
    STUDENT_TOTAL['分数.5'],
    STUDENT_TOTAL['分数.6'],
    STUDENT_TOTAL['分数.7']], axis=1)


def subject(path):
    print("学科信息计算中......\n")
    for x in range(0, 7):
        subject_name = SUBJECT_LIST[x]
        subject_full_score = SUBJECT_FULL_SCORE_LIST[x]
        subject_total = SUBJECT_TOTAL_LIST[x]
        subject_prob_total = SUBJECT_PROB_TOTAL[x]

        this_subject = subject_anaylsis.Analysis(subject_name, SCHOOL_NAME, subject_full_score,
                                                 subject_prob_total, subject_total, ITEM_SCORE,
                                                 STUDENT_TOTAL, SUBJECT_TOTAL_SCORE, x)

        problem_analysis = this_subject.problem_analysis()

        print(subject_name + "科目报表计算完成")

        this_subject_export_path = path + "{}/".format(subject_name)
        if not os.path.exists(this_subject_export_path):
            os.makedirs(this_subject_export_path)
        for each in problem_analysis.keys():
            excel_name = "{}.xlsx".format(each)
            this_excel = problem_analysis[each]
            with pd.ExcelWriter(this_subject_export_path + excel_name, engine='xlsxwriter') as writer:
                if type(this_excel) == dict:
                    for content in this_excel.keys():
                        this_excel[content].to_excel(writer, index=False, sheet_name=content)
                        if each == "2-1-3-2试题难度分布":
                            if content == "难度分布":
                                workbook = writer.book
                                worksheet = writer.sheets[content]

                                chart = workbook.add_chart({'type': 'column'})
                                max_row = len(this_excel[content])
                                col_x = this_excel[content].columns.get_loc('难度范围')
                                col_y = this_excel[content].columns.get_loc("比例")

                                chart.add_series({
                                    'name': '总分占比',
                                    'categories': [content, 1, col_x, max_row, col_x],
                                    'values': [content, 1, col_y, max_row, col_y],
                                    'data_labels': {'value': True},
                                })

                                chart.set_title({'name': "难度比例图"})
                                chart.set_x_axis({'name': '难度范围', 'text_axis': True})
                                chart.set_y_axis({'name': '比例'})
                                chart.set_size({'width': 720, 'height': 576})
                                worksheet.insert_chart('A8', chart)

                        if each == "2-1-3-3试题区分度分布":
                            if content == "区分度分布":
                                workbook = writer.book
                                worksheet = writer.sheets[content]

                                chart = workbook.add_chart({'type': 'column'})
                                max_row = len(this_excel[content])
                                col_x = this_excel[content].columns.get_loc('区分度范围')
                                col_y = this_excel[content].columns.get_loc("比例")

                                chart.add_series({
                                    'name': '总分占比',
                                    'categories': [content, 1, col_x, max_row, col_x],
                                    'values': [content, 1, col_y, max_row, col_y],
                                    'data_labels': {'value': True},
                                })

                                chart.set_title({'name': "区分度比例图"})
                                chart.set_x_axis({'name': '区分度范围', 'text_axis': True})
                                chart.set_y_axis({'name': '比例'})
                                chart.set_size({'width': 720, 'height': 576})
                                worksheet.insert_chart('A8', chart)

                        if each == "2-1-3-9全卷及各题难度曲线图":
                            workbook = writer.book
                            worksheet = writer.sheets[content]
                            chart = workbook.add_chart({'type': 'line'})
                            max_row = len(this_excel[content])
                            col_x = this_excel[content].columns.get_loc('分数段')
                            col_y = this_excel[content].columns.get_loc('难度值')

                            chart.add_series({
                                'name': '难度值',
                                'categories': [content, 1, col_x, max_row, col_x],
                                'values': [content, 1, col_y, max_row, col_y],
                                'marker': {
                                    'type': 'square',
                                    'size': 4,
                                    'border': {'color': 'black'},
                                    'fill': {'color': 'red'},
                                },
                                'data_labels': {'value': True, 'position': 'left'},
                            })
                            if content == "全卷":
                                chart.set_title({'name': content + '难度曲线'})
                            else:
                                chart.set_title({'name': '第' + str(content) + '题难度曲线'})
                            chart.set_x_axis({'name': '分数段'})
                            chart.set_y_axis({'name': '难度值', 'min': 0, 'max': 1})
                            chart.set_size({'width': 720, 'height': 576})
                            worksheet.insert_chart('D2', chart)
                else:
                    problem_analysis[each].to_excel(writer, index=False, sheet_name=each)
                    if each == "2-1-3-1各题目指标":
                        workbook = writer.book
                        worksheet = writer.sheets[each]
                        chart = workbook.add_chart({'type': 'scatter'})
                        max_row = len(problem_analysis[each])
                        col_x = problem_analysis[each].columns.get_loc('难度')
                        col_y = problem_analysis[each].columns.get_loc('区分度')

                        prob_list = problem_analysis[each]["题号"]
                        custom_labels = []
                        for prob in prob_list:
                            custom_labels.append({'value': prob})

                        chart.add_series({
                            'name': '难度，区分度',
                            'categories': [each, 1, col_x, max_row, col_x],
                            'values': [each, 1, col_y, max_row, col_y],
                            'marker': {'type': 'circle', 'size': 4, 'fill': {'color': 'orange'}},
                            'data_labels': {'value': True, 'custom': custom_labels},
                        })

                        chart.set_title({'name': '题目难度与区分度分布图'})
                        chart.set_x_axis({'name': '难度', 'min': 0, 'max': 1})
                        chart.set_y_axis({'name': '区分度', 'min': -1, 'max': 1})
                        chart.set_size({'width': 920, 'height': 860})
                        worksheet.insert_chart('A30', chart)

                    if each == "2-1-4-1各学校得分情况":
                        problem_analysis[each].to_excel(writer, index=False, sheet_name=each)
                        workbook = writer.book
                        worksheet = writer.sheets[each]

                        # 平均分
                        chart = workbook.add_chart({'type': 'column'})
                        max_row = len(problem_analysis[each])
                        col_x = problem_analysis[each].columns.get_loc('学校')
                        col_y = problem_analysis[each].columns.get_loc('全体平均分')

                        chart.add_series({
                            'name': '平均分',
                            'categories': [each, 1, col_x, max_row, col_x],
                            'values': [each, 1, col_y, max_row, col_y],
                            'data_labels': {'value': True},
                        })

                        chart.set_title({'name': "各学校平均分分布图"})
                        chart.set_x_axis({'name': '学校', 'text_axis': True})
                        chart.set_y_axis({'name': '平均分'})
                        chart.set_size({'width': 720, 'height': 560})
                        worksheet.insert_chart('A20', chart)

                        # 得分率
                        chart = workbook.add_chart({'type': 'column'})
                        max_row = len(problem_analysis[each])
                        col_x = problem_analysis[each].columns.get_loc('学校')
                        col_y = problem_analysis[each].columns.get_loc('得分率')

                        chart.add_series({
                            'name': '得分率',
                            'categories': [each, 1, col_x, max_row, col_x],
                            'values': [each, 1, col_y, max_row, col_y],
                            'data_labels': {'value': True},
                        })

                        chart.set_title({'name': "各学校得分率分布图"})
                        chart.set_x_axis({'name': '学校', 'text_axis': True})
                        chart.set_y_axis({'name': '得分率'})
                        chart.set_size({'width': 720, 'height': 560})
                        worksheet.insert_chart('M20', chart)

                        # 超优率
                        chart = workbook.add_chart({'type': 'column'})
                        max_row = len(problem_analysis[each])
                        col_x = problem_analysis[each].columns.get_loc('学校')
                        col_y = problem_analysis[each].columns.get_loc('超优率')

                        chart.add_series({
                            'name': '超优率',
                            'categories': [each, 1, col_x, max_row, col_x],
                            'values': [each, 1, col_y, max_row, col_y],
                            'data_labels': {'value': True},
                        })

                        chart.set_title({'name': "各学校超优率分布图"})
                        chart.set_x_axis({'name': '学校', 'text_axis': True})
                        chart.set_y_axis({'name': '超优率'})
                        chart.set_size({'width': 720, 'height': 560})
                        worksheet.insert_chart('A50', chart)

                        # 优秀率
                        chart = workbook.add_chart({'type': 'column'})
                        max_row = len(problem_analysis[each])
                        col_x = problem_analysis[each].columns.get_loc('学校')
                        col_y = problem_analysis[each].columns.get_loc('超优率')

                        chart.add_series({
                            'name': '优秀率',
                            'categories': [each, 1, col_x, max_row, col_x],
                            'values': [each, 1, col_y, max_row, col_y],
                            'data_labels': {'value': True},
                        })

                        chart.set_title({'name': "各学校优秀率分布图"})
                        chart.set_x_axis({'name': '学校', 'text_axis': True})
                        chart.set_y_axis({'name': '优秀率'})
                        chart.set_size({'width': 720, 'height': 560})
                        worksheet.insert_chart('M50', chart)

                        # 及格率
                        chart = workbook.add_chart({'type': 'column'})
                        max_row = len(problem_analysis[each])
                        col_x = problem_analysis[each].columns.get_loc('学校')
                        col_y = problem_analysis[each].columns.get_loc('及格率')

                        chart.add_series({
                            'name': '及格率',
                            'categories': [each, 1, col_x, max_row, col_x],
                            'values': [each, 1, col_y, max_row, col_y],
                            'data_labels': {'value': True},
                        })

                        chart.set_title({'name': "各学校及格率分布图"})
                        chart.set_x_axis({'name': '学校', 'text_axis': True})
                        chart.set_y_axis({'name': '及格率'})
                        chart.set_size({'width': 720, 'height': 560})
                        worksheet.insert_chart('A80', chart)

                        # 低分率
                        chart = workbook.add_chart({'type': 'column'})
                        max_row = len(problem_analysis[each])
                        col_x = problem_analysis[each].columns.get_loc('学校')
                        col_y = problem_analysis[each].columns.get_loc('低分率')

                        chart.add_series({
                            'name': '低分率',
                            'categories': [each, 1, col_x, max_row, col_x],
                            'values': [each, 1, col_y, max_row, col_y],
                            'data_labels': {'value': True},
                        })

                        chart.set_title({'name': "各学校低分率分布图"})
                        chart.set_x_axis({'name': '学校', 'text_axis': True})
                        chart.set_y_axis({'name': '低分率'})
                        chart.set_size({'width': 720, 'height': 560})
                        worksheet.insert_chart('M80', chart)

    print("学科报表输出完成\n")


def sector(path):
    print("全区信息计算中......\n")
    sector = sector_analysis.Analysis(SCHOOL_NAME, SUBJECT_LIST, SUBJECT_FULL_SCORE_LIST,
                                      SUBJECT_TOTAL_LIST, SUBJECT_TOTAL_SCORE)

    sector_analysis_output = sector.sector_output()
    print("全区报表计算完成")
    if not os.path.exists(path):
        os.makedirs(path)
    for each in sector_analysis_output.keys():
        excel_name = "{}.xlsx".format(each)
        this_excel = sector_analysis_output[each]
        if each == "2-2-2总分及各科上线的有效分数线":
            effective_score_analysis = sector.subject_effective_score_analysis(this_excel)
            for subject in effective_score_analysis.keys():
                score_line_path = "./OUTPUT/学科/{}/".format(subject)
                subject_dict = effective_score_analysis[subject]
                for goals in subject_dict.keys():
                    with pd.ExcelWriter(score_line_path + goals, engine='xlsxwriter') as writer:
                        subject_dict[goals].to_excel(writer, index=False, sheet_name=goals)
                        workbook = writer.book
                        worksheet = writer.sheets[goals]

                        # 总分上线
                        chart = workbook.add_chart({'type': 'column'})
                        max_row = len(subject_dict[goals]) - 1
                        col_x = subject_dict[goals].columns.get_loc('学校')
                        col_y = subject_dict[goals].columns.get_loc("总分上线人数")

                        chart.add_series({
                            'name': '总分上线人数',
                            'categories': [goals, 1, col_x, max_row, col_x],
                            'values': [goals, 1, col_y, max_row, col_y],
                            'data_labels': {'value': True},
                        })

                        chart.set_title({'name': "总分上线各校分布图"})
                        chart.set_x_axis({'name': '学校', 'text_axis': True})
                        chart.set_y_axis({'name': '总分上线人数'})
                        chart.set_size({'width': 720, 'height': 560})
                        worksheet.insert_chart('G2', chart)

                        # 单上线
                        chart = workbook.add_chart({'type': 'column'})
                        max_row = len(subject_dict[goals]) - 1
                        col_x = subject_dict[goals].columns.get_loc('学校')
                        col_y = subject_dict[goals].columns.get_loc(subject+'单上线')

                        chart.add_series({
                            'name': '单上线人数',
                            'categories': [goals, 1, col_x, max_row, col_x],
                            'values': [goals, 1, col_y, max_row, col_y],
                            'data_labels': {'value': True},
                        })

                        chart.set_title({'name': subject + "单科上线各校分布图"})
                        chart.set_x_axis({'name': '学校', 'text_axis': True})
                        chart.set_y_axis({'name': '单上线人数'})
                        chart.set_size({'width': 720, 'height': 560})
                        worksheet.insert_chart('A30', chart)

                        # 双上线
                        chart = workbook.add_chart({'type': 'column'})
                        max_row = len(subject_dict[goals]) - 1
                        col_x = subject_dict[goals].columns.get_loc('学校')
                        col_y = subject_dict[goals].columns.get_loc(subject+'双上线')

                        chart.add_series({
                            'name': '双上线人数',
                            'categories': [goals, 1, col_x, max_row, col_x],
                            'values': [goals, 1, col_y, max_row, col_y],
                            'data_labels': {'value': True},
                        })

                        chart.set_title({'name': subject + "双上线各校分布图"})
                        chart.set_x_axis({'name': '学校', 'text_axis': True})
                        chart.set_y_axis({'name': '双上线人数'})
                        chart.set_size({'width': 720, 'height': 560})
                        worksheet.insert_chart('M30', chart)

                        # M1
                        chart = workbook.add_chart({'type': 'column'})
                        max_row = len(subject_dict[goals]) - 1
                        col_x = subject_dict[goals].columns.get_loc('学校')
                        col_y = subject_dict[goals].columns.get_loc(subject + 'M1')

                        chart.add_series({
                            'name': 'M1人数',
                            'categories': [goals, 1, col_x, max_row, col_x],
                            'values': [goals, 1, col_y, max_row, col_y],
                            'data_labels': {'value': True},
                        })

                        chart.set_title({'name': subject + "M1各校分布图"})
                        chart.set_x_axis({'name': '学校', 'text_axis': True})
                        chart.set_y_axis({'name': 'M1人数'})
                        chart.set_size({'width': 720, 'height': 560})
                        worksheet.insert_chart('A60', chart)

                        # M2
                        chart = workbook.add_chart({'type': 'column'})
                        max_row = len(subject_dict[goals]) - 1
                        col_x = subject_dict[goals].columns.get_loc('学校')
                        col_y = subject_dict[goals].columns.get_loc(subject + 'M2')

                        chart.add_series({
                            'name': 'M2人数',
                            'categories': [goals, 1, col_x, max_row, col_x],
                            'values': [goals, 1, col_y, max_row, col_y],
                            'data_labels': {'value': True},
                        })

                        chart.set_title({'name': subject + "M2各校分布图"})
                        chart.set_x_axis({'name': '学校', 'text_axis': True})
                        chart.set_y_axis({'name': 'M2人数'})
                        chart.set_size({'width': 720, 'height': 560})
                        worksheet.insert_chart('A60', chart)

        with pd.ExcelWriter(path + excel_name, engine='xlsxwriter') as writer:
            if type(this_excel) == dict:
                for content in this_excel.keys():
                    this_excel[content].to_excel(writer, index=False, sheet_name=content)
            else:
                sector_analysis_output[each].to_excel(writer, index=False, sheet_name=each)

    print("全区报表输出完成\n")


def school_output(path):
    print("各学校学科信息计算中......\n")
    for school in SCHOOL_NAME[:-1]:
        read_path = "./考试数据/学校数据/{}/".format(school)

        this_school_export_path = path + "{}/".format(school)

        item_score = pd.read_excel(read_path + "item_score.xlsx")

        school_student_total = pd.read_excel(read_path + "{}-学生信息.xlsx".format(school))

        school_subject_total_score = pd.concat([
            school_student_total[['班级', '分数']],
            school_student_total['分数.1'],
            school_student_total['分数.2'],
            school_student_total['分数.3'],
            school_student_total['分数.4'],
            school_student_total['分数.5'],
            school_student_total['分数.6'],
            school_student_total['分数.7']], axis=1)

        Class_name = []
        [Class_name.append(x) for x in school_student_total['班级'] if x not in Class_name]
        Class_name.append("全校")

        print("当前学校：" + school)
        school_subject_total_list = []
        for index in range(7):
            subject_name = SUBJECT_LIST[index]
            subject_full_score = SUBJECT_FULL_SCORE_LIST[index]
            subject_total = pd.read_excel(read_path + "/小题分/小题分({}).xls".format(subject_name))
            school_subject_total_list.append(subject_total)
            subject_prob_total = SUBJECT_PROB_TOTAL[index]
            this_subject = school_subject_analysis.Analysis(subject_name, Class_name, subject_full_score,
                                                            subject_prob_total, subject_total, item_score,
                                                            school_student_total, school_subject_total_score, index)
            problem_analysis = this_subject.problem_analysis()

            print(subject_name + "科目报表计算完成")

            this_subject_export_path = this_school_export_path + "学科/{}/".format(subject_name)
            if not os.path.exists(this_subject_export_path):
                os.makedirs(this_subject_export_path)
            for each in problem_analysis.keys():
                excel_name = "{}.xlsx".format(each)
                this_excel = problem_analysis[each]
                with pd.ExcelWriter(this_subject_export_path + excel_name, engine='xlsxwriter') as writer:
                    if type(this_excel) == dict:
                        for content in this_excel.keys():
                            this_excel[content].to_excel(writer, index=False, sheet_name=content)
                            if each == "2-1-3-2试题难度分布":
                                if content == "难度分布":
                                    workbook = writer.book
                                    worksheet = writer.sheets[content]

                                    chart = workbook.add_chart({'type': 'column'})
                                    max_row = len(this_excel[content])
                                    col_x = this_excel[content].columns.get_loc('难度范围')
                                    col_y = this_excel[content].columns.get_loc("比例")

                                    chart.add_series({
                                        'name': '总分占比',
                                        'categories': [content, 1, col_x, max_row, col_x],
                                        'values': [content, 1, col_y, max_row, col_y],
                                        'data_labels': {'value': True},
                                    })

                                    chart.set_title({'name': "难度比例图"})
                                    chart.set_x_axis({'name': '难度范围', 'text_axis': True})
                                    chart.set_y_axis({'name': '比例'})
                                    chart.set_size({'width': 720, 'height': 576})
                                    worksheet.insert_chart('A8', chart)

                            if each == "2-1-3-3试题区分度分布":
                                if content == "区分度分布":
                                    workbook = writer.book
                                    worksheet = writer.sheets[content]

                                    chart = workbook.add_chart({'type': 'column'})
                                    max_row = len(this_excel[content])
                                    col_x = this_excel[content].columns.get_loc('区分度范围')
                                    col_y = this_excel[content].columns.get_loc("比例")

                                    chart.add_series({
                                        'name': '总分占比',
                                        'categories': [content, 1, col_x, max_row, col_x],
                                        'values': [content, 1, col_y, max_row, col_y],
                                        'data_labels': {'value': True},
                                    })

                                    chart.set_title({'name': "区分度比例图"})
                                    chart.set_x_axis({'name': '区分度范围', 'text_axis': True})
                                    chart.set_y_axis({'name': '比例'})
                                    chart.set_size({'width': 720, 'height': 576})
                                    worksheet.insert_chart('A8', chart)

                            if each == "2-1-3-9全卷及各题难度曲线图":
                                workbook = writer.book
                                worksheet = writer.sheets[content]
                                chart = workbook.add_chart({'type': 'line'})
                                max_row = len(this_excel[content])
                                col_x = this_excel[content].columns.get_loc('分数段')
                                col_y = this_excel[content].columns.get_loc('难度值')

                                chart.add_series({
                                    'name': '难度值',
                                    'categories': [content, 1, col_x, max_row, col_x],
                                    'values': [content, 1, col_y, max_row, col_y],
                                    'marker': {
                                        'type': 'square',
                                        'size': 4,
                                        'border': {'color': 'black'},
                                        'fill': {'color': 'red'},
                                    },
                                    'data_labels': {'value': True, 'position': 'left'},
                                })
                                if content == "全卷":
                                    chart.set_title({'name': content + '难度曲线'})
                                else:
                                    chart.set_title({'name': '第' + str(content) + '题难度曲线'})
                                chart.set_x_axis({'name': '分数段'})
                                chart.set_y_axis({'name': '难度值', 'min': 0, 'max': 1})
                                chart.set_size({'width': 720, 'height': 576})
                                worksheet.insert_chart('D2', chart)
                    else:
                        problem_analysis[each].to_excel(writer, index=False, sheet_name=each)
                        if each == "2-1-3-1各题目指标":
                            workbook = writer.book
                            worksheet = writer.sheets[each]
                            chart = workbook.add_chart({'type': 'scatter'})
                            max_row = len(problem_analysis[each])
                            col_x = problem_analysis[each].columns.get_loc('难度')
                            col_y = problem_analysis[each].columns.get_loc('区分度')

                            prob_list = problem_analysis[each]["题号"]
                            custom_labels = []
                            for prob in prob_list:
                                custom_labels.append({'value': prob})

                            chart.add_series({
                                'name': '题号',
                                'categories': [each, 1, col_x, max_row, col_x],
                                'values': [each, 1, col_y, max_row, col_y],
                                'marker': {'type': 'circle', 'size': 4, 'fill': {'color': 'orange'}},
                                'data_labels': {'value': True, 'custom': custom_labels},
                            })

                            chart.set_title({'name': '题目难度与区分度分布图'})
                            chart.set_x_axis({'name': '难度', 'min': 0, 'max': 1})
                            chart.set_y_axis({'name': '区分度', 'min': -1, 'max': 1})
                            chart.set_size({'width': 920, 'height': 860})
                            worksheet.insert_chart('A30', chart)

                        if each == "2-1-4-1各班级得分情况":
                            problem_analysis[each].to_excel(writer, index=False, sheet_name=each)
                            workbook = writer.book
                            worksheet = writer.sheets[each]

                            # 平均分
                            chart = workbook.add_chart({'type': 'column'})
                            max_row = len(problem_analysis[each])
                            col_x = problem_analysis[each].columns.get_loc('班级')
                            col_y = problem_analysis[each].columns.get_loc('全体平均分')

                            chart.add_series({
                                'name': '平均分',
                                'categories': [each, 1, col_x, max_row, col_x],
                                'values': [each, 1, col_y, max_row, col_y],
                                'data_labels': {'value': True},
                            })

                            chart.set_title({'name': "各班级平均分分布图"})
                            chart.set_x_axis({'name': '班级', 'text_axis': True})
                            chart.set_y_axis({'name': '平均分'})
                            chart.set_size({'width': 720, 'height': 560})
                            worksheet.insert_chart('A20', chart)

                            # 得分率
                            chart = workbook.add_chart({'type': 'column'})
                            max_row = len(problem_analysis[each])
                            col_x = problem_analysis[each].columns.get_loc('班级')
                            col_y = problem_analysis[each].columns.get_loc('得分率')

                            chart.add_series({
                                'name': '得分率',
                                'categories': [each, 1, col_x, max_row, col_x],
                                'values': [each, 1, col_y, max_row, col_y],
                                'data_labels': {'value': True},
                            })

                            chart.set_title({'name': "各班级得分率分布图"})
                            chart.set_x_axis({'name': '班级', 'text_axis': True})
                            chart.set_y_axis({'name': '得分率'})
                            chart.set_size({'width': 720, 'height': 560})
                            worksheet.insert_chart('M20', chart)

                            # 超优率
                            chart = workbook.add_chart({'type': 'column'})
                            max_row = len(problem_analysis[each])
                            col_x = problem_analysis[each].columns.get_loc('班级')
                            col_y = problem_analysis[each].columns.get_loc('超优率')

                            chart.add_series({
                                'name': '超优率',
                                'categories': [each, 1, col_x, max_row, col_x],
                                'values': [each, 1, col_y, max_row, col_y],
                                'data_labels': {'value': True},
                            })

                            chart.set_title({'name': "各班级超优率分布图"})
                            chart.set_x_axis({'name': '班级', 'text_axis': True})
                            chart.set_y_axis({'name': '超优率'})
                            chart.set_size({'width': 720, 'height': 560})
                            worksheet.insert_chart('A50', chart)

                            # 优秀率
                            chart = workbook.add_chart({'type': 'column'})
                            max_row = len(problem_analysis[each])
                            col_x = problem_analysis[each].columns.get_loc('班级')
                            col_y = problem_analysis[each].columns.get_loc('超优率')

                            chart.add_series({
                                'name': '优秀率',
                                'categories': [each, 1, col_x, max_row, col_x],
                                'values': [each, 1, col_y, max_row, col_y],
                                'data_labels': {'value': True},
                            })

                            chart.set_title({'name': "各班级优秀率分布图"})
                            chart.set_x_axis({'name': '班级', 'text_axis': True})
                            chart.set_y_axis({'name': '优秀率'})
                            chart.set_size({'width': 720, 'height': 560})
                            worksheet.insert_chart('M50', chart)

                            # 及格率
                            chart = workbook.add_chart({'type': 'column'})
                            max_row = len(problem_analysis[each])
                            col_x = problem_analysis[each].columns.get_loc('班级')
                            col_y = problem_analysis[each].columns.get_loc('及格率')

                            chart.add_series({
                                'name': '及格率',
                                'categories': [each, 1, col_x, max_row, col_x],
                                'values': [each, 1, col_y, max_row, col_y],
                                'data_labels': {'value': True},
                            })

                            chart.set_title({'name': "各班级及格率分布图"})
                            chart.set_x_axis({'name': '班级', 'text_axis': True})
                            chart.set_y_axis({'name': '及格率'})
                            chart.set_size({'width': 720, 'height': 560})
                            worksheet.insert_chart('A80', chart)

                            # 低分率
                            chart = workbook.add_chart({'type': 'column'})
                            max_row = len(problem_analysis[each])
                            col_x = problem_analysis[each].columns.get_loc('班级')
                            col_y = problem_analysis[each].columns.get_loc('低分率')

                            chart.add_series({
                                'name': '低分率',
                                'categories': [each, 1, col_x, max_row, col_x],
                                'values': [each, 1, col_y, max_row, col_y],
                                'data_labels': {'value': True},
                            })

                            chart.set_title({'name': "各班级低分率分布图"})
                            chart.set_x_axis({'name': '班级', 'text_axis': True})
                            chart.set_y_axis({'name': '低分率'})
                            chart.set_size({'width': 720, 'height': 560})
                            worksheet.insert_chart('M80', chart)

        print("\n{}学科报表输出完成\n".format(school))

        print(school + "整体报表计算中")
        sector = school_sector_analysis.Analysis(Class_name, SUBJECT_LIST, SUBJECT_FULL_SCORE_LIST,
                                                 school_subject_total_list, school_subject_total_score)

        sector_analysis_output = sector.sector_output()
        print(school + "整体报表计算完成\n")
        if not os.path.exists(path):
            os.makedirs(path)
        for each in sector_analysis_output.keys():
            excel_name = "{}.xlsx".format(each)
            this_excel = sector_analysis_output[each]
            if each == "2-2-2总分及各科上线的有效分数线":
                effective_score_analysis = sector.subject_effective_score_analysis(this_excel)
                for subject in effective_score_analysis.keys():
                    score_line_path = "./OUTPUT/学校/{}/学科/{}/".format(school, subject)
                    subject_dict = effective_score_analysis[subject]
                    for goals in subject_dict.keys():
                        with pd.ExcelWriter(score_line_path + goals, engine='xlsxwriter') as writer:
                            subject_dict[goals].to_excel(writer, index=False, sheet_name=goals)
                            workbook = writer.book
                            worksheet = writer.sheets[goals]

                            # 总分上线
                            chart = workbook.add_chart({'type': 'column'})
                            max_row = len(subject_dict[goals]) - 1
                            col_x = subject_dict[goals].columns.get_loc('班级')
                            col_y = subject_dict[goals].columns.get_loc("总分上线人数")

                            chart.add_series({
                                'name': '总分上线人数',
                                'categories': [goals, 1, col_x, max_row, col_x],
                                'values': [goals, 1, col_y, max_row, col_y],
                                'data_labels': {'value': True},
                            })

                            chart.set_title({'name': "总分上线各班分布图"})
                            chart.set_x_axis({'name': '班级', 'text_axis': True})
                            chart.set_y_axis({'name': '总分上线人数'})
                            chart.set_size({'width': 720, 'height': 560})
                            worksheet.insert_chart('G2', chart)

                            # 单上线
                            chart = workbook.add_chart({'type': 'column'})
                            max_row = len(subject_dict[goals]) - 1
                            col_x = subject_dict[goals].columns.get_loc('班级')
                            col_y = subject_dict[goals].columns.get_loc(subject + '单上线')

                            chart.add_series({
                                'name': '单上线人数',
                                'categories': [goals, 1, col_x, max_row, col_x],
                                'values': [goals, 1, col_y, max_row, col_y],
                                'data_labels': {'value': True},
                            })

                            chart.set_title({'name': subject + "单科上线各班分布图"})
                            chart.set_x_axis({'name': '班级', 'text_axis': True})
                            chart.set_y_axis({'name': '单上线人数'})
                            chart.set_size({'width': 720, 'height': 560})
                            worksheet.insert_chart('A30', chart)

                            # 双上线
                            chart = workbook.add_chart({'type': 'column'})
                            max_row = len(subject_dict[goals]) - 1
                            col_x = subject_dict[goals].columns.get_loc('班级')
                            col_y = subject_dict[goals].columns.get_loc(subject + '双上线')

                            chart.add_series({
                                'name': '双上线人数',
                                'categories': [goals, 1, col_x, max_row, col_x],
                                'values': [goals, 1, col_y, max_row, col_y],
                                'data_labels': {'value': True},
                            })

                            chart.set_title({'name': subject + "双上线各班分布图"})
                            chart.set_x_axis({'name': '班级', 'text_axis': True})
                            chart.set_y_axis({'name': '双上线人数'})
                            chart.set_size({'width': 720, 'height': 560})
                            worksheet.insert_chart('M30', chart)

                            # M1
                            chart = workbook.add_chart({'type': 'column'})
                            max_row = len(subject_dict[goals]) - 1
                            col_x = subject_dict[goals].columns.get_loc('班级')
                            col_y = subject_dict[goals].columns.get_loc(subject + 'M1')

                            chart.add_series({
                                'name': 'M1人数',
                                'categories': [goals, 1, col_x, max_row, col_x],
                                'values': [goals, 1, col_y, max_row, col_y],
                                'data_labels': {'value': True},
                            })

                            chart.set_title({'name': subject + "M2各班分布图"})
                            chart.set_x_axis({'name': '班级', 'text_axis': True})
                            chart.set_y_axis({'name': 'M2人数'})
                            chart.set_size({'width': 720, 'height': 560})
                            worksheet.insert_chart('A60', chart)

                            # M2
                            chart = workbook.add_chart({'type': 'column'})
                            max_row = len(subject_dict[goals]) - 1
                            col_x = subject_dict[goals].columns.get_loc('班级')
                            col_y = subject_dict[goals].columns.get_loc(subject + 'M2')

                            chart.add_series({
                                'name': 'M2人数',
                                'categories': [goals, 1, col_x, max_row, col_x],
                                'values': [goals, 1, col_y, max_row, col_y],
                                'data_labels': {'value': True},
                            })

                            chart.set_title({'name': subject + "M2各班分布图"})
                            chart.set_x_axis({'name': '班级', 'text_axis': True})
                            chart.set_y_axis({'name': 'M2人数'})
                            chart.set_size({'width': 720, 'height': 560})
                            worksheet.insert_chart('M60', chart)

            with pd.ExcelWriter(path +"{}/".format(school) + excel_name, engine='xlsxwriter') as writer:
                if type(this_excel) == dict:
                    for content in this_excel.keys():
                        this_excel[content].to_excel(writer, index=False, sheet_name=content)
                else:
                    sector_analysis_output[each].to_excel(writer, index=False, sheet_name=each)

    print("各学校报表输出完成\n")


if __name__ == '__main__':
    export_path = "./OUTPUT/"
    subject_export_path = export_path + "学科/"
    school_export_path = export_path + "学校/"
    sector_export_path = export_path + "全区/"
    school_data_path = "./考试数据/学校数据/"

    # 计算并输出学科报表
    subject(subject_export_path)

    # 计算并输出全区报表
    sector(sector_export_path)

    # 按学校拆分各表
    print("按学校拆分各表中......")
    print("学校名单:" + ",".join(SCHOOL_NAME))
    seperate_school_data.student_total_sector_to_school(STUDENT_TOTAL, SCHOOL_NAME, school_data_path)
    for index in range(7):
        subject_total = SUBJECT_TOTAL_LIST[index]
        subject_name = SUBJECT_LIST[index]
        seperate_school_data.subject_total_sector_to_school(subject_total, SCHOOL_NAME, school_data_path, subject_name)
    print("拆分完成\n")

    # 计算并输出学校报表
    school_output(school_export_path)
