import os
import pandas as pd
import numpy as np

import overall_analysis
import subject_analysis
import seperate_school_data
import school_subject_analysis
import school_overall_analysis

# 读取小题分
ITEM_SCORE = pd.read_excel("./考试数据/item_score.xlsx").fillna(0)

# 读取学生总分
SCHOOL_NAME = []
STUDENT_TOTAL = pd.read_excel("./考试数据/报名信息-学生成绩.xlsx").fillna(0).replace(['弘金国际地学校'], '弘金地国际学校')
[SCHOOL_NAME.append(x) for x in STUDENT_TOTAL['学校'] if x not in SCHOOL_NAME]
SCHOOL_NAME.append("全体")

# 读取各科总分和小题满分
CHIN_TOTAL = pd.read_excel("./考试数据/小题分/小题分(语文).xls").fillna(0).replace(['弘金国际地学校'], '弘金地国际学校')
for index, value in STUDENT_TOTAL.iterrows():
    student = STUDENT_TOTAL.iloc[index]
    if student['学号'] not in CHIN_TOTAL['学号'].values:
        missing_student = {'学号': student['学号'], '考号': student['考号'], '姓名': student['姓名'], '班级': student['班级'],
                           '学校': student['学校']}
        CHIN_TOTAL = CHIN_TOTAL.append(missing_student, ignore_index=True).fillna(0)
CHIN_PROB_TOTAL = pd.read_excel("./考试数据/小题满分.xlsx", sheet_name="语文").fillna(0)

MATH_TOTAL = pd.read_excel("./考试数据/小题分/小题分(数学).xls").fillna(0).replace(['弘金国际地学校'], '弘金地国际学校')
for index, value in STUDENT_TOTAL.iterrows():
    student = STUDENT_TOTAL.iloc[index]
    if student['学号'] not in MATH_TOTAL['学号'].values:
        missing_student = {'学号': student['学号'], '考号': student['考号'], '姓名': student['姓名'], '班级': student['班级'],
                           '学校': student['学校']}
        MATH_TOTAL = MATH_TOTAL.append(missing_student, ignore_index=True).fillna(0)
MATH_PROB_TOTAL = pd.read_excel("./考试数据/小题满分.xlsx", sheet_name="数学").fillna(0)

ENGL_TOTAL = pd.read_excel("./考试数据/小题分/小题分(英语).xls").fillna(0).replace(['弘金国际地学校'], '弘金地国际学校')
for index, value in STUDENT_TOTAL.iterrows():
    student = STUDENT_TOTAL.iloc[index]
    if student['学号'] not in ENGL_TOTAL['学号'].values:
        missing_student = {'学号': student['学号'], '考号': student['考号'], '姓名': student['姓名'], '班级': student['班级'],
                           '学校': student['学校']}
        ENGL_TOTAL = ENGL_TOTAL.append(missing_student, ignore_index=True).fillna(0)
ENGL_PROB_TOTAL = pd.read_excel("./考试数据/小题满分.xlsx", sheet_name="英语").fillna(0)

HIST_TOTAL = pd.read_excel("./考试数据/小题分/小题分(历史).xls").fillna(0).replace(['弘金国际地学校'], '弘金地国际学校')
for index, value in STUDENT_TOTAL.iterrows():
    student = STUDENT_TOTAL.iloc[index]
    if student['学号'] not in HIST_TOTAL['学号'].values:
        missing_student = {'学号': student['学号'], '考号': student['考号'], '姓名': student['姓名'], '班级': student['班级'],
                           '学校': student['学校']}
        HIST_TOTAL = HIST_TOTAL.append(missing_student, ignore_index=True).fillna(0)
HIST_PROB_TOTAL = pd.read_excel("./考试数据/小题满分.xlsx", sheet_name="历史").fillna(0)

PHYS_TOTAL = pd.read_excel("./考试数据/小题分/小题分(物理).xls").fillna(0).replace(['弘金国际地学校'], '弘金地国际学校')
for index, value in STUDENT_TOTAL.iterrows():
    student = STUDENT_TOTAL.iloc[index]
    if student['学号'] not in PHYS_TOTAL['学号'].values:
        missing_student = {'学号': student['学号'], '考号': student['考号'], '姓名': student['姓名'], '班级': student['班级'],
                           '学校': student['学校']}
        PHYS_TOTAL = PHYS_TOTAL.append(missing_student, ignore_index=True).fillna(0)
PHYS_PROB_TOTAL = pd.read_excel("./考试数据/小题满分.xlsx", sheet_name="物理").fillna(0)

CHEM_TOTAL = pd.read_excel("./考试数据/小题分/小题分(化学).xls").fillna(0).replace(['弘金国际地学校'], '弘金地国际学校')
for index, value in STUDENT_TOTAL.iterrows():
    student = STUDENT_TOTAL.iloc[index]
    if student['学号'] not in CHEM_TOTAL['学号'].values:
        missing_student = {'学号': student['学号'], '考号': student['考号'], '姓名': student['姓名'], '班级': student['班级'],
                           '学校': student['学校']}
        CHEM_TOTAL = CHEM_TOTAL.append(missing_student, ignore_index=True).fillna(0)
CHEM_PROB_TOTAL = pd.read_excel("./考试数据/小题满分.xlsx", sheet_name="化学").fillna(0)

POLI_TOTAL = pd.read_excel("./考试数据/小题分/小题分(道德与法治).xls").fillna(0).replace(['弘金国际地学校'], '弘金地国际学校')
for index, value in STUDENT_TOTAL.iterrows():
    student = STUDENT_TOTAL.iloc[index]
    if student['学号'] not in POLI_TOTAL['学号'].values:
        missing_student = {'学号': student['学号'], '考号': student['考号'], '姓名': student['姓名'], '班级': student['班级'],
                           '学校': student['学校']}
        POLI_TOTAL = POLI_TOTAL.append(missing_student, ignore_index=True).fillna(0)
POLI_PROB_TOTAL = pd.read_excel("./考试数据/小题满分.xlsx", sheet_name="道德与法治").fillna(0)


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

SUBJECT_PROB_TOTAL_LIST = [
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


def ability_analysis(item_score, subject_name, subject_total, subject_prob_total):
    ability_name_list = list(subject_prob_total.columns[4:-2])
    prob_name_list = list(subject_total.columns[8:])
    this_subject = item_score.groupby('sskm_mc').get_group(subject_name)

    ability_score_rate = {"能力层次": [], "满分": [], "平均分": [], "标准差": [], "得分率": []}

    for ability_name in ability_name_list:
        ability_total = subject_prob_total[ability_name].sum()
        ability_ave = 0
        ability_prob = []
        for prob_index in range(len(subject_prob_total.index)):
            this_prob_total = subject_prob_total.groupby('sj_tid').get_group(prob_index + 1)

            target = int(this_prob_total[ability_name])
            total = int(this_prob_total['score'])
            if target != 0:
                prob_name = prob_name_list[prob_index]
                this_prob = subject_total[prob_name]
                this_prob_item_score = this_subject.groupby('sj_tid').get_group(prob_index + 1)
                ability_prob.append(this_prob_item_score)
                coefficient = target / total
                this_prob_ave = this_prob.mean()
                this_prob_ability_ave = this_prob_ave * coefficient
                ability_ave += this_prob_ability_ave
        ability_score_rate['能力层次'].append(ability_name)
        if len(ability_prob) != 0:
            ability_concated = pd.concat(ability_prob)
            ability_std = ability_concated['score'].std()
        else:
            ability_std = 0
        if ability_ave == 0:
            score_rate = 0
        else:
            score_rate = np.round(ability_ave / ability_total * 100, 2)
        ability_score_rate['满分'].append(ability_total)
        ability_score_rate['平均分'].append(np.round(ability_ave, 2))
        ability_score_rate['标准差'].append(np.round(ability_std, 2))
        ability_score_rate['得分率'].append(score_rate)

    ability_score_rate_df = pd.DataFrame(ability_score_rate)

    return ability_score_rate_df


def knowledge_field_and_point(item_score, subject_name, subject_total, subject_prob_total):
    this_subject = item_score.groupby("sskm_mc").get_group(subject_name)

    knowledge_field_list = list(subject_prob_total['知识模块'])
    knowledge_field_list_compact = []
    [knowledge_field_list_compact.append(x)
     for x in knowledge_field_list
     if x not in knowledge_field_list_compact and x != 0]

    knowledge_point_list = list(subject_prob_total['知识点'])
    knowledge_point_list_compact = []
    [knowledge_point_list_compact.append(x)
     for x in knowledge_point_list
     if x not in knowledge_point_list_compact and x != 0]

    knowledge_field_dict = {}
    knowledge_point_dict = {}

    concated_field = {}
    concated_point = {}

    if len(knowledge_field_list_compact) != 0:
        knowledge_field_dict["知识模块"] = knowledge_field_list_compact
        knowledge_field_dict["知识模块满分"] = []
        knowledge_field_dict["知识模块平均分"] = []
        knowledge_field_dict["知识模块标准差"] = []
        knowledge_field_dict["知识模块得分率"] = []

        for length in range(len(knowledge_field_list_compact)):
            knowledge_field_dict["知识模块满分"].append(0)
            knowledge_field_dict["知识模块平均分"].append(0)
            knowledge_field_dict["知识模块标准差"].append(0)
            knowledge_field_dict["知识模块得分率"].append(0)

        prob_name_list = list(subject_total.columns[8:])
        for index in range(len(subject_prob_total)):
            prob_name = prob_name_list[index]

            this_prob_total = subject_prob_total.groupby('sj_tid').get_group(index + 1)
            this_prob_item_score = this_subject.groupby('sj_tid').get_group(index+1)
            knowledge_field_name = this_prob_total['知识模块'].values
            if knowledge_field_name != 0:
                knowledge_field_index = knowledge_field_list_compact.index(knowledge_field_name)

                if knowledge_field_name[0] not in concated_field:
                    concated_field[knowledge_field_name[0]] = []
                    concated_field[knowledge_field_name[0]].append(this_prob_item_score)
                else:
                    concated_field[knowledge_field_name[0]].append(this_prob_item_score)

                prob_ave = subject_total[prob_name].mean()
                total = int(this_prob_total['score'].values)
                knowledge_field_dict["知识模块满分"][knowledge_field_index] += total

                if prob_ave != 0:
                    knowledge_field_dict["知识模块平均分"][knowledge_field_index] += prob_ave

        for x in range(len(knowledge_field_list_compact)):
            this_ave = knowledge_field_dict["知识模块平均分"][x]
            this_total = knowledge_field_dict["知识模块满分"][x]
            this_field_all = pd.concat(concated_field[knowledge_field_list_compact[x]])
            knowledge_field_dict["知识模块标准差"][x] = np.round(this_field_all['score'].std(), 2)
            knowledge_field_dict["知识模块得分率"][x] = np.round(this_ave / this_total * 100, 2)
    knowledge_field_df = pd.DataFrame(knowledge_field_dict)

    if len(knowledge_point_list_compact) != 0:
        knowledge_point_dict["知识点"] = knowledge_point_list_compact
        knowledge_point_dict["知识点满分"] = []
        knowledge_point_dict["知识点平均分"] = []
        knowledge_point_dict["知识点标准差"] = []
        knowledge_point_dict["知识点得分率"] = []

        for length in range(len(knowledge_point_list_compact)):
            knowledge_point_dict["知识点满分"].append(0)
            knowledge_point_dict["知识点平均分"].append(0)
            knowledge_point_dict["知识点标准差"].append(0)
            knowledge_point_dict["知识点得分率"].append(0)

        prob_name_list = list(subject_total.columns[8:])
        for index in range(len(subject_prob_total)):
            prob_name = prob_name_list[index]

            this_prob_total = subject_prob_total.groupby('sj_tid').get_group(index + 1)
            this_prob_item_score = this_subject.groupby('sj_tid').get_group(index + 1)
            knowledge_point_name = this_prob_total['知识点'].values
            if knowledge_point_name != 0:
                knowledge_point_index = knowledge_point_list_compact.index(knowledge_point_name)

                if knowledge_point_name[0] not in concated_point:
                    concated_point[knowledge_point_name[0]] = []
                    concated_point[knowledge_point_name[0]].append(this_prob_item_score)
                else:
                    concated_point[knowledge_point_name[0]].append(this_prob_item_score)

                prob_ave = subject_total[prob_name].mean()
                total = int(this_prob_total['score'].values)
                knowledge_point_dict["知识点满分"][knowledge_point_index] += total

                if prob_ave != 0:
                    knowledge_point_dict["知识点平均分"][knowledge_point_index] += prob_ave

        for x in range(len(knowledge_point_list_compact)):
            this_ave = knowledge_point_dict["知识点平均分"][x]
            this_total = knowledge_point_dict["知识点满分"][x]
            this_point_all = pd.concat(concated_point[knowledge_point_list_compact[x]])
            knowledge_point_dict["知识点标准差"][x] = np.round(this_point_all['score'].std(), 2)
            knowledge_point_dict["知识点得分率"][x] = np.round(this_ave / this_total * 100, 2)
    knowledge_point_df = pd.DataFrame(knowledge_point_dict)

    output_dict = {"知识模块分析.xlsx": knowledge_field_df, "知识点分析.xlsx": knowledge_point_df}
    return output_dict


def subject(path):
    print("学科信息计算中......\n")
    for x in range(0, 7):
        subject_name = SUBJECT_LIST[x]
        subject_full_score = SUBJECT_FULL_SCORE_LIST[x]
        subject_total = SUBJECT_TOTAL_LIST[x]
        subject_prob_total = SUBJECT_PROB_TOTAL_LIST[x]

        this_subject = subject_analysis.Analysis(subject_name, SCHOOL_NAME, subject_full_score,
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
                        # 难度与区分度分布图
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

                        # 预估难度与实测难度图
                        chart = workbook.add_chart({'type': 'line'})
                        max_row = len(problem_analysis[each])
                        col_x = problem_analysis[each].columns.get_loc('题号')
                        col_y_actual_difficulty = problem_analysis[each].columns.get_loc('预估难度')
                        col_y_estimate_difficulty = problem_analysis[each].columns.get_loc('难度')

                        chart.add_series({
                            'name': '预估难度',
                            'categories': [each, 1, col_x, max_row, col_x],
                            'values': [each, 1, col_y_actual_difficulty, max_row, col_y_actual_difficulty],
                            'data_labels': {'value': True},
                        })

                        chart.add_series({
                            'name': '实测难度',
                            'categories': [each, 1, col_x, max_row, col_x],
                            'values': [each, 1, col_y_estimate_difficulty, max_row, col_y_estimate_difficulty],
                            'data_labels': {'value': True},
                        })

                        chart.set_title({'name': subject_name + '预估难度与实测难度对比图'})
                        chart.set_x_axis({'name': '题号', 'text_axis': True})
                        chart.set_y_axis({'name': '难度值', 'min': 0, 'max': 1})
                        chart.set_size({'width': 920, 'height': 860})
                        worksheet.insert_chart('P30', chart)

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

    print("学科知识领域与能力层次分析中......")
    for x in range(6):
        subject_name = SUBJECT_LIST[x]
        print("当前科目：{}".format(subject_name))
        subject_total = SUBJECT_TOTAL_LIST[x]
        subject_prob_total = SUBJECT_PROB_TOTAL_LIST[x]
        ability_analysis_result = ability_analysis(ITEM_SCORE, subject_name, subject_total, subject_prob_total)

        ability_and_knowledge_output_path = "{}/{}/{}".format(path, subject_name, subject_name)

        with pd.ExcelWriter(ability_and_knowledge_output_path+'能力层次得分率.xlsx', engine='xlsxwriter') as writer:
            ability_analysis_result.to_excel(writer, index=False, sheet_name=subject_name+"能力层次得分率")
            workbook = writer.book
            worksheet = writer.sheets[subject_name+"能力层次得分率"]
            chart = workbook.add_chart({'type': 'column'})
            max_row = len(ability_analysis_result)
            col_x = ability_analysis_result.columns.get_loc('能力层次')
            col_y = ability_analysis_result.columns.get_loc("得分率")

            chart.add_series({
                'name': '能力层次得分率',
                'categories': [subject_name+"能力层次得分率", 1, col_x, max_row, col_x],
                'values': [subject_name+"能力层次得分率", 1, col_y, max_row, col_y],
                'data_labels': {'value': True},
            })

            chart.set_title({'name': subject_name + "能力层次得分率图"})
            chart.set_x_axis({'name': '能力层次', 'text_axis': True})
            chart.set_y_axis({'name': '得分率'})
            chart.set_size({'width': 720, 'height': 576})
            worksheet.insert_chart('A7', chart)

        knowledge_analysis = knowledge_field_and_point(ITEM_SCORE, subject_name, subject_total, subject_prob_total)
        for excel_name in knowledge_analysis.keys():
            with pd.ExcelWriter(ability_and_knowledge_output_path+excel_name, engine='xlsxwriter') as writer:
                knowledge_analysis[excel_name].to_excel(writer, index=False, sheet_name=excel_name[:-5])
                if excel_name == "知识模块分析.xlsx":
                    if len(knowledge_analysis[excel_name]) != 0:
                        workbook = writer.book
                        worksheet = writer.sheets["知识模块分析"]
                        chart = workbook.add_chart({'type': 'column'})
                        max_row = len(knowledge_analysis[excel_name])
                        col_x = knowledge_analysis[excel_name].columns.get_loc('知识模块')
                        col_y = knowledge_analysis[excel_name].columns.get_loc("知识模块得分率")

                        chart.add_series({
                            'name': '知识模块得分率',
                            'categories': ["知识模块分析", 1, col_x, max_row, col_x],
                            'values': ["知识模块分析", 1, col_y, max_row, col_y],
                            'data_labels': {'value': True},
                        })

                        chart.set_title({'name': subject_name + "知识模块得分率图"})
                        chart.set_x_axis({'name': '知识模块', 'text_axis': True})
                        chart.set_y_axis({'name': '得分率'})
                        chart.set_size({'width': 720, 'height': 576})
                        worksheet.insert_chart('F2', chart)
                if excel_name == "知识点分析.xlsx":
                    if len(knowledge_analysis[excel_name]) != 0:
                        workbook = writer.book
                        worksheet = writer.sheets["知识点分析"]
                        chart = workbook.add_chart({'type': 'column'})
                        max_row = len(knowledge_analysis[excel_name])
                        col_x = knowledge_analysis[excel_name].columns.get_loc('知识点')
                        col_y = knowledge_analysis[excel_name].columns.get_loc("知识点得分率")

                        chart.add_series({
                            'name': '知识点得分率',
                            'categories': ["知识点分析", 1, col_x, max_row, col_x],
                            'values': ["知识点分析", 1, col_y, max_row, col_y],
                            'data_labels': {'value': True},
                        })

                        chart.set_title({'name': subject_name + "知识点得分率图"})
                        chart.set_x_axis({'name': '知识点', 'text_axis': True})
                        chart.set_y_axis({'name': '得分率'})
                        chart.set_size({'width': 720, 'height': 576})
                        worksheet.insert_chart('F2', chart)
        print("输出完成")
    print("学科知识领域与能力层次分析完成\n")


def overall(path):
    print("全区信息计算中......\n")
    overall = overall_analysis.Analysis(SCHOOL_NAME, SUBJECT_LIST, SUBJECT_FULL_SCORE_LIST,
                                      SUBJECT_TOTAL_LIST, SUBJECT_TOTAL_SCORE)

    overall_analysis_output = overall.overall_output()
    print("全区报表计算完成")
    if not os.path.exists(path):
        os.makedirs(path)
    for each in overall_analysis_output.keys():
        excel_name = "{}.xlsx".format(each)
        this_excel = overall_analysis_output[each]
        if each == "2-2-2总分及各科上线的有效分数线":
            effective_score_analysis = overall.subject_effective_score_analysis(this_excel)
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
                overall_analysis_output[each].to_excel(writer, index=False, sheet_name=each)

    print("全区报表输出完成\n")


def school_output(path):
    print("各学校学科信息计算中......\n")
    for school in SCHOOL_NAME[:-1]:
        read_path = "./考试数据/学校数据/{}/".format(school)

        this_school_export_path = path + "{}/".format(school)

        item_score = pd.read_excel(read_path + "item_score.xlsx").fillna(0)

        school_student_total = pd.read_excel(read_path + "{}-学生信息.xlsx".format(school)).fillna(0)

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
        for index in range(0, 7):
            subject_name = SUBJECT_LIST[index]
            subject_full_score = SUBJECT_FULL_SCORE_LIST[index]
            subject_total = pd.read_excel(read_path + "/小题分/小题分({}).xls".format(subject_name)).fillna(0)
            school_subject_total_list.append(subject_total)
            subject_prob_total = SUBJECT_PROB_TOTAL_LIST[index]
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

        print(school + "学科知识领域与能力层次分析中......")
        for x in range(6):
            subject_name = SUBJECT_LIST[x]
            subject_total = school_subject_total_list[x]
            subject_prob_total = SUBJECT_PROB_TOTAL_LIST[x]
            ability_analysis_result = ability_analysis(item_score, subject_name, subject_total, subject_prob_total)

            ability_and_knowledge_output_path = "{}/学科/{}/{}".format(this_school_export_path, subject_name, subject_name)
            with pd.ExcelWriter(ability_and_knowledge_output_path + '能力层次得分率.xlsx', engine='xlsxwriter') as writer:
                ability_analysis_result.to_excel(writer, index=False, sheet_name=subject_name + "能力层次得分率")
                workbook = writer.book
                worksheet = writer.sheets[subject_name + "能力层次得分率"]
                chart = workbook.add_chart({'type': 'column'})
                max_row = len(ability_analysis_result)
                col_x = ability_analysis_result.columns.get_loc('能力层次')
                col_y = ability_analysis_result.columns.get_loc("得分率")

                chart.add_series({
                    'name': '能力层次得分率',
                    'categories': [subject_name + "能力层次得分率", 1, col_x, max_row, col_x],
                    'values': [subject_name + "能力层次得分率", 1, col_y, max_row, col_y],
                    'data_labels': {'value': True},
                })

                chart.set_title({'name': subject_name + "能力层次得分率图"})
                chart.set_x_axis({'name': '能力层次', 'text_axis': True})
                chart.set_y_axis({'name': '得分率'})
                chart.set_size({'width': 720, 'height': 576})
                worksheet.insert_chart('A7', chart)

            knowledge_analysis = knowledge_field_and_point(item_score, subject_name, subject_total, subject_prob_total)
            for excel_name in knowledge_analysis.keys():
                with pd.ExcelWriter(ability_and_knowledge_output_path + excel_name, engine='xlsxwriter') as writer:
                    knowledge_analysis[excel_name].to_excel(writer, index=False, sheet_name=excel_name[:-5])
                    if excel_name == "知识模块分析.xlsx":
                        if len(knowledge_analysis[excel_name]) != 0:
                            workbook = writer.book
                            worksheet = writer.sheets["知识模块分析"]
                            chart = workbook.add_chart({'type': 'column'})
                            max_row = len(knowledge_analysis[excel_name])
                            col_x = knowledge_analysis[excel_name].columns.get_loc('知识模块')
                            col_y = knowledge_analysis[excel_name].columns.get_loc("知识模块得分率")

                            chart.add_series({
                                'name': '知识模块得分率',
                                'categories': ["知识模块分析", 1, col_x, max_row, col_x],
                                'values': ["知识模块分析", 1, col_y, max_row, col_y],
                                'data_labels': {'value': True},
                            })

                            chart.set_title({'name': subject_name + "知识模块得分率图"})
                            chart.set_x_axis({'name': '知识模块', 'text_axis': True})
                            chart.set_y_axis({'name': '得分率'})
                            chart.set_size({'width': 720, 'height': 576})
                            worksheet.insert_chart('F2', chart)
                    if excel_name == "知识点分析.xlsx":
                        if len(knowledge_analysis[excel_name]) != 0:
                            workbook = writer.book
                            worksheet = writer.sheets["知识点分析"]
                            chart = workbook.add_chart({'type': 'column'})
                            max_row = len(knowledge_analysis[excel_name])
                            col_x = knowledge_analysis[excel_name].columns.get_loc('知识点')
                            col_y = knowledge_analysis[excel_name].columns.get_loc("知识点得分率")

                            chart.add_series({
                                'name': '知识点得分率',
                                'categories': ["知识点分析", 1, col_x, max_row, col_x],
                                'values': ["知识点分析", 1, col_y, max_row, col_y],
                                'data_labels': {'value': True},
                            })

                            chart.set_title({'name': subject_name + "知识点得分率图"})
                            chart.set_x_axis({'name': '知识点', 'text_axis': True})
                            chart.set_y_axis({'name': '得分率'})
                            chart.set_size({'width': 720, 'height': 576})
                            worksheet.insert_chart('F2', chart)

        print(school + "学科知识领域与能力层次分析完成")

        print(school + "整体报表计算中")
        overall = school_overall_analysis.Analysis(Class_name, SUBJECT_LIST, SUBJECT_FULL_SCORE_LIST,
                                                 school_subject_total_list, school_subject_total_score)

        overall_analysis_output = overall.overall_output()
        print(school + "整体报表计算完成\n")
        if not os.path.exists(path):
            os.makedirs(path)
        for each in overall_analysis_output.keys():
            excel_name = "{}.xlsx".format(each)
            this_excel = overall_analysis_output[each]
            if each == "2-2-2总分及各科上线的有效分数线":
                effective_score_analysis = overall.subject_effective_score_analysis(this_excel)
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
                    overall_analysis_output[each].to_excel(writer, index=False, sheet_name=each)

    print("各学校报表输出完成\n")


if __name__ == '__main__':
    export_path = "./OUTPUT/"
    subject_export_path = export_path + "学科/"
    school_export_path = export_path + "学校/"
    overall_export_path = export_path + "全区/"
    school_data_path = "./考试数据/学校数据/"

    # 计算并输出学科报表
    subject(subject_export_path)

    # 计算并输出全区报表
    overall(overall_export_path)

    # 按学校拆分各表
    print("按学校拆分各表中......")
    print("学校名单:" + ",".join(SCHOOL_NAME))
    seperate_school_data.student_total_overall_to_school(STUDENT_TOTAL, SCHOOL_NAME, school_data_path)
    for index in range(7):
        subject_total = SUBJECT_TOTAL_LIST[index]
        subject_name = SUBJECT_LIST[index]
        seperate_school_data.subject_total_overall_to_school(subject_total, SCHOOL_NAME, school_data_path, subject_name)
    print("拆分完成\n")

    # 计算并输出学校报表
    school_output(school_export_path)
