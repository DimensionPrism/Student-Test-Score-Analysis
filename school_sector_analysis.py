import os
import pandas as pd
import numpy as np


class Analysis:
    def __init__(self, Class_name, subject_list, subject_full_score_list, subject_total_list, subject_total_score):

        self.Class_name = Class_name
        self.subject_list = subject_list
        self.subject_full_score_list = subject_full_score_list
        self.subject_total_list = subject_total_list

        self.subject_total_score = subject_total_score

    def sector_output(self):

        Class_name = self.Class_name
        subject_list = self.subject_list
        subject_full_score_list = self.subject_full_score_list
        subject_total_list = self.subject_total_list

        subject_total_score = self.subject_total_score

        output_dict = {}

        total_score_analysis = self.total_score_analysis(Class_name, subject_total_score, subject_full_score_list)
        output_dict["1-3-2总分得分情况"] = total_score_analysis

        subject_score_analysis = self.subject_score_analysis(subject_list, subject_total_list,
                                                             subject_total_score, subject_full_score_list)
        output_dict["1-3-1科目得分情况"] = subject_score_analysis

        total_score_distribution_analysis = self.total_score_distribution_analysis(Class_name, subject_total_score,
                                                                                   subject_full_score_list)
        output_dict["1-5-2总分分数段分布表"] = total_score_distribution_analysis

        total_score_line_analysis = self.total_score_line_analysis(Class_name, subject_total_score)
        output_dict["1-4总分上线人数（目标1-5）"] = total_score_line_analysis

        score_line = self.score_line(subject_total_score, subject_list)
        output_dict["2-2-2总分及各科上线的有效分数线"] = score_line

        return output_dict

    def total_score_analysis(self, Class_name, subject_total_score, subject_full_score_list):
        full_score = subject_full_score_list[-1]
        score_dict = {"班级": Class_name, "最高分": [], "最低分": [], "中位数": [], "平均分": [], "标准差": [],
                      "满分率": [], "超优率": [], "优秀率": [], "良好率": [], "及格率": [], "待及格": [], "低分率": [],
                      "全距": []}

        append_list = ["超优率", "优秀率", "良好率", "及格率"]
        for Class in Class_name:
            if Class == "全校":
                this_Class = subject_total_score
            else:
                this_Class = subject_total_score.groupby("班级").get_group(Class)

            student_score = this_Class['分数.7']
            stu_num = len(student_score)

            score_dict["最高分"].append(student_score.max())
            score_dict["最低分"].append(student_score.min())
            score_dict["中位数"].append(np.round(student_score.median(), 2))
            score_dict["平均分"].append(np.round(student_score.mean(), 2))
            score_dict["标准差"].append(np.round(student_score.std(), 2))

            full_score_student = student_score.where(student_score == full_score).dropna()
            full_score_stu_num = len(full_score_student)
            full_score_rate = str(np.round(full_score_stu_num / stu_num * 100, 2))
            full_score_rate = "{}%".format(full_score_rate)
            score_dict["满分率"].append(full_score_rate)

            for interval in range(10, 6, -1):
                append_index = 10 - interval
                append_name = append_list[append_index]

                coefficient_high = interval / 10
                coefficient_low = (interval - 1) / 10

                this_rate_student = student_score.where(student_score < full_score * coefficient_high).dropna()
                this_rate_student = this_rate_student.where(this_rate_student >= full_score * coefficient_low).dropna()
                this_rate_stu_num = len(this_rate_student)

                this_rate = str(np.round(this_rate_stu_num / stu_num * 100, 2))
                this_rate = "{}%".format(this_rate)

                score_dict[append_name].append(this_rate)

            almost_pass_student = student_score.where(student_score < full_score * 0.6).dropna()
            almost_pass_student = student_score.where(almost_pass_student >= full_score * 0.4).dropna()
            almost_pass_stu_num = len(almost_pass_student)

            almost_pass_rate = str(np.round(almost_pass_stu_num / stu_num * 100, 2))
            almost_pass_rate = "{}%".format(almost_pass_rate)

            score_dict["待及格"].append(almost_pass_rate)

            low_student = student_score.where(student_score < full_score * 0.4).dropna()
            low_student = low_student.where(low_student >= 0).dropna()
            low_stu_num = len(low_student)

            low_rate = str(np.round(low_stu_num / stu_num * 100, 2))
            low_rate = "{}%".format(low_rate)

            score_dict["低分率"].append(low_rate)

            score_dict["全距"].append(np.round((student_score.max() - student_score.min()), 1))

        score_df = pd.DataFrame(score_dict)
        return score_df

    def subject_score_analysis(self, subject_list, subject_total_list, subject_total_score, subject_full_score_list):

        score_dict = {"学科": subject_list, "总分": subject_full_score_list, "最高分": [],
                      "最低分": [], "中位数": [], "平均分": [],
                      "标准差": [], "满分率": [], "超优率": [],
                      "优秀率": [], "良好率": [], "及格率": [],
                      "待及格": [], "低分率": [], "全距": []}

        append_list = ["超优率", "优秀率", "良好率", "及格率"]
        for index in range(8):
            if index == 7:
                this_subject = subject_total_score['分数.7']
            else:
                subject = subject_total_list[index]
                this_subject = subject['全卷']
            full_score = subject_full_score_list[index]

            stu_num = len(this_subject)

            score_dict["最高分"].append(this_subject.max())
            score_dict["最低分"].append(this_subject.min())
            score_dict["中位数"].append(np.round(this_subject.median(), 2))
            score_dict["平均分"].append(np.round(this_subject.mean(), 2))
            score_dict["标准差"].append(np.round(this_subject.std(), 2))

            full_score_student = this_subject.where(this_subject == full_score).dropna()
            full_score_stu_num = len(full_score_student)
            full_score_rate = str(np.round(full_score_stu_num / stu_num * 100, 2))
            full_score_rate = "{}%".format(full_score_rate)
            score_dict["满分率"].append(full_score_rate)

            for interval in range(10, 6, -1):
                append_index = 10 - interval
                append_name = append_list[append_index]

                coefficient_high = interval / 10
                coefficient_low = (interval - 1) / 10

                this_rate_student = this_subject.where(this_subject < full_score * coefficient_high).dropna()
                this_rate_student = this_rate_student.where(this_rate_student >= full_score * coefficient_low).dropna()
                this_rate_stu_num = len(this_rate_student)

                this_rate = str(np.round(this_rate_stu_num / stu_num * 100, 2))
                this_rate = "{}%".format(this_rate)

                score_dict[append_name].append(this_rate)

            almost_pass_student = this_subject.where(this_subject < full_score * 0.6).dropna()
            almost_pass_student = this_subject.where(almost_pass_student >= full_score * 0.4).dropna()
            almost_pass_stu_num = len(almost_pass_student)

            almost_pass_rate = str(np.round(almost_pass_stu_num / stu_num * 100, 2))
            almost_pass_rate = "{}%".format(almost_pass_rate)

            score_dict["待及格"].append(almost_pass_rate)

            low_student = this_subject.where(this_subject < full_score * 0.4).dropna()
            low_student = low_student.where(low_student >= 0).dropna()
            low_stu_num = len(low_student)

            low_rate = str(np.round(low_stu_num / stu_num * 100, 2))
            low_rate = "{}%".format(low_rate)

            score_dict["低分率"].append(low_rate)

            score_dict["全距"].append(np.round((this_subject.max() - this_subject.min()), 1))

        score_df = pd.DataFrame(score_dict)
        return score_df

    def total_score_distribution_analysis(self, Class_name, subject_total_score, subject_full_score_list):
        full_score = subject_full_score_list[-1]
        analysis_dict = {"班级": Class_name, "【0,10%）": [], "【10%,20%）": [], "【20%,30%）": [],
                         "【30%,40%）": [], "【40%,50%）": [], "【50%,60%）": [],
                         "【60%,70%）": [], "【70%,80%）": [], "【80%,90%）": [],
                         "【90%-100%）": [], "人数": []}

        append_list = ["【0,10%）", "【10%,20%）", "【20%,30%）", "【30%,40%）", "【40%,50%）", "【50%,60%）",
                       "【60%,70%）", "【70%,80%）", "【80%,90%）", "【90%-100%）"]
        for Class in Class_name:
            if Class == "全校":
                this_Class = subject_total_score['分数.7']
            else:
                this_Class = subject_total_score.groupby('班级').get_group(Class)
                this_Class = this_Class['分数.7']

            stu_num = len(this_Class)
            analysis_dict["人数"].append(stu_num)

            for interval in range(0, 10):
                append_name = append_list[interval]
                high_boundary = full_score * (interval + 1) / 10
                low_boundary = full_score * interval / 10

                student = this_Class.where(this_Class < high_boundary).dropna()
                student = student.where(this_Class >= low_boundary).dropna()
                num = len(student)

                analysis_dict[append_name].append(num)
        analysis_df = pd.DataFrame(analysis_dict)
        return analysis_df

    def total_score_line_analysis(self, Class_name, subject_total_score):
        score_line_dict = {"班级": Class_name, "人数": [], "目标1人数": [], "目标1上线率": [], "目标2人数": [],
                           "目标2上线率": [], "目标3人数": [], "目标3上线率": [], "目标4人数": [],
                           "目标4上线率": [], "目标5人数": [], "目标5上线率": []}

        total_score_line_list = [562, 508, 449, 382, 314]
        append_list_num = ["目标1人数", "目标2人数", "目标3人数", "目标4人数", "目标5人数"]
        append_list_rate = ["目标1上线率", "目标2上线率", "目标3上线率", "目标4上线率", "目标5上线率"]

        for Class in Class_name:
            if Class == "全校":
                this_Class = subject_total_score['分数.7']
            else:
                this_Class = subject_total_score.groupby('班级').get_group(Class)
                this_Class = this_Class['分数.7']
            stu_num = len(this_Class)
            score_line_dict["人数"].append(stu_num)

            for interval in range(5):
                score_line = total_score_line_list[interval]
                append_name_num = append_list_num[interval]
                append_name_rate = append_list_rate[interval]

                student = this_Class.where(this_Class >= score_line).dropna()
                num = len(student)
                rate = np.round(num / stu_num, 2)
                score_line_dict[append_name_num].append(num)
                score_line_dict[append_name_rate].append(rate)

        score_line_df = pd.DataFrame(score_line_dict)
        return score_line_df

    def score_line(self, subject_total_score, subject_list):
        total_score_line_list = [562, 508, 449, 382, 314]
        score_line_dict = {"总分上线分数": total_score_line_list}

        for subject in subject_list[:-1]:
            score_line_dict[subject+"上线分数"] = []

        for index in range(5):
            score_line = total_score_line_list[index]
            student = subject_total_score.where(subject_total_score['分数.7'] >= score_line).dropna()
            stu_ave = student['分数.7'].mean()
            score_line_coefficient = score_line / stu_ave

            for suffix in range(7):
                if suffix == 0:
                    subject_name = '分数'
                else:
                    subject_name = '分数.{}'.format(suffix)

                subject_ave = student[subject_name].mean()
                subject_effective_score = np.round(subject_ave * score_line_coefficient, 1)
                append_name = subject_list[suffix]+"上线分数"
                score_line_dict[append_name].append(subject_effective_score)

        score_line_df = pd.DataFrame(score_line_dict)
        return score_line_df

    def subject_effective_score_analysis(self, score_line):
        Class_name = self.Class_name
        subject_list = self.subject_list
        subject_total_score = self.subject_total_score

        total_score_line_list = [562, 508, 449, 382, 314]

        total_dict = {}

        for suffix in range(7):
            if suffix == 0:
                subject_code = '分数'
            else:
                subject_code = '分数.{}'.format(suffix)
            subject_name = subject_list[suffix]

            subject_dict = {}
            for total_score_line in total_score_line_list:
                score_line_index = total_score_line_list.index(total_score_line)
                subject_score_line_name = subject_name+"上线分数"
                name = "{}目标{}（{}）命中率分析.xlsx".format(subject_name, str(score_line_index + 1), str(total_score_line))

                subject_effective_score = score_line[subject_score_line_name][score_line_index]

                dict = {"班级": Class_name, "总分上线人数": [],
                        subject_name+"单上线": [], subject_name+"双上线": [],
                        subject_name+"M1": [], subject_name+"M2": []}
                for Class in Class_name:
                    if Class == "全校":
                        this_Class = subject_total_score
                    else:
                        this_Class = subject_total_score.groupby("班级").get_group(Class)

                    this_Class_above_line = this_Class.where(this_Class['分数.7'] >= total_score_line).dropna()
                    this_Class_above_line_stu_num = len(this_Class_above_line)
                    dict["总分上线人数"].append(this_Class_above_line_stu_num)

                    this_Class_single = this_Class.where(this_Class[subject_code] >= subject_effective_score).dropna()
                    num_single = len(this_Class_single)
                    dict[subject_name+"单上线"].append(num_single)

                    this_subject = this_Class_above_line[subject_code]
                    this_Class_double = this_subject.where(this_subject >= subject_effective_score).dropna()
                    num_double = len(this_Class_double)
                    dict[subject_name+"双上线"].append(num_double)

                    if num_single == 0:
                        this_Class_hit = 0
                    else:
                        this_Class_hit = num_double / num_single
                    dict[subject_name+"M1"].append(np.round(this_Class_hit*100, 1))

                    if this_Class_above_line_stu_num == 0:
                        this_Class_contribution = 0
                    else:
                        this_Class_contribution = num_double / this_Class_above_line_stu_num
                    dict[subject_name+"M2"].append(np.round(this_Class_contribution*100, 1))

                df = pd.DataFrame(dict)
                subject_dict[name] = df

            total_dict[subject_name] = subject_dict
        return total_dict


