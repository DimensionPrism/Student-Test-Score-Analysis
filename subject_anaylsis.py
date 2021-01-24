import os
import pandas as pd
import numpy as np


class Analysis:
    def __init__(self, subject_name, school_name, subject_full_score,
                 subject_prob_total, subject_total, item_score,
                 student_total, subject_total_score, index):
        item_score = item_score.groupby(item_score.sjmc).get_group(subject_name)
        self.item_score = item_score
        self.student_total = student_total
        self.subject_total_score = subject_total_score
        self.school_name = school_name

        difficulty_curve_block = []
        for interval in range(0, 10, 1):
            if interval == 0:
                high_group_boundary = "0"
                low_group_boundary = int(subject_full_score * ((interval + 1) / 10))
            else:
                high_group_boundary = int(subject_full_score * (interval / 10) + 1)
                low_group_boundary = int(subject_full_score * ((interval + 1) / 10))

            curve_block = "{}-{}".format(high_group_boundary, low_group_boundary)
            difficulty_curve_block.append(curve_block)

        sub_num = 0
        if "客观题" in self.item_score.values:
            self.sub_prob = subject_prob_total.groupby(subject_prob_total.sskm_dt).get_group("客观题")

            sub_data = item_score.groupby(item_score.sskm_dt).get_group("客观题")
            sub_data = sub_data[["xh", "sjmc", "sj_tid", "sskm_dt", "sskm_th", "score"]]
            self.sub_data = sub_data

            sub_num = sub_data['sj_tid'].max()
            self.sub_tid_grouped = sub_data.groupby(sub_data.sj_tid)

        obj_num = 0
        if "主观题" in self.item_score.values:
            self.obj_prob = subject_prob_total.groupby(subject_prob_total.sskm_dt).get_group("主观题")

            obj_data = item_score.groupby(item_score.sskm_dt).get_group("主观题")
            obj_data = obj_data[["xh", "sjmc", "sj_tid", "sskm_dt", "sskm_th", "score"]]
            self.obj_data = obj_data

            obj_num = obj_data['sj_tid'].max() - sub_num
            self.obj_tid_grouped = obj_data.groupby(obj_data.sj_tid)

        prob_num = item_score['sj_tid'].max()
        stu_num = len(item_score.index) / prob_num
        difficulty = np.round(subject_total['全卷'].mean() / 100, 2)

        prob_name = []
        for prob in subject_total.columns:
            prob_name.append(prob)
        del prob_name[0:8]

        subject_sorted = subject_total.sort_values(by=['全卷'], ascending=False)
        subject_sorted_reverse = subject_total.sort_values(by=['全卷'])

        boundary = int(np.round(len(subject_sorted.index) * 0.27))
        high_group = subject_sorted.head(boundary)
        low_group = subject_sorted.tail(boundary)

        self.index = index
        self.subject_name = subject_name
        self.subject_full_score = subject_full_score
        self.difficulty_curve_block = difficulty_curve_block

        self.sub_num = sub_num
        self.obj_num = obj_num
        self.prob_num = prob_num
        self.stu_num = stu_num
        self.difficulty = difficulty

        self.prob_name = prob_name
        self.prob_name_id = prob_name.copy()

        self.subject_total = subject_total
        self.subject_sorted = subject_sorted
        self.subject_sorted_reverse = subject_sorted_reverse

        self.boundary = boundary
        self.high_group = high_group
        self.low_group = low_group

        self.diff_range = ["[0, 0.2)", "[0.2, 0.4)", "[0.4, 0.6)", "[0.6, 0.8)", "[0.8, 1.0]"]
        self.diff_evaluate = ["难", "较难", "中等", "较易", "易"]
        self.dist_range = ["[0.4, *)", "[0.3, 0.4)", "[0.2, 0.3)", "[-1, 0.2"]
        self.dist_evaluate = ["优秀", "良好", "尚可", "差"]

    def problem_analysis(self):
        index = self.index

        high_group = self.high_group
        low_group = self.low_group

        stu_num = self.stu_num
        school_name = self.school_name

        test_difficulty = self.difficulty

        student_total = self.student_total
        subject_full_score = self.subject_full_score
        subject_total_score = self.subject_total_score
        subject_sorted_reverse = self.subject_sorted_reverse['全卷']

        sub_analysis = {"题号": [], "题型": [], "人数": [], "最高分": [],
                        "最低分": [], "平均分": [], "标准差": [], "得分率": [],
                        "满分率": [], "零分率": [], "难度": [], "区分度": [],
                        }
        obj_analysis = {"题号": [], "题型": [], "人数": [], "最高分": [],
                        "最低分": [], "平均分": [], "标准差": [], "得分率": [],
                        "满分率": [], "零分率": [], "难度": [], "区分度": [],
                        }
        all_prob_analysis = {"题号": [], "题型": [], "人数": [], "最高分": [],
                             "最低分": [], "平均分": [], "标准差": [], "得分率": [],
                             "满分率": [], "零分率": [], "难度": [], "区分度": [],
                             }

        diff_sub_num = [0, 0, 0, 0, 0]
        diff_sub_score = [0, 0, 0, 0, 0]
        diff_obj_num = [0, 0, 0, 0, 0]
        diff_obj_score = [0, 0, 0, 0, 0]
        diff_total_num = [0, 0, 0, 0, 0]
        diff_total_score = [0, 0, 0, 0, 0]
        diff_sub_name = [[], [], [], [], []]
        diff_obj_name = [[], [], [], [], []]
        diff_rate = [0, 0, 0, 0, 0]

        dist_sub_num = [0, 0, 0, 0]
        dist_sub_score = [0, 0, 0, 0]
        dist_obj_num = [0, 0, 0, 0]
        dist_obj_score = [0, 0, 0, 0]
        dist_total_num = [0, 0, 0, 0]
        dist_total_score = [0, 0, 0, 0]
        dist_sub_name = [[], [], [], []]
        dist_obj_name = [[], [], [], []]
        dist_rate = [0, 0, 0, 0]

        test_structure_title = []
        test_structure_score = []
        test_structure_num = []
        test_structure_name = []
        test_structure_ave = []
        test_structure_std = []
        test_structure_score_rate = []

        # 全卷分段难度计算
        difficulty_list = []
        difficulty_curve = {}
        for interval in range(0, 10):
            low_group_boundary = subject_full_score * ((interval + 1) / 10)
            if interval == 0:
                high_group_boundary = subject_full_score * (interval / 10)
            else:
                high_group_boundary = subject_full_score * (interval / 10) + 1
            block = subject_sorted_reverse.where(subject_sorted_reverse <= low_group_boundary).dropna()
            block = block.where(block > high_group_boundary).dropna()
            if len(block) == 0:
                ave = 0
            else:
                ave = block.mean()
            difficulty = np.round(ave / subject_full_score, 2)
            difficulty_list.append(difficulty)

        test_diff_block_dict = {'分数段': self.difficulty_curve_block, '难度值': difficulty_list}
        test_diff_block_df = pd.DataFrame(test_diff_block_dict)
        difficulty_curve['全卷'] = test_diff_block_df

        if "客观题" in self.item_score.values:
            sub_prob = self.sub_prob.groupby(self.sub_prob.sj_tid)
            sub_num = self.sub_num

            test_structure_title.append("客观题")

            sub_score = self.sub_prob['score'].sum()
            test_structure_score.append(sub_score)

            test_structure_num.append(sub_num)

            test_structure_sub_name = []
            for sub in range(0, sub_num):
                # 小题名称
                prob_name = self.prob_name[sub]
                sub_analysis["题号"].append(prob_name)
                test_structure_sub_name.append(prob_name)

                # 提取所有人的这一小题信息
                sub_inc = sub + 1
                this_sub = self.sub_tid_grouped.get_group(sub_inc)

                sub_analysis["题型"].append("客观题")

                # 作答学生人数
                sub_analysis["人数"].append(int(stu_num))

                # 小题满分
                highest = sub_prob.get_group(sub_inc)
                highest = int(highest['score'].max())

                # 最高分与最低分计算
                sub_max = int(this_sub['score'].max())
                sub_analysis["最高分"].append(sub_max)
                sub_min = int(this_sub['score'].min())
                sub_analysis["最低分"].append(sub_min)

                # 平均分计算
                ave = np.round(this_sub['score'].mean(), 2)
                sub_analysis["平均分"].append(ave)

                # 标准差计算
                std = np.round(this_sub['score'].std(), 2)
                sub_analysis["标准差"].append(std)

                # 得分率
                score_rate = np.round(this_sub['score'].mean() / highest * 100, 2)
                sub_analysis["得分率"].append(score_rate)

                # 满分率
                if highest in this_sub['score'].values:
                    max_group = this_sub.groupby(this_sub.score).get_group(highest)
                    full_score_stu = len(max_group.index)
                    full_score_rate = np.round(full_score_stu / stu_num * 100, 2)
                else:
                    full_score_rate = 0
                sub_analysis["满分率"].append(full_score_rate)

                # 零分率
                if 0 in this_sub['score'].values:
                    zero_group = this_sub.groupby(this_sub.score).get_group(0)
                    zero_score_stu = len(zero_group.index)
                    zero_score_rate = np.round(zero_score_stu / stu_num * 100, 2)
                else:
                    zero_score_rate = 0
                sub_analysis["零分率"].append(zero_score_rate)

                # 难度计算
                if sub_max == 0:
                    sub_difficulty = 0
                else:
                    sub_difficulty = np.round(this_sub['score'].mean() / sub_max, 2)
                sub_analysis["难度"].append(sub_difficulty)
                if 0 <= sub_difficulty < 0.2:
                    diff_sub_num[0] += 1
                    diff_sub_score[0] += highest
                    diff_total_num[0] += 1
                    diff_total_score[0] += highest
                    diff_sub_name[0].append(prob_name)
                elif 0.2 <= sub_difficulty < 0.4:
                    diff_sub_num[1] += 1
                    diff_sub_score[1] += highest
                    diff_total_num[1] += 1
                    diff_total_score[1] += highest
                    diff_sub_name[1].append(prob_name)
                elif 0.4 <= sub_difficulty < 0.6:
                    diff_sub_num[2] += 1
                    diff_sub_score[2] += highest
                    diff_total_num[2] += 1
                    diff_total_score[2] += highest
                    diff_sub_name[2].append(prob_name)
                elif 0.6 <= sub_difficulty < 0.8:
                    diff_sub_num[3] += 1
                    diff_sub_score[3] += highest
                    diff_total_num[3] += 1
                    diff_total_score[3] += highest
                    diff_sub_name[3].append(prob_name)
                elif 0.8 <= sub_difficulty < 1:
                    diff_sub_num[4] += 1
                    diff_sub_score[4] += highest
                    diff_total_num[4] += 1
                    diff_total_score[4] += highest
                    diff_sub_name[4].append(prob_name)

                # 区分度计算
                if 0 in high_group[prob_name].values:
                    high_group_zero_rate = int(len(high_group.groupby(prob_name).get_group(0).index)) / self.boundary
                else:
                    high_group_zero_rate = 0
                high_group_diff = np.round(1 - high_group_zero_rate, 2)
                if 0 in low_group[prob_name].values:
                    low_group_zero_rate = int(len(low_group.groupby(prob_name).get_group(0).index)) / self.boundary
                else:
                    low_group_zero_rate = 0
                low_group_diff = np.round(1 - low_group_zero_rate, 2)
                difference = high_group_diff - low_group_diff
                if difference == 0:
                    distinction = -1
                else:
                    distinction = np.round(high_group_diff - low_group_diff, 2)
                sub_analysis["区分度"].append(distinction)
                if 0.4 <= distinction:
                    dist_sub_num[0] += 1
                    dist_sub_score[0] += highest
                    dist_total_num[0] += 1
                    dist_total_score[0] += highest
                    dist_sub_name[0].append(prob_name)
                elif 0.3 <= distinction < 0.4:
                    dist_sub_num[1] += 1
                    dist_sub_score[1] += highest
                    dist_total_num[1] += 1
                    dist_total_score[1] += highest
                    dist_sub_name[1].append(prob_name)
                elif 0.2 <= distinction < 0.3:
                    dist_sub_num[2] += 1
                    dist_sub_score[2] += highest
                    dist_total_num[2] += 1
                    dist_total_score[2] += highest
                    dist_sub_name[2].append(prob_name)
                elif -1 <= distinction < 0.2:
                    dist_sub_num[3] += 1
                    dist_sub_score[3] += highest
                    dist_total_num[3] += 1
                    dist_total_score[3] += highest
                    dist_sub_name[3].append(prob_name)

                # 分段难度计算
                difficulty_list = []
                subject_sorted_reverse = self.subject_sorted_reverse[['全卷', prob_name]]

                for interval in range(0, 10):
                    low_group_boundary = subject_full_score * ((interval + 1) / 10)
                    if interval == 0:
                        high_group_boundary = subject_full_score * (interval / 10)
                    else:
                        high_group_boundary = subject_full_score * (interval / 10) + 1
                    block = subject_sorted_reverse.where(subject_sorted_reverse['全卷'] <= low_group_boundary).dropna()
                    block = block.where(block['全卷'] > high_group_boundary).dropna()
                    if len(block) == 0:
                        ave = 0
                    else:
                        ave = block[prob_name].mean()
                    difficulty = np.round(ave / highest, 2)
                    difficulty_list.append(difficulty)

                prob_diff_block_dict = {'分数段': self.difficulty_curve_block, '难度值': difficulty_list}
                prob_diff_block_df = pd.DataFrame(prob_diff_block_dict)
                difficulty_curve[prob_name] = prob_diff_block_df

            test_structure_name.append(','.join(test_structure_sub_name))
            test_structure_ave.append(np.round(self.sub_data['score'].sum() / stu_num, 2))
            test_structure_score_rate.append(np.round(np.mean(sub_analysis["得分率"]), 2))
            students = []
            sum_of_score = []
            [students.append(x) for x in self.sub_data['xh'] if x not in students]
            for student in students:
                stu_grouped = self.sub_data.groupby(self.sub_data.xh).get_group(student)
                sum_of_score.append(stu_grouped['score'].sum())
            test_structure_std.append(np.round(np.std(sum_of_score), 2))

        if "主观题" in self.item_score.values:
            obj_prob = self.obj_prob.groupby(self.obj_prob.sj_tid)

            if self.sub_num == 0:
                first_obj = 0
            else:
                first_obj = self.sub_num
            prob_num = self.prob_num
            obj_num = prob_num - first_obj

            test_structure_title.append("主观题")
            obj_score = self.obj_prob['score'].sum()
            test_structure_score.append(obj_score)

            test_structure_num.append(obj_num)

            test_structure_obj_name = []
            for obj in range(first_obj, prob_num):
                # 小题名称
                prob_name = self.prob_name[obj]
                obj_analysis["题号"].append(prob_name)
                test_structure_obj_name.append(prob_name)

                # 提取所有人的这一小题信息
                obj_inc = obj + 1
                this_obj = self.obj_tid_grouped.get_group(obj_inc)

                obj_analysis["题型"].append("主观题")

                # 作答学生人数
                obj_analysis["人数"].append(int(stu_num))

                # 小题满分
                highest = obj_prob.get_group(obj_inc)
                highest = int(highest['score'].max())

                # 最高分与最低分计算
                obj_max = int(this_obj['score'].max())
                obj_analysis["最高分"].append(obj_max)
                obj_min = int(this_obj['score'].min())
                obj_analysis["最低分"].append(obj_min)

                # 平均分计算
                ave = np.round(this_obj['score'].mean(), 2)
                obj_analysis["平均分"].append(ave)

                # 标准差计算
                std = np.round(this_obj['score'].std(), 2)
                obj_analysis["标准差"].append(std)

                # 得分率
                score_rate = np.round(this_obj['score'].mean() / highest * 100, 2)
                obj_analysis["得分率"].append(score_rate)

                # 满分率
                if highest in this_obj['score'].values:
                    max_group = this_obj.groupby(this_obj.score).get_group(highest)
                    full_score_stu = len(max_group.index)
                    full_score_rate = np.round(full_score_stu / stu_num * 100, 2)
                else:
                    full_score_rate = 0
                obj_analysis["满分率"].append(full_score_rate)

                # 零分率
                if 0 in this_obj['score'].values:
                    zero_group = this_obj.groupby(this_obj.score).get_group(0)
                    zero_score_stu = len(zero_group.index)
                    zero_score_rate = np.round(zero_score_stu / stu_num * 100, 2)
                else:
                    zero_score_rate = 0
                obj_analysis["零分率"].append(zero_score_rate)

                # 难度计算
                if obj_max == 0:
                    obj_difficulty = 0
                else:
                    obj_difficulty = np.round(this_obj['score'].mean() / obj_max, 2)
                obj_analysis["难度"].append(obj_difficulty)
                if 0 <= obj_difficulty < 0.2:
                    diff_obj_num[0] += 1
                    diff_obj_score[0] += highest
                    diff_total_num[0] += 1
                    diff_total_score[0] += highest
                    diff_obj_name[0].append(prob_name)
                elif 0.2 <= obj_difficulty < 0.4:
                    diff_obj_num[1] += 1
                    diff_obj_score[1] += highest
                    diff_total_num[1] += 1
                    diff_total_score[1] += highest
                    diff_obj_name[1].append(prob_name)
                elif 0.4 <= obj_difficulty < 0.6:
                    diff_obj_num[2] += 1
                    diff_obj_score[2] += highest
                    diff_total_num[2] += 1
                    diff_total_score[2] += highest
                    diff_obj_name[2].append(prob_name)
                elif 0.6 <= obj_difficulty < 0.8:
                    diff_obj_num[3] += 1
                    diff_obj_score[3] += highest
                    diff_total_num[3] += 1
                    diff_total_score[3] += highest
                    diff_obj_name[3].append(prob_name)
                elif 0.8 <= obj_difficulty < 1:
                    diff_obj_num[4] += 1
                    diff_obj_score[4] += highest
                    diff_total_num[4] += 1
                    diff_total_score[4] += highest
                    diff_obj_name[4].append(prob_name)

                # 区分度计算
                high_group_total = high_group[prob_name].sum()
                low_group_total = low_group[prob_name].sum()
                prob_max = self.subject_sorted[prob_name].max()
                prob_min = self.subject_sorted[prob_name].min()
                difference = prob_max - prob_min
                if difference == 0:
                    distinction = -1
                else:
                    distinction = (high_group_total - low_group_total) / (self.boundary * (prob_max - prob_min))
                distinction = np.round(distinction, 2)
                obj_analysis["区分度"].append(distinction)
                if 0.4 <= distinction:
                    dist_obj_num[0] += 1
                    dist_obj_score[0] += highest
                    dist_total_num[0] += 1
                    dist_total_score[0] += highest
                    dist_obj_name[0].append(prob_name)
                elif 0.3 <= distinction < 0.4:
                    dist_obj_num[1] += 1
                    dist_obj_score[1] += highest
                    dist_total_num[1] += 1
                    dist_total_score[1] += highest
                    dist_obj_name[1].append(prob_name)
                elif 0.2 <= distinction < 0.3:
                    dist_obj_num[2] += 1
                    dist_obj_score[2] += highest
                    dist_total_num[2] += 1
                    dist_total_score[2] += highest
                    dist_obj_name[2].append(prob_name)
                elif -1 <= distinction < 0.2:
                    dist_obj_num[3] += 1
                    dist_obj_score[3] += highest
                    dist_total_num[3] += 1
                    dist_total_score[3] += highest
                    dist_obj_name[3].append(prob_name)

                # 分段难度计算
                difficulty_list = []
                subject_sorted_reverse = self.subject_sorted_reverse[['全卷', prob_name]]

                for interval in range(0, 10):
                    low_group_boundary = subject_full_score * ((interval + 1) / 10)
                    if interval == 0:
                        high_group_boundary = -1
                    else:
                        high_group_boundary = subject_full_score * (interval / 10) + 1
                    block = subject_sorted_reverse.where(subject_sorted_reverse['全卷'] <= low_group_boundary).dropna()
                    block = block.where(block['全卷'] > high_group_boundary).dropna()
                    if len(block) == 0:
                        ave = 0
                    else:
                        ave = block[prob_name].mean()
                    difficulty = np.round(ave / highest, 2)
                    difficulty_list.append(difficulty)

                prob_diff_block_dict = {'分数段': self.difficulty_curve_block, '难度值': difficulty_list}
                prob_diff_block_df = pd.DataFrame(prob_diff_block_dict)
                difficulty_curve[prob_name] = prob_diff_block_df

            test_structure_name.append(','.join(test_structure_obj_name))
            test_structure_score_rate.append(np.round(np.mean(obj_analysis["得分率"]), 2))
            test_structure_ave.append(np.round(self.obj_data['score'].sum() / stu_num, 2))
            students = []
            sum_of_score = []
            [students.append(x) for x in self.obj_data['xh'] if x not in students]
            for student in students:
                stu_grouped = self.obj_data.groupby(self.obj_data.xh).get_group(student)
                sum_of_score.append(stu_grouped['score'].sum())
            test_structure_std.append(np.round(np.std(sum_of_score), 2))

        for key in all_prob_analysis.keys():
            all_prob_analysis[key] = sub_analysis[key] + obj_analysis[key]
        prob_analysis_df = pd.DataFrame(all_prob_analysis)

        output_dict = {"2-1-3-1各题目指标": prob_analysis_df, "2-1-3-9全卷及各题难度曲线图": difficulty_curve}

        summary = self.summary_analysis(test_difficulty, all_prob_analysis["区分度"],
                                        subject_total_score, subject_full_score, index)
        output_dict["2-1-1试卷概况"] = summary

        for index in range(5):
            rate = diff_total_score[index] / subject_full_score
            diff_rate[index] = np.round(rate*100, 2)
        diff_distribution = self.diff_analysis(diff_sub_num, diff_sub_score, diff_sub_name,
                                               diff_obj_num, diff_obj_score, diff_obj_name,
                                               diff_total_num, diff_total_score, diff_rate)
        output_dict["2-1-3-2试题难度分布"] = diff_distribution

        for index in range(4):
            rate = dist_total_score[index] / subject_full_score
            dist_rate[index] = np.round(rate*100, 2)
        dist_distribution = self.dist_analysis(dist_sub_num, dist_sub_score, dist_sub_name,
                                               dist_obj_num, dist_obj_score, dist_obj_name,
                                               dist_total_num, dist_total_score, dist_rate)
        output_dict["2-1-3-3试题区分度分布"] = dist_distribution

        test_structure = self.structure(test_structure_title, test_structure_score,
                                        test_structure_num, test_structure_name,
                                        test_structure_ave, test_structure_std,
                                        test_structure_score_rate)
        output_dict["2-1-2-4题型分析"] = test_structure

        subject_score_analysis = self.subject_score_analysis(school_name, subject_total_score, subject_full_score, index)
        output_dict["2-1-4-1各学校得分情况"] = subject_score_analysis

        if index == 0:
            subject = '分数'
        else:
            subject = '分数.' + str(index)
        student_total_subject = pd.concat([student_total.iloc[:, 1:7], student_total[subject]], axis=1)
        student_total_subject = student_total_subject.sort_values(by=subject, ascending=False)
        student_total_subject['排名'] = student_total_subject[subject].rank(method='min', ascending=False, )

        first_ten = self.first_ten(student_total_subject)
        output_dict["科目前十名"] = first_ten
        last_ten = self.last_ten(student_total_subject)
        output_dict["科目后十名"] = last_ten

        score_segment_analysis = self.score_segment_analysis(subject_full_score,
                                                             subject_total_score, school_name, subject)
        output_dict["2-1-4-3各分数段人数统计"] = score_segment_analysis

        ranked_score_segment_analysis = self.ranked_score_segment_analysis(school_name, subject_total_score, subject)
        output_dict["2-1-4-5全区前N名各校分布"] = ranked_score_segment_analysis

        return output_dict

    def summary_analysis(self, test_difficulty, prob_analysis_distinction, subject_total_score,
                         subject_full_score, index):
        test_std = self.subject_total['全卷'].std()
        test_std = np.round(test_std, 2)
        test_ave = self.subject_total['全卷'].mean()
        test_ave = np.round(test_ave, 2)
        distinction = np.round(np.mean(prob_analysis_distinction), 2)
        test_score_rate = np.round(test_ave / subject_full_score * 100, 2)

        corr_list = subject_total_score.corr()['分数.7'].tolist()

        summary_dict = {"科目": [self.subject_name], "分值": [subject_full_score], "预估难度": [np.nan],
                        "实测难度": [test_difficulty], "区分度": [distinction], "平均分": [test_ave], "标准差": [test_std],
                        "得分率": test_score_rate, "单科与总分相关系数": np.round(corr_list[index], 4)}

        summary = pd.DataFrame(summary_dict)
        return summary

    def diff_analysis(self,
                      diff_sub_num, diff_sub_score, diff_sub_name,
                      diff_obj_num, diff_obj_score, diff_obj_name,
                      diff_total_num, diff_total_score, diff_rate):

        diff_dict_score = {"难度范围": self.diff_range, "难度评价": self.diff_evaluate,
                           "客观题题数": diff_sub_num, "客观题分值": diff_sub_score,
                           "主观题题数": diff_obj_num, "主观题分值": diff_obj_score,
                           "全卷题数": diff_total_num, "全卷分值": diff_total_score}
        diff_df_score = pd.DataFrame(diff_dict_score)

        diff_sub_name = [','.join(name) for name in diff_sub_name]
        diff_obj_name = [','.join(name) for name in diff_obj_name]
        diff_dict_name = {"难度范围": self.diff_range, "难度评价": self.diff_evaluate,
                          "客观题题数": diff_sub_num, "客观题题号": diff_sub_name,
                          "主观题题数": diff_obj_num, "主观题题号": diff_obj_name,
                          "总题数": diff_total_num, "全卷分值": diff_total_score, "比例": diff_rate}
        diff_df_name = pd.DataFrame(diff_dict_name)

        diff_analysis = {"难度题型分布": diff_df_score, "难度分布": diff_df_name}
        return diff_analysis

    def dist_analysis(self,
                      dist_sub_num, dist_sub_score, dist_sub_name,
                      dist_obj_num, dist_obj_score, dist_obj_name,
                      dist_total_num, dist_total_score, dist_rate):
        dist_dict_score = {"区分度范围": self.dist_range, "区分度评价": self.dist_evaluate,
                           "客观题题数": dist_sub_num, "客观题分值": dist_sub_score,
                           "主观题题数": dist_obj_num, "主观题分值": dist_obj_score,
                           "全卷题数": dist_total_num, "全卷分值": dist_total_score}
        dist_df_score = pd.DataFrame(dist_dict_score)

        dist_sub_name = [','.join(name) for name in dist_sub_name]
        dist_obj_name = [','.join(name) for name in dist_obj_name]
        dist_dict_name = {"区分度范围": self.dist_range, "区分度评价": self.dist_evaluate,
                          "客观题题数": dist_sub_num, "客观题题号": dist_sub_name,
                          "主观题题数": dist_obj_num, "主观题题号": dist_obj_name,
                          "总题数": dist_total_num, "全卷分值": dist_total_score, "比例": dist_rate}
        dist_df_name = pd.DataFrame(dist_dict_name)

        dist_analysis = {"区分度题型分布": dist_df_score, "区分度分布": dist_df_name}
        return dist_analysis

    def structure(self,
                  test_structure_title, test_structure_score,
                  test_structure_num, test_structure_name,
                  test_structure_ave, test_structure_std,
                  test_structure_score_rate):

        structure_dict = {"题型": test_structure_title, "分值": test_structure_score,
                          "题目数量": test_structure_num, "对应题目": test_structure_name,
                          "平均分": test_structure_ave, "标准差": test_structure_std,
                          "得分率": test_structure_score_rate}
        structure_df = pd.DataFrame(structure_dict)
        return structure_df

    def subject_score_analysis(self, school_name, subject_total_score, subject_full_score, index):
        subject_score_analysis = {"学校": school_name, "总人数": [], "最高分": [], "最低分": [],
                                  "全体平均分": [], "高分段平均分": [], "低分段平均分": [],
                                  "标准差": [], "满分率": [], "超优率": [], "优秀率": [],
                                  "良好率": [], "及格率": [], "低分率": [], "超均率": [],
                                  "比均率": [], "难度": [], "得分率": [], "单科与总分的相关系数": []}
        for school in school_name:
            subject_list = list(subject_total_score.columns)
            this_subject = subject_list[index + 1]
            total_ave = subject_total_score[this_subject].mean()
            if school == "全体":
                this_school = subject_total_score[[this_subject, '分数.7']]
                this_school_score = this_school
            else:
                this_school = subject_total_score.groupby("学校").get_group(school)
                this_school_score = this_school[[this_subject, '分数.7']]

            score_sorted = this_school[this_subject].sort_values(ascending=False)
            school_boundary = int(np.round(len(score_sorted) * 0.27, 0))
            high_group = score_sorted.head(school_boundary)
            low_group = score_sorted.tail(school_boundary)

            high_ave = high_group.mean()
            subject_score_analysis["高分段平均分"].append(np.round(high_ave, 2))
            low_ave = low_group.mean()
            subject_score_analysis["低分段平均分"].append(np.round(low_ave, 2))

            correlation_list = this_school_score.corr()
            correlation = np.round(correlation_list[this_subject][1], 4)
            subject_score_analysis["单科与总分的相关系数"].append(correlation)

            school_stu_num = len(score_sorted)
            subject_score_analysis["总人数"].append(school_stu_num)

            school_max = score_sorted.max()
            subject_score_analysis["最高分"].append(np.round(school_max, 2))

            school_min = score_sorted.min()
            subject_score_analysis["最低分"].append(np.round(school_min, 2))

            school_ave = score_sorted.mean()
            subject_score_analysis["全体平均分"].append(np.round(school_ave, 2))

            school_std = score_sorted.std()
            subject_score_analysis["标准差"].append(np.round(school_std, 2))

            rate_list = ["满分率", "超优率", "优秀率", "良好率", "及格率"]
            for interval in range(10, 5, -1):
                rate_coefficient = interval / 10
                rate_index = 10 - interval
                if interval == 10:
                    rate_num = len(score_sorted.where(score_sorted == subject_full_score * rate_coefficient).dropna())
                else:
                    rate_num = len(score_sorted.where(score_sorted >= subject_full_score * rate_coefficient).dropna())
                rate = np.round(rate_num / school_stu_num * 100, 2)
                rate_name = rate_list[rate_index]
                subject_score_analysis[rate_name].append(rate)

            rate_num = len(score_sorted.where(score_sorted < subject_full_score * 0.4).dropna())
            rate = np.round(rate_num / school_stu_num * 100, 2)
            subject_score_analysis["低分率"].append(rate)

            if school == "全体":
                subject_score_analysis["比均率"].append(np.nan)
                subject_score_analysis["超均率"].append(np.nan)
            else:
                on_ave_rate = school_ave / total_ave * 100
                over_ave_rate = on_ave_rate - 100
                subject_score_analysis["比均率"].append(np.round(on_ave_rate, 2))
                subject_score_analysis["超均率"].append(np.round(over_ave_rate, 2))

            school_difficulty = school_ave / subject_full_score
            subject_score_analysis["难度"].append(np.round(school_difficulty, 2))

            school_score_rate = school_difficulty * 100
            subject_score_analysis["得分率"].append(np.round(school_score_rate, 2))

        subject_score_analysis_df = pd.DataFrame(subject_score_analysis)
        return subject_score_analysis_df

    def first_ten(self, student_total_subject):
        first_ten_rank_list = []
        for rank in range(1, 11):
            if rank in student_total_subject['排名'].values:
                first_ten_rank_list.append(student_total_subject.groupby('排名').get_group(rank))
        first_ten = pd.concat(first_ten_rank_list, ignore_index=True)
        return first_ten

    def last_ten(self, student_total_subject):
        last_ten_rank_list = []
        rank_last = student_total_subject['排名'].max()
        stu_rank_last = student_total_subject.groupby('排名').get_group(rank_last)
        stu_num_rank_last = len(stu_rank_last)

        if stu_num_rank_last > 10:
            last_ten_rank_list.append(stu_rank_last)
        else:
            last_ten_rank_list.append(stu_rank_last)
            remaining = 10 - stu_num_rank_last
            while remaining > 0:
                rank_last = rank_last - 1
                if rank_last in student_total_subject['排名'].values:
                    stu_rank_last = student_total_subject.groupby('排名').get_group(rank_last)
                    stu_num_rank_last = len(stu_rank_last)
                    last_ten_rank_list.append(stu_rank_last)
                    remaining = remaining - stu_num_rank_last
        last_ten = pd.concat(last_ten_rank_list, ignore_index=True)
        last_ten = last_ten.iloc[::-1]
        return last_ten

    def score_segment_analysis(self, subject_full_score, subject_total_score, school_name, subject):
        score_segment = []
        for interval in range(0, subject_full_score, int(subject_full_score * 0.1)):
            score_segment.append("[{}, {})".format(str(interval), str(int(interval + subject_full_score * 0.1))))
        score_segment.append("[{}, *]".format(subject_full_score))
        score_segment.reverse()
        score_segment.append("人数")

        segment_dict = {"分数段": score_segment}
        for school in school_name[:-1]:
            this_school_data = []
            this_school = subject_total_score.groupby("学校").get_group(school)

            for interval in range(0, subject_full_score, int(subject_full_score * 0.1)):
                high_group = this_school.where(this_school[subject] >= interval).dropna()
                low_group = high_group.where(high_group[subject] < interval + subject_full_score * 0.1).dropna()
                if len(low_group) == 0:
                    this_school_data.append(np.nan)
                else:
                    this_school_data.append(len(low_group))
            high_group = this_school.where(this_school[subject] >= subject_full_score).dropna()
            if len(high_group) == 0:
                this_school_data.append(np.nan)
            else:
                this_school_data.append(len(high_group))
            this_school_data.reverse()
            this_school_data.append(len(this_school))
            segment_dict[school] = this_school_data
        segment_df = pd.DataFrame(segment_dict)
        return segment_df

    def ranked_score_segment_analysis(self, school_name, subject_total_score, subject):
        rank_segment = ["[*, 10]", "[*, 30]", "[*, 50]", "[*, 100]", "[*, 200]", "[*, 300]", "[*, 500]", "[*, 1000]",
                        "人数", "最高分", "最低分"]
        segments = [10, 30, 50, 100, 200, 300, 500, 1000]
        this_subject = subject_total_score[['学校', subject]].sort_values(by=subject, ascending=False)

        ranked_dict = {"名次段": rank_segment}

        for school in school_name[:-1]:
            ranked_dict[school] = []
            for segment in segments:
                student = this_subject[0:segment]
                if school in student["学校"].values:
                    this_school = student.groupby("学校").get_group(school)
                    ranked_dict[school].append(len(this_school))
                else:
                    ranked_dict[school].append(np.nan)
            this_school = subject_total_score.groupby("学校").get_group(school)
            ranked_dict[school].append(len(this_school))
            ranked_dict[school].append(this_school[subject].max())
            ranked_dict[school].append(this_school[subject].min())

        ranked_df = pd.DataFrame(ranked_dict)
        return ranked_df

