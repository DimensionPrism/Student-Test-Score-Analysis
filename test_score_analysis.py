import os
import pandas as pd
import numpy as np

# 读取小题分
ITEM_SCORE = pd.read_excel("tb_cj_itemscore.xlsx")

# 读取各科总分和小题满分
CHIN_TOTAL = pd.read_excel("小题分(语文).xls")
CHIN_PROB_TOTAL = pd.read_excel("小题满分.xlsx", sheet_name="语文")

MATH_TOTAL = pd.read_excel("小题分(数学).xls")
MATH_PROB_TOTAL = pd.read_excel("小题满分.xlsx", sheet_name="数学")

ENGL_TOTAL = pd.read_excel("小题分(英语全卷).xls")
ENGL_PROB_TOTAL = pd.read_excel("小题满分.xlsx", sheet_name="英语")

HIST_TOTAL = pd.read_excel("小题分(历史).xls")
HIST_PROB_TOTAL = pd.read_excel("小题满分.xlsx", sheet_name="历史")

PHYS_TOTAL = pd.read_excel("小题分(物理).xls")
PHYS_PROB_TOTAL = pd.read_excel("小题满分.xlsx", sheet_name="物理")

CHEM_TOTAL = pd.read_excel("小题分(化学).xls")
CHEM_PROB_TOTAL = pd.read_excel("小题满分.xlsx", sheet_name="化学")

POLI_TOTAL = pd.read_excel("小题分(道德与法治).xls")
POLI_PROB_TOTAL = pd.read_excel("小题满分.xlsx", sheet_name="道德与法治")

# 读取学生总分
STUDENT_TOTAL = pd.read_excel("报名信息-学生成绩.xls")

SUBJECT_LIST = ["语文", "数学", "英语", "历史", "物理", "化学", "道德与法治", "总分"]
FULL_SCORE = [120, 100, 100, 70, 70, 50, 50, 560]
SUBJECT_TOTAL_LIST = [CHIN_TOTAL, MATH_TOTAL, ENGL_TOTAL, HIST_TOTAL, PHYS_TOTAL, CHEM_TOTAL, POLI_TOTAL]
SUBJECT_PROB_TOTAL = [CHIN_PROB_TOTAL, MATH_PROB_TOTAL, ENGL_PROB_TOTAL, HIST_PROB_TOTAL,
                      PHYS_PROB_TOTAL, CHEM_PROB_TOTAL, POLI_PROB_TOTAL]


TOTAL_SCORE = pd.concat([STUDENT_TOTAL[['学校', '分数']], STUDENT_TOTAL['分数.4'],
                         STUDENT_TOTAL['分数.1'], STUDENT_TOTAL[ '分数.2'],
                         STUDENT_TOTAL['分数.5'], STUDENT_TOTAL['分数.6'],
                         STUDENT_TOTAL['分数.3'], STUDENT_TOTAL['分数.7']], axis=1)


class Exam:
    def __init__(self, item_score, total, prob_total, name, test_score):
        self.stu_arr = []
        self.prob_id = []
        self.prob_max = []
        self.prob_min = []
        self.prob_ave = []
        self.prob_std = []
        self.prob_scr = []
        self.prob_fsr = []
        self.prob_nsr = []
        self.prob_dif = []
        self.prob_dist = []
        self.prob_type = []

        self.item_score = item_score.groupby(item_score.sjmc).get_group(name)
        self.name = name
        self.total = total

        self.prob_diff_curve = {}
        self.test_score = test_score
        b1 = '0-'+str(int(self.test_score*0.1))
        b2 = str(int(self.test_score*0.1+1))+'-'+str(int(self.test_score*0.2))
        b3 = str(int(self.test_score*0.2+1))+'-'+str(int(self.test_score*0.3))
        b4 = str(int(self.test_score*0.3+1))+'-'+str(int(self.test_score*0.4))
        b5 = str(int(self.test_score*0.4+1))+'-'+str(int(self.test_score*0.5))
        b6 = str(int(self.test_score*0.5+1))+'-'+str(int(self.test_score*0.6))
        b7 = str(int(self.test_score*0.6+1))+'-'+str(int(self.test_score*0.7))
        b8 = str(int(self.test_score*0.7+1))+'-'+str(int(self.test_score*0.8))
        b9 = str(int(self.test_score*0.8+1))+'-'+str(int(self.test_score*0.9))
        b10 = str(int(self.test_score*0.9+1))+'-'+str(int(self.test_score))
        self.curve_list = [b1, b2, b3, b4, b5, b6, b7, b8, b9, b10]

        if "客观题" in self.item_score.values:
            self.sub_prob_score = prob_total.groupby(prob_total.sskm_dt).get_group("客观题")
            self.sub_data = self.item_score.groupby(self.item_score.sskm_dt).get_group("客观题")
            self.sub_data = self.sub_data[["xh", "sjmc", "sj_tid", "sskm_dt", "sskm_th", "score", "AnswerStr"]]
            self.total_sub = self.sub_data['sj_tid'].max()
            self.sub_tid_grouped = self.sub_data.groupby(self.sub_data.sj_tid)
        else:
            self.total_sub = 0

        if "主观题" in self.item_score.values:
            self.obj_prob_score = prob_total.groupby(prob_total.sskm_dt).get_group("主观题")
            self.obj_data = self.item_score.groupby(self.item_score.sskm_dt).get_group("主观题")
            self.obj_data = self.obj_data[["xh", "sjmc", "sj_tid", "sskm_dt", "sskm_th", "score", "AnswerStr"]]
            self.obj_tid_grouped = self.obj_data.groupby(self.obj_data.sj_tid)

        self.total_prob = self.item_score['sj_tid'].max()
        self.total_stu = len(self.item_score.index) / self.total_prob

        self.difficulty = np.round(total['全卷'].mean() / 100, 2)

        self.prob_name = []
        for x in total.columns:
            self.prob_name.append(x)
        del self.prob_name[0:8]

        subtract_list = []
        if name == "英语" or name == "历史" or name == "道德与法治":
            for x in range(0, 2 * self.total_sub):
                if x % 2 != 0:
                    subtract_list.append(self.prob_name[x])
        self.prob_name = [item for item in self.prob_name if item not in subtract_list]
        self.prob_name_id = self.prob_name.copy()

        self.total_sorted = total.sort_values(by=['全卷'], ascending=False)
        self.total_sorted_reverse = total.sort_values(by=['全卷'])
        self.boundary = int(np.round(len(self.total_sorted.index) * 0.27))
        self.left = self.total_sorted.head(self.boundary)
        self.right = self.total_sorted.tail(self.boundary)

        self.diff_range = ["[0, 0.2)", "[0.2, 0.4)", "[0.4, 0.6)", "[0.6, 0.8)", "[0.8, 1.0]"]
        self.diff_evaluate = ["难", "较难", "中等", "较易", "易"]
        self.dist_range = ["[0.4, *)", "[0.3, 0.4)", "[0.2, 0.3)", "[0, 0.2"]
        self.dist_evaluate = ["优秀", "良好", "尚可", "差"]

        self.diff1_num_sub = 0
        self.diff1_total_sub = 0
        self.diff2_num_sub = 0
        self.diff2_total_sub = 0
        self.diff3_num_sub = 0
        self.diff3_total_sub = 0
        self.diff4_num_sub = 0
        self.diff4_total_sub = 0
        self.diff5_num_sub = 0
        self.diff5_total_sub = 0
        self.diff1_list_sub = []
        self.diff2_list_sub = []
        self.diff3_list_sub = []
        self.diff4_list_sub = []
        self.diff5_list_sub = []

        self.dist1_num_sub = 0
        self.dist1_total_sub = 0
        self.dist2_num_sub = 0
        self.dist2_total_sub = 0
        self.dist3_num_sub = 0
        self.dist3_total_sub = 0
        self.dist4_num_sub = 0
        self.dist4_total_sub = 0
        self.dist1_list_sub = []
        self.dist2_list_sub = []
        self.dist3_list_sub = []
        self.dist4_list_sub = []

        self.diff1_num_obj = 0
        self.diff1_total_obj = 0
        self.diff2_num_obj = 0
        self.diff2_total_obj = 0
        self.diff3_num_obj = 0
        self.diff3_total_obj = 0
        self.diff4_num_obj = 0
        self.diff4_total_obj = 0
        self.diff5_num_obj = 0
        self.diff5_total_obj = 0
        self.diff1_list_obj = []
        self.diff2_list_obj = []
        self.diff3_list_obj = []
        self.diff4_list_obj = []
        self.diff5_list_obj = []

        self.dist1_num_obj = 0
        self.dist1_total_obj = 0
        self.dist2_num_obj = 0
        self.dist2_total_obj = 0
        self.dist3_num_obj = 0
        self.dist3_total_obj = 0
        self.dist4_num_obj = 0
        self.dist4_total_obj = 0
        self.dist1_list_obj = []
        self.dist2_list_obj = []
        self.dist3_list_obj = []
        self.dist4_list_obj = []

    def total_difficulty_curve(self):
        diff_list = []
        highest = self.test_score
        total_sorted_reverse = self.total_sorted_reverse[['全卷']]

        block_one = total_sorted_reverse.where(total_sorted_reverse['全卷'] <= self.test_score * 0.1).dropna()
        b1_diff = np.round(block_one['全卷'].mean() / highest, 2)
        diff_list.append(b1_diff)

        block_two = total_sorted_reverse.where(total_sorted_reverse['全卷'] <= self.test_score * 0.2).dropna()
        block_two = block_two.where(block_two['全卷'] > self.test_score * 0.1 + 1).dropna()
        b2_diff = np.round(block_two['全卷'].mean() / highest, 2)
        diff_list.append(b2_diff)

        block_three = total_sorted_reverse.where(total_sorted_reverse['全卷'] <= self.test_score * 0.3).dropna()
        block_three = block_three.where(block_three['全卷'] > self.test_score * 0.2 + 1).dropna()
        b3_diff = np.round(block_three['全卷'].mean() / highest, 2)
        diff_list.append(b3_diff)

        block_four = total_sorted_reverse.where(total_sorted_reverse['全卷'] <= self.test_score * 0.4).dropna()
        block_four = block_four.where(block_four['全卷'] > self.test_score * 0.3 + 1).dropna()
        b4_diff = np.round(block_four['全卷'].mean() / highest, 2)
        diff_list.append(b4_diff)

        block_five = total_sorted_reverse.where(total_sorted_reverse['全卷'] <= self.test_score * 0.5).dropna()
        block_five = block_five.where(block_five['全卷'] > self.test_score * 0.4 + 1).dropna()
        b5_diff = np.round(block_five['全卷'].mean() / highest, 2)
        diff_list.append(b5_diff)

        block_six = total_sorted_reverse.where(total_sorted_reverse['全卷'] <= self.test_score * 0.6).dropna()
        block_six = block_six.where(block_six['全卷'] > self.test_score * 0.5 + 1).dropna()
        b6_diff = np.round(block_six['全卷'].mean() / highest, 2)
        diff_list.append(b6_diff)

        block_seven = total_sorted_reverse.where(total_sorted_reverse['全卷'] <= self.test_score * 0.7).dropna()
        block_seven = block_seven.where(block_seven['全卷'] > self.test_score * 0.6 + 1).dropna()
        b7_diff = np.round(block_seven['全卷'].mean() / highest, 2)
        diff_list.append(b7_diff)

        block_eight = total_sorted_reverse.where(total_sorted_reverse['全卷'] <= self.test_score * 0.8).dropna()
        block_eight = block_eight.where(block_eight['全卷'] > self.test_score * 0.7 + 1).dropna()
        b8_diff = np.round(block_eight['全卷'].mean() / highest, 2)
        diff_list.append(b8_diff)

        block_nine = total_sorted_reverse.where(total_sorted_reverse['全卷'] <= self.test_score * 0.9).dropna()
        block_nine = block_nine.where(block_nine['全卷'] > self.test_score * 0.8 + 1).dropna()
        b9_diff = np.round(block_nine['全卷'].mean() / highest, 2)
        diff_list.append(b9_diff)

        block_ten = total_sorted_reverse.where(total_sorted_reverse['全卷'] <= self.test_score).dropna()
        block_ten = block_ten.where(block_ten['全卷'] > self.test_score * 0.9 + 1).dropna()
        b10_diff = np.round(block_ten['全卷'].mean() / highest, 2)
        diff_list.append(b10_diff)

        dict = {'分数段': self.curve_list, '难度值': diff_list}
        df = pd.DataFrame(dict)
        self.prob_diff_curve['全卷'] = df

    def calc_sub(self):
        left = self.left
        right = self.right
        sub_arr = []
        sub_prob_score = self.sub_prob_score.groupby(self.sub_prob_score.sj_tid)
        for x in range(1, self.total_sub + 1):
            sub_arr.append(x)
            self.stu_arr.append(self.total_stu)
            self.prob_type.append("客观题")

        for x in sub_arr:
            prob_name = self.prob_name[0]
            sub = self.sub_tid_grouped.get_group(x)

            # 小题满分
            highest = sub_prob_score.get_group(x)
            highest = int(highest['score'].max())

            # 最高分与最低分计算
            sub_max = int(sub['score'].max())
            self.prob_max.append(sub_max)
            sub_min = int(sub['score'].min())
            self.prob_min.append(sub_min)

            # 平均分计算
            total = sub['score'].sum()
            ave = total / self.total_stu
            ave = np.round(ave, 2)
            self.prob_ave.append(ave)

            # 标准差计算
            std = np.round(sub['score'].std(), 2)
            self.prob_std.append(std)

            # 得分率
            score_rate = np.round((total / self.total_stu) / highest * 100, 1)
            self.prob_scr.append(score_rate)

            # 满分率
            if highest in sub['score'].values:
                max_group = sub.groupby(sub.score).get_group(highest)
                full_score_stu = len(max_group.index)
                full_score_rate = np.round(full_score_stu / self.total_stu * 100, 1)
            else:
                full_score_rate = 0
            self.prob_fsr.append(full_score_rate)

            # 零分率
            if 0 in sub['score'].values:
                zero_group = sub.groupby(sub.score).get_group(0)
                zero_score_stu = len(zero_group.index)
                zero_score_rate = np.round(zero_score_stu / self.total_stu * 100, 1)
            else:
                zero_score_rate = 0
            self.prob_nsr.append(zero_score_rate)

            # 难度计算
            sub_difficulty = np.round((total / self.total_stu) / sub_max, 2)
            self.prob_dif.append(sub_difficulty)
            if 0 <= sub_difficulty < 0.2:
                self.diff1_num_sub += 1
                self.diff1_total_sub += highest
                self.diff1_list_sub.append(prob_name)
            elif 0.2 <= sub_difficulty < 0.4:
                self.diff2_num_sub += 1
                self.diff2_total_sub += highest
                self.diff2_list_sub.append(prob_name)
            elif 0.4 <= sub_difficulty < 0.6:
                self.diff3_num_sub += 1
                self.diff3_total_sub += highest
                self.diff3_list_sub.append(prob_name)
            elif 0.6 <= sub_difficulty < 0.8:
                self.diff4_num_sub += 1
                self.diff4_total_sub += highest
                self.diff4_list_sub.append(prob_name)
            elif 0.8 <= sub_difficulty < 1:
                self.diff5_num_sub += 1
                self.diff5_total_sub += highest
                self.diff5_list_sub.append(prob_name)

            # 区分度计算
            if 0 in left[prob_name].values:
                left_zero_rate = int(len(left.groupby(prob_name).get_group(0).index)) / self.boundary
            else:
                left_zero_rate = 0
            left_diff = np.round(1 - left_zero_rate, 2)
            right_zero_rate = int(len(right.groupby(prob_name).get_group(0).index)) / self.boundary
            right_diff = np.round(1 - right_zero_rate, 2)
            distinction = np.round(left_diff - right_diff, 2)
            self.prob_dist.append(distinction)
            if 0.4 <= distinction:
                self.dist1_num_sub += 1
                self.dist1_total_sub += highest
                self.dist1_list_sub.append(prob_name)
            elif 0.3 <= distinction < 0.4:
                self.dist2_num_sub += 1
                self.dist2_total_sub += highest
                self.dist2_list_sub.append(prob_name)
            elif 0.2 <= distinction < 0.3:
                self.dist3_num_sub += 1
                self.dist3_total_sub += highest
                self.dist3_list_sub.append(prob_name)
            elif 0.2 < distinction <= 0:
                self.dist4_num_sub += 1
                self.dist4_total_sub += highest
                self.dist4_list_sub.append(prob_name)

            # 分段难度计算
            diff_list = []
            total_sorted_reverse = self.total_sorted_reverse[['全卷', prob_name]]

            block_one = total_sorted_reverse.where(total_sorted_reverse['全卷'] <= self.test_score * 0.1).dropna()
            b1_diff = np.round((block_one[prob_name].sum() / len(block_one)) / highest, 2)
            diff_list.append(b1_diff)

            block_two = total_sorted_reverse.where(total_sorted_reverse['全卷'] <= self.test_score * 0.2).dropna()
            block_two = block_two.where(block_two['全卷'] > self.test_score * 0.1 + 1).dropna()
            b2_diff = np.round((block_two[prob_name].sum() / len(block_two)) / highest, 2)
            diff_list.append(b2_diff)

            block_three = total_sorted_reverse.where(total_sorted_reverse['全卷'] <= self.test_score * 0.3).dropna()
            block_three = block_three.where(block_three['全卷'] > self.test_score * 0.2 + 1).dropna()
            b3_diff = np.round((block_three[prob_name].sum() / len(block_three)) / highest, 2)
            diff_list.append(b3_diff)

            block_four = total_sorted_reverse.where(total_sorted_reverse['全卷'] <= self.test_score * 0.4).dropna()
            block_four = block_four.where(block_four['全卷'] > self.test_score * 0.3 + 1).dropna()
            b4_diff = np.round((block_four[prob_name].sum() / len(block_four)) / highest, 2)
            diff_list.append(b4_diff)

            block_five = total_sorted_reverse.where(total_sorted_reverse['全卷'] <= self.test_score * 0.5).dropna()
            block_five = block_five.where(block_five['全卷'] > self.test_score * 0.4 + 1).dropna()
            b5_diff = np.round((block_five[prob_name].sum() / len(block_five)) / highest, 2)
            diff_list.append(b5_diff)

            block_six = total_sorted_reverse.where(total_sorted_reverse['全卷'] <= self.test_score * 0.6).dropna()
            block_six = block_six.where(block_six['全卷'] > self.test_score * 0.5 + 1).dropna()
            b6_diff = np.round((block_six[prob_name].sum() / len(block_six)) / highest, 2)
            diff_list.append(b6_diff)

            block_seven = total_sorted_reverse.where(total_sorted_reverse['全卷'] <= self.test_score * 0.7).dropna()
            block_seven = block_seven.where(block_seven['全卷'] > self.test_score * 0.6 + 1).dropna()
            b7_diff = np.round((block_seven[prob_name].sum() / len(block_seven)) / highest, 2)
            diff_list.append(b7_diff)

            block_eight = total_sorted_reverse.where(total_sorted_reverse['全卷'] <= self.test_score * 0.8).dropna()
            block_eight = block_eight.where(block_eight['全卷'] > self.test_score * 0.7 + 1).dropna()
            b8_diff = np.round((block_eight[prob_name].sum() / len(block_eight)) / highest, 2)
            diff_list.append(b8_diff)

            block_nine = total_sorted_reverse.where(total_sorted_reverse['全卷'] <= self.test_score * 0.9).dropna()
            block_nine = block_nine.where(block_nine['全卷'] > self.test_score * 0.8 + 1).dropna()
            b9_diff = np.round((block_nine[prob_name].sum() / len(block_nine)) / highest, 2)
            diff_list.append(b9_diff)

            block_ten = total_sorted_reverse.where(total_sorted_reverse['全卷'] <= self.test_score).dropna()
            block_ten = block_ten.where(block_ten['全卷'] > self.test_score * 0.9 + 1).dropna()
            b10_diff = np.round((block_ten[prob_name].sum() / len(block_ten)) / highest, 2)
            diff_list.append(b10_diff)

            dict = {'分数段': self.curve_list, '难度值': diff_list}
            df = pd.DataFrame(dict)
            self.prob_diff_curve[str(x)] = df

            del self.prob_name[0]

        self.prob_id += sub_arr

    def calc_obj(self):
        left = self.left
        right = self.right
        obj_arr = []
        obj_prob_score = self.obj_prob_score.groupby(self.obj_prob_score.sj_tid)
        for tid in range(self.total_sub + 1, self.total_prob + 1):
            obj_arr.append(tid)
            self.stu_arr.append(self.total_stu)
            self.prob_type.append("主观题")

        for prob in obj_arr:
            prob_name = self.prob_name[0]
            obj = self.obj_tid_grouped.get_group(prob)

            # 小题满分
            highest = obj_prob_score.get_group(prob)
            highest = int(highest['score'].max())

            # 最高分与最低分计算
            obj_max = int(obj['score'].max())
            self.prob_max.append(obj_max)
            obj_min = int(obj['score'].min())
            self.prob_min.append(obj_min)

            # 平均分计算
            total = obj['score'].sum()
            ave = total / self.total_stu
            ave = np.round(ave, 2)
            self.prob_ave.append(ave)

            # 标准差计算
            std = np.round(obj['score'].std(), 2)
            self.prob_std.append(std)

            # 得分率计算
            score_rate = np.round((total / self.total_stu) / highest * 100, 1)
            self.prob_scr.append(score_rate)

            # 满分率
            if highest in obj['score'].values:
                max_group = obj.groupby(obj.score).get_group(highest)
                full_score_stu = len(max_group.index)
                full_score_rate = np.round(full_score_stu / self.total_stu * 100, 1)
            else:
                full_score_rate = 0
            self.prob_fsr.append(full_score_rate)

            # 零分率
            if 0 in obj['score'].values:
                zero_group = obj.groupby(obj.score).get_group(0)
                zero_score_stu = len(zero_group.index)
                zero_score_rate = np.round(zero_score_stu / self.total_stu * 100, 1)
            else:
                zero_score_rate = 0
            self.prob_nsr.append(zero_score_rate)

            # 难度计算
            obj_difficulty = np.round((total / self.total_stu) / obj_max, 2)
            self.prob_dif.append(obj_difficulty)
            if 0 <= obj_difficulty < 0.2:
                self.diff1_num_obj += 1
                self.diff1_total_obj += highest
                self.diff1_list_obj.append(prob_name)
            elif 0.2 <= obj_difficulty < 0.4:
                self.diff2_num_obj += 1
                self.diff2_total_obj += highest
                self.diff2_list_obj.append(prob_name)
            elif 0.4 <= obj_difficulty < 0.6:
                self.diff3_num_obj += 1
                self.diff3_total_obj += highest
                self.diff3_list_obj.append(prob_name)
            elif 0.6 <= obj_difficulty < 0.8:
                self.diff4_num_obj += 1
                self.diff4_total_obj += highest
                self.diff4_list_obj.append(prob_name)
            elif 0.8 <= obj_difficulty < 1:
                self.diff5_num_obj += 1
                self.diff5_total_obj += highest
                self.diff5_list_obj.append(prob_name)

            # 区分度计算
            left_total = left[prob_name].sum()
            right_total = right[prob_name].sum()
            prob_max = self.total_sorted[prob_name].max()
            prob_min = self.total_sorted[prob_name].min()
            distinction = (left_total - right_total) / (self.boundary * (prob_max - prob_min))
            distinction = np.round(distinction, 2)
            self.prob_dist.append(distinction)
            if 0.4 <= distinction:
                self.dist1_num_obj += 1
                self.dist1_total_obj += highest
                self.dist1_list_obj.append(prob_name)
            elif 0.3 <= distinction < 0.4:
                self.dist2_num_obj += 1
                self.dist2_total_obj += highest
                self.dist2_list_obj.append(prob_name)
            elif 0.2 <= distinction < 0.3:
                self.dist3_num_obj += 1
                self.dist3_total_obj += highest
                self.dist3_list_obj.append(prob_name)
            elif 0.2 < distinction <= 0:
                self.dist4_num_obj += 1
                self.dist4_total_obj += highest
                self.dist4_list_obj.append(prob_name)

            # 小题分段难度计算
            diff_list = []
            total_sorted_reverse = self.total_sorted_reverse[['全卷', prob_name]]

            block_one = total_sorted_reverse.where(total_sorted_reverse['全卷'] <= self.test_score * 0.1).dropna()
            b1_diff = np.round(block_one[prob_name].mean() / highest, 2)
            diff_list.append(b1_diff)

            block_two = total_sorted_reverse.where(total_sorted_reverse['全卷'] <= self.test_score * 0.2).dropna()
            block_two = block_two.where(block_two['全卷'] > self.test_score * 0.1 + 1).dropna()
            b2_diff = np.round(block_two[prob_name].mean() / highest, 2)
            diff_list.append(b2_diff)

            block_three = total_sorted_reverse.where(total_sorted_reverse['全卷'] <= self.test_score * 0.3).dropna()
            block_three = block_three.where(block_three['全卷'] > self.test_score * 0.2 + 1).dropna()
            b3_diff = np.round(block_three[prob_name].mean() / highest, 2)
            diff_list.append(b3_diff)

            block_four = total_sorted_reverse.where(total_sorted_reverse['全卷'] <= self.test_score * 0.4).dropna()
            block_four = block_four.where(block_four['全卷'] > self.test_score * 0.3 + 1).dropna()
            b4_diff = np.round(block_four[prob_name].mean() / highest, 2)
            diff_list.append(b4_diff)

            block_five = total_sorted_reverse.where(total_sorted_reverse['全卷'] <= self.test_score * 0.5).dropna()
            block_five = block_five.where(block_five['全卷'] > self.test_score * 0.4 + 1).dropna()
            b5_diff = np.round(block_five[prob_name].mean() / highest, 2)
            diff_list.append(b5_diff)

            block_six = total_sorted_reverse.where(total_sorted_reverse['全卷'] <= self.test_score * 0.6).dropna()
            block_six = block_six.where(block_six['全卷'] > self.test_score * 0.5 + 1).dropna()
            b6_diff = np.round(block_six[prob_name].mean() / highest, 2)
            diff_list.append(b6_diff)

            block_seven = total_sorted_reverse.where(total_sorted_reverse['全卷'] <= self.test_score * 0.7).dropna()
            block_seven = block_seven.where(block_seven['全卷'] > self.test_score * 0.6 + 1).dropna()
            b7_diff = np.round(block_seven[prob_name].mean() / highest, 2)
            diff_list.append(b7_diff)

            block_eight = total_sorted_reverse.where(total_sorted_reverse['全卷'] <= self.test_score * 0.8).dropna()
            block_eight = block_eight.where(block_eight['全卷'] > self.test_score * 0.7 + 1).dropna()
            b8_diff = np.round(block_eight[prob_name].mean() / highest, 2)
            diff_list.append(b8_diff)

            block_nine = total_sorted_reverse.where(total_sorted_reverse['全卷'] <= self.test_score * 0.9).dropna()
            block_nine = block_nine.where(block_nine['全卷'] > self.test_score * 0.8 + 1).dropna()
            b9_diff = np.round(block_nine[prob_name].mean() / highest, 2)
            diff_list.append(b9_diff)

            block_ten = total_sorted_reverse.where(total_sorted_reverse['全卷'] <= self.test_score).dropna()
            block_ten = block_ten.where(block_ten['全卷'] > self.test_score * 0.9 + 1).dropna()
            b10_diff = np.round(block_ten[prob_name].mean() / highest, 2)
            diff_list.append(b10_diff)

            dict = {'分数段': self.curve_list, '难度值': diff_list}
            df = pd.DataFrame(dict)
            self.prob_diff_curve[str(prob)] = df

            del self.prob_name[0]

        self.prob_id += obj_arr

    def to_dataframe(self):
        output_dict = {"题号": self.prob_id, "题型": self.prob_type, "人数": self.total_stu, "最高分": self.prob_max,
                       "最低分": self.prob_min, "平均分": self.prob_ave, "标准差": self.prob_std, "得分率": self.prob_scr,
                       "满分率": self.prob_fsr, "零分率": self.prob_nsr, "难度": self.prob_dif, "区分度": self.prob_dist}
        output_df = pd.DataFrame(output_dict)
        return output_df

    def summary(self):
        std = self.total['全卷'].std()
        std = np.round(std, 2)
        ave_score = self.total['全卷'].mean()
        ave_score = np.round(ave_score, 2)
        difference = np.round(np.mean(self.prob_dist), 2)
        summary_dict = {}

        corr_list = TOTAL_SCORE.corr()['分数.7'].tolist()

        if self.name == "语文":
            summary_dict = {"科目": [self.name], "分值": ["120"], "预估难度": [np.nan], "实测难度": [self.difficulty],
                            "区分度": [difference], "平均分": [ave_score], "标准差": [std], "单科与总分相关系数": corr_list[0]}
        elif self.name == "数学":
            summary_dict = {"科目": [self.name], "分值": ["100"], "预估难度": [np.nan], "实测难度": [self.difficulty],
                            "区分度": [difference], "平均分": [ave_score], "标准差": [std], "单科与总分相关系数": corr_list[1]}
        elif self.name == "英语":
            summary_dict = {"科目": [self.name], "分值": ["100"], "预估难度": [np.nan], "实测难度": [self.difficulty],
                            "区分度": [difference], "平均分": [ave_score], "标准差": [std], "单科与总分相关系数": corr_list[2]}
        elif self.name == "历史":
            summary_dict = {"科目": [self.name], "分值": ["70"], "预估难度": [np.nan], "实测难度": [self.difficulty],
                            "区分度": [difference], "平均分": [ave_score], "标准差": [std], "单科与总分相关系数": corr_list[3]}
        elif self.name == "物理":
            summary_dict = {"科目": [self.name], "分值": ["70"], "预估难度": [np.nan], "实测难度": [self.difficulty],
                            "区分度": [difference], "平均分": [ave_score], "标准差": [std], "单科与总分相关系数": corr_list[4]}
        elif self.name == "化学":
            summary_dict = {"科目": [self.name], "分值": ["50"], "预估难度": [np.nan], "实测难度": [self.difficulty],
                            "区分度": [difference], "平均分": [ave_score], "标准差": [std], "单科与总分相关系数": corr_list[5]}
        elif self.name == "道德与法治":
            summary_dict = {"科目": [self.name], "分值": ["50"], "预估难度": [np.nan], "实测难度": [self.difficulty],
                            "区分度": [difference], "平均分": [ave_score], "标准差": [std], "单科与总分相关系数": corr_list[6]}

        output = pd.DataFrame(summary_dict)
        return output

    def test_structure(self):
        title = []
        score = []
        number = []
        name = []
        average = []
        std = []

        students = []
        sum_of_score = []
        if "客观题" in self.item_score.values:
            name_subset = []
            title.append("客观题")
            score.append(self.sub_prob_score['score'].sum())
            number.append(len(self.sub_prob_score.index))
            for x in range(1, len(self.sub_prob_score.index) + 1):
                name_subset.append(str(x))
            name.append(','.join(name_subset))
            average.append(np.round(self.sub_data['score'].sum() / self.total_stu, 2))
            [students.append(x) for x in self.sub_data['xh'] if x not in students]
            for x in students:
                stu_grouped = self.sub_data.groupby(self.sub_data.xh).get_group(x)
                sum_of_score.append(stu_grouped['score'].sum())
            std.append(np.round(np.std(sum_of_score), 2))

        students = []
        sum_of_score = []
        if "主观题" in self.item_score.values:
            name_subset = []
            title.append("主观题")
            score.append(self.obj_prob_score['score'].sum())
            number.append(len(self.obj_prob_score.index))
            for x in range(self.total_prob - len(self.obj_prob_score.index) + 1, self.total_prob + 1):
                name_subset.append(str(x))
            name.append(','.join(name_subset))
            average.append(np.round(self.obj_data['score'].sum() / self.total_stu, 2))
            [students.append(x) for x in self.obj_data['xh'] if x not in students]
            for x in students:
                stu_grouped = self.obj_data.groupby(self.obj_data.xh).get_group(x)
                sum_of_score.append(stu_grouped['score'].sum())
            std.append(np.round(np.std(sum_of_score), 2))

        structure_dict = {"题型": title, "分值": score, "题目数量": number, "对应题目": name, "平均分": average, "标准差": std}
        output = pd.DataFrame(structure_dict)
        return output

    def calc_process(self):
        self.total_difficulty_curve()
        if "客观题" in self.item_score.values:
            self.calc_sub()
        if "主观题" in self.item_score.values:
            self.calc_obj()

    def difficulty_curve(self):
        return self.prob_diff_curve

    def distribution(self):
        num_diff_sub = [self.diff1_num_sub, self.diff2_num_sub, self.diff3_num_sub, self.diff4_num_sub,
                        self.diff5_num_sub]
        total_diff_sub = [self.diff1_total_sub, self.diff2_total_sub, self.diff3_total_sub, self.diff4_total_sub,
                        self.diff5_total_sub]
        num_diff_obj = [self.diff1_num_obj, self.diff2_num_obj, self.diff3_num_obj, self.diff4_num_obj,
                        self.diff5_num_obj]
        total_diff_obj = [self.diff1_total_obj, self.diff2_total_obj, self.diff3_total_obj, self.diff4_total_obj,
                          self.diff5_total_obj]

        num_dist_sub = [self.dist1_num_sub, self.dist2_num_sub, self.dist3_num_sub, self.dist4_num_sub]
        total_dist_sub = [self.dist1_total_sub, self.dist2_total_sub, self.dist3_total_sub, self.dist4_total_sub]
        num_dist_obj = [self.dist1_num_obj, self.dist2_num_obj, self.dist3_num_obj, self.dist4_num_obj]
        total_dist_obj = [self.dist1_total_obj, self.dist2_total_obj, self.dist3_total_obj, self.dist4_total_obj]

        total_diff = []
        total1_diff = self.diff1_total_sub + self.diff1_total_obj
        total_diff.append(total1_diff)
        total2_diff = self.diff2_total_sub + self.diff2_total_obj
        total_diff.append(total2_diff)
        total3_diff = self.diff3_total_sub + self.diff3_total_obj
        total_diff.append(total3_diff)
        total4_diff = self.diff4_total_sub + self.diff4_total_obj
        total_diff.append(total4_diff)
        total5_diff = self.diff5_total_sub + self.diff5_total_obj
        total_diff.append(total5_diff)

        num_diff = []
        num1_diff = self.diff1_num_sub + self.diff1_num_obj
        num_diff.append(num1_diff)
        num2_diff = self.diff2_num_sub + self.diff2_num_obj
        num_diff.append(num2_diff)
        num3_diff = self.diff3_num_sub + self.diff3_num_obj
        num_diff.append(num3_diff)
        num4_diff = self.diff4_num_sub + self.diff4_num_obj
        num_diff.append(num4_diff)
        num5_diff = self.diff5_num_sub + self.diff5_num_obj
        num_diff.append(num5_diff)

        total_dist = []
        total1_dist = self.dist1_total_sub + self.dist1_total_obj
        total_dist.append(total1_dist)
        total2_dist = self.dist2_total_sub + self.dist2_total_obj
        total_dist.append(total2_dist)
        total3_dist = self.dist3_total_sub + self.dist3_total_obj
        total_dist.append(total3_dist)
        total4_dist = self.dist4_total_sub + self.dist4_total_obj
        total_dist.append(total4_dist)

        num_dist = []
        num1_dist = self.dist1_num_sub + self.dist1_num_obj
        num_dist.append(num1_dist)
        num2_dist = self.dist2_num_sub + self.dist2_num_obj
        num_dist.append(num2_dist)
        num3_dist = self.dist3_num_sub + self.dist3_num_obj
        num_dist.append(num3_dist)
        num4_dist = self.dist4_num_sub + self.dist4_num_obj
        num_dist.append(num4_dist)

        name_diff_sub = []
        name1_diff_sub = (','.join(self.diff1_list_sub))
        name_diff_sub.append(name1_diff_sub)
        name2_diff_sub = (','.join(self.diff2_list_sub))
        name_diff_sub.append(name2_diff_sub)
        name3_diff_sub = (','.join(self.diff3_list_sub))
        name_diff_sub.append(name3_diff_sub)
        name4_diff_sub = (','.join(self.diff4_list_sub))
        name_diff_sub.append(name4_diff_sub)
        name5_diff_sub = (','.join(self.diff5_list_sub))
        name_diff_sub.append(name5_diff_sub)

        name_diff_obj = []
        name1_diff_obj = (','.join(self.diff1_list_obj))
        name_diff_obj.append(name1_diff_obj)
        name2_diff_obj = (','.join(self.diff2_list_obj))
        name_diff_obj.append(name2_diff_obj)
        name3_diff_obj = (','.join(self.diff3_list_obj))
        name_diff_obj.append(name3_diff_obj)
        name4_diff_obj = (','.join(self.diff4_list_obj))
        name_diff_obj.append(name4_diff_obj)
        name5_diff_obj = (','.join(self.diff5_list_obj))
        name_diff_obj.append(name5_diff_obj)

        name_dist_sub = []
        name1_dist_sub = (','.join(self.dist1_list_sub))
        name_dist_sub.append(name1_dist_sub)
        name2_dist_sub = (','.join(self.dist2_list_sub))
        name_dist_sub.append(name2_dist_sub)
        name3_dist_sub = (','.join(self.dist3_list_sub))
        name_dist_sub.append(name3_dist_sub)
        name4_dist_sub = (','.join(self.dist4_list_sub))
        name_dist_sub.append(name4_dist_sub)

        name_dist_obj = []
        name1_dist_obj = (','.join(self.dist1_list_obj))
        name_dist_obj.append(name1_dist_obj)
        name2_dist_obj = (','.join(self.dist2_list_obj))
        name_dist_obj.append(name2_dist_obj)
        name3_dist_obj = (','.join(self.dist3_list_obj))
        name_dist_obj.append(name3_dist_obj)
        name4_dist_obj = (','.join(self.dist4_list_obj))
        name_dist_obj.append(name4_dist_obj)

        diff_dict1 = {"难度范围": self.diff_range, "难度评价": self.diff_evaluate, "客观题题数": num_diff_sub,
                      "客观题分值": total_diff_sub, "主观题题数": num_diff_obj, "主观题分值": total_diff_obj,
                      "全卷题数": num_diff, "全卷分值": total_diff}
        diff_dict2 = {"难度范围": self.diff_range, "难度评价": self.diff_evaluate, "客观题题数": num_diff_sub,
                      "客观题题号": name_diff_sub, "主观题题数": num_diff_obj, "主观题题号": name_diff_obj,
                      "总题数": num_diff}
        dist_dict1 = {"区分度范围": self.dist_range, "区分度评价": self.dist_evaluate, "客观题题数": num_dist_sub,
                      "客观题分值": total_dist_sub, "主观题题数": num_dist_obj, "主观题分值": total_dist_obj,
                      "全卷题数": num_dist, "全卷分值": total_dist}
        dist_dict2 = {"难度范围": self.dist_range, "难度评价": self.dist_evaluate, "客观题题数": num_dist_sub,
                      "客观题题号": name_dist_sub, "主观题题数": num_dist_obj, "主观题题号": name_dist_obj,
                      "总题数": num_dist}

        diff_df1 = pd.DataFrame(diff_dict1)
        diff_df2 = pd.DataFrame(diff_dict2)
        dist_df1 = pd.DataFrame(dist_dict1)
        dist_df2 = pd.DataFrame(dist_dict2)

        dict = {"难度题型分布": diff_df1, "难度分布": diff_df2, "区分度题型分布": dist_df1, "区分度分布": dist_df2}
        return dict


SUBJECT_MAX_SCORE = []
SUBJECT_MIN_SCORE = []
SUBJECT_MED_SCORE = []
SUBJECT_AVE_SCORE = []
SUBJECT_STD_SCORE = []
SUBJECT_FULL_SCORE_RATE = []
SUBJECT_A_RATE = []
SUBJECT_B_RATE = []
SUBJECT_C_RATE = []
SUBJECT_D_RATE = []
SUBJECT_E_RATE = []
SUBJECT_DISTANCE = []
SCHOOL_NAME = []
[SCHOOL_NAME.append(x) for x in STUDENT_TOTAL['学校'] if x not in SCHOOL_NAME]


class Score_situaiton:
    def __init__(self, school, total_score):
        self.school_name = school
        self.score = total_score
        self.score_list = list(total_score.columns)
        self.dict_of_all = {}
        self.dict_of_score_rate = {}

    def statistics(self):
        school_name = self.school_name.copy()
        school_name.append('全体')

        for score in self.score_list[1:8]:
            index = self.score_list.index(score) -1
            subject_name = SUBJECT_LIST[index]
            full_score = FULL_SCORE[index]

            total_ave = self.score[score].mean()

            total_stu_list = []
            max_list = []
            min_list = []
            ave_list_all = []
            ave_list_left = []
            ave_list_right = []
            std_list = []
            full_rate = []
            a_rate = []
            b_rate = []
            c_rate = []
            d_rate = []
            e_rate = []
            f_rate = []
            g_rate = []
            difficulty = []
            correlation = []

            sr_list = []

            for school in self.school_name:
                this_school = self.score.groupby("学校").get_group(school)
                school_subject_score = this_school[score].sort_values(ascending=False)
                boundary = int(np.round(len(school_subject_score)*0.27, 0))
                left = school_subject_score.head(boundary)
                right = school_subject_score.tail(boundary)
                left_ave = left.mean()
                right_ave = right.mean()
                ave_list_left.append(np.round(left_ave, 2))
                ave_list_right.append(np.round(right_ave, 2))

                for_corr = this_school[[score, '分数.7']]
                corr = for_corr.corr()
                correlation.append(np.round(corr[score][1], 4))

                total_stu = len(school_subject_score)
                total_stu_list.append(total_stu)

                max_score = school_subject_score.max()
                max_list.append(np.round(max_score, 2))
                min_score = school_subject_score.min()
                min_list.append(np.round(min_score, 2))
                ave_score = school_subject_score.mean()
                ave_list_all.append(np.round(ave_score, 2))
                std_score = school_subject_score.std()
                std_list.append(np.round(std_score, 2))

                fr_stu = len(school_subject_score.where(school_subject_score == full_score).dropna())
                fr = np.round(fr_stu / total_stu * 100, 2)
                full_rate.append(fr)

                ar_stu = len(school_subject_score.where(school_subject_score >= full_score * 0.9).dropna())
                ar = np.round(ar_stu / total_stu * 100, 2)
                a_rate.append(ar)

                br_stu = school_subject_score.where(school_subject_score >= full_score * 0.8).dropna()
                br_stu = len(br_stu.where(br_stu < full_score * 0.9).dropna())
                br = np.round(br_stu / total_stu * 100, 2)

                b_rate.append(br)

                cr_stu = school_subject_score.where(school_subject_score >= full_score * 0.7).dropna()
                cr_stu = len(cr_stu.where(cr_stu < full_score * 0.8).dropna())
                cr = np.round(cr_stu / total_stu * 100, 2)
                c_rate.append(cr)

                dr_stu = school_subject_score.where(school_subject_score >= full_score * 0.6).dropna()
                dr_stu = len(dr_stu.where(dr_stu < full_score * 0.7).dropna())
                dr = np.round(dr_stu / total_stu * 100, 2)
                d_rate.append(dr)

                er_stu = len(school_subject_score.where(school_subject_score < full_score * 0.4).dropna())
                er = np.round(er_stu / total_stu * 100, 2)
                e_rate.append(er)

                gr = ave_score / total_ave * 100
                g_rate.append(np.round(gr, 2))

                fr = gr - 100
                f_rate.append(np.round(fr, 2))

                school_diff = ave_score / full_score
                difficulty.append(np.round(school_diff, 2))

                sr_list.append(np.round(school_diff * 100, 2))

                if self.school_name.index(school) == len(self.school_name) - 1:
                    all = self.score[[score, '分数.7']]
                    school_subject_score = all[score].sort_values(ascending=False)
                    boundary = int(np.round(len(school_subject_score) * 0.27, 0))
                    left = school_subject_score.head(boundary)
                    right = school_subject_score.tail(boundary)
                    left_ave = left.mean()
                    right_ave = right.mean()
                    ave_list_left.append(np.round(left_ave, 2))
                    ave_list_right.append(np.round(right_ave, 2))

                    corr = all.corr()
                    correlation.append(np.round(corr[score][1], 4))

                    total_stu = len(all)
                    total_stu_list.append(total_stu)

                    max_score = all[score].max()
                    max_list.append(np.round(max_score, 2))
                    min_score = all[score].min()
                    min_list.append(np.round(min_score, 2))
                    ave_score = all[score].mean()
                    ave_list_all.append(np.round(ave_score, 2))
                    std_score = all[score].std()
                    std_list.append(np.round(std_score, 2))

                    fr_stu = len(all[score].where(all[score] == full_score).dropna())
                    fr = np.round(fr_stu / total_stu * 100, 2)
                    full_rate.append(fr)

                    ar_stu = len(all[score].where(all[score] >= full_score * 0.9).dropna())
                    ar = np.round(ar_stu / total_stu * 100, 2)
                    a_rate.append(ar)

                    br_stu = all[score].where(all[score] >= full_score * 0.8).dropna()
                    br_stu = len(br_stu.where(br_stu < full_score * 0.9).dropna())
                    br = np.round(br_stu / total_stu * 100, 2)

                    b_rate.append(br)

                    cr_stu = all[score].where(all[score] >= full_score * 0.7).dropna()
                    cr_stu = len(cr_stu.where(cr_stu < full_score * 0.8).dropna())
                    cr = np.round(cr_stu / total_stu * 100, 2)
                    c_rate.append(cr)

                    dr_stu = all[score].where(all[score] >= full_score * 0.6).dropna()
                    dr_stu = len(dr_stu.where(dr_stu < full_score * 0.7).dropna())
                    dr = np.round(dr_stu / total_stu * 100, 2)
                    d_rate.append(dr)

                    er_stu = len(all[score].where(all[score] < full_score * 0.4).dropna())
                    er = np.round(er_stu / total_stu * 100, 2)
                    e_rate.append(er)

                    g_rate.append(np.nan)

                    f_rate.append(np.nan)

                    school_diff = ave_score / full_score
                    difficulty.append(np.round(school_diff, 2))

                    sr_list.append(np.round(school_diff*100, 2))

                    dict = {"学校": school_name, "总人数": total_stu_list, "最高分": max_list, "最低分": min_list,
                            "全体平均分": ave_list_all, "高分段平均分": ave_list_left, "低分段平均分": ave_list_right,
                            "标准差": std_list, "满分率": full_rate, "超优率": a_rate, "优秀率": b_rate, "良好率": c_rate,
                            "及格率": d_rate, "低分率": e_rate, "超均率": f_rate, "比均率": g_rate, "难度": difficulty,
                            "单科与总分的相关系数": correlation}
                    df = pd.DataFrame(dict)
                    self.dict_of_all[subject_name] = df

                    sn = [x[0:2] for x in school_name]
                    score_dict = {"学校": school_name, "得分率": sr_list}
                    score_df = pd.DataFrame(score_dict)
                    self.dict_of_score_rate[subject_name] = score_df

        return self.dict_of_all

    def score_rate(self):
        return self.dict_of_score_rate

    def rank(self):
        rank_dict = {}
        score_block_all = []
        for x in range(7):
            full_score = FULL_SCORE[x]
            score_block = []
            for y in range(0, full_score, int(full_score*0.1)):
                score_block.append("[{}, {})".format(str(y), str(int(y + full_score*0.1))))
            score_block.append("[{}, *]".format(full_score))
            score_block.reverse()
            score_block.append("人数")
            score_block_all.append(score_block)

        for score in self.score_list[1:8]:
            index = self.score_list.index(score) - 1
            subject_name = SUBJECT_LIST[index]
            full_score = FULL_SCORE[index]

            subject_dict = {"分数段": score_block_all[index]}

            for school in self.school_name:
                school_data = []
                this_school = self.score.groupby("学校").get_group(school)

                for x in range(0, full_score, int(full_score*0.1)):
                    left_bound = this_school.where(this_school[score] >= x).dropna()
                    right_bound = left_bound.where(left_bound[score] < x+full_score*0.1).dropna()
                    if len(right_bound) == 0:
                        school_data.append(np.nan)
                    else:
                        school_data.append(len(right_bound))
                left_bound = this_school.where(this_school[score] >= full_score).dropna()
                if len(left_bound) == 0:
                    school_data.append(np.nan)
                else:
                    school_data.append(len(left_bound))
                school_data.reverse()
                school_data.append(len(this_school))
                subject_dict[school] = school_data

            subject_df = pd.DataFrame(subject_dict)
            rank_dict[subject_name] = subject_df

        return rank_dict

    def sorted_rank(self):
        all = {}
        rank_block = ["[*, 10]", "[*, 30]", "[*, 50]", "[*, 100]", "[*, 200]", "[*, 300]", "[*, 500]", "[*, 1000]",
                      "人数", "最高分", "最低分"]
        blocks = [10, 30, 50, 100, 200, 300, 500, 1000]

        for score in self.score_list[1:8]:
            subject_dict = {}

            index = self.score_list.index(score) - 1
            subject_name = SUBJECT_LIST[index]
            full_score = FULL_SCORE[index]

            this_subject = self.score[['学校', score]].sort_values(by=score, ascending=False)

            subject_dict["名次段"] = rank_block

            for school in self.school_name:
                subject_dict[school] = []

            for x in blocks:
                interval = this_subject[0:x]
                for school in self.school_name:
                    if school in interval["学校"].values:
                        this_school = interval.groupby("学校").get_group(school)
                        subject_dict[school].append(len(this_school))
                    else:
                        subject_dict[school].append(np.nan)

            for school in self.school_name:
                this_school = self.score.groupby("学校").get_group(school)
                subject_dict[school].append(len(this_school))
                subject_dict[school].append(this_school[score].max())
                subject_dict[school].append(this_school[score].min())

            subject_df = pd.DataFrame(subject_dict)
            all[subject_name] = subject_df

        return all


class Analysis:
    def __init__(self, stu_total, total):
        self.student = stu_total
        self.maximum = 560
        self.subject_total = total

    def total_score_analysis(self):
        school_grouped = self.student.groupby(self.student.学校)

        max_score = []
        min_score = []
        med_score = []
        ave_score = []
        std_score = []
        full_score_rate = []
        a_rate = []
        b_rate = []
        c_rate = []
        d_rate = []
        e_rate = []
        distance = []

        for x in SCHOOL_NAME:
            school = school_grouped.get_group(x)
            stu_list = school['分数.7']
            total_stu = len(stu_list)

            max_score.append(stu_list.max())
            min_score.append(stu_list.min())
            med_score.append(np.round(stu_list.median(), 1))
            ave_score.append(np.round(stu_list.mean(), 2))
            std_score.append(np.round(stu_list.std(), 2))

            full_score_stu = len(stu_list.where(stu_list == self.maximum).dropna())
            fsr = np.round(full_score_stu / total_stu * 100, 2)
            fsr = str(fsr) + "%"
            full_score_rate.append(fsr)

            ar_stu = len(stu_list.where(stu_list >= self.maximum * 0.9).dropna())
            ar = np.round(ar_stu / total_stu * 100, 2)
            ar = str(ar) + "%"
            a_rate.append(ar)

            br_stu = stu_list.where(stu_list >= self.maximum * 0.8).dropna()
            br_stu = len(br_stu.where(br_stu < self.maximum * 0.9).dropna())
            br = np.round(br_stu / total_stu * 100, 2)
            br = str(br) + "%"
            b_rate.append(br)

            cr_stu = stu_list.where(stu_list >= self.maximum * 0.7).dropna()
            cr_stu = len(cr_stu.where(cr_stu < self.maximum * 0.8).dropna())
            cr = np.round(cr_stu / total_stu * 100, 2)
            cr = str(cr) + "%"
            c_rate.append(cr)

            dr_stu = stu_list.where(stu_list >= self.maximum * 0.6).dropna()
            dr_stu = len(dr_stu.where(dr_stu < self.maximum * 0.7).dropna())
            dr = np.round(dr_stu / total_stu * 100, 2)
            dr = str(dr) + "%"
            d_rate.append(dr)

            er_stu = len(stu_list.where(stu_list < self.maximum * 0.4).dropna())
            er = np.round(er_stu / total_stu * 100, 2)
            er = str(er) + "%"
            e_rate.append(er)

            distance.append(np.round((stu_list.max() - stu_list.min()), 1))

        score_dict = {"学校": SCHOOL_NAME, "最高分": max_score, "最低分": min_score, "中位数": med_score, "平均分": ave_score,
                      "满分率": full_score_rate, "超优率": a_rate, "优秀率": b_rate, "良好率": c_rate, "及格率": d_rate,
                      "低分率": e_rate, "标准差": std_score,  "全距": distance}
        output = pd.DataFrame(score_dict)
        return output

    def subject_score_analysis(self):
        scores = self.subject_total['全卷']
        total_stu = len(scores)
        SUBJECT_MAX_SCORE.append(scores.max())
        SUBJECT_MIN_SCORE.append(scores.min())
        SUBJECT_MED_SCORE.append(np.round(scores.median(), 1))
        SUBJECT_AVE_SCORE.append(np.round(scores.mean()))
        SUBJECT_STD_SCORE.append(np.round(scores.std(), 2))

        full_score_stu = len(scores.where(scores == self.student).dropna())
        fsr = np.round(full_score_stu / total_stu * 100, 2)
        fsr = str(fsr) + "%"
        SUBJECT_FULL_SCORE_RATE.append(fsr)

        ar_stu = len(scores.where(scores >= self.student * 0.9).dropna())
        ar = np.round(ar_stu / total_stu * 100, 2)
        ar = str(ar) + "%"
        SUBJECT_A_RATE.append(ar)

        br_stu = scores.where(scores >= self.student * 0.8).dropna()
        br_stu = len(br_stu.where(br_stu < self.student * 0.9).dropna())
        br = np.round(br_stu / total_stu * 100, 2)
        br = str(br) + "%"
        SUBJECT_B_RATE.append(br)

        cr_stu = scores.where(scores >= self.student * 0.7).dropna()
        cr_stu = len(cr_stu.where(cr_stu < self.student * 0.8).dropna())
        cr = np.round(cr_stu / total_stu * 100, 2)
        cr = str(cr) + "%"
        SUBJECT_C_RATE.append(cr)

        dr_stu = scores.where(scores >= self.student * 0.6).dropna()
        dr_stu = len(dr_stu.where(dr_stu < self.student * 0.7).dropna())
        dr = np.round(dr_stu / total_stu * 100, 2)
        dr = str(dr) + "%"
        SUBJECT_D_RATE.append(dr)

        er_stu = len(scores.where(scores < self.student * 0.4).dropna())
        er = np.round(er_stu / total_stu * 100, 2)
        er = str(er) + "%"
        SUBJECT_E_RATE.append(er)

        SUBJECT_DISTANCE.append(np.round(scores.max() - scores.min(), 2))

    def score_block_analysis(self):
        block_one = []
        block_two = []
        block_three = []
        block_four = []
        block_five = []
        block_six = []
        block_seven = []
        block_eight = []
        block_nine = []
        block_ten = []

        stu_num = []

        school_grouped = self.student.groupby(self.student.学校)
        for x in SCHOOL_NAME:
            school = school_grouped.get_group(x)
            stu_list = school['分数.7']

            school_stu = len(school)
            stu_num.append(school_stu)

            b1 = len(stu_list.where(stu_list < self.maximum * 0.1).dropna())
            block_one.append(b1)

            b2 = stu_list.where(stu_list < self.maximum * 0.2).dropna()
            b2 = len(b2.where(b2 >= self.maximum * 0.1).dropna())
            block_two.append(b2)

            b3 = stu_list.where(stu_list < self.maximum * 0.3).dropna()
            b3 = len(b3.where(b3 >= self.maximum * 0.2).dropna())
            block_three.append(b3)

            b4 = stu_list.where(stu_list < self.maximum * 0.4).dropna()
            b4 = len(b4.where(b4 >= self.maximum * 0.3).dropna())
            block_four.append(b4)

            b5 = stu_list.where(stu_list < self.maximum * 0.5).dropna()
            b5 = len(b5.where(b5 >= self.maximum * 0.4).dropna())
            block_five.append(b5)

            b6 = stu_list.where(stu_list < self.maximum * 0.6).dropna()
            b6 = len(b6.where(b6 >= self.maximum * 0.5).dropna())
            block_six.append(b6)

            b7 = stu_list.where(stu_list < self.maximum * 0.7).dropna()
            b7 = len(b7.where(b7 >= self.maximum * 0.6).dropna())
            block_seven.append(b7)

            b8 = stu_list.where(stu_list < self.maximum * 0.8).dropna()
            b8 = len(b8.where(b8 >= self.maximum * 0.7).dropna())
            block_eight.append(b8)

            b9 = stu_list.where(stu_list < self.maximum * 0.9).dropna()
            b9 = len(b9.where(b9 >= self.maximum * 0.8).dropna())
            block_nine.append(b9)

            b10 = stu_list.where(stu_list < self.maximum).dropna()
            b10 = len(b10.where(b10 >= self.maximum * 0.9).dropna())
            block_ten.append(b10)

        school_list = SCHOOL_NAME[:]
        school_list.append("全区总计")

        stu_list = self.student['分数.7']

        school_stu = len(self.student)
        stu_num.append(school_stu)

        b1 = len(stu_list.where(stu_list < self.maximum * 0.1).dropna())
        block_one.append(b1)

        b2 = stu_list.where(stu_list < self.maximum * 0.2).dropna()
        b2 = len(b2.where(b2 >= self.maximum * 0.1).dropna())
        block_two.append(b2)

        b3 = stu_list.where(stu_list < self.maximum * 0.3).dropna()
        b3 = len(b3.where(b3 >= self.maximum * 0.2).dropna())
        block_three.append(b3)

        b4 = stu_list.where(stu_list < self.maximum * 0.4).dropna()
        b4 = len(b4.where(b4 >= self.maximum * 0.3).dropna())
        block_four.append(b4)

        b5 = stu_list.where(stu_list < self.maximum * 0.5).dropna()
        b5 = len(b5.where(b5 >= self.maximum * 0.4).dropna())
        block_five.append(b5)

        b6 = stu_list.where(stu_list < self.maximum * 0.6).dropna()
        b6 = len(b6.where(b6 >= self.maximum * 0.5).dropna())
        block_six.append(b6)

        b7 = stu_list.where(stu_list < self.maximum * 0.7).dropna()
        b7 = len(b7.where(b7 >= self.maximum * 0.6).dropna())
        block_seven.append(b7)

        b8 = stu_list.where(stu_list < self.maximum * 0.8).dropna()
        b8 = len(b8.where(b8 >= self.maximum * 0.7).dropna())
        block_eight.append(b8)

        b9 = stu_list.where(stu_list < self.maximum * 0.9).dropna()
        b9 = len(b9.where(b9 >= self.maximum * 0.8).dropna())
        block_nine.append(b9)

        b10 = stu_list.where(stu_list < self.maximum).dropna()
        b10 = len(b10.where(b10 >= self.maximum * 0.9).dropna())
        block_ten.append(b10)

        dict = {"学校": school_list,"【0,10%）": block_one, "【10%,20%）": block_two, "【20%,30%）": block_three,
                "【30%,40%）": block_four, "【40%,50%）": block_five, "【50%,60%）": block_six, "【60%,70%）": block_seven,
                "【70%,80%）": block_eight, "【80%,90%）": block_nine, "【90%-100%）": block_ten, "小计": stu_num}
        output = pd.DataFrame(dict)
        return output

    def school_score_analysis(self):
        stu_num = []
        a_goal = []
        a_goal_rate = []
        b_goal = []
        b_goal_rate = []
        c_goal = []
        c_goal_rate = []
        d_goal = []
        d_goal_rate = []
        e_goal = []
        e_goal_rate = []

        school_grouped = self.student.groupby(self.student.学校)
        for x in SCHOOL_NAME:
            school = school_grouped.get_group(x)
            stu_list = school['分数.7']

            school_stu = len(school)
            stu_num.append(school_stu)

            ag = len(stu_list.where(stu_list >= 462.0).dropna())
            agr = np.round(ag / school_stu * 100, 1)
            a_goal.append(ag)
            a_goal_rate.append(agr)

            bg = len(stu_list.where(stu_list >= 430.7).dropna())
            bgr = np.round(bg / school_stu * 100, 1)
            b_goal.append(bg)
            b_goal_rate.append(bgr)

            cg = len(stu_list.where(stu_list >= 404.5).dropna())
            cgr = np.round(cg / school_stu * 100, 1)
            c_goal.append(cg)
            c_goal_rate.append(cgr)

            dg = len(stu_list.where(stu_list >= 371.1).dropna())
            dgr = np.round(dg / school_stu * 100, 1)
            d_goal.append(dg)
            d_goal_rate.append(dgr)

            eg = len(stu_list.where(stu_list >= 302.9).dropna())
            egr = np.round(eg / school_stu * 100, 1)
            e_goal.append(eg)
            e_goal_rate.append(egr)

        school_list = SCHOOL_NAME[:]
        school_list.append("全体")

        stu_list = self.student['分数.7']

        school_stu = len(self.student)
        stu_num.append(school_stu)

        ag = len(stu_list.where(stu_list >= 462.0).dropna())
        agr = np.round(ag / school_stu * 100, 1)
        a_goal.append(ag)
        a_goal_rate.append(agr)

        bg = len(stu_list.where(stu_list >= 430.7).dropna())
        bgr = np.round(bg / school_stu * 100, 1)
        b_goal.append(bg)
        b_goal_rate.append(bgr)

        cg = len(stu_list.where(stu_list >= 404.5).dropna())
        cgr = np.round(cg / school_stu * 100, 1)
        c_goal.append(cg)
        c_goal_rate.append(cgr)

        dg = len(stu_list.where(stu_list >= 371.1).dropna())
        dgr = np.round(dg / school_stu * 100, 1)
        d_goal.append(dg)
        d_goal_rate.append(dgr)

        eg = len(stu_list.where(stu_list >= 302.9).dropna())
        egr = np.round(eg / school_stu * 100, 1)
        e_goal.append(eg)
        e_goal_rate.append(egr)

        dict = {"学校": school_list, "人数": stu_num, "目标1人数": a_goal, "目标1上线率": a_goal_rate, "目标2人数": b_goal,
                "目标2上线率": b_goal_rate, "目标3人数": c_goal, "目标3上线率": c_goal_rate, "目标4人数": d_goal,
                "目标4上线率": d_goal_rate, "目标5人数": e_goal, "目标5上线率": e_goal_rate}

        output = pd.DataFrame(dict)
        return output


def full_score_analysis():
    total_stu = len(STUDENT_TOTAL['分数.7'])
    SUBJECT_MAX_SCORE.append(STUDENT_TOTAL['分数.7'].max())
    SUBJECT_MIN_SCORE.append(STUDENT_TOTAL['分数.7'].min())
    SUBJECT_MED_SCORE.append(np.round(STUDENT_TOTAL['分数.7'].median(), 1))
    SUBJECT_AVE_SCORE.append(np.round(STUDENT_TOTAL['分数.7'].mean()))
    SUBJECT_STD_SCORE.append(np.round(STUDENT_TOTAL['分数.7'].std(), 2))

    full_score_stu = len(STUDENT_TOTAL['分数.7'].where(STUDENT_TOTAL['分数.7'] == 560).dropna())
    fsr = np.round(full_score_stu / total_stu * 100, 2)
    fsr = str(fsr) + "%"
    SUBJECT_FULL_SCORE_RATE.append(fsr)

    ar_stu = len(STUDENT_TOTAL['分数.7'].where(STUDENT_TOTAL['分数.7'] >= 560 * 0.9).dropna())
    ar = np.round(ar_stu / total_stu * 100, 2)
    ar = str(ar) + "%"
    SUBJECT_A_RATE.append(ar)

    br_stu = STUDENT_TOTAL['分数.7'].where(STUDENT_TOTAL['分数.7'] >= 560 * 0.8).dropna()
    br_stu = len(br_stu.where(br_stu < 560 * 0.9).dropna())
    br = np.round(br_stu / total_stu * 100, 2)
    br = str(br) + "%"
    SUBJECT_B_RATE.append(br)

    cr_stu = STUDENT_TOTAL['分数.7'].where(STUDENT_TOTAL['分数.7'] >= 560 * 0.7).dropna()
    cr_stu = len(cr_stu.where(cr_stu < 560 * 0.8).dropna())
    cr = np.round(cr_stu / total_stu * 100, 2)
    cr = str(cr) + "%"
    SUBJECT_C_RATE.append(cr)

    dr_stu = STUDENT_TOTAL['分数.7'].where(STUDENT_TOTAL['分数.7'] >= 560 * 0.6).dropna()
    dr_stu = len(dr_stu.where(dr_stu < 560 * 0.7).dropna())
    dr = np.round(dr_stu / total_stu * 100, 2)
    dr = str(dr) + "%"
    SUBJECT_D_RATE.append(dr)

    er_stu = len(STUDENT_TOTAL['分数.7'].where(STUDENT_TOTAL['分数.7'] < 560* 0.4).dropna())
    er = np.round(er_stu / total_stu * 100, 2)
    er = str(er) + "%"
    SUBJECT_E_RATE.append(er)

    SUBJECT_DISTANCE.append(np.round(STUDENT_TOTAL['分数.7'].max() - STUDENT_TOTAL['分数.7'].min(), 2))


def subject_to_dataframe():
    score_dict = {"学科": SUBJECT_LIST, "总分": FULL_SCORE, "最高分": SUBJECT_MAX_SCORE,
                  "最低分": SUBJECT_MIN_SCORE, "中位数": SUBJECT_MED_SCORE, "平均分": SUBJECT_AVE_SCORE,
                  "标准差": SUBJECT_STD_SCORE, "满分率": SUBJECT_FULL_SCORE_RATE, "超优率": SUBJECT_A_RATE,
                  "优秀率": SUBJECT_B_RATE, "良好率": SUBJECT_C_RATE, "及格率": SUBJECT_D_RATE, "低分率": SUBJECT_E_RATE,
                  "全距": SUBJECT_DISTANCE}

    output = pd.DataFrame(score_dict)
    return output


class Effective_Score:
    def __init__(self, school, total_score, full_score, subject_list):
        self.score_line = [462, 430.7, 404.5, 371.1, 302.9]
        self.school_name = school
        self.full_score = full_score
        self.total_score = total_score
        self.score_list = list(total_score.columns)
        self.subject_list = subject_list
        all_stu_total = total_score['分数.7']

        block_one_stu = total_score.where(total_score['分数.7'] >= self.score_line[0]).dropna()
        self.b1_stu = len(block_one_stu)
        self.b1_ave = block_one_stu['分数.7'].mean()
        self.b1_co = self.score_line[0] / self.b1_ave

        block_two_stu = total_score.where(total_score['分数.7'] >= self.score_line[1]).dropna()
        self.b2_stu = len(block_two_stu)
        self.b2_ave = block_two_stu['分数.7'].mean()
        self.b2_co = self.score_line[1] / self.b2_ave

        block_three_stu = total_score.where(total_score['分数.7'] >= self.score_line[2]).dropna()

        self.b3_stu = len(block_three_stu)
        self.b3_ave = block_three_stu['分数.7'].mean()
        self.b3_co = self.score_line[2] / self.b3_ave

        block_four_stu = total_score.where(total_score['分数.7'] >= self.score_line[3]).dropna()
        self.b4_stu = len(block_four_stu)
        self.b4_ave = block_four_stu['分数.7'].mean()
        self.b4_co = self.score_line[3] / self.b4_ave

        block_five_stu = total_score.where(total_score['分数.7'] >= self.score_line[4]).dropna()
        self.b5_stu = len(block_five_stu)
        self.b5_ave = block_five_stu['分数.7'].mean()
        self.b5_co = self.score_line[4] / self.b5_ave

        self.co_list = [self.b1_co, self.b2_co, self.b3_co, self.b4_co, self.b5_co]

        self.subject_above_rate = {}

    def subject_ES(self):
        score_line_dict = {}
        for score_line in self.score_line:
            score_line_index = self.score_line.index(score_line)

            total = []
            single = []
            double = []
            hit_rate = []
            contribution_rate = []

            for score in self.score_list[1:8]:
                score_list_index = self.score_list.index(score) - 1
                subject_name = SUBJECT_LIST[score_list_index]
                full_score = FULL_SCORE[score_list_index]

                subject_dict = {}
                subject_total = []
                subject_single = []
                subject_double = []
                subject_hit_rate = []
                subject_contribution_rate = []

                for school in self.school_name:
                    this_school = self.total_score.groupby("学校").get_group(school)
                    above_line = this_school.where(this_school['分数.7'] >= score_line).dropna()

                    total_above_line = len(above_line)
                    subject_total.append(total_above_line)

                    this_subject = above_line[score]
                    average = this_subject.mean()
                    subject_effect = average * self.co_list[score_line_index]

                    this_school_single = this_school.where(this_school[score] >= subject_effect).dropna()
                    num_single = len(this_school_single)
                    subject_single.append(num_single)

                    this_school_double = this_subject.where(this_subject >= subject_effect).dropna()
                    num_double = len(this_school_double)
                    subject_double.append(num_double)

                    if num_single == 0:
                        this_school_hit = 0
                    else:
                        this_school_hit = num_double / num_single
                    subject_hit_rate.append(np.round(this_school_hit*100, 1))
                    if total_above_line == 0:
                        this_school_contribution = 0
                    else:
                        this_school_contribution = num_double / total_above_line
                    subject_contribution_rate.append(np.round(this_school_contribution*100, 1))

                subject_dict["学校"] = self.school_name
                subject_dict["上线人数"] = subject_single
                subject_df = pd.DataFrame(subject_dict)
                self.subject_above_rate[subject_name] = subject_df

                above_line = self.total_score.where(self.total_score['分数.7'] >= score_line).dropna()
                total_above_line = len(above_line)

                subject_total.append(total_above_line)

                this_subject = above_line[score]
                average = this_subject.mean()
                subject_effect = average * self.co_list[score_line_index]

                this_school_single = self.total_score.where(self.total_score[score] >= subject_effect).dropna()
                num_single = len(this_school_single)
                subject_single.append(num_single)

                this_school_double = this_subject.where(this_subject >= subject_effect).dropna()
                num_double = len(this_school_double)
                subject_double.append(num_double)

                if num_single == 0:
                    this_school_hit = 0
                else:
                    this_school_hit = num_double / num_single
                subject_hit_rate.append(np.round(this_school_hit * 100, 1))
                if total_above_line == 0:
                    this_school_contribution = 0
                else:
                    this_school_contribution = num_double / total_above_line
                subject_contribution_rate.append(np.round(this_school_contribution * 100, 1))

                total.append(subject_total)
                single.append(subject_single)
                double.append(subject_double)
                hit_rate.append(subject_hit_rate)
                contribution_rate.append(subject_contribution_rate)

            this_goal_dict = {}
            school_name = self.school_name.copy()
            school_name.append("全体")
            this_goal_dict["学校"] = school_name
            for index in range(len(SUBJECT_LIST)-1):
                this_subject = SUBJECT_LIST[index]
                this_goal_dict["总分上线人数"] = total[index]
                this_goal_dict[this_subject+"单上线"] = single[index]
                this_goal_dict[this_subject + "双上线"] = double[index]
                this_goal_dict[this_subject + "命中率"] = hit_rate[index]
                this_goal_dict[this_subject + "贡献率"] = contribution_rate[index]

            this_goal_df = pd.DataFrame(this_goal_dict)
            score_line_dict[str(score_line)] = this_goal_df

        return score_line_dict

    def distribution_graph(self):
        return self.subject_above_rate


if __name__ == "__main__":
    summary_output = {}
    target_output = {}
    structure_summary = {}
    export_data_path = "D:/Data/app/output/"
    difficulty_curve_data_path = "D:/Data/app/output/2-1-3-9分段难度曲线/"
    diff_and_dist_distribution = "D:/Data/app/output/2-1-3-6试题难度与区分度分布/"
    if not os.path.exists(export_data_path):
        os.makedirs(export_data_path)
    if not os.path.exists(difficulty_curve_data_path):
        os.makedirs(difficulty_curve_data_path)
    if not os.path.exists(diff_and_dist_distribution):
        os.makedirs(diff_and_dist_distribution)

    analysis = Analysis(STUDENT_TOTAL, 0)
    school_analysis = analysis.total_score_analysis()
    print("学校得分情况（总分）计算完成")
    score_line_analysis = analysis.school_score_analysis()
    print("分数线预测计算完成")
    score_block_analysis = analysis.score_block_analysis()
    print("分数段分布表计算完成")
    effective_score = Effective_Score(SCHOOL_NAME, TOTAL_SCORE, FULL_SCORE, SUBJECT_LIST)
    effect_score_dict = effective_score.subject_ES()
    above_rate_graph = effective_score.distribution_graph()
    print("上线率计算完成")

    for x in range(0, 7):
        analysis = Analysis(FULL_SCORE[x], SUBJECT_TOTAL_LIST[x])
        analysis.subject_score_analysis()
    full_score_analysis()
    subject_analysis = subject_to_dataframe()
    print("科目得分情况计算完成\n")

    situation = Score_situaiton(SCHOOL_NAME, TOTAL_SCORE)
    subject_sorted_rank = situation.sorted_rank()
    print("全区前N名各校分布及名单（科目）计算完成")
    situ_stats = situation.statistics()
    print("各学校得分情况（科目）计算完成")
    score_rate = situation.score_rate()
    print("各学校得分率分布图（科目）计算完成")
    subject_score_block = situation.rank()
    print("各校分数段人数统计（科目）计算完成\n")

    with pd.ExcelWriter(export_data_path+"2-1-4-5全区前N名各校分布及名单（科目）.xlsx", engine='xlsxwriter') as writer:
        for subject in subject_sorted_rank.keys():
            subject_sorted_rank[subject].to_excel(writer, index=False, sheet_name=subject)

    with pd.ExcelWriter(export_data_path+"2-1-4-1-3学校得分率分布图（科目）.xlsx", engine='xlsxwriter') as writer:
        for each in score_rate.keys():
            score_rate[each].to_excel(writer, index=False, sheet_name=each)
            workbook = writer.book
            worksheet = writer.sheets[each]
            chart = workbook.add_chart({'type': 'column'})
            max_row = len(score_rate[each])
            col_x = score_rate[each].columns.get_loc('学校')
            col_y = score_rate[each].columns.get_loc('得分率')

            chart.add_series({
                'name': '得分率',
                'categories': [each, 1, col_x, max_row, col_x],
                'values': [each, 1, col_y, max_row, col_y],
                'data_labels': {'value': True},
            })

            chart.set_title({'name': "各学校" + each + "得分率分布图"})
            chart.set_x_axis({'name': '学校', 'text_axis': True})
            chart.set_y_axis({'name': '得分率'})
            chart.set_size({'width': 1080, 'height': 864})
            worksheet.insert_chart('D2', chart)

    with pd.ExcelWriter(export_data_path+"2-1-4-3各校分数段人数统计.xlsx", engine='xlsxwriter') as writer:
        for subject in subject_score_block.keys():
            subject_score_block[subject].to_excel(writer, index=False, sheet_name=subject)
    print("各校分数段人数统计输出完成（科目）")

    for x in range(0, 7):
        subject = Exam(ITEM_SCORE, SUBJECT_TOTAL_LIST[x], SUBJECT_PROB_TOTAL[x], SUBJECT_LIST[x], FULL_SCORE[x])
        subject.calc_process()
        df = subject.to_dataframe()
        sheet_name = SUBJECT_LIST[x] + "题目指标.xlsx"
        target_output[sheet_name] = df
        print(SUBJECT_LIST[x]+"题目指标计算完成")
        subject_summary = subject.summary()
        summary_output[SUBJECT_LIST[x]] = subject_summary
        print(SUBJECT_LIST[x]+"概况计算完成")
        structure = subject.test_structure()
        structure_summary[SUBJECT_LIST[x]] = structure
        print(SUBJECT_LIST[x]+"题型分析计算完成\n")
        diff_curve = subject.difficulty_curve()
        print(SUBJECT_LIST[x]+"小题难度计算完成\n")
        distribution = subject.distribution()
        print(SUBJECT_LIST[x]+"难度与区分度分布计算完成")
        with pd.ExcelWriter(diff_and_dist_distribution + SUBJECT_LIST[x] + "难度与区分度分布.xlsx", engine='xlsxwriter') as writer:
            for each in distribution.keys():
                distribution[each].to_excel(writer, sheet_name=each)
        with pd.ExcelWriter(difficulty_curve_data_path + SUBJECT_LIST[x] + "分段小题难度.xlsx", engine='xlsxwriter') as writer:
            for prob in diff_curve.keys():
                diff_curve[prob].to_excel(writer, index=False, sheet_name=prob)
                workbook = writer.book
                worksheet = writer.sheets[prob]
                chart = workbook.add_chart({'type': 'line'})
                max_row = len(diff_curve[prob])
                col_x = diff_curve[prob].columns.get_loc('分数段')
                col_y = diff_curve[prob].columns.get_loc('难度值')

                chart.add_series({
                    'name': '难度值',
                    'categories': [prob, 1, col_x, max_row, col_x],
                    'values': [prob, 1, col_y, max_row, col_y],
                    'marker': {
                        'type': 'square',
                        'size': 4,
                        'border': {'color': 'black'},
                        'fill': {'color': 'red'},
                    },
                    'data_labels': {'value': True, 'position': 'left'},
                })
                if prob == "全卷":
                    chart.set_title({'name': prob + '难度曲线'})
                else:
                    chart.set_title({'name': '第'+str(prob)+'题难度曲线'})
                chart.set_x_axis({'name': '分数段'})
                chart.set_y_axis({'name': '难度值', 'min': 0, 'max': 1})
                chart.set_size({'width': 720, 'height': 576})
                worksheet.insert_chart('D2', chart)
        print(SUBJECT_LIST[x]+"分段小题难度输出完成\n")

    with pd.ExcelWriter(export_data_path+"2-1-4-1学校得分情况（科目）.xlsx", engine='xlsxwriter') as writer:
        for each in situ_stats.keys():
            situ_stats[each].to_excel(writer, index=False, sheet_name=each)
    print("学校得分情况（科目）输出完成")
    with pd.ExcelWriter(export_data_path+"1-3-2学校得分情况（总分）.xlsx", engine='xlsxwriter') as writer:
        school_analysis.to_excel(writer, index=False)
    print("学校得分情况（总分）输出完成")
    with pd.ExcelWriter(export_data_path+"1-4分数线预测.xlsx", engine='xlsxwriter') as writer:
        score_line_analysis.to_excel(writer, index=False)
    print("分数线预测输出完成")
    with pd.ExcelWriter(export_data_path + "1-5-2分数段分布表.xlsx", engine='xlsxwriter') as writer:
        score_block_analysis.to_excel(writer, index=False)
    print("分数段分布表输出完成")
    with pd.ExcelWriter(export_data_path+"1-3-1科目得分情况.xlsx", engine='xlsxwriter') as writer:
        subject_analysis.to_excel(writer, index=False)
    print("科目得分情况输出完成")

    with pd.ExcelWriter(export_data_path+"2-1-3-1题目指标.xlsx", engine='xlsxwriter') as writer:
        for each in target_output.keys():
            target_output[each].to_excel(writer, index=False, sheet_name=each)
            workbook = writer.book
            worksheet = writer.sheets[each]
            chart = workbook.add_chart({'type': 'scatter'})
            max_row = len(target_output[each])
            col_x = target_output[each].columns.get_loc('难度')
            col_y = target_output[each].columns.get_loc('区分度')

            chart.add_series({
                'name': '难度，区分度',
                'categories': [each, 1, col_x, max_row, col_x],
                'values': [each, 1, col_y, max_row, col_y],
                'marker': {'type': 'circle', 'size': 4, 'fill': {'color': 'orange'}},
                'data_labels': {'value': True, 'category': True, 'position': 'left'},
            })

            chart.set_title({'name': '题目难度与区分度分布图'})
            chart.set_x_axis({'name': '难度', 'min': 0, 'max': 1})
            chart.set_y_axis({'name': '区分度', 'min': 0, 'max': 1})
            chart.set_size({'width': 720, 'height': 576})
            worksheet.insert_chart('N2', chart)
    print("题目指标输出完成")

    with pd.ExcelWriter(export_data_path+"2-1-0-0概况.xlsx", engine='xlsxwriter') as writer:
        for summary in summary_output.keys():
            summary_output[summary].to_excel(writer, sheet_name=summary, index=False)
    print("概括输出完成")
    with pd.ExcelWriter(export_data_path+"2-1-2-4题型分析.xlsx", engine='xlsxwriter') as writer:
        for structure in structure_summary.keys():
            structure_summary[structure].to_excel(writer, sheet_name=structure, index=False)
    print("题型分析输出完成")
    with pd.ExcelWriter(export_data_path+"2-2-5单科上线命中率及贡献率.xlsx", engine='xlsxwriter') as writer:
        for score_line in effect_score_dict.keys():
            effect_score_dict[score_line].to_excel(writer, index=False, sheet_name=score_line)
    print("上线率输出完成")
    with pd.ExcelWriter(export_data_path+"2-2-3学科上线分数对比各校分布图.xlsx", engine='xlsxwriter') as writer:
        for subject in above_rate_graph.keys():
            above_rate_graph[subject].to_excel(writer, index=False, sheet_name=subject)
            workbook = writer.book
            worksheet = writer.sheets[subject]
            chart = workbook.add_chart({'type': 'column'})
            max_row = len(above_rate_graph[subject])
            col_x = above_rate_graph[subject].columns.get_loc('学校')
            col_y = above_rate_graph[subject].columns.get_loc('上线人数')

            chart.add_series({
                'name': '上线人数',
                'categories': [subject, 1, col_x, max_row, col_x],
                'values': [subject, 1, col_y, max_row, col_y],
                'data_labels': {'value': True},
            })

            chart.set_title({'name': subject + "单科上线各校分布图"})
            chart.set_x_axis({'name': '学校', 'text_axis': True})
            chart.set_y_axis({'name': '上线人数'})
            chart.set_size({'width': 1080, 'height': 864})
            worksheet.insert_chart('D2', chart)
    print("学科上线分数对比各校分布图")