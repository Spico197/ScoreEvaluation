import argparse
import os

import pandas as pd
import numpy as np
import json


class Excel(object):
    def __init__(self, program, file_name):
        self.file_name = file_name
        self.data = pd.read_excel(self.file_name)
        self.normalize()
        self.program = {}
        with open(program, 'r', encoding='utf-8') as f:
            json_data = json.load(f)
            self.program = json_data
        if len(self.program.items()) == 0:
            raise IOError("培养方案导入失败")
        
        if self.program:
            self.program_version = self.program['培养方案版本']
            self.program_major = self.program['专业名称']
            self.program_students = self.program['适用人群']
            self.program_graduation_credit = self.program['毕业要求学分']
            self.program_requests = self.program['分类要求']
    
    def normalize(self):
        """将成绩中的等级转换为分数"""
        self.data.loc[self.data['成绩'] == '优秀', '成绩'] = 90
        self.data.loc[self.data['成绩'] == '优', '成绩'] = 90
        self.data.loc[self.data['成绩'] == '良好', '成绩'] = 80
        self.data.loc[self.data['成绩'] == '良', '成绩'] = 80        
        self.data.loc[self.data['成绩'] == '中等', '成绩'] = 70
        self.data.loc[self.data['成绩'] == '中', '成绩'] = 70        
        self.data.loc[self.data['成绩'] == '及格', '成绩'] = 60
        self.data.loc[self.data['成绩'] == '不及格', '成绩'] = 50

    def weighted_score(self):
        """获取平均学分绩点和加权平均成绩"""
        credit = self.data['学分'].values
        gpa = self.data['绩点'].values
        score = self.data['成绩'].values
        
        if self.data['学分'].isnull().any():
            raise ValueError("学分中有空值, 请检查成绩单格式")
        if self.data['绩点'].isnull().any():
            raise ValueError("绩点中有空值, 请检查成绩单格式")
        if self.data['成绩'].isnull().any():
            raise ValueError("成绩中有空值, 请检查成绩单格式")
        
        weighted_gpa = np.dot(credit, gpa)/np.sum(credit)
        weighted_score = np.dot(credit, score)/np.sum(credit)
        
        return weighted_score, weighted_gpa

    def course_count(self):
        total_number = self.data['课程代码'].count()    # 总门数
        makeup_number = total_number - self.data.loc[pd.isnull(self.data['补考成绩']), '课程代码'].count() # 补考门数
        makeup_non_pass_number = self.data.loc[self.data['补考成绩'] < 60, '课程代码'].count() # 补考仍未通过
        retake_number = self.data['重修标记'].sum() # 重修科目
        
        return total_number, makeup_number, makeup_non_pass_number, retake_number

    def credit_stat(self):
        credit_sum = self.data['学分'].sum()    # 总学分
        pass_credit = self.data.loc[(self.data['成绩'] >= 60) | (self.data['补考成绩'] >= 60), '学分'].sum()   # 已经通过的学分

        return credit_sum, pass_credit, self.program_graduation_credit

    def module_course_stat(self):
        less_than_requirement = []
        greater_than_requirement = []
        for module in self.program_requests:
            # module example : {"课程性质": "创新实践", "学分要求": 3},
            module_name = module['课程性质']
            module_credit_requirement = module['学分要求']
            module_credit_autually = self.data.loc[
                                        (self.data['课程性质'] == module_name)
                                        & ((self.data['成绩'] >= 60) | (self.data['补考成绩'] >= 60))
                                    , "学分"].sum()
            if module_credit_autually > module_credit_requirement:
                greater_than_requirement.append((module_name, module_credit_requirement, module_credit_autually - module_credit_requirement))
            elif module_credit_autually < module_credit_requirement:
                less_than_requirement.append((module_name, module_credit_requirement, module_credit_requirement - module_credit_autually))
        
        return less_than_requirement, greater_than_requirement

    def optional_course(self):
        # 您的通识拓展选修课共通过了 {} 门, 分别为:
        # 课程代码: {}, 课程名称: {}, 课程归属: {}, 学分: {}
        report = "\n\n您的通识拓展选修课共通过了 {} 门, 总学分为: {}, 共有 {} 个类别, 分别为:\n".format(
            self.data.loc[(self.data['课程性质'] == '通识拓展选修课') & (self.data['成绩'] >= 60), '课程代码'].count(),
            self.data.loc[(self.data['课程性质'] == '通识拓展选修课') & (self.data['成绩'] >= 60), '学分'].sum(),
            len(set(self.data.loc[(self.data['课程性质'] == '通识拓展选修课') & (self.data['成绩'] >= 60), '课程归属'])))
        optional_courses = self.data.loc[(self.data['课程性质'] == '通识拓展选修课') & (self.data['成绩'] >= 60)]
        for ind in optional_courses.index:
            report += "\n\t课程代码: {}, 课程名称: {}, 课程归属: {}, 学分: {}".format(
                optional_courses.loc[ind, '课程代码'],
                optional_courses.loc[ind, '课程名称'],
                optional_courses.loc[ind, '课程归属'],
                optional_courses.loc[ind, '学分']
            )
        return report

    def __str__(self):
        return self.data.values


def main(args):

    if args.program == 'ls':
        files = os.listdir('./programs/')
        for file in files:
            if file[-5:] == '.json':
                with open(os.path.join('./programs/', file), 'r', encoding='utf-8') as f:
                    json_data = json.load(f)
                    print("方案名: {} | 适用专业: {} | 版本: {} | 适用人群: {}".format(
                        file, json_data['专业名称'], json_data['培养方案版本'], json_data['适用人群']
                    ))
        return
    
    if args.program is None or args.scorefile is None:
        raise ValueError("输入参数错误")
    else:
        files = os.listdir('./programs/')
        if args.program not in files:
            raise FileExistsError("培养方案不存在, 请检查输入名和方案名是否相同")
        if args.scorefile[-5:] != '.xlsx' and args.scorefile[-4:] != '.xls':
            raise FileExistsError("成绩单格式错误, 请检查成绩单后缀名是否为`.xlsx`或`.xls`")
        
        excel = Excel(os.path.join('./programs/', args.program), args.scorefile)
        report_string = """
    {} 专业的同学您好, 您正在使用的培养方案版本为 {}, 适用于 {} 使用. 

    若有统计错误之处, 烦请将本人成绩单(excel文件)和错误信息发送至邮箱: spico1026@gmail.com
    若培养方案列表中暂不支持您的培养方案, 也请您将相关培养方案发送至邮箱, 我将为您更新.

    重要说明: 最终统计结果仅供参考, 请以教学科出具的相关文件为准!!!

基本概要:

    您的加权平均分为 {}, 您的平均学分绩点为 {}.
    您共修了 {} 门课, 其中挂科门数为 {} 门, 补考仍未通过的门数为 {} 门, 重修门数为 {} 门.
    您共修了 {} 个学分, 其中通过的学分为 {}, 毕业要求学分为 {} .
        """
        # 培养方案中 {} 模块应修 {} 学分, 目前还差 {} 个学分.
        # 培养方案中 {} 模块应修 {} 学分, 超修了 {} 个学分.

        weighted_score, weighted_gpa = excel.weighted_score()
        total_number, makeup_number, makeup_non_pass_number, retake_number = excel.course_count()
        credit_sum, pass_credit, graduation_credit = excel.credit_stat()
        less_than_requirement, greater_than_requirement = excel.module_course_stat()

        report = report_string.format(
            excel.program_major, excel.program_version, excel.program_students,
            weighted_score, weighted_gpa, total_number, makeup_number, makeup_non_pass_number, retake_number,
            credit_sum, pass_credit, graduation_credit
        )

        report += "\n未修满学分模块:\n"

        if not less_than_requirement: 
            report += "您没有未修满学分的模块"
        else:
            for item in less_than_requirement:
                report += "\n\t培养方案中 {} 模块应修 {} 学分, 目前还差 {} 个学分.".format(item[0], item[1], item[2])

        report += "\n\n超修学分模块:\n"

        if not greater_than_requirement:
            report += "您没有超修的模块"
        else:
            for item in greater_than_requirement:
                report += "\n\t培养方案中 {} 模块应修 {} 学分, 超修了 {} 个学分.".format(item[0], item[1], item[2])

        report += excel.optional_course()

        # print(report)
        # print("\n报告已保存在 report.txt 文件中. 最终统计结果仅供参考, 请以教学科发放的具体文件为准!!!")
        
        # with open('report.txt', 'w', encoding="utf-8") as file:
        #     file.writelines(report) 



if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument("program", help="培养方案的选择, 或输入`ls`查看目前支持的培养方案")
    parser.add_argument("scorefile", nargs='?', help="Excel版成绩单(*.xls|*.xlsx)")
    args = parser.parse_args()
    main(args)
