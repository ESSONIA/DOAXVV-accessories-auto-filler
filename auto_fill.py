#交互方式:控制台程序,程序自动读取同目录内的excel文档并且解析数据
# 接下来需要完成的工作:
# 3)自动对比效能,给出最大值后打印最优结果
# 4)优化交互和数据检测

import os
import sys
import pandas as pd
import numpy as np
from tabulate import tabulate
from prettytable import PrettyTable
from scipy.optimize import linprog
from itertools import permutations

# 指定Excel文件的路径
script_dir = os.path.dirname(sys.argv[0])
excel_file_path = os.path.join(script_dir, 'inputdata.xlsx')

# 从Excel文件加载数据到DataFrame
df = pd.read_excel(excel_file_path)

def solver_kernel(name, values, costs, capacity, upperbound):
    # 线性规划 这是一个包装函数,将传入的原始数组首先加工成为合适的格式之后,后调用线性规划函数

    # 目标函数为求最大值,各变量系数即各潜力的实际效能.改变方案为求该最大值的相反数的最小值,即求所有系数取反后的最小值
    obj = -1 * values
    # 约束条件1:PP_Cost总和不能超过首饰的PP容量
    lhs_ineq = [costs]
    # 生成一个等长的全零数组,作为约束变量的下限
    zero_array = np.zeros_like(upperbound, dtype=int)
    # 使用 np.column_stack 将两个一维数组垂直堆叠成一个二维数组 构建成边界条件需要的格式
    bnd = np.column_stack((zero_array, upperbound))
    # 生成一个等长的全1数组,用于传参,指明约束变量必须使用整数
    ones_array = np.ones_like(upperbound, dtype=int)
    opt = linprog(c=obj, A_ub=lhs_ineq, b_ub=capacity, bounds=bnd,integrality=ones_array,method="highs")

    # 打印结果

    if(opt.success == True):
        df2 = pd.DataFrame({'潜力名称': name, '选取数量': opt.x.astype(int)})
        condition = df2['选取数量'] > 0
        filtered_df = df2[condition]
        table = PrettyTable()
        table.field_names = filtered_df.columns.tolist()

        for row in filtered_df.itertuples(index=False):
            table.add_row(row)

        # 打印PrettyTable对象
        print("PP容量为" + str(capacity) + "的潜力自动填充方案:")
        print(table)
        print()
        # 将约束变量作为结果返回.用于计算库存减少
        return opt.x.astype(int), (int)(opt.fun)


# 该函数需要完成连续三次调用
# 初步设想是在该函数内部连续调用三次solver_kernel(),并且在每次调用后进行库存的修正
# 将pandas::dataframe对象直接传入.对于数据的初加工在函数内部进行
def auto_fill(df):
    # 数据解析
    # 使用 iloc 获取第一列从第二行开始的内容并转换为 Python 数组
    P_Skill_names = df.iloc[:, 0]
    P_Skill_PP_Costs = df.iloc[:, 1].dropna().astype(int).to_numpy()
    P_Skill_Values = df.iloc[:, 2].dropna().astype(int).to_numpy()
    P_Skill_Limit = df.iloc[:, 3].dropna().astype(int).to_numpy()
    P_Skill_Stock = df.iloc[:, 4].dropna().astype(int).to_numpy()



    # 约束条件2:选取数量必须小于等于每个首饰内的使用次数
    # 约束条件3:选取数量必须小于等于剩余库存
    # 这两个条件不涉及到约束变量的系数,因此将其作为边界条件设置
    # 这两个条件可以首先求一次最小值.作为约束变量的上限
    P_Skill_available = np.minimum(P_Skill_Limit, P_Skill_Stock)

    Accessories_PP_capacity = df.iloc[:, 5].dropna().astype(int).to_numpy()
    # 使用 permutations 函数列出所有可能的排列
    Accessories_all_permutations = list(permutations(Accessories_PP_capacity))
    # 使用 set 去除重复项
    Accessories_unique_permutations = set(Accessories_all_permutations)
    # 打印去除重复项后的排列
    for unique_permutation in Accessories_unique_permutations:
        print("||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||")
        print("当前饰品填充顺序: " + str(unique_permutation))
        # 连续调用三次函数
        result_first_fill, value_first_fill = solver_kernel(P_Skill_names, P_Skill_Values,P_Skill_PP_Costs,unique_permutation[0],P_Skill_available)
        P_Skill_available_after_first_fill = np.minimum(P_Skill_Limit, P_Skill_Stock - result_first_fill)
        # print(result_first_fill)
        result_second_fill, value_second_fill = solver_kernel(P_Skill_names, P_Skill_Values,P_Skill_PP_Costs,unique_permutation[1],P_Skill_available_after_first_fill)
        P_Skill_available_after_second_fill = np.minimum(P_Skill_Limit, P_Skill_Stock - result_first_fill - result_second_fill)
        # print(result_second_fill)
        result_thrid_fill, value_thrid_fill = solver_kernel(P_Skill_names, P_Skill_Values,P_Skill_PP_Costs,unique_permutation[2],P_Skill_available_after_second_fill)
        total_value = value_first_fill + value_second_fill + value_thrid_fill
        print("潜力填充方案得到的总效能为:" + str(-1 * total_value) + "%")
auto_fill(df)

    








