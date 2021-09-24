import pandas as pd
import numpy as np
import sys


# ---Helper functions
def generate_time(year, month, k):
    """
    用于生成 y年m月 之后 k 期的月份名称
    :param k: 是k期，不包括本月
    :return: 返回的是一个list
    """
    date = []
    for i in range(k):
        if month < 12:
            month = month + 1
        elif month == 12:
            year = year + 1
            month = 1
        if month < 10:
            m = '0' + str(month)
            date.append('%s年%s月' % (year, m))
        else:
            date.append('%s年%s月' % (year, month))
    return date



def output(file_name, frequency, y, rst):
    """
    :param file_name:
    :param frequency:
    :param y:
    :param rst:
    :return:
    """
    with pd.ExcelWriter('./result/' + file_name + '.xlsx') as writer:
        y.to_excel(writer, sheet_name='y值', index=False)
        temp = pd.DataFrame()
        temp.to_excel(writer, sheet_name='年度数据', index=False)
        if frequency == '月度数据':
            temp.to_excel(writer, sheet_name='季度数据', index=False)
            rst.to_excel(writer, sheet_name='月度数据', index=False)
        elif frequency == '季度数据':
            rst.to_excel(writer, sheet_name='季度数据', index=False)
            temp.to_excel(writer, sheet_name='月度数据', index=False)
        temp.to_excel(writer, sheet_name='周度数据', index=False)
        temp.to_excel(writer, sheet_name='日度数据', index=False)
    print(file_name + '输出完毕')


def process_y(data):
    """
    处理解释变量y的函数，找到y，也就是GDP
    :param data: 即来源于 zb_excel:sheet 原始数据
    :return:
    """
    y = pd.DataFrame(data.iloc[:, [0, 1]])    # 选取['日期'， '地区生产总值']两列，为什么这里只提取了第一列GDP这个指标？
    y_drop = y.dropna(subset=[y.columns[1]])  # drop NA的信息 输出的y值。注意中间可能缺了一个月的情况没处理
    y1 = pd.DataFrame(columns=y.columns, index=[0])
    y1.loc[0, '日期'] = '特征数据滞后期数'        # 添加首行名称
    y_drop = y1.append(y_drop)                # 首行是说明，其他是
    return y_drop



def verify_freq(data, col_name):
    """
    判断数据的频率
    :param data: 原始数据
    :param col_name: 需要检测的变量名称
    :return: freq:频率
    """
    # col_name = "人身险"
    col_names = ['日期', col_name]
    data_used = data[col_names]
    data_used = data_used.dropna(subset=[col_name])
    date_list = list(data_used["日期"])
    month_list = [item[-3:-1] for item in date_list]
    if ("10" not in month_list) and ("03" in month_list):
        freq = "季"    # 如果10月不在，3月在，就是季
    if ("10" not in month_list) and ("03" not in month_list):
        freq = "年"    # 如果10月不在，3月也不再，就是年
    if ("10" in month_list) and ("03" in month_list):
        freq = "月"    # 如果10月在，3月也在，就是月
    return freq



# ---处理先行指标
def before_hand(data_name):
    """
    处理先行数据
    :param data_name:
    :return:
    """

    # data_name = '地区生产总值-江门'
    # data_name = '地区生产总值-中山'

    # Question 为什么要用V4
    # zb 和 data 来源于同一个 excel，zb 是一个筛选出来的指标，data是原始数据
    zb = pd.read_excel('./data/' + data_name + '先行性指标筛选结果_自定义相关系数.xlsx', sheet_name='V4')
    data = pd.read_excel('./data/' + data_name + '先行性指标筛选结果_自定义相关系数.xlsx', sheet_name='原始数据')

    # x指标名称   V4这个sheet里指标的名称
    zbmc = list(zb['指标名称'])[1:]

    # x指标期数   V4这个sheet里的指标的先行期数 qs 是期数的意思
    qs = list(zb['先行期数'])[1:]

    # 处理y
    y = process_y(data)
    start_month = list(y['日期'])[1]                # y 开始的那一个月
    i = data[data['日期'] == start_month].index[0]  # 获得y开始月份的index，注意第一行是说明"特征数据滞后期数"

    fst = pd.DataFrame(columns=data.columns, index=[0])
    fst.loc[0, '日期'] = '特征数据滞后期数'    # fst 是第一行是说明将"日期"这一列的首行填充为"特征数据滞后期数"

    # 处理data
    data = data.iloc[i:]       # 只要从y的首个index，即月份开始截取有效数据
    data = fst.append(data)    # 将fst说明添加在首行

    date = list(data['日期'])   # date 是截取后data的所有信息
    date = [i for i in date if '01月' not in i]
    data = data[data['日期'].isin(date)]  # 去掉1月份的数据，不需要使用1月份用来预测



    # 判定季度/月度，季度数据只有3/6/9/12月没有10月
    frequency = '季度数据'
    for i in list(y['日期']):
        if '10月' in i:
            frequency = '月度数据' # ???

    # 加上所有先行日期
    k = max(qs)
    if frequency == '季度数据':
        k = k * 3
    final_date = date[len(date) - 1]
    ym = final_date.replace('月', '').split('年')
    year = int(ym[0])
    month = int(ym[1])

    addtional_date = generate_time(year, month, k)
    addtional_date = [i for i in addtional_date if '01月' not in i]
    date.extend(addtional_date)

    # 先行期数处理。
    rst = pd.DataFrame()
    rst['日期'] = date
    if frequency == '季度数据':
        qs = [i * 3 for i in qs]

    for i in range(len(qs)):
        temp = list(data[zbmc[i]])
        temp[0] = qs[i]
        for j in range(qs[i]):
            temp.insert(1, np.nan)
        temp = pd.Series(temp)
        rst[zbmc[i]] = temp
    rst.loc[0, '日期'] = '特征数据滞后期数'

    # 输出结果
    output(data_name + 'V1先行指标', frequency, y, rst)
    return rst


# ---处理解释变量
def predict(data_name, before):
    """
    :param data_name: 指标名字
    :param before: before 是 before_hand 的返回值
    :return: 返回的是预测的结果吗，为什么要用 predict 的函数
    """
    zb = pd.read_excel('./data/' + data_name + '解释变量筛选结果.xlsx', sheet_name='V3')
    data = pd.read_excel('./data/' + data_name + '解释变量筛选结果.xlsx', sheet_name='原始数据')
    zbmc = list(zb['指标名称'])[1:]
    y = process_y(data)
    start_month = list(y['日期'])[1]  # y的开始时间
    i = data[data['日期'] == start_month].index[0]  # 获得开始月份的index

    fst = pd.DataFrame(data.iloc[[0], :])
    fst = pd.DataFrame(columns=fst.columns, index=[0])
    fst.loc[0, '日期'] = '特征数据滞后期数'

    data = data.iloc[i:]
    data = fst.append(data)

    date = list(data['日期'])
    date = [i for i in date if '01月' not in i]
    data = data[data['日期'].isin(date)]

    before_cols = list(before.columns)
    zbmc = [i for i in zbmc if i not in before_cols]  # 去掉先行指标里有的

    # 计算y的最新日期位置
    y_date = list(y['日期'])
    y_final = y_date[len(y_date) - 1]
    y_pos = 0
    for d in list(data['日期']):
        if d == y_final:
            break
        y_pos += 1

    rst = pd.DataFrame()
    rst['日期'] = date
    for i in range(len(zbmc)):
        temp = list(data[zbmc[i]])
        # 计算滞后期数
        zb_pos = None
        for j in range(len(temp) - 1, -1, -1):
            if not pd.isna(temp[j]):
                zb_pos = j
                break
        zhqs = y_pos - zb_pos
        temp.insert(0, zhqs)
        temp = pd.Series(temp)
        rst[zbmc[i]] = temp

    rst.drop('日期', axis=1, inplace=True)
    rst = pd.concat([before, rst], axis=1)
    # 输出结果

    # 判定季度/月度，季度数据只有3/6/9/12月没有10月
    frequency = '季度数据'
    for i in list(y['日期']):
        if '10月' in i:
            frequency = '月度数据'

    output(data_name + 'V1解释变量', frequency, y, rst)

    # 输出逻辑表
    logical = pd.DataFrame()
    gx = ['Y值']
    for i in range(len(list(rst.columns)[1:])):
        gx.append('X值')
    temp = list(rst.columns)[1:]
    temp.insert(0, y.columns[1])
    logical['指标名称'] = temp
    logical['关系'] = gx
    logical['起始年份'] = start_month.split('年')[0]
    logical['万德编号'] = np.nan

    temp = list(rst.iloc[0])
    temp[0] = 0
    # 更新一下解释变量的先行期数
    for i in range(1, len(zbmc) + 1, 1):
        temp[-i] = 0
    logical['先行期数'] = temp
    logical = logical[['起始年份', '关系', '指标名称', '万德编号', '先行期数']]
    logical.to_excel('./result/' + data_name + 'V1-逻辑表.xlsx', index=False)
    print('逻辑表输出完毕')


def main(data_name):
    """

    :param data_name: 就是zb 指标的名字
    :return: 返回预测的结果？
    """
    before = before_hand(data_name)   # 调用 before_hand 函数
    predict(data_name, before)        # 将 before_hand 函数调用的结果用于 predict 函数


if __name__ == '__main__':

    """
    zb 是指标的意思
    这里采用的方式是 zb + "解释变量筛选结果"
    """

    zb = '地区生产总值-江门'
    main(zb)
