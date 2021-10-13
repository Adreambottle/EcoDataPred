import pandas as pd
import numpy as np


# ---Helper functions
def generate_time(year, month, k):
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
    y = pd.DataFrame(data.iloc[:, [0, 1]])
    y_drop = y.dropna(subset=[y.columns[1]])  # 输出的y值。注意中间可能缺了一个月的情况没处理
    y1 = pd.DataFrame(columns=y.columns, index=[0])
    y1.loc[0, '日期'] = '特征数据滞后期数'
    y_drop = y1.append(y_drop)
    return y_drop


# ---处理先行指标
def before_hand(data_name):
    zb = pd.read_excel('./data/' + data_name + '先行性指标筛选结果_自定义相关系数.xlsx', sheet_name='V4')
    data = pd.read_excel('./data/' + data_name + '先行性指标筛选结果_自定义相关系数.xlsx', sheet_name='原始数据')
    # x指标名称
    zbmc = list(zb['指标名称'])[1:]
    # x指标期数
    qs = list(zb['先行期数'])[1:]
    # 处理y
    y = process_y(data)
    start_month = list(y['日期'])[1]
    i = data[data['日期'] == start_month].index[0]  # 获得y开始月份的index

    fst = pd.DataFrame(data.iloc[[0], :])
    fst = pd.DataFrame(columns=fst.columns, index=[0])
    fst.loc[0, '日期'] = '特征数据滞后期数'

    data = data.iloc[i:]
    data = fst.append(data)

    date = list(data['日期'])
    date = [i for i in date if '01月' not in i]
    data = data[data['日期'].isin(date)]  # 去掉1月份

    # 判定季度/月度，季度数据只有3/6/9/12月没有10月
    frequency = '季度数据'
    for i in list(y['日期']):
        if '10月' in i:
            frequency = '月度数据'

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
    before = before_hand(data_name)
    predict(data_name, before)


if __name__ == '__main__':
    zb = '地区生产总值-江门'
    main(zb)
