import pandas as pd
import numpy as np

##
# 指标为：
# {1 地区生产总值， 2 第三产业增加值， 3 固定资产投资额， 4 规上工业， 5 进出口总额， 6 社会消费品零售总额}


def generate_time(year, month, k, frequency):
    """
    生成时间
    :param year:
    :param month:
    :param k:
    :param frequency:
    :return:
    """
    if frequency == '月度数据':
        date = []
        for i in range(k):
            if month < 12:
                month = month + 1

            elif month == 12:
                year = year + 1
                month = 1
            date.append('%s年%s月' % (year, month))
        return date

    elif frequency == '季度数据':
        date = []
        for i in range(k):
            if month < 12:
                month = month + 3

            elif month == 12:
                year = year + 1
                month = 3
            date.append('%s年%s月' % (year, month))
        return date


def before_hand(data_name):
    zb = pd.read_excel('./data/' + data_name + '先行性指标筛选结果_自定义相关系数.xlsx', sheet_name='v1')
    data = pd.read_excel('./data/' + data_name + '先行性指标筛选结果_自定义相关系数.xlsx', sheet_name='原始数据')
    # 指标名称的拼音
    zbmc = list(zb['指标名称'])[1:]
    # 期数的拼音
    qs = list(zb['先行期数'])[1:]
    date = list(data['日期'])
    y = pd.DataFrame(data.iloc[:, [0, 1]])
    y.loc[0, '日期'] = '特征数据滞后期数'
    y = y.fillna(method='pad')
    # 判定季度/月度，季度数据只有3/6/9/12月没有10月
    frequency = '季度数据'
    for i in date:
        if '10月' in i:
            frequency = '月度数据'

    # 加上所有先行日期    
    k = max(qs)
    final_date = date[len(date) - 1]
    ym = final_date.replace('月', '').split('年')
    year = int(ym[0])
    month = int(ym[1])
    addtional_date = generate_time(year, month, k, frequency)
    date.extend(addtional_date)

    rst = pd.DataFrame()
    rst['日期'] = date
    for i in range(len(qs)):
        temp = list(data[zbmc[i]])
        temp[0] = qs[i]
        for j in range(qs[i]):
            temp.insert(1, np.nan)
        temp = pd.Series(temp)
        rst[zbmc[i]] = temp
    rst.loc[0, '日期'] = '特征数据滞后期数'

    with pd.ExcelWriter('./result/' + data_name + 'v1先行指标.xlsx') as writer:
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
    print('先行指标输出完毕')
    return rst


def predict(data_name, before):
    zb = pd.read_excel('./data/' + data_name + '解释变量筛选结果.xlsx', sheet_name='v1')
    data = pd.read_excel('./data/' + data_name + '解释变量筛选结果.xlsx', sheet_name='原始数据')
    zbmc = list(zb['指标名称'])[1:]
    date = list(data['日期'])

    y = pd.DataFrame(data.iloc[:, [0, 1]])
    y1 = pd.DataFrame(columns=y.columns, index=[0])
    y1.loc[0, '日期'] = '特征数据滞后期数'
    y = y1.append(y, ignore_index=True, sort=False)
    y = y.fillna(method='pad')
    frequency = '季度数据'
    for i in date:
        if '10月' in i:
            frequency = '月度数据'

    rst = pd.DataFrame()
    date2 = list(date)
    date2.insert(0, '特征数据滞后期数')
    rst['日期'] = date2
    for i in range(len(zbmc)):
        temp = data[zbmc[i]]
        # 计算滞后期数
        num = 0
        for j in range(len(temp) - 1, -1, -1):
            if np.isnan(temp.iloc[j]):
                num = num + 1
            else:
                break
        temp = list(temp)
        temp.insert(0, num)
        temp = pd.Series(temp)
        rst[zbmc[i]] = temp

    rst.drop('日期', axis=1, inplace=True)
    rst = pd.concat([before, rst], axis=1)

    with pd.ExcelWriter('./result/' + data_name + 'v1解释变量.xlsx') as writer:
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
    print('解释变量输出完毕')

    # 输出逻辑表
    logical = pd.DataFrame()
    gx = ['Y值']
    for i in range(len(list(rst.columns)[1:])):
        gx.append('X值')
    temp = list(rst.columns)[1:]
    temp.insert(0, y.columns[1])
    logical['指标名称'] = temp
    logical['关系'] = gx
    logical['起始年份'] = 2015
    logical['万德编号'] = np.nan

    temp = list(rst.iloc[0])
    temp[0] = 0
    # 更新一下解释变量的先行期数
    for i in range(1, len(zbmc) + 1, 1):
        temp[-i] = 0
    logical['先行期数'] = temp
    logical = logical[['起始年份', '关系', '指标名称', '万德编号', '先行期数']]
    logical.to_excel('./result/' + data_name + 'v1-逻辑表.xlsx', index=False)
    print('逻辑表输出完毕')


def main(data_name):
    before = before_hand(data_name)
    predict(data_name, before)


if __name__ == '__main__':
    zb = '进出口总额-中山'
    main(zb)
