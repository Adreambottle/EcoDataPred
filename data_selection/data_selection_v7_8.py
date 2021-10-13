import numpy as np
import pandas as pd
import re
from datetime import datetime, timedelta

import cProfile
# cProfile.run('main(data_name_list=["规模以上工业增加值-许昌", "地区生产总值-中山"])')

def zh_to_datetime(date):
    """
    将中文的 str "2020年3月" 转换成 datetime.datetime 格式
    :param date: 中文版本的日期
    :return: 返回 date.datetime
    """
    year, month = re.split('年|月', date)[:2]
    year = int(year)
    if len(month) == 1:
        month = int(month)
    elif (len(month) == 2) and (month[0] == '0'):
        month = int(month[1])
    elif (len(month) == 2) and (month[0] == '1'):
        month = int(month)
    date = datetime(year=year, month=month, day=1)
    return date


def datetime_to_zh(date):
    """
    将 datetime.datetime 转换成中文的 str "2020年3月" 格式
    :param date: datetime.datetime 格式
    :return: 返回中文str
    """
    if date.month < 10:
        Month = "0" + str(date.month)
    else:
        Month = str(date.month)
    return (f"{date.year}年{Month}月")



def generate_time(year, month, k:int):
    """
    用于生成 y年m月 之后 k 期的月份名称
    :param k: 是k期，不包括本月
    :return: 返回的是一个list
    """
    date_index_list = []
    for i in range(k):


        month_total = month + i
        year_add = month_total // 12
        month_real = month_total % 12
        print(month_real)

        date = datetime(year=year+year_add, month=month_real, day=1)
        date_index_list.append(date)

    return date_index_list




def get_zb_zhihou(ser):
    """
    根据结尾判断滞后期数
    :param ser:
    :return:
    """
    # ser = org_data["工业销售产值"]
    i = len(ser)
    while pd.isna(ser[i-1]):
        i -= 1
    return -i


def get_zb_freq(ser):
    """
    判断单个指标的频率
    :param org_data: 原始数据
    :param
    :return: freq:频率
    """
    # ser = org_data["工业销售产值"]

    ser = ser.dropna()
    month_list = ser.index.month

    if (10 not in month_list) and (3 in month_list):
        freq = "Q"
    if (10 not in month_list) and (3 not in month_list):
        freq = "Y"
    if (10 in month_list) and (3 in month_list):
        freq = "M"
    return freq


def get_zb_start_year(ser):
    """
    """
    # ser = org_data["工业销售产值"]

    ser = ser.dropna()
    start_year = ser.index[0].year
    return start_year


def output_2type_tb(data_name, type, y_tb, M_tb, S_tb, v_xx=None, v_js=None):
    """
    生成 解释变量 和 先行指标 两种表格
    :param data_name: 数据的名称，一般是 y + 城市名字
    :param sheet_Vn: 第几张表格
    :param type: 是 解释变量 还是 先行指标
    :param y_tb: 构建 sheet y
    :param M_tb: 构建 sheet 月度数据
    :param S_tb: 构建 sheet 季度数据
    """
    if type == "先行指标":
        name = f'./result/{data_name}-{type}{v_xx}.xlsx'
    if type == "解释变量":
        name = f'./result/{data_name}-{type}{v_js}-先行{v_xx}.xlsx'

    with pd.ExcelWriter(name) as writer:
        blank_tb = pd.DataFrame()
        y_tb.to_excel(writer, sheet_name='y值', index=False)
        blank_tb.to_excel(writer, sheet_name='年度数据', index=False)

        if len(S_tb.columns) <= 1:
            blank_tb.to_excel(writer, sheet_name='季度数据', index=False)
        else:
            S_tb.to_excel(writer, sheet_name='季度数据', index=False)

        if len(M_tb.columns) <= 1:
            blank_tb.to_excel(writer, sheet_name='月度数据', index=False)
        else:
            M_tb.to_excel(writer, sheet_name='月度数据', index=False)

        blank_tb.to_excel(writer, sheet_name='周度数据', index=False)
        blank_tb.to_excel(writer, sheet_name='日度数据', index=False)

    print("Finish output excel ", name)



def process_org_data(data_name, type, sheet_Vn):
    """
    筛选原始数据
    :param data_name:  数据的名称，一般是 y + 城市名字
    :param type: 是 解释变量 or 先行指标
    :param sheet_Vn 读取的 解释变量 or 先行指标 是哪个 sheet
    :return:
    """


    # data_name = '地区生产总值-江门'
    # sheet_Vn = 'V1'
    # type = '先行指标'
    # y_name = '地区生产总值'
    # org_data = process_org_data(data_name, type, sheet_Vn)

    # type 是"解释变量"还是"先行指标" 确定文件名
    if type == "解释变量":
        type_full = "解释变量筛选结果"
    elif type == "先行指标":
        type_full = "先行性指标筛选结果_自定义相关系数"

    org_data = pd.read_excel('./data/' + data_name + type_full + '.xlsx', sheet_name='原始数据')

    if "滞后" in org_data.iloc[0,0]:
        org_data = org_data[1:]
    org_data['日期'] = org_data['日期'].apply(zh_to_datetime)
    org_data.set_index(["日期"], inplace=True)

    # print(type, sheet_Vn)
    zb_use = list(pd.read_excel('./data/' + data_name + type_full + '.xlsx', sheet_name=sheet_Vn)["指标名称"])
    org_data = org_data[zb_use]

    return org_data



def build_table(data_name, type, sheet_Vn):
    """
    构建逻辑表格
    :param data_name: 表格的名字 '地区生产总值-江门'
    :param type: "解释变量" "先行指标"
    :param sheet_Vn: 是第几个 sheet
    :return: 返回逻辑表
    """

    # eg
    # data_name = '地区生产总值-江门'
    # sheet_Vn = 'V1'
    # type = '先行指标'
    # type = '解释变量'
    # y_name = '地区生产总值'

    if type == "解释变量":
        type_full = "解释变量筛选结果"
    if type == "先行指标":
        type_full = "先行性指标筛选结果_自定义相关系数"

    org_data = process_org_data(data_name, type, sheet_Vn)
    org_data = org_data[org_data.index.month != 1]
    # 这里的都是加上日期这一列的

    # 记录 org_data 的先行期数，读取自表格
    zb_qishu = pd.read_excel('./data/' + data_name + type_full + '.xlsx', sheet_name=sheet_Vn)["先行期数"]


    # 记录 org_data 的特征滞后期数，这个是根据数据最新的日期计算的
    zb_zhihou = org_data.apply(get_zb_zhihou)
    zb_zhihou = zb_zhihou - zb_zhihou[0]

    zb_freq = org_data.apply(get_zb_freq)
    zb_start_year = org_data.apply(get_zb_start_year)

    zb_names = org_data.columns

    zb_guanxi = ["X值"] * len(zb_names)
    zb_guanxi[0] = "Y值"

    # # 这里转换成 np.array 的格式，因为后面要花式索引
    # zb_names = np.array(zb_names)
    # zb_qishu = np.array(zb_qishu)
    # zb_freq = np.array(zb_freq)


    # 这里最好不要分开写

    logical_tb = pd.DataFrame(columns=['起始年份', '关系', '指标名称', 'ID', 'ID_city', '先行期数', '频率'])
    logical_tb['起始年份'] = list(zb_start_year)
    logical_tb['关系'] = list(zb_guanxi)
    logical_tb['指标名称'] = list(zb_names)
    logical_tb['先行期数'] = list(zb_qishu)
    logical_tb['滞后期数'] = list(zb_zhihou)
    logical_tb['频率'] = list(zb_freq)
    logical_tb = logical_tb.reset_index(drop=True)

    return logical_tb, org_data


def add_first_row(tb, ser):
    """
    在 table 的第一行添加之后期数这一行
    :param tb:
    :param ser:
    :return:
    """
    new_tb = pd.DataFrame(columns=tb.columns)
    first_row = ["特征滞后期数"] + list(ser)
    new_tb.loc[0,:] = first_row
    new_tb = pd.concat([new_tb, tb])
    return new_tb


def add_index(data:pd.DataFrame, k):
    """
    因为需要根据先行期数将数据进行shift操作
    需要将后续的日期补全
    :param data:
    :param k:
    :return:
    """
    # data = org_data
    data_index = data.index
    now = data_index[-1]
    index_add = pd.date_range(start=now, periods=k+1, freq='MS')[1:]

    for i in range(len(index_add)):
        data.loc[index_add[i]] = np.nan
    return data


def get_season_month(data):
    """
    只取月份是 3， 6， 9， 12 的数据，作为季度数据
    :param data:
    :return:
    """
    index = data.index
    id_369 = []
    for id in index:
        if id.month in [3, 6, 9, 12]:
            id_369.append(id)
    return data.loc[id_369]


def drop_y(data:pd.DataFrame, y_name):
    """
    删除变量中的 y
    :param data:
    :param y_name:
    :return:
    """
    if y_name in data.columns:
        data = data.drop([y_name], axis=1)
    return data


def drop_1_month(data):
    """
    删除 1 月份的数据
    :param data:
    :return:
    """
    data = data[data.index.month != 1]
    return data


def drop_na_tail(data):
    """
    删去尾部全部都是na的数据
    :param data:
    :return:
    """
    if data.shape[0] > 0:
        i = data.shape[0] - 1
        while all(pd.isna(data.iloc[i, :])):
            data = data[:i]
            i = i - 1
    return data


def build_xxzb(data_name, v_xx, type, logical_tb, org_data, y_name):
    """
    构建先行指标这个 table
    :param data_name: 数据的名称
    :param v_xx: 先行指标的 sheet Number
    :param type: "先行指标" or "解释变量"
    :param logical_tb: 传入的逻辑表
    :param org_data: 处理好的原始数据
    :param y_name: 作为y的变量
    :return: 先行指标的季度和月度表格，需要传给构建解释变量的函数
    """

    # logical_tb, org_data = build_table(data_name, type, sheet_Vn)
    org_data = drop_1_month(org_data)
    y_tb = org_data[y_name]
    y_tb = y_tb.dropna()
    y_start_month = y_tb.index[0]
    y_freq = get_zb_freq(y_tb)


    y_tb_ad = y_tb.reset_index()
    y_tb_ad["日期"] = y_tb_ad["日期"].apply(datetime_to_zh)
    y_tb_ad = add_first_row(y_tb_ad, [np.nan])

    xianxing_max = max(logical_tb['先行期数']) * 3
    org_data = add_index(org_data, xianxing_max + 10)
    org_data = drop_1_month(org_data)


    M_zb = list(logical_tb[logical_tb['频率'] == 'M']['指标名称'])
    M_xianxing = list(logical_tb[logical_tb['频率'] == 'M']['先行期数'])

    M_tb = pd.DataFrame(index=org_data.index)

    for i, zb in enumerate(M_zb):
        ser = org_data[zb]

        if y_freq == "Q":
            ser = ser.shift(M_xianxing[i] * 3)
        if y_freq == "M":
            ser = ser.shift(M_xianxing[i])
        M_tb[zb] = ser

    M_tb = M_tb.loc[y_start_month:]
    try:  # 我也不知道为什么会报错，那就这样吧
        M_tb = drop_na_tail(M_tb)
    except:
        M_tb = M_tb

    M_tb_ad = M_tb.reset_index(drop=False)
    M_tb_ad["日期"] = M_tb_ad["日期"].apply(datetime_to_zh)
    if y_freq == "Q":
        M_xianxing = [i * 3 for i in M_xianxing]

    M_xx = add_first_row(M_tb_ad, M_xianxing)
    M_xx = drop_y(M_xx, y_name)



    S_zb = list(logical_tb[logical_tb['频率'] == 'S']['指标名称'])
    S_xianxing = list(logical_tb[logical_tb['频率'] == 'S']['先行期数'])

    S_tb = pd.DataFrame(index=org_data.index)
    for i, zb in enumerate(S_zb):
        ser = org_data[zb]
        ser = ser.shift(S_xianxing[i])
        S_tb[zb] = ser

    S_tb = S_tb.loc[y_start_month:]
    try:
        S_tb = drop_na_tail(S_tb)
    except:
        S_tb = S_tb

    S_tb = get_season_month(S_tb)
    S_tb_ad = S_tb.reset_index(drop=False)
    S_tb_ad["日期"] = S_tb_ad["日期"].apply(datetime_to_zh)
    S_xx = add_first_row(S_tb_ad, S_xianxing)
    S_xx = drop_y(S_xx, y_name)



    output_2type_tb(data_name, type, y_tb_ad, M_xx, S_xx, v_xx=v_xx, v_js=None)
    return M_xx, S_xx




def build_jsbl(data_name, v_xx, v_js, type, logical_tb, org_data, y_name, M_xx, S_xx):
    """
    构建解释变量这个表格
    :param data_name: 表格的名称
    :param v_xx: 先行指标的 sheet_number
    :param v_js: 解释变量的 sheet_number
    :param type: "解释变量" or "先行指标"
    :param logical_tb: 传入的逻辑表
    :param org_data: 传入的处理好的原始数据
    :param y_name: 变量 y 的名字
    :param M_xx: 先行指标表格的月份指标 table M
    :param S_xx: 线性指标表格的月份指标 table S
    :return:
    """

    y_tb = org_data[y_name]
    y_tb = drop_1_month(y_tb)
    y_tb = y_tb.dropna()
    y_start_month = y_tb.index[0]
    y_tb_ad = y_tb.reset_index()
    y_tb_ad["日期"] = y_tb_ad["日期"].apply(datetime_to_zh)
    y_tb_ad = add_first_row(y_tb_ad, [np.nan])


    # org_data = org_data.loc[y_start_month:]

    M_zb = list(logical_tb[logical_tb['频率'] == 'M']['指标名称'])
    M_zhihou = list(logical_tb[logical_tb['频率'] == 'M']['滞后期数'])

    org_data = drop_1_month(org_data)
    M_tb = org_data[M_zb]
    M_tb = M_tb.loc[y_start_month:]
    # M_tb = drop_1_month(M_tb)
    M_tb_ad = M_tb.reset_index(drop=False)
    M_tb_ad["日期"] = M_tb_ad["日期"].apply(datetime_to_zh)
    M_js = add_first_row(M_tb_ad, M_zhihou)
    M_js = drop_y(M_js, y_name)



    S_zb = list(logical_tb[logical_tb['频率'] == 'S']['指标名称'])
    S_zhihou = list(logical_tb[logical_tb['频率'] == 'S']['滞后期数'])

    S_tb = org_data[S_zb]
    S_tb = S_tb.loc[y_start_month:]
    S_tb = get_season_month(S_tb)

    S_tb_ad = S_tb.reset_index(drop=False)
    S_tb_ad["日期"] = S_tb_ad["日期"].apply(datetime_to_zh)
    S_js = add_first_row(S_tb_ad, S_zhihou)
    S_js = drop_y(S_js, y_name)



    M_js = M_js[['日期'] + list(set(M_js.columns) - set(M_xx.columns))]
    S_js = S_js[['日期'] + list(set(S_js.columns) - set(S_xx.columns))]

    M_tb_cb = pd.merge(M_xx, M_js, how="outer", on="日期")
    S_tb_cb = pd.merge(S_xx, S_js, how="outer", on="日期")

    output_2type_tb(data_name, type, y_tb_ad, M_tb_cb, S_tb_cb, v_xx, v_js)




def check_Vn(v_xx_list: list, v_js_list: list):
    """
    查看先行指标和解释变量的 sheet number 能否对应上
    如果能对应上只是用单独的 sheet number
    如果对应不上则使用 排列组合
    :param v_xx_list:
    :param v_js_list:
    :return: 返回一个先行指标和解释变量的pair
    """
    if v_js_list == v_xx_list:
        return list(zip(v_js_list, v_xx_list))

    else:
        permutation = []
        for i in v_xx_list:
            for j in v_js_list:
                permutation.append((i, j))
        return permutation



def get_outcome(data_name_list):
    """
    构建表格
    :param data_name_list: 包括 data_name 的 list
    """

    # data_name_list = ["地区生产总值-北京"]
    for data_name in data_name_list:

        # data_name = "地区生产总值-北京"
        type_full = "解释变量筛选结果"
        v_js_list = pd.ExcelFile('./data/' + data_name + type_full + '.xlsx').sheet_names
        v_js_list = [i for i in v_js_list if i[0] in ['v', 'V']]


        type_full = "先行性指标筛选结果_自定义相关系数"
        v_xx_list = pd.ExcelFile('./data/' + data_name + type_full + '.xlsx').sheet_names
        v_xx_list = [i for i in v_xx_list if i[0] in ['v', 'V']]

        Vn_pairs = check_Vn(v_xx_list, v_js_list)

        for (v_xx, v_js) in Vn_pairs:
            # v_js = "v1"
            # v_xx = "v1"


            logical_tbs = []


            type = "先行指标"
            logical_tb, org_data = build_table(data_name, type, v_xx)
            y_name = org_data.columns[0]
            M_xx, S_xx = build_xxzb(data_name, v_xx, type, logical_tb, org_data, y_name)
            logical_tbs.append(logical_tb)

            type = "解释变量"
            logical_tb, org_data = build_table(data_name, type, v_js)
            y_name = org_data.columns[0]
            build_jsbl(data_name, v_xx, v_js, type, logical_tb, org_data, y_name, M_xx, S_xx)

            logical_tb = logical_tb[1:]
            logical_tbs.append(logical_tb)


            logical_cat_tb = pd.concat(logical_tbs)
            logical_cat_tb = logical_cat_tb.drop_duplicates(subset=["指标名称"], keep='first')
            logical_cat_tb = logical_cat_tb.drop(["滞后期数"], axis=1)
            logical_tb_name = f"./result/{data_name}-逻辑表-解释变量{v_js}-先行指标{v_xx}.xlsx"
            logical_cat_tb.to_excel(logical_tb_name, index=False)
            print(f"Finish output excel ", logical_tb_name)


def main(data_name_list):
    get_outcome(data_name_list)

if __name__ == '__main__':
    main(data_name_list=["规模以上工业增加值-许昌", "地区生产总值-中山"])
