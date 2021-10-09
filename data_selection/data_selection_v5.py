import numpy as np
import pandas as pd
import re


def get_zb_zhihou(ser):
    """
    :return: 前面空了几格
    """
    # ser = org_data["工业销售产值"]
    i = len(ser)
    while pd.isna(ser[i-1]):
        i -= 1
    return i



def drop_tb_na(tb: pd.DataFrame):
    """
    用于 y_tb 删除月份这一列上的空缺值
    :param tb: y_tb
    :return: 删除之后的表格
    """
    cols = tb.columns

    # 因为第一行可能是"特征滞后期数"
    if "滞" in tb.iloc[0, 0]:
        fst_row = tb.iloc[[0], :]
        tb_without_fst = tb.iloc[1:, :]
        tb_without_fst = tb_without_fst.dropna(subset=[cols[1]])
        tb_new = pd.concat([fst_row, tb_without_fst], axis=0)
        return tb_new
    else:
        tb_new = tb.dropna(subset=[cols[1]])
        return tb_new


def drop_no_369_month(tb: pd.DataFrame):
    """
    删去非 3月 6月 9月 的行，这里用于季度数据
    :param tb: 是月份齐全的表格
    :return: 删除非 3月 6月 9月之后的表格
    """
    month_col = list(tb['日期'].str[-3:-1])
    need_delete = ['01', '02', '04', '05', '07', '08', '10', '11']
    month_index = []
    for month in month_col:
        if month not in need_delete:
            month_index.append(True)
        else:
            month_index.append(False)
    tb_new = tb[month_index]
    return tb_new


def output_2type_tb(data_name, sheet_Vn, type, y_tb, M_tb, S_tb):
    """
    生成 解释变量 和 先行指标 两种表格
    :param data_name: 数据的名称，一般是 y + 城市名字
    :param sheet_Vn: 第几张表格
    :param type: 是 解释变量 还是 先行指标
    :param y_tb: 构建 sheet y
    :param M_tb: 构建 sheet 月度数据
    :param S_tb: 构建 sheet 季度数据
    """
    with pd.ExcelWriter(f'./result/{data_name}-{type}-{sheet_Vn}.xlsx') as writer:
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

    print("Finish output excel ", f'{data_name}-{type}-{sheet_Vn}.xlsx')



def get_freq(ser, org_data):
    """
    判断单个指标的频率
    :param org_data: 原始数据
    :param
    :return: freq:频率
    """
    # ser = org_data["工业销售产值"]

    data_used = pd.concat([org_data['日期'], ser], axis=1)
    data_used = data_used.dropna()
    month_str_list = list(data_used["日期"])
    month_int_list = []
    for month_str in month_str_list:
        """
        有的数据中用的是 "2016年05月"
        有的数据中用的是 "2016年5月"
        所以要正则提取月份
        """
        month = re.split('年|月', month_str)[1]
        if len(month) == 1:
            month = int(month)
        elif (len(month) == 2) and (month[0] == '0'):
            month = int(month[1])
        elif (len(month) == 2) and (month[0] == '1'):
            month = int(month)
        month_int_list.append(month)

    # 如果10月不在，3月在，就是 S 季
    # 如果10月不在，3月也不在，就是 Y 年
    # 如果10月在，3月也在，就是 M 月
    if (10 not in month_int_list) and (3 in month_int_list):
        freq = "S"
    if (10 not in month_int_list) and (3 not in month_int_list):
        freq = "Y"
    if (10 in month_int_list) and (3 in month_int_list):
        freq = "M"
    return freq



def get_start_year(ser, org_data):
    """
    判断单个指标的频率
    :param org_data: 原始数据
    :param
    :return: freq:频率
    """
    # ser = org_data["工业销售产值"]

    data_used = pd.concat([org_data['日期'], ser], axis=1)
    data_used = data_used.dropna()
    start_year = re.split('年|月', data_used.iloc[0, 0])[0]
    return start_year




def process_org_data(data_name, type, sheet_Vn):
    """
    筛选原始数据
    :param data_name:  数据的名称，一般是 y + 城市名字
    :param type: 是 解释变量 or 先行指标
    :param y_name: y 的名称，需要传入
    :return:
    """

    # eg
    # data_name = '地区生产总值-江门'
    # sheet_Vn = 'V4'
    # type = '解释变量'
    # y_name = '地区生产总值'
    # org_data = process_org_data(data_name, '地区生产总值')

    # type 是"解释变量"还是"先行指标" 确定文件名
    if type == "解释变量":
        type_full = "解释变量筛选结果"
    elif type == "先行指标":
        type_full = "先行性指标筛选结果_自定义相关系数"


    org_data = pd.read_excel('./data/' + data_name + type_full + '.xlsx', sheet_name='原始数据')
    zb_use = list(pd.read_excel('./data/' + data_name + type_full + '.xlsx', sheet_name=sheet_Vn)["指标名称"])
    org_data = org_data[['日期']+zb_use]

    return org_data



def build_table(data_name, type, sheet_Vn, y_name):

    # eg
    # data_name = '地区生产总值-江门'
    # sheet_Vn = 'V1'
    # type = '先行指标'
    # type = '解释变量'
    # y_name = '地区生产总值'

    if type == "解释变量":
        type_full = "解释变量筛选结果"
    elif type == "先行指标":
        type_full = "先行性指标筛选结果_自定义相关系数"

    org_data = process_org_data(data_name, type, sheet_Vn)


    # 这里的都是加上日期这一列的

    # 记录 org_data 的先行期数，读取自表格
    zb_qishu = pd.read_excel('./data/' + data_name + type_full + '.xlsx', sheet_name=sheet_Vn)["先行期数"]
    zb_qishu = ["None"] + list(zb_qishu)

    # 记录 org_data 的特征滞后期数，这个是根据数据最新的日期计算的
    zb_zhihou = org_data.apply(get_zb_zhihou)
    zb_zhihou = zb_zhihou - zb_zhihou[1]

    # 记录频率
    zb_freq = org_data.iloc[:, 1:].apply(get_freq, org_data=org_data)
    zb_freq = ["None"] + list(zb_freq)

    # 记录起始年份
    zb_start_Y = org_data.iloc[:, 1:].apply(get_start_year, org_data=org_data)
    zb_start_Y = ["None"] + list(zb_start_Y)

    zb_names = list(org_data.columns)

    zb_guanxi = ["X值"] * len(zb_names)
    zb_guanxi[0] = "None"
    zb_guanxi[1] = "Y值"

    # 这里转换成 np.array 的格式，因为后面要花式索引
    zb_names = np.array(zb_names)
    zb_qishu = np.array(zb_qishu)
    zb_freq = np.array(zb_freq)

    # 这里用于构建逻辑表
    logical_tb = pd.DataFrame(columns=['起始年份', '关系', '指标名称', 'ID', 'ID_city', '先行期数', '频率'])
    logical_tb['起始年份'] = zb_start_Y
    logical_tb['关系'] = zb_guanxi
    logical_tb['指标名称'] = zb_names
    logical_tb['先行期数'] = zb_qishu
    logical_tb['频率'] = zb_freq

    logical_tb = logical_tb[1:]



    # 这里构建 sheet y值
    y_tb = org_data.loc[:,['日期', y_name]]
    y_tb = drop_tb_na(y_tb)
    y_start_index = get_zb_zhihou(org_data[y_name])

    # 先按照 zb_freq_list 区分 M 和 S
    M_name = []
    S_name = []
    M_index = []
    S_index = []
    for i, x_name in enumerate(zb_names):
        if x_name != y_name:
            if zb_freq[i] == 'M':
                M_name.append(x_name)
                M_index.append(i)
            elif zb_freq[i] == 'S':
                S_name.append(x_name)
                S_index.append(i)


    # 处理 M 的数据
    M_tb = org_data[["日期"] + M_name]
    data_start_index = y_start_index - max(zb_qishu[M_index])
    M_tb = M_tb[data_start_index:]

    M_qishu = pd.Series(np.hstack(["特征数据滞后期数", zb_qishu[M_index]])).to_frame().T
    M_qishu.columns = M_tb.columns

    M_tb = pd.concat([M_qishu, M_tb])








    # 因为需要向后推移数据，所以记录可能出现的推移的长度
    # zb_qishu[M_index] 有可能是空 array，即没有月度指标的情况
    if not zb_qishu[M_index].size:
        len_add = 0
    else:
        len_add = max(zb_qishu[M_index])

    # 月度指标最终的长度是 元数据的长度 + 推移的长度
    len_max_M = len(org_data) + len_add

    # 新建一个dictionary用来储存每个指标，因为长度可能不一致，不能直接变成 DataFrame
    M_tb = {}
    M_tb['日期'] = list(org_data['日期']) + [np.nan] * (len_max_M - len(org_data))



    qishu_M = []      # 记录先行期数
    for i, x_name in enumerate(zb_names[M_index]):
        qishu_M.append(zb_qishu[i])                              # 记录月度数据的先行期数
        x_data = list(org_data[x_name])                          # 将指标数据转化成list，方便添加nan
        x_data = [np.nan]*zb_qishu[i] + x_data                   # 根据先行期数在某一指标数据前添加nan
        x_data = x_data + [np.nan] * (len_max_M - len(x_data))   # 在后面补全nan
        M_tb[x_name] = x_data                                    # 将处理好的某一指标数据添加到 dict M_tb 中

    M_tb = pd.DataFrame(M_tb)               # 将 dictionary 改成 pd.DataFrame
    M_tb = M_tb.dropna(subset=['日期'])      # 删除日期中的空值
    M_tb.iloc[0,1:] = qishu_M               # 将第一行改成先行期数




    # 处理 S 数据，需要先对 org_data 进行3 6 9月的筛选
    org_data_S = drop_no_369_month(org_data)
    if not zb_qishu[S_index].size:
        len_add = 0
    else:
        len_add = max(zb_qishu[S_index])
    len_max_S = len(org_data_S) + len_add
    S_tb = {}
    S_tb['日期'] = list(org_data_S['日期']) + [np.nan] * (len_max_S - len(org_data_S))
    qs_S = []
    for i, x_name in enumerate(zb_names[S_index]):
        qs_S.append(zb_qishu[i])
        x_data = list(org_data_S[x_name])
        x_data = [np.nan] * zb_qishu[i] + x_data
        x_data = x_data + [np.nan] * (len_max_S - len(x_data))
        S_tb[x_name] = x_data
    S_tb = pd.DataFrame(S_tb)
    S_tb = S_tb.dropna(subset=['日期'])
    S_tb.iloc[0, 1:] = qs_S


    # 这里将 解释变量 和 先行指标 两个表格输出
    output_2type_tb(data_name, sheet_Vn, type, y_tb, M_tb, S_tb)

    return logical_tb






def build_table_old(data_name, type, sheet_Vn, y_name):
    """
    对于每个 data_name 生成和本地保存 先行指标 和 解释变量 这两个表
                     返回逻辑表，因为要在后面进行拼接

    :param data_name: 数据的名称，一般是 y-城市名字 eg:'地区生产总值-江门'
    :param sheet_Vn: 来源于哪张sheet eg:'V4'
    :param org_data: 在内存中已经读取的 org_data。来源于 process_org_data 的返回值
    :param y_name: 这里就是 地区生产总值
    :return:
    """

    # eg
    # data_name = '地区生产总值-江门'
    # sheet_Vn = 'V4'
    # type = '解释变量'
    # y_name = '地区生产总值'
    # org_data = process_org_data(data_name, '地区生产总值')




    if type == "解释变量":
        type_full = "解释变量筛选结果"
    elif type == "先行指标":
        type_full = "先行性指标筛选结果_自定义相关系数"

    zb_data = pd.read_excel('./data/' + data_name + type_full +'.xlsx', sheet_name=sheet_Vn)


    zb_names = list(zb_data['指标名称'])              # zb_names指标名称
    zb_qishu = list(zb_data['先行期数'])              # qs是期数的意思，指先行期数，是读取的，没有调整单位

    org_data = org_data.loc[:, ['日期'] + zb_names]  # 从原始数据中筛选出需要用的指标数据
    org_data = org_data.reset_index()               # 重新定义index

    zb_freq_list = []                               # 记录了每个指标的频率
    zb_start_Y_list = []                            # 记录了每个指标的起始年份
    zb_guanxi_list = []                             # 记录了每个指标的关系
    for x_name in zb_names:
        if x_name != '日期':

            # 这里用来确认 指标 x 的频率
            # 调用了 verify_freq 函数
            freq = get_freq(org_data, x_name)
            zb_freq_list.append(freq)


            # 这里用来确认指标的关系，是X还是Y
            if x_name == y_name:
                zb_guanxi_list.append("Y值")
            else:
                zb_guanxi_list.append("X值")


            # 这里用来记录指标的起始年份
            # start_ix 会从0递增，直到出现数据，不再是nan
            start_ix = 0
            while np.isnan(org_data[x_name][start_ix]):
                start_ix += 1
            start_Y = org_data['日期'][start_ix][:4]
            zb_start_Y_list.append(start_Y)


    # 这里好像不需要将季度的先行期数*3
    # qs_real = zb_qishu.copy()
    # for i, name in enumerate(zb_names):
    #     if zb_freq_list[i] == 'S':
    #         qs_real[i] = zb_qishu[i] * 3

    # 这里转换成 np.array 的格式，因为后面要花式索引
    zb_names = np.array(zb_names)
    zb_qishu = np.array(zb_qishu)
    zb_freq_list = np.array(zb_freq_list)


    # 这里用于构建逻辑表
    logical_tb = pd.DataFrame(columns=['起始年份', '关系', '指标名称', 'ID', 'ID_city', '先行期数', '频率'])
    logical_tb['起始年份'] = zb_start_Y_list
    logical_tb['关系'] = zb_guanxi_list
    logical_tb['指标名称'] = zb_names
    logical_tb['先行期数'] = zb_qishu
    logical_tb['频率'] = zb_freq_list

    """
    这里构建 先行指标 和 解释变量两个表格
    """

    # 这里构建 sheet y值
    y_tb = org_data.loc[:,['日期', y_name]]
    y_tb = drop_tb_na(y_tb)

    # 先按照 zb_freq_list 区分 M 和 S
    M_index = []
    S_index = []
    for i, x_name in enumerate(zb_names):
        if x_name != y_name:
            if zb_freq_list[i] == 'M':
                M_index.append(i)
            elif zb_freq_list[i] == 'S':
                S_index.append(i)

    # 处理 M 的数据

    # 因为需要向后推移数据，所以记录可能出现的推移的长度
    # zb_qishu[M_index] 有可能是空 array，即没有月度指标的情况
    if not zb_qishu[M_index].size:
        len_add = 0
    else:
        len_add = max(zb_qishu[M_index])

    # 月度指标最终的长度是 元数据的长度 + 推移的长度
    len_max_M = len(org_data) + len_add

    # 新建一个dictionary用来储存每个指标，因为长度可能不一致，不能直接变成 DataFrame
    M_tb = {}
    M_tb['日期'] = list(org_data['日期']) + [np.nan] * (len_max_M - len(org_data))

    qishu_M = []      # 记录先行期数
    for i, x_name in enumerate(zb_names[M_index]):
        qishu_M.append(zb_qishu[i])                              # 记录月度数据的先行期数
        x_data = list(org_data[x_name])                          # 将指标数据转化成list，方便添加nan
        x_data = [np.nan]*zb_qishu[i] + x_data                   # 根据先行期数在某一指标数据前添加nan
        x_data = x_data + [np.nan] * (len_max_M - len(x_data))   # 在后面补全nan
        M_tb[x_name] = x_data                                    # 将处理好的某一指标数据添加到 dict M_tb 中

    M_tb = pd.DataFrame(M_tb)               # 将 dictionary 改成 pd.DataFrame
    M_tb = M_tb.dropna(subset=['日期'])      # 删除日期中的空值
    M_tb.iloc[0,1:] = qishu_M               # 将第一行改成先行期数



    # 处理 S 数据，需要先对 org_data 进行3 6 9月的筛选
    org_data_S = drop_no_369_month(org_data)
    if not zb_qishu[S_index].size:
        len_add = 0
    else:
        len_add = max(zb_qishu[S_index])
    len_max_S = len(org_data_S) + len_add
    S_tb = {}
    S_tb['日期'] = list(org_data_S['日期']) + [np.nan] * (len_max_S - len(org_data_S))
    qs_S = []
    for i, x_name in enumerate(zb_names[S_index]):
        qs_S.append(zb_qishu[i])
        x_data = list(org_data_S[x_name])
        x_data = [np.nan] * zb_qishu[i] + x_data
        x_data = x_data + [np.nan] * (len_max_S - len(x_data))
        S_tb[x_name] = x_data
    S_tb = pd.DataFrame(S_tb)
    S_tb = S_tb.dropna(subset=['日期'])
    S_tb.iloc[0, 1:] = qs_S


    # 这里将 解释变量 和 先行指标 两个表格输出
    output_2type_tb(data_name, sheet_Vn, type, y_tb, M_tb, S_tb)

    return logical_tb