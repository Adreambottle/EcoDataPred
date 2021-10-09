import pandas as pd
import numpy as np
import re
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


def process_y(data, y_name="地区生产总值"):
    """
    处理解释变量y的函数，找到y，默认是地区生产总值
    Question_Remained 添加了地区生产总值
    :param data: 即来源于 zb_excel:sheet 原始数据
    :return:
    """

    # y_name = "地区生产总值"
    y = pd.DataFrame(data.loc[:, ['日期', y_name]])    # 选取['日期'， y_name]两列
    y_drop = y.dropna(subset=[y.columns[1]])          # drop NA的信息 输出的y值
    y1 = pd.DataFrame(columns=y.columns, index=[0])
    y1.loc[0, '日期'] = '特征数据滞后期数'                # 添加首行名称
    y_drop = y1.append(y_drop)                        # 首行是说明，其他是
    return y_drop


def verify_freq(org_data, col_name):
    """
    判断单个指标的频率
    :param org_data: 原始数据
    :param col_name: 需要检测的变量名称
    :return: freq:频率
    """
    # col_name = "人身险"
    # data = pd.read_excel("./data/地区生产总值-中山解释变量筛选结果.xlsx", sheet_name="原始数据")
    # col_name = "六大高耗能行业-石油、煤炭及其他燃料加工业+EX_ST_57_GYNY"
    
    
    # 选出需要用到的指标和日期这两列
    col_names = ['日期', col_name]
    data_used = org_data[col_names]
    data_used = data_used.dropna(subset=[col_name])
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



def process_org_data(data_name, type):
    """
    筛选原始数据
    :param data_name:  数据的名称，一般是 y + 城市名字
    :param type: 是 解释变量 or 先行指标
    :param y_name: y 的名称，需要传入
    :return: 
    """
    
    # type 是"解释变量"还是"先行指标" 确定文件名
    if type == "解释变量":
        type_full = "解释变量筛选结果"
    elif type == "先行指标":
        type_full = "先行性指标筛选结果_自定义相关系数"
    org_data = pd.read_excel('./data/' + data_name + type_full + '.xlsx', sheet_name='原始数据')

    # 这里定义y_name是除了"日期"之外的第一列
    if org_data.columns[0] == '日期':
        y_name = org_data.columns[1]
    else:
        "第0列不是日期，y_name会是第1列吗"
    
    # 如果 org_data 原始数据的第一行是滞后期数，将其删除就好
    if "滞" in org_data.iloc[0, 0]:
        org_data = org_data[1:]

    # 处理y，因为不需要所有的y，按照y进行一次筛选
    y = process_y(org_data, y_name)
    start_month = list(y['日期'])[1]                        # y 开始的那一个月
    i = org_data[org_data['日期'] == start_month].index[0]  # 获得y开始月份的index，注意第一行是说明"特征数据滞后期数"

    # 在第一行添加 '特征数据滞后期数' 
    fst = pd.DataFrame(columns=org_data.columns, index=[0])
    fst.loc[0, '日期'] = '特征数据滞后期数'                    # fst 是第一行是说明将"日期"这一列的首行填充为"特征数据滞后期数"

    # 处理 org_data
    org_data = org_data.iloc[i:]         # 只要从y的首个index，即月份开始截取有效数据
    org_data = fst.append(org_data)      # 将fst说明添加在首行

    date = list(org_data['日期'])         # date 是截取后data的所有信息
    date = [i for i in date if '01月' not in i]
    org_data = org_data[org_data['日期'].isin(date)]  # 去掉1月份的数据，不需要使用1月份用来预测

    return org_data

def get_zb_freq(ser):
    """
    用于处理两种表格里面包含的指标
    :param org_data: 原始表
    :zb_name: 储存了这个表格需要用到的名称
    :return:
    """
    i = 0
    while np.isnan(ser[i]):
        i += 1


def get_zb_qishu(ser):
    """
    :return: 前面空了几格
    """
    # ser = org_data["工业销售产值"]
    i = len(ser)
    while pd.isna(ser[i-1]):
        i -= 1
    return i


def drop_tb_na(tb:pd.DataFrame):
    """
    用于 y_tb 删除月份这一列上的空缺值
    :param tb: y_tb
    :return: 删除之后的表格
    """
    cols = tb.columns
    
    # 因为第一行可能是"特征滞后期数"
    if "滞" in tb.iloc[0, 0]:
        fst_row = tb.iloc[[0],:]
        tb_without_fst = tb.iloc[1:,:]
        tb_without_fst = tb_without_fst.dropna(subset=[cols[1]])
        tb_new = pd.concat([fst_row, tb_without_fst], axis=0)
        return tb_new
    else:
        tb_new = tb.dropna(subset=[cols[1]])
        return tb_new

def drop_no_369_month(tb:pd.DataFrame):
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


def build_table(data_name, type, sheet_Vn, org_data, y_name):
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
    if type == "先行指标":
        type_full = "先行性指标筛选结果_自定义相关系数"

    zb_data = pd.read_excel('./data/' + data_name + type_full +'.xlsx', sheet_name=sheet_Vn)
    zb_names = list(zb_data['指标名称'])              # zb_names指标名称
    org_data = org_data.loc[:, ['日期'] + zb_names]  # 从原始数据中筛选出需要用的指标数据
    org_data = org_data.reset_index()               # 重新定义index

    zb_qishu = list(zb_data['先行期数'])  # qs是期数的意思，指先行期数，是读取的，没有调整单位

    if type == "先行指标":
        zb_zhihou = zb_qishu
    if type == "解释变量":
        zb_zhihou = list(org_data.apply(get_zb_qishu))[1:]



    zb_freq_list = []                               # 记录了每个指标的频率
    zb_start_Y_list = []                            # 记录了每个指标的起始年份
    zb_guanxi_list = []                             # 记录了每个指标的关系
    for x_name in zb_names:
        if x_name != '日期':

            # 这里用来确认 指标 x 的频率
            # 调用了 verify_freq 函数
            freq = verify_freq(org_data, x_name)
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

    if type == "先行期数":
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
        qishu_S = []
        for i, x_name in enumerate(zb_names[S_index]):
            qishu_S.append(zb_qishu[i])
            x_data = list(org_data_S[x_name])
            x_data = [np.nan] * zb_qishu[i] + x_data
            x_data = x_data + [np.nan] * (len_max_S - len(x_data))
            S_tb[x_name] = x_data
        S_tb = pd.DataFrame(S_tb)
        S_tb = S_tb.dropna(subset=['日期'])
        S_tb.iloc[0, 1:] = qishu_S


        # 这里将 解释变量 和 先行指标 两个表格输出
        output_2type_tb(data_name, sheet_Vn, type, y_tb, M_tb, S_tb)

        return logical_tb




    if type == "解释变量":
        # 这里构建 sheet y值
        y_tb = org_data.loc[:, ['日期', y_name]]
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

        qishu_M = []  # 记录先行期数
        zhihou_M = []
        for i, x_name in enumerate(zb_names[M_index]):
            qishu_M.append(zb_qishu[i])  # 记录月度数据的先行期数
            zhihou_M.append(zb_zhihou[i])
            x_data = list(org_data[x_name])  # 将指标数据转化成list，方便添加nan
            x_data = [np.nan] * zb_qishu[i] + x_data  # 根据先行期数在某一指标数据前添加nan
            x_data = x_data + [np.nan] * (len_max_M - len(x_data))  # 在后面补全nan
            M_tb[x_name] = x_data  # 将处理好的某一指标数据添加到 dict M_tb 中

        M_tb = pd.DataFrame(M_tb)  # 将 dictionary 改成 pd.DataFrame
        M_tb = M_tb.dropna(subset=['日期'])  # 删除日期中的空值
        M_tb.iloc[0, 1:] = zhihou_M  # 将第一行改成滞后期数

        # 处理 S 数据，需要先对 org_data 进行3 6 9月的筛选
        org_data_S = drop_no_369_month(org_data)
        if not zb_qishu[S_index].size:
            len_add = 0
        else:
            len_add = max(zb_qishu[S_index])
        len_max_S = len(org_data_S) + len_add
        S_tb = {}
        S_tb['日期'] = list(org_data_S['日期']) + [np.nan] * (len_max_S - len(org_data_S))
        qishu_S = []
        zhihou_S = []
        for i, x_name in enumerate(zb_names[S_index]):
            qishu_S.append(zb_qishu[i])
            zhihou_S.append(zb_zhihou[i])
            x_data = list(org_data_S[x_name])
            x_data = [np.nan] * zb_qishu[i] + x_data
            x_data = x_data + [np.nan] * (len_max_S - len(x_data))
            S_tb[x_name] = x_data
        S_tb = pd.DataFrame(S_tb)
        S_tb = S_tb.dropna(subset=['日期'])
        S_tb.iloc[0, 1:] = zhihou_S


def check_Vn(v_jsbl:list, v_xxzb:list):
    if v_jsbl == v_xxzb:
        return list(zip(v_jsbl, v_xxzb))
    else:
        permutation = []
        for i in v_jsbl:
            for j in v_xxzb:
                permutation.append((i, j))
        return permutation


def get_outcome(data_name_list):
    """
    构建表格
    :param data_name_list: 包括 data_name 的 list
    :param sheet_Vn_list:  包括 sheet_Vn 的 list
    :param type_list:      ['解释变量', '先行指标']
    """
    type_list = ['解释变量', '先行指标']
    for data_name in data_name_list:
        for type in type_list:
            if type == "解释变量":
                type_full = "解释变量筛选结果"
                v_jsbl_list = pd.ExcelFile('./data/' + data_name + type_full + '.xlsx').sheet_names
                v_jsbl_list = [i for i in v_jsbl_list if i[0] in ['v', 'V']]

            elif type == "先行指标":
                type_full = "先行性指标筛选结果_自定义相关系数"
                v_xxzb_list = pd.ExcelFile('./data/' + data_name + type_full + '.xlsx').sheet_names
                v_xxzb_list = [i for i in v_xxzb_list if i[0] in ['v', 'V']]

        Vn_pairs = check_Vn(v_jsbl_list, v_xxzb_list)


        for (v_jsbl, v_xxzb) in Vn_pairs:

            logical_tbs = []

            type = "先行指标"
            org_data = process_org_data(data_name, type)
            y_name = org_data.columns[1]
            logical_tb = build_table(data_name, type, v_jsbl, org_data, y_name)
            logical_tbs.append(logical_tb)

            type = "解释变量"
            org_data = process_org_data(data_name, type)
            y_name = org_data.columns[1]
            logical_tb = build_table(data_name, type, v_xxzb, org_data, y_name)
            # logical_tb = logical_tb[1:]
            logical_tbs.append(logical_tb)

            logical_cat_tb = pd.concat(logical_tbs)
            logical_cat_tb = logical_cat_tb.drop_duplicates(subset=["指标名称"], keep='first')
            logical_cat_tb.to_excel(f"./result/{data_name}-逻辑表-解释变量{v_jsbl}-先行指标{v_xxzb}.xlsx", index=False)
            print(f"Finish output excel", f"{data_name}-逻辑表-解释变量{v_jsbl}-先行指标{v_xxzb}.xlsx")


def main(data_name_list):
    get_outcome(data_name_list)


# def main():
#
#     # 处理中山地区的数据
#     data_name_list = ["地区生产总值-中山"]
#     type_list = ['解释变量', '先行指标']
#     sheet_Vn_list = ['v1', 'v2', 'v3', 'v4', 'v5']
#     get_outcome(data_name_list, sheet_Vn_list, type_list)
#
#
#     # 处理江门地区的数据
#     data_name_list = ["地区生产总值-江门"]
#     type_list = ['解释变量', '先行指标']
#     sheet_Vn_list = ['V1', 'V2', 'V3', 'V4']
#     y_name = '地区生产总值'
#     get_outcome(data_name_list, sheet_Vn_list, type_list)


if __name__ == '__main__':
    main(data_name_list= ["地区生产总值-中山", "地区生产总值-江门"])
