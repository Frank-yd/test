import pandas as pd
import os
# try:
#     import global_setting.global_dic as glv
# except:
#     print("global_setting cannot be found, pls check")
import global_setting.global_dic as glv
import warnings
warnings.filterwarnings("ignore")
import global_tools_func.global_tools as gt

#abc——colomuns

def extract(product_code,time,filelist):
    # 文件名检索 日期为特征值 定向提取单一文件 但输出为list
    time = gt.intdate_transfer(time)
    list1 = []
    list2 = []
    for i in filelist:
        if product_code == 'SVX619':
            time = gt.strdate_transfer(time)
        if product_code =='SLA626':
            time = gt.strdate_transfer(time)
        if product_code in i:
            list1.append(i)
    for j in list1:
        if time in j:
            list2.append(j)
    if len(list1) == 0 or len(list2) == 0:
        print(product_code+"在"+time+"存在数据缺失。")
    return list2


def SF1000_option_name_transfer(option_name):
    # 输入为 盛丰1000下期权产品名
    # 输出为 对应产品代码
    if option_name[0:5] == '沪深300':
        if option_name[5] == '沽':
            option_name_new = 'IO24' + option_name[6:8] + '-P-' + option_name[-4:] + '.CFE'
        else:
            option_name_new = 'IO24' + option_name[6:8] + '-C-' + option_name[-4:] + '.CFE'
    elif option_name[0:4] == '上证50':
        if option_name[4] == '沽':
            option_name_new = 'HO24' + option_name[5:7] + '-P-' + option_name[-4:] + '.CFE'
        else:
            option_name_new = 'HO24' + option_name[5:7] + '-C-' + option_name[-4:] + '.CFE'
    elif option_name[0:6] == '中证1000':
        if option_name[6] == '沽':
            option_name_new = 'MO24' + option_name[7:9] + '-P-' + option_name[-4:] + '.CFE'
        else:
            option_name_new = 'MO24' + option_name[7:9] + '-C-' + option_name[-4:] + '.CFE'
    else:
        print('期权名字格式特殊转换失败，请手动改正')
        option_name_new = option_name
    return option_name_new


def option_name_transfer_NJ300(option_name):
    # 输入为 盛丰1000下期权产品名
    # 输出为 对应产品代码
    if option_name[0:5] == '沪深300':
        if option_name[5] == '沽':
            if option_name[6] == '1':
                option_name_new = 'IO241' + option_name[7] + '-P-' + option_name[-4:] + '.CFE'
            else:
                option_name_new = 'IO240' + option_name[6] + '-P-' + option_name[-4:] + '.CFE'
        else:
            if option_name[6] == '1':
                option_name_new = 'IO241' + option_name[7] + '-C-' + option_name[-4:] + '.CFE'
            else:
                option_name_new = 'IO240' + option_name[6] + '-C-' + option_name[-4:] + '.CFE'
    elif option_name[0:4] == '上证50':
        if option_name[5] == '沽':
            if option_name[5] == '1':
                option_name_new = 'IO241' + option_name[6] + '-P-' + option_name[-4:] + '.CFE'
            else:
                option_name_new = 'IO240' + option_name[5] + '-P-' + option_name[-4:] + '.CFE'
        else:
            if option_name[5] == '1':
                option_name_new = 'IO241' + option_name[6] + '-C-' + option_name[-4:] + '.CFE'
            else:
                option_name_new = 'IO240' + option_name[5] + '-C-' + option_name[-4:] + '.CFE'
    elif option_name[0:6] == '中证1000':
        if option_name[6] == '沽':
            if option_name[7] == '1':
                option_name_new = 'IO241' + option_name[8] + '-P-' + option_name[-4:] + '.CFE'
            else:
                option_name_new = 'IO240' + option_name[7] + '-P-' + option_name[-4:] + '.CFE'
        else:
            if option_name[7] == '1':
                option_name_new = 'IO241' + option_name[8] + '-C-' + option_name[-4:] + '.CFE'
            else:
                option_name_new = 'IO240' + option_name[7] + '-C-' + option_name[-4:] + '.CFE'
    else:
        print('期权名字格式特殊转换失败，请手动改正')
        option_name_new = option_name
    return option_name_new


def option_name_transfer(option_name):
    # 输入为 盛丰1000外其他产品下期权名
    # 输出为 对应产品代码
    if option_name[0:5] == '沪深300':
        option_name_new = 'IO' + option_name[-11:] + '.CFE'
    elif option_name[0:4] == '上证50':
        option_name_new = 'HO' + option_name[-11:] + '.CFE'
    elif option_name[0:6] == '中证1000':
        option_name_new = 'MO' + option_name[-11:] + '.CFE'
    else:
        print('期权名字格式特殊转换失败，请手动改正')
        option_name_new = option_name
    return option_name_new


def RR500(file_name):
    # 输入为 瑞锐中证500文件名
    # 输出为 解析后的四级估值表 分为股票 可转债 期货 期权 国债五个模块
    folder_path = glv.get('folder_path')
    for i in file_name:
        current_file = i
        date = current_file[-12:-4]
        product_name = current_file[0:6]
        a = os.path.join(folder_path, current_file)
        a = a.replace('\\', '\\\\')                                      #路径拼接


        df = pd.read_excel(a, header=None)                                           #sheet1 stock_tracking
        df.reset_index(drop=True, inplace=True)
        new_columns = df.iloc[3]
        df = df.iloc[4:]
        df.columns = new_columns
        df.set_index('科目代码', inplace=True)
        df1 = df.dropna(subset=['停牌信息'])
        df1 = df1[df1['科目名称'].str.len() <= 5]                                      #df抓取所有股票，转债和其他五字符以内产品
        df1 = df1.drop(df1[df1['科目名称'].str.contains('T2')].index)                  #df删除以T2开头的国债
        df2 = df1[df1['科目名称'].str.contains('转')]                                   #df抓取所有转债
        stock_df = df1[~df1.index.isin(df2.index)]
        stock_index = stock_df.index.tolist()
        stock_code_list = []
        for i in stock_index:
            stock_code = i[-6:]
            stock_code_list.append(stock_code)
        day = pd.Series([date] * len(stock_df), index=stock_df.index)
        stock_name = stock_df['科目名称'].tolist()
        stock_amount = stock_df['数量'].tolist()
        stock_close = stock_df['市价'].tolist()
        stock_value = stock_df['市值'].tolist()
        result = pd.DataFrame()
        try:
            result.insert(0, '日期', day)
            result.insert(1, '产品名称', product_name)
            result.insert(2, '股票代码', stock_code_list)
            result.insert(3, '股票名称', stock_name)
            result.insert(4, '数量', stock_amount)
            result.insert(5, '市价', stock_close)
            result.insert(6, '市值', stock_value)
        except:
            print(day+product_name+"产品sheet1出现变动,请更正")

        df = pd.read_excel(a, header=None)                                      # sheet2 future_tracking
        df.reset_index(drop=True, inplace=True)
        new_columns = df.iloc[3]
        df = df.iloc[4:]
        df.columns = new_columns
        df1 = df.dropna(subset=['停牌信息'])
        df1 = df1[df1['科目名称'].str.len() >= 6]
        df1 = df1.drop(df1[df1['科目名称'].str.contains('国债')].index)
        df2_1 = df1[df1['科目名称'].str.contains('期货')]
        df2_2 = df1[df1['科目名称'].str.contains('I')]
        future_df = pd.concat([df2_1, df2_2], axis=0)
        code = future_df['科目代码'].tolist()
        future_name = future_df['科目代码'].tolist()
        for i in range(len(future_name)):
            current_element = future_name[i]
            scode = current_element[-6:]
            future_name[i] = scode
        future_amount = future_df['数量'].tolist()
        unit_cost = future_df['单位成本'].tolist()
        cost = future_df['成本'].tolist()
        pure_relative_cost = future_df['成本占净值%'].tolist()
        market_price = future_df['市价'].tolist()
        future_value = future_df['市值'].tolist()
        pure_relative_future_value = future_df['市值占净值%'].tolist()
        valuation_appreciation = future_df['估值增值'].tolist()
        trade_info = future_df['停牌信息'].tolist()
        type1 = []
        for i in code:
            if '3102' in i:
                s = '衍生工具'
                type1.append(s)
        type2 = []
        for j in code:
            if '310201' in j:
                d = '中金所_投机_买方_股指'
                type2.append(d)
            if '310203' in j:
                d = '中金所_多头'
                type2.append(d)
            if '310204' in j:
                d = '中金所_空头'
                type2.append(d)
        direction = []
        for k in future_value:
            if k > 0:
                b = 'long'
                direction.append(b)
            if k < 0:
                b = 'short'
                direction.append(b)
        day = pd.Series([date] * len(future_df), index=future_df.index)
        result1 = pd.DataFrame()
        try:
            result1.insert(0, '日期', day)
            result1.insert(1, '产品名称', product_name)
            result1.insert(2, '种类', type1)
            result1.insert(3, '种类名称', type2)
            result1.insert(4, '代码', code)
            result1.insert(5, '方向', direction)
            result1.insert(6, '科目名称', future_name)
            result1.insert(7, '数量', future_amount)
            result1.insert(8, '单位成本', unit_cost)
            result1.insert(9, '成本', cost)
            result1.insert(10, '成本占净值%', pure_relative_cost)
            result1.insert(11, '市价', market_price)
            result1.insert(12, '市值', future_value)
            result1.insert(13, '市值占净值%', pure_relative_future_value)
            result1.insert(14, '估值增值', valuation_appreciation)
            result1.insert(15, '停牌信息', trade_info)
        except:
            print(day + product_name + "产品sheet2出现变动,请更正")


        df = pd.read_excel(a, header=None)                                      #sheet3 c_bond_tracking
        df.reset_index(drop=True, inplace=True)
        new_columns = df.iloc[3]
        df = df.iloc[4:]
        df.columns = new_columns
        df.set_index('科目代码', inplace=True)

        df1 = df.dropna(subset=['停牌信息'])
        df1 = df1[df1['科目名称'].str.len() <= 5]
        df2 = df1[df1['科目名称'].str.contains('转')]
        cbond_df = df2
        cbond_index = cbond_df.index.tolist()
        cbond_code_list = []
        for i in cbond_index:
            cbond_code = i[-6:]
            cbond_code_list.append(cbond_code)
        day = pd.Series([date] * len(cbond_df), index=cbond_df.index)
        cbond_name = cbond_df['科目名称'].tolist()
        cbond_amount = cbond_df['数量'].tolist()
        cbond_close = cbond_df['市价'].tolist()
        cbond_value = cbond_df['市值'].tolist()
        result2 = pd.DataFrame()
        try:
            result2.insert(0, '日期', day)
            result2.insert(1, '产品名称', product_name)
            result2.insert(2, '债券代码', cbond_code_list)
            result2.insert(3, '债券名称', cbond_name)
            result2.insert(4, '数量', cbond_amount)
            result2.insert(5, '市价', cbond_close)
            result2.insert(6, '市值', cbond_value)
        except:
            print(day+product_name+"产品sheet3出现变动,请更正")

        df = pd.read_excel(a, header=None)
        df.reset_index(drop=True, inplace=True)
        new_columns = df.iloc[3]
        df = df.iloc[4:]
        df.columns = new_columns
        df1 = df.dropna(subset=['停牌信息'])                                     #sheet4 option_tracking
        df1 = df1[df1['科目名称'].str.len() >= 6]
        df3 = df1[df1['科目名称'].str.contains('-')]
        df3.loc[:, '科目名称'] = df3.loc[:, '科目名称'].apply(option_name_transfer)
        code = df3['科目代码'].tolist()
        type1 = []
        for i in code:
            if '3102' in i:
                s = '衍生工具'
                type1.append(s)
        option_value = df3['市值'].tolist()
        direction = []
        for j in option_value:
            if j > 0:
                z = 'long'
                direction.append(z)
            if j < 0:
                z = 'short'
                direction.append(z)
        try:
            df3['日期'] = date
            df3['产品名称'] = product_name
            df3['种类'] = type1
            df3['代码'] = code
            df3['方向'] = direction
            df3 = df3[
                ['日期', '产品名称', '种类', '代码', '科目名称','方向', '数量', '单位成本', '成本', '成本占净值%', '市价', '市值',
                 '市值占净值%',
                 '估值增值', '停牌信息']]
        except:
            print(day+product_name+"产品sheet4出现变动,请更正")

        df = pd.read_excel(a, header=None)                                                  # sheet5 bond_tracking
        df.reset_index(drop=True, inplace=True)
        new_columns = df.iloc[3]
        df = df.iloc[4:]
        df.columns = new_columns
        df.set_index('科目代码', inplace=True)
        df1 = df.dropna(subset=['停牌信息'])
        df1 = df1[df1['科目名称'].str.len() >= 6]
        df2 = df1[df1['科目名称'].str.contains('国债')]
        bond_df = df2
        bond_index = bond_df.index.tolist()
        bond_code_list = []
        for i in bond_index:
            bond_code = i[-6:]
            bond_code_list.append(bond_code)
        day = pd.Series([date] * len(bond_df), index=bond_df.index)
        bond_name = bond_df['科目名称'].tolist()
        bond_amount = bond_df['数量'].tolist()
        bond_close = bond_df['市价'].tolist()
        bond_value = bond_df['市值'].tolist()
        result3 = pd.DataFrame()
        try:
            result3.insert(0, '日期', day)
            result3.insert(1, '产品名称', product_name)
            result3.insert(2, '债券代码', bond_code_list)
            result3.insert(3, '债券名称', bond_name)
            result3.insert(4, '数量', bond_amount)
            result3.insert(5, '市价', bond_close)
            result3.insert(6, '市值', bond_value)
        except:
            print(day + product_name + "产品sheet5出现变动,请更正")

        result_path = glv.get('outputpath_RR500')
        result_path_final = os.path.join(result_path, date + '_瑞锐500指增产品跟踪.xlsx')
        gt.create_file_directory(result_path_final)
        with pd.ExcelWriter(result_path_final, engine='openpyxl') as writer:
            result.to_excel(writer, sheet_name='stock_tracking',index=False)
            result2.to_excel(writer, sheet_name='c_bond_tracking', index=False)
            result1.to_excel(writer, sheet_name='future_tracking',index=False)
            df3.to_excel(writer, sheet_name='option_tracking',index=False)
            result3.to_excel(writer, sheet_name='bond_tracking', index=False)


def RRJX(file_name):
    # 输入为 瑞锐精选产品文件名
    # 输出为 解析后的四级估值表 分为股票 可转债 期货 期权 国债五个模块
    folder_path = glv.get('folder_path')
    for i in file_name:
        current_file = i
        date = current_file[-12:-4]
        product_name = current_file[0:6]
        a = os.path.join(folder_path, current_file)
        a = a.replace('\\', '\\\\')
        df = pd.read_excel(a, header=None)                                             #sheet1 stock_tracking
        df.reset_index(drop=True, inplace=True)
        new_columns = df.iloc[3]
        df = df.iloc[4:]
        df.columns = new_columns
        df.set_index('科目代码', inplace=True)
        df1 = df.dropna(subset=['停牌信息'])
        df1 = df1[df1['科目名称'].str.len() <= 5]
        df1 = df1.drop(df1[df1['科目名称'].str.contains('T2')].index)
        df2 = df1[df1['科目名称'].str.contains('转')]
        stock_df = df1[~df1.index.isin(df2.index)]
        stock_index = stock_df.index.tolist()
        stock_code_list = []
        for i in stock_index:
            stock_code = i[-6:]
            stock_code_list.append(stock_code)
        day = pd.Series([date] * len(stock_df), index=stock_df.index)
        stock_name = stock_df['科目名称'].tolist()
        stock_amount = stock_df['数量'].tolist()
        stock_close = stock_df['市价'].tolist()
        stock_value = stock_df['市值'].tolist()
        result = pd.DataFrame()
        try:
            result.insert(0, '日期', day)
            result.insert(1, '产品名称', product_name)
            result.insert(2, '股票代码', stock_code_list)
            result.insert(3, '股票名称', stock_name)
            result.insert(4, '数量', stock_amount)
            result.insert(5, '市价', stock_close)
            result.insert(6, '市值', stock_value)
        except:
            print(day+product_name+"产品sheet1出现变动,请更正")

        df = pd.read_excel(a, header=None)                                               #sheet2 future_tracking
        df.reset_index(drop=True, inplace=True)
        new_columns = df.iloc[3]
        df = df.iloc[4:]
        df.columns = new_columns
        df1 = df.dropna(subset=['停牌信息'])
        df1 = df1[df1['科目名称'].str.len() >= 6]
        df1 = df1.drop(df1[df1['科目名称'].str.contains('国债')].index)
        df2_1 = df1[df1['科目名称'].str.contains('期货')]
        df2_2 = df1[df1['科目名称'].str.contains('I')]
        future_df = pd.concat([df2_1, df2_2], axis=0)
        code = future_df['科目代码'].tolist()
        future_name = future_df['科目代码'].tolist()
        for i in range(len(future_name)):
            current_element = future_name[i]
            scode = current_element[-6:]
            future_name[i] = scode
        future_amount = future_df['数量'].tolist()
        unit_cost = future_df['单位成本'].tolist()
        cost = future_df['成本'].tolist()
        pure_relative_cost = future_df['成本占净值%'].tolist()
        market_price = future_df['市价'].tolist()
        future_value = future_df['市值'].tolist()
        pure_relative_future_value = future_df['市值占净值%'].tolist()
        valuation_appreciation = future_df['估值增值'].tolist()
        trade_info = future_df['停牌信息'].tolist()
        type1 = []
        for i in code:
            if '3102' in i:
                s = '衍生工具'
                type1.append(s)
        type2 = []
        for j in code:
            if '310201' in j:
                d = '中金所_投机_买方_股指'
                type2.append(d)
            if '310203' in j:
                d = '中金所_多头'
                type2.append(d)
            if '310204' in j:
                d = '中金所_空头'
                type2.append(d)
        direction = []
        for k in future_value:
            if k > 0:
                b = 'long'
                direction.append(b)
            if k < 0:
                b = 'short'
                direction.append(b)

        product_name = current_file[:6]
        day = pd.Series([date] * len(future_df), index=future_df.index)
        result1 = pd.DataFrame()
        try:
            result1.insert(0, '日期', day)
            result1.insert(1, '产品名称', product_name)
            result1.insert(2, '种类', type1)
            result1.insert(3, '种类名称', type2)
            result1.insert(4, '代码', code)
            result1.insert(5, '方向', direction)
            result1.insert(6, '科目名称', future_name)
            result1.insert(7, '数量', future_amount)
            result1.insert(8, '单位成本', unit_cost)
            result1.insert(9, '成本', cost)
            result1.insert(10, '成本占净值%', pure_relative_cost)
            result1.insert(11, '市价', market_price)
            result1.insert(12, '市值', future_value)
            result1.insert(13, '市值占净值%', pure_relative_future_value)
            result1.insert(14, '估值增值', valuation_appreciation)
            result1.insert(15, '停牌信息', trade_info)
        except:
            print(day+product_name+"产品sheet2出现变动,请更正")

        df = pd.read_excel(a, header=None)                                                  # sheet3 c_bond_tracking
        df.reset_index(drop=True, inplace=True)
        new_columns = df.iloc[3]
        df = df.iloc[4:]
        df.columns = new_columns
        df.set_index('科目代码', inplace=True)

        df1 = df.dropna(subset=['停牌信息'])
        df1 = df1[df1['科目名称'].str.len() <= 5]
        df2 = df1[df1['科目名称'].str.contains('转')]
        cbond_df = df2
        cbond_index = cbond_df.index.tolist()
        cbond_code_list = []
        for i in cbond_index:
            cbond_code = i[-6:]
            cbond_code_list.append(cbond_code)
        day = pd.Series([date] * len(cbond_df), index=cbond_df.index)
        cbond_name = cbond_df['科目名称'].tolist()
        cbond_amount = cbond_df['数量'].tolist()
        cbond_close = cbond_df['市价'].tolist()
        cbond_value = cbond_df['市值'].tolist()
        result2 = pd.DataFrame()
        try:
            result2.insert(0, '日期', day)
            result2.insert(1, '产品名称', product_name)
            result2.insert(2, '债券代码', cbond_code_list)
            result2.insert(3, '债券名称', cbond_name)
            result2.insert(4, '数量', cbond_amount)
            result2.insert(5, '市价', cbond_close)
            result2.insert(6, '市值', cbond_value)
        except:
            print(day + product_name + "产品sheet3出现变动,请更正")

        df = pd.read_excel(a, header=None)
        df.reset_index(drop=True, inplace=True)
        new_columns = df.iloc[3]
        df = df.iloc[4:]
        df.columns = new_columns
        df1 = df.dropna(subset=['停牌信息'])                                             #sheet4 option_tracking
        df1 = df1[df1['科目名称'].str.len() >= 6]
        df3 = df1[df1['科目名称'].str.contains('-')]
        df3.loc[:, '科目名称'] = df3.loc[:, '科目名称'].apply(option_name_transfer)
        day = pd.Series([date] * len(df3), index=df3.index)
        product_name = current_file[:6]
        code = df3['科目代码'].tolist()
        type1 = []
        for i in code:
            if '3102' in i:
                s = '衍生工具'
                type1.append(s)
        direction = []
        option_value = df3['市值'].tolist()
        for j in option_value:
            if j > 0:
                z = 'long'
                direction.append(z)
            if j < 0:
                z = 'short'
                direction.append(z)
        try:
            df3['方向'] = direction
            df3['日期'] = day
            df3['产品名称'] = product_name
            df3['种类'] = type1
            df3['代码'] = code
            df3 = df3[
                ['日期', '产品名称', '种类', '代码', '科目名称','方向', '数量', '单位成本', '成本', '成本占净值%', '市价', '市值',
                 '市值占净值%',
                 '估值增值', '停牌信息']]
        except:
            print(day+product_name+"产品sheet4出现变动,请更正")

        df = pd.read_excel(a, header=None)                                                  # sheet5 bond_tracking
        df.reset_index(drop=True, inplace=True)
        new_columns = df.iloc[3]
        df = df.iloc[4:]
        df.columns = new_columns
        df.set_index('科目代码', inplace=True)
        df1 = df.dropna(subset=['停牌信息'])
        df1 = df1[df1['科目名称'].str.len() >= 6]
        df2 = df1[df1['科目名称'].str.contains('国债')]
        bond_df = df2
        bond_index = bond_df.index.tolist()
        bond_code_list = []
        for i in bond_index:
            bond_code = i[-6:]
            bond_code_list.append(bond_code)
        day = pd.Series([date] * len(bond_df), index=bond_df.index)
        bond_name = bond_df['科目名称'].tolist()
        bond_amount = bond_df['数量'].tolist()
        bond_close = bond_df['成本'].tolist()
        bond_value = bond_df['市值'].tolist()
        result3 = pd.DataFrame()
        try:
            result3.insert(0, '日期', day)
            result3.insert(1, '产品名称', product_name)
            result3.insert(2, '债券代码', bond_code_list)
            result3.insert(3, '债券名称', bond_name)
            result3.insert(4, '数量', bond_amount)
            result3.insert(5, '市价', bond_close)
            result3.insert(6, '市值', bond_value)
        except:
            print(day + product_name + "产品sheet5出现变动,请更正")
        result_path = glv.get('outputpath_RRJX')
        result_path_final = os.path.join(result_path, date + '_瑞锐精选产品跟踪.xlsx')
        gt.create_file_directory(result_path_final)
        with pd.ExcelWriter(result_path_final, engine='openpyxl') as writer:
            result.to_excel(writer, sheet_name='stock_tracking', index=False)
            result2.to_excel(writer, sheet_name='c_bond_tracking', index=False)
            result1.to_excel(writer, sheet_name='future_tracking', index=False)
            df3.to_excel(writer, sheet_name='option_tracking', index=False)
            result3.to_excel(writer, sheet_name='bond_tracking', index=False)


def SF500_N08(file_name):
    # 输入为 盛丰盛元中证500指数增强8号文件名
    # 输出为 解析后的四级估值表 分为股票 可转债 期货 期权 国债五个模块
    folder_path = glv.get('folder_path')
    for i in file_name:
        current_file = i
        date = current_file[-12:-4]
        product_name = current_file[0:6]
        a = os.path.join(folder_path, current_file)
        a = a.replace('\\', '\\\\')
        print(a)
        df = pd.read_excel(a, header=None)                     #sheet1 stock_tracking

        df.reset_index(drop=True, inplace=True)
        new_columns = df.iloc[3]
        df = df.iloc[4:]
        df.columns = new_columns
        df.set_index('科目代码', inplace=True)
        df1 = df.dropna(subset=['停牌信息'])
        df1 = df1[df1['科目名称'].str.len() <= 5]
        df1 = df1.drop(df1[df1['科目名称'].str.contains('T2')].index)
        df2 = df1[df1['科目名称'].str.contains('转')]
        stock_df = df1[~df1.index.isin(df2.index)]
        stock_index = stock_df.index.tolist()
        stock_code_list = []
        for i in stock_index:
            stock_code = i[-6:]
            stock_code_list.append(stock_code)
        day = pd.Series([date] * len(stock_df), index=stock_df.index)
        stock_name = stock_df['科目名称'].tolist()
        stock_amount = stock_df['数量'].tolist()
        stock_close = stock_df['市价'].tolist()
        stock_value = stock_df['市值'].tolist()
        result = pd.DataFrame()
        try:
            result.insert(0, '日期', day)
            result.insert(1, '产品名称', product_name)
            result.insert(2, '股票代码', stock_code_list)
            result.insert(3, '股票名称', stock_name)
            result.insert(4, '数量', stock_amount)
            result.insert(5, '市价', stock_close)
            result.insert(6, '市值', stock_value)
        except:
            print(day+product_name+"产品sheet1出现变动,请更正")

        df = pd.read_excel(a, header=None)                                            #sheet2 future_tracking
        df.reset_index(drop=True, inplace=True)
        new_columns = df.iloc[3]
        df = df.iloc[4:]
        df.columns = new_columns
        df1 = df.dropna(subset=['停牌信息'])
        df1 = df1[df1['科目名称'].str.len() >= 6]
        df1 = df1.drop(df1[df1['科目名称'].str.contains('国债')].index)
        df2_1 = df1[df1['科目名称'].str.contains('期货')]
        df2_2 = df1[df1['科目名称'].str.contains('I')]
        future_df = pd.concat([df2_1, df2_2], axis=0)
        code = future_df['科目代码'].tolist()
        future_name = future_df['科目代码'].tolist()
        for i in range(len(future_name)):
            current_element = future_name[i]
            scode = current_element[8:]
            future_name[i] = scode
        future_amount = future_df['数量'].tolist()
        unit_cost = future_df['单位成本'].tolist()
        cost = future_df['成本'].tolist()
        pure_relative_cost = future_df['成本占净值%'].tolist()
        market_price = future_df['市价'].tolist()
        future_value = future_df['市值'].tolist()
        pure_relative_future_value = future_df['市值占净值%'].tolist()
        valuation_appreciation = future_df['估值增值'].tolist()
        trade_info = future_df['停牌信息'].tolist()
        type1 = []
        for i in code:
            if '3102' in i:
                s = '衍生工具'
                type1.append(s)
        type2 = []
        for j in code:
            if '310201' in j:
                d = '中金所_投机_买方_股指'
                type2.append(d)
            if '310203' in j:
                d = '中金所_多头'
                type2.append(d)
            if '310204' in j:
                d = '中金所_空头'
                type2.append(d)
        direction = []
        for k in future_value:
            if k > 0:
                b = 'long'
                direction.append(b)
            if k < 0:
                b = 'short'
                direction.append(b)
        product_name = current_file[:6]
        day = pd.Series([date] * len(future_df), index=future_df.index)
        result1 = pd.DataFrame()
        try:
            result1.insert(0, '日期', day)
            result1.insert(1, '产品名称', product_name)
            result1.insert(2, '种类', type1)
            result1.insert(3, '种类名称', type2)
            result1.insert(4, '代码', code)
            result1.insert(5, '方向', direction)
            result1.insert(6, '科目名称', future_name)
            result1.insert(7, '数量', future_amount)
            result1.insert(8, '单位成本', unit_cost)
            result1.insert(9, '成本', cost)
            result1.insert(10, '成本占净值%', pure_relative_cost)
            result1.insert(11, '市价', market_price)
            result1.insert(12, '市值', future_value)
            result1.insert(13, '市值占净值%', pure_relative_future_value)
            result1.insert(14, '估值增值', valuation_appreciation)
            result1.insert(15, '停牌信息', trade_info)
        except:
            print(day+product_name+"产品sheet2出现变动,请更正")

        df = pd.read_excel(a, header=None)                                                  # sheet3 c_bond_tracking
        df.reset_index(drop=True, inplace=True)
        new_columns = df.iloc[3]
        df = df.iloc[4:]
        df.columns = new_columns
        df.set_index('科目代码', inplace=True)
        df1 = df.dropna(subset=['停牌信息'])
        df1 = df1[df1['科目名称'].str.len() <= 5]
        df2 = df1[df1['科目名称'].str.contains('转')]
        cbond_df = df2
        cbond_index = cbond_df.index.tolist()
        cbond_code_list = []
        for i in cbond_index:
            cbond_code = i[-6:]
            cbond_code_list.append(cbond_code)
        day = pd.Series([date] * len(cbond_df), index=cbond_df.index)
        cbond_name = cbond_df['科目名称'].tolist()
        cbond_amount = cbond_df['数量'].tolist()
        cbond_close = cbond_df['市价'].tolist()
        cbond_value = cbond_df['市值'].tolist()
        result2 = pd.DataFrame()
        try:
            result2.insert(0, '日期', day)
            result2.insert(1, '产品名称', product_name)
            result2.insert(2, '债券代码', cbond_code_list)
            result2.insert(3, '债券名称', cbond_name)
            result2.insert(4, '数量', cbond_amount)
            result2.insert(5, '市价', cbond_close)
            result2.insert(6, '市值', cbond_value)
        except:
            print(day + product_name + "产品sheet3出现变动,请更正")

        df = pd.read_excel(a, header=None)
        df.reset_index(drop=True, inplace=True)
        new_columns = df.iloc[3]
        df = df.iloc[4:]
        df.columns = new_columns
        df1 = df.dropna(subset=['停牌信息'])                                           #sheet4 option_tracking
        df1 = df1[df1['科目名称'].str.len() >= 6]
        df3 = df1[df1['科目名称'].str.contains('-')]
        df3.loc[:, '科目名称'] = df3.loc[:, '科目名称'].apply(option_name_transfer)
        day = pd.Series([date] * len(df3), index=df3.index)
        product_name = current_file[:6]
        code = df3['科目代码'].tolist()
        type1 = []
        for i in code:
            if '3102' in i:
                s = '衍生工具'
                type1.append(s)
        option_value = df3['市值'].tolist()
        direction = []
        for j in option_value:
            if j > 0:
                z = 'long'
                direction.append(z)
            if j < 0:
                z = 'short'
                direction.append(z)
        try:
            df3['方向'] = direction
            df3['日期'] = day
            df3['产品名称'] = product_name
            df3['种类'] = type1
            df3['代码'] = code
            df3 = df3[
                ['日期', '产品名称', '种类', '代码', '科目名称','方向','数量', '单位成本', '成本', '成本占净值%', '市价', '市值',
                 '市值占净值%',
                 '估值增值', '停牌信息']]
        except:
            print(day+product_name+"产品sheet4出现变动,请更正")

        result3 = pd.DataFrame(data=None, columns=['日期', '产品名称', '债券代码', '债券名称', '数量', '市价', '市值'])


        result_path = glv.get('outputpath_SF500_N08')
        result_path_final = os.path.join(result_path, date + '_盛元8号指增产品跟踪.xlsx')
        gt.create_file_directory(result_path_final)
        with pd.ExcelWriter(result_path_final, engine='openpyxl') as writer:
            result.to_excel(writer, sheet_name='stock_tracking', index=False)
            result2.to_excel(writer, sheet_name='c_bond_tracking', index=False)
            result1.to_excel(writer, sheet_name='future_tracking', index=False)
            df3.to_excel(writer, sheet_name='option_tracking', index=False)
            result3.to_excel(writer, sheet_name='bond_tracking', index=False)


def XYHY_N01(file_name):
    # 输入为 宣夜惠盈一号文件名
    # 输出为 解析后的四级估值表 分为股票 可转债 期货 期权 国债五个模块
    folder_path = glv.get('folder_path')
    for i in file_name:
        current_file = i
        date = current_file[-12:-4]
        product_name = current_file[0:6]
        a = os.path.join(folder_path, current_file)
        a = a.replace('\\', '\\\\')
        print(date)
        df = pd.read_excel(a, header=None)                                           #sheet1 stock_tracking
        df.reset_index(drop=True, inplace=True)
        new_columns = df.iloc[4]
        df = df.iloc[8:]
        df.columns = new_columns
        df.set_index('科目代码', inplace=True)
        df1 = df.dropna(subset=['停牌信息'])
        df1 = df1[df1['科目名称'].str.len() <= 5]
        df1 = df1.drop(df1[df1['科目名称'].str.contains('T2')].index)
        df2 = df1[df1['科目名称'].str.contains('转')]
        stock_df = df1[~df1.index.isin(df2.index)]
        stock_index = stock_df.index.tolist()
        stock_code_list = []
        for i in stock_index:
            stock_code = i[-9:-3]
            stock_code_list.append(stock_code)
        day = pd.Series([date] * len(stock_df), index=stock_df.index)
        stock_name = stock_df['科目名称'].tolist()
        stock_amount = stock_df['数量'].tolist()
        stock_close = stock_df['行情'].tolist()
        stock_value = stock_df['市值-本币'].tolist()
        result = pd.DataFrame()
        try:
            result.insert(0, '日期', day)
            result.insert(1, '产品名称', product_name)
            result.insert(2, '股票代码', stock_code_list)
            result.insert(3, '股票名称', stock_name)
            result.insert(4, '数量', stock_amount)
            result.insert(5, '市价', stock_close)
            result.insert(6, '市值', stock_value)
        except:
            print(day+product_name+"产品sheet1出现变动,请更正")

        df = pd.read_excel(a, header=None)                                            #sheet2 future_tracking
        df.reset_index(drop=True, inplace=True)
        new_columns = df.iloc[4]
        df = df.iloc[8:]
        df.columns = new_columns
        df1 = df.dropna(subset=['停牌信息'])
        df1 = df1[df1['科目名称'].str.len() >= 6]
        df1 = df1.drop(df1[df1['科目名称'].str.contains('国债')].index)
        df2_1 = df1[df1['科目名称'].str.contains('期货')]
        df2_2 = df1[df1['科目名称'].str.contains('I')]
        future_df = pd.concat([df2_1, df2_2], axis=0)
        code = future_df['科目代码'].tolist()
        future_name = future_df['科目名称'].tolist()
        future_amount = future_df['数量'].tolist()
        unit_cost = future_df['单位成本'].tolist()
        cost = future_df['成本-本币'].tolist()
        pure_relative_cost = future_df['成本占比'].tolist()
        market_price = future_df['行情'].tolist()
        future_value = future_df['市值-本币'].tolist()
        pure_relative_future_value = future_df['市值占比'].tolist()
        valuation_appreciation = [cost[m]-future_value[m] for m in range(0,len(cost))]
        trade_info = future_df['停牌信息'].tolist()
        type1 = []
        for i in code:
            if '3102' in i:
                s = '衍生工具'
                type1.append(s)
        type2 = []
        for j in code:
            if '3102.01' in j:
                d = '中金所_投机_买方_股指'
                type2.append(d)
            if '3102.03' in j:
                d = '中金所_投机_卖方_股指'
                type2.append(d)
        direction = []
        for k in future_value:
            if k > 0:
                b = 'long'
                direction.append(b)
            if k < 0:
                b = 'short'
                direction.append(b)
        day = pd.Series([date] * len(future_df), index=future_df.index)
        product_name = current_file[:6]
        result1 = pd.DataFrame()
        try:
            result1.insert(0, '日期', day)
            result1.insert(1, '产品名称', product_name)
            result1.insert(2, '种类', type1)
            result1.insert(3, '种类名称', type2)
            result1.insert(4, '代码', code)
            result1.insert(5, '方向', direction)
            result1.insert(6, '科目名称', future_name)
            result1.insert(7, '数量', future_amount)
            result1.insert(8, '单位成本', unit_cost)
            result1.insert(9, '成本', cost)
            result1.insert(10, '成本占净值%', pure_relative_cost)
            result1.insert(11, '市价', market_price)
            result1.insert(12, '市值', future_value)
            result1.insert(13, '市值占净值%', pure_relative_future_value)
            result1.insert(14, '估值增值', valuation_appreciation)
            result1.insert(15, '停牌信息', trade_info)
        except:
            print(day+product_name+"产品sheet2出现变动,请更正")

        df = pd.read_excel(a, header=None)                                                  # sheet3 c_bond_tracking
        df.reset_index(drop=True, inplace=True)
        new_columns = df.iloc[4]
        df = df.iloc[8:]
        df.columns = new_columns
        df.set_index('科目代码', inplace=True)
        df1 = df.dropna(subset=['停牌信息'])
        df1 = df1[df1['科目名称'].str.len() <= 5]
        df2 = df1[df1['科目名称'].str.contains('转')]
        cbond_df = df2
        cbond_index = cbond_df.index.tolist()
        cbond_code_list = []
        for i in cbond_index:
            cbond_code = i[-9:-3]
            cbond_code_list.append(cbond_code)
        day = pd.Series([date] * len(cbond_df), index=cbond_df.index)
        cbond_name = cbond_df['科目名称'].tolist()
        cbond_amount = cbond_df['数量'].tolist()
        cbond_close = cbond_df['成本-本币'].tolist()
        cbond_value = cbond_df['市值-本币'].tolist()
        result2 = pd.DataFrame()
        try:
            result2.insert(0, '日期', day)
            result2.insert(1, '产品名称', product_name)
            result2.insert(2, '债券代码', cbond_code_list)
            result2.insert(3, '债券名称', cbond_name)
            result2.insert(4, '数量', cbond_amount)
            result2.insert(5, '市价', cbond_close)
            result2.insert(6, '市值', cbond_value)
        except:
            print(day + product_name + "产品sheet3出现变动,请更正")

        df = pd.read_excel(a, header=None)
        df.reset_index(drop=True, inplace=True)
        new_columns = df.iloc[4]
        df = df.iloc[8:]
        df.columns = new_columns
        df1 = df.dropna(subset=['停牌信息'])                                             #sheet4 option_tracking
        df1 = df1[df1['科目名称'].str.len() >= 6]
        df3 = df1[df1['科目名称'].str.contains('-')]
        df3.loc[:, '科目名称'] = df3.loc[:, '科目名称'].apply(option_name_transfer)
        df.set_index('科目代码', inplace=True)
        day = pd.Series([date] * len(df3), index=df3.index)
        product_name = current_file[:6]
        code = df3['科目代码'].tolist()
        type1 = []
        for i in code:
            if '3102' in i:
                s = '衍生工具'
                type1.append(s)
            else:
                s = ''
                type1.append(s)
        option_value = df3['市值-本币'].tolist()
        direction = []
        for j in option_value:
            if j > 0:
                z = 'long'
                direction.append(z)
            if j < 0:
                z = 'short'
                direction.append(z)
        try:
            df3['方向'] = direction
            df3['日期'] = day
            df3['产品名称'] = product_name
            df3['种类'] = type1
            df3['代码'] = code
            df3.rename(columns={'成本-本币' : '成本'},inplace=True)
            df3.rename(columns={'市值-本币': '市值'}, inplace=True)
            df3.rename(columns={'估值增值-本币': '估值增值'}, inplace=True)
            df3.rename(columns={'成本占比': '成本占净值%'}, inplace=True)
            df3.rename(columns={'市值占比': '市值占净值%'}, inplace=True)
            df3.rename(columns={'行情': '市价'}, inplace=True)
            df3 = df3[
                ['日期', '产品名称', '种类', '代码', '科目名称','方向', '数量', '单位成本', '成本', '成本占净值%', '市价',
                 '市值',
                 '市值占净值%',
                 '估值增值', '停牌信息']]
        except:
            print(day+product_name+"产品sheet4出现变动,请更正")

        df = pd.read_excel(a, header=None)                                                  # sheet5 bond_tracking
        df.reset_index(drop=True, inplace=True)
        new_columns = df.iloc[4]
        df = df.iloc[8:]
        df.columns = new_columns
        df.set_index('科目代码', inplace=True)
        df1 = df.dropna(subset=['停牌信息'])
        df1 = df1[df1['科目名称'].str.len() <= 5]
        df2 = df1[df1['科目名称'].str.contains('T2')]
        bond_df = df2
        bond_index = bond_df.index.tolist()
        bond_code_list = []
        for i in bond_index:
            bond_code = i[-9:-3]
            bond_code_list.append(bond_code)
        day = pd.Series([date] * len(bond_df), index=bond_df.index)
        bond_name = bond_df['科目名称'].tolist()
        bond_amount = bond_df['数量'].tolist()
        bond_close = bond_df['行情'].tolist()
        bond_value = bond_df['市值-本币'].tolist()
        result3 = pd.DataFrame()
        try:
            result3.insert(0, '日期', day)
            result3.insert(1, '产品名称', product_name)
            result3.insert(2, '债券代码', bond_code_list)
            result3.insert(3, '债券名称', bond_name)
            result3.insert(4, '数量', bond_amount)
            result3.insert(5, '市价', bond_close)
            result3.insert(6, '市值', bond_value)
        except:
            print(day + product_name + "产品sheet5出现变动,请更正")

        result_path = glv.get('outputpath_XYHY_N01')
        result_path_final = os.path.join(result_path, date + '_惠盈一号指增产品跟踪.xlsx')
        gt.create_file_directory(result_path_final)
        with pd.ExcelWriter(result_path_final, engine='openpyxl') as writer:
            result.to_excel(writer, sheet_name='stock_tracking', index=False)
            result2.to_excel(writer, sheet_name='c_bond_tracking', index=False)
            result1.to_excel(writer, sheet_name='future_tracking', index=False)
            df3.to_excel(writer, sheet_name='option_tracking', index=False)
            result3.to_excel(writer, sheet_name='bond_tracking', index=False)


def SF1000_N01(file_name):
    # 输入为 盛丰中证1000指数增强1号文件名
    # 输出为 解析后的四级估值表 分为股票 可转债 期货 期权 国债五个模块
    folder_path = glv.get('folder_path')
    for i in file_name:
        current_file = i
        date = current_file[0:10]
        date = date[0:4] + date[5:7] + date[8:]
        product_name = current_file[11:17]
        a = os.path.join(folder_path, current_file)
        a = a.replace('\\', '\\\\')
        df = pd.read_excel(a, header=None)                                            #sheet1 stock_tracking
        df.reset_index(drop=True, inplace=True)
        new_columns = df.iloc[2]
        df = df.iloc[3:]
        df.columns = new_columns
        df.set_index('科目代码', inplace=True)
        df1 = df.dropna(subset=['停牌信息'])
        df1 = df1[df1['科目名称'].str.len() <= 5]
        df1 = df1.drop(df1[df1['科目名称'].str.contains('T2')].index)
        df2 = df1[df1['科目名称'].str.contains('转')]
        stock_df = df1[~df1.index.isin(df2.index)]
        stock_index = stock_df.index.tolist()
        stock_code_list = []
        for i in stock_index:
            stock_code = i[-6:]
            stock_code_list.append(stock_code)
        day = pd.Series([date] * len(stock_df), index=stock_df.index)
        stock_name = stock_df['科目名称'].tolist()
        stock_amount = stock_df['数量'].tolist()
        stock_close = stock_df['市价'].tolist()
        stock_value = stock_df['市值'].tolist()
        result = pd.DataFrame()
        try:
            result.insert(0, '日期', day)
            result.insert(1, '产品名称', product_name)
            result.insert(2, '股票代码', stock_code_list)
            result.insert(3, '股票名称', stock_name)
            result.insert(4, '数量', stock_amount)
            result.insert(5, '市价', stock_close)
            result.insert(6, '市值', stock_value)
        except:
            print(day+product_name+"产品sheet1出现变动,请更正")

        df = pd.read_excel(a, header=None)                                             #sheet2 future_tracking
        df.reset_index(drop=True, inplace=True)
        new_columns = df.iloc[2]
        df = df.iloc[3:]
        df.columns = new_columns
        df1 = df.dropna(subset=['停牌信息'])
        df1 = df1[df1['科目名称'].str.len() >= 6]
        df1 = df1.drop(df1[df1['科目名称'].str.contains('国债')].index)
        df2_1 = df1[df1['科目名称'].str.contains('期货')]
        df2_2 = df1[df1['科目名称'].str.contains('I')]
        future_df = pd.concat([df2_1, df2_2], axis=0)
        code = future_df['科目代码'].tolist()
        future_name = future_df['科目代码'].tolist()
        for i in range(len(future_name)):
            current_element = future_name[i]
            scode = current_element[8:]
            future_name[i] = scode
        future_amount = future_df['数量'].tolist()
        unit_cost = future_df['单位成本'].tolist()
        cost = future_df['成本'].tolist()
        pure_relative_cost = future_df['成本占净值(%)'].tolist()
        market_price = future_df['市价'].tolist()
        future_value = future_df['市值'].tolist()
        pure_relative_future_value = future_df['市值占净值(%)'].tolist()
        valuation_appreciation = future_df['估值增值'].tolist()
        trade_info = future_df['停牌信息'].tolist()
        type1 = []
        for i in code:
            if '3102' in i:
                s = '衍生工具'
                type1.append(s)
        type2 = []
        for j in code:
            if '310201' in j:
                d = '中金所_投机_买方_股指'
                type2.append(d)
            if '310203' in j:
                d = '中金所_多头'
                type2.append(d)
            if '310204' in j:
                d = '中金所_空头'
                type2.append(d)
        direction = []
        for k in future_value:
            if k > 0:
                b = 'long'
                direction.append(b)
            if k < 0:
                b = 'short'
                direction.append(b)
        product_name = current_file[11:17]
        day = pd.Series([date] * len(future_df), index=future_df.index)
        result1 = pd.DataFrame()
        try:
            result1.insert(0, '日期', day)
            result1.insert(1, '产品名称', product_name)
            result1.insert(2, '种类', type1)
            result1.insert(3, '种类名称', type2)
            result1.insert(4, '代码', code)
            result1.insert(5, '方向', direction)
            result1.insert(6, '科目名称', future_name)
            result1.insert(7, '数量', future_amount)
            result1.insert(8, '单位成本', unit_cost)
            result1.insert(9, '成本', cost)
            result1.insert(10, '成本占净值%', pure_relative_cost)
            result1.insert(11, '市价', market_price)
            result1.insert(12, '市值', future_value)
            result1.insert(13, '市值占净值%', pure_relative_future_value)
            result1.insert(14, '估值增值', valuation_appreciation)
            result1.insert(15, '停牌信息', trade_info)
        except:
            print(day+product_name+"产品sheet2出现变动,请更正")

        df = pd.read_excel(a, header=None)                                     # sheet3 c_bond_tracking
        df.reset_index(drop=True, inplace=True)
        new_columns = df.iloc[2]
        df = df.iloc[3:]
        df.columns = new_columns
        df.set_index('科目代码', inplace=True)
        df1 = df.dropna(subset=['停牌信息'])
        df1 = df1[df1['科目名称'].str.len() <= 5]
        df2 = df1[df1['科目名称'].str.contains('转')]
        cbond_df = df2
        cbond_index = cbond_df.index.tolist()
        cbond_code_list = []
        for i in cbond_index:
            cbond_code = i[-6:]
            cbond_code_list.append(cbond_code)
        day = pd.Series([date] * len(cbond_df), index=cbond_df.index)
        cbond_name = cbond_df['科目名称'].tolist()
        cbond_amount = cbond_df['数量'].tolist()
        cbond_close = cbond_df['市价'].tolist()
        cbond_value = cbond_df['市值'].tolist()
        result2 = pd.DataFrame()
        try:
            result2.insert(0, '日期', day)
            result2.insert(1, '产品名称', product_name)
            result2.insert(2, '债券代码', cbond_code_list)
            result2.insert(3, '债券名称', cbond_name)
            result2.insert(4, '数量', cbond_amount)
            result2.insert(5, '市价', cbond_close)
            result2.insert(6, '市值', cbond_value)
        except:
            print(day + product_name + "产品sheet3出现变动,请更正")

        df = pd.read_excel(a, header=None)
        df.reset_index(drop=True, inplace=True)
        new_columns = df.iloc[2]
        df = df.iloc[3:]
        df.columns = new_columns

        df1 = df.dropna(subset=['停牌信息'])                                             #sheet4 option_tracking
        df1 = df1[df1['科目名称'].str.len() >= 6]
        df3_1 = df1[df1['科目名称'].str.contains('沽')]
        df3_2 = df1[df1['科目名称'].str.contains('购')]
        df3 = pd.concat([df3_1, df3_2], axis=0)
        df3.loc[:, '科目名称'] = df3.loc[:, '科目名称'].apply(SF1000_option_name_transfer)
        day = pd.Series([date] * len(df3), index=df3.index)
        product_name = current_file[11:17]
        code = df3['科目代码'].tolist()
        type1 = []
        for i in code:
            if '3102' in i:
                s = '衍生工具'
                type1.append(s)
        option_value = df3['市值'].tolist()
        direction = []
        for j in option_value:
            if j > 0:
                z = 'long'
                direction.append(z)
            if j < 0:
                z = 'short'
                direction.append(z)
        try:
            df3['方向'] = direction
            df3['日期'] = day
            df3['产品名称'] = product_name
            df3['种类'] = type1
            df3['代码'] = code
            df3.rename(columns={'成本占净值(%)': '成本占净值%'}, inplace=True)
            df3.rename(columns={'市值占净值(%)': '市值占净值%'}, inplace=True)
            df3 = df3[
                ['日期', '产品名称', '种类', '代码', '科目名称','方向', '数量', '单位成本', '成本', '成本占净值%', '市价', '市值',
                 '市值占净值%',
                 '估值增值', '停牌信息']]
        except:
            print(day+product_name+"产品sheet4出现变动,请更正")

        result3 = pd.DataFrame(data=None, columns=['日期', '产品名称', '债券代码', '债券名称', '数量', '市价', '市值'])

        result_path = glv.get('outputpath_SF1000_N01')
        result_path_final = os.path.join(result_path, date + '_盛丰1000指增产品跟踪.xlsx')
        gt.create_file_directory(result_path_final)
        with pd.ExcelWriter(result_path_final, engine='openpyxl') as writer:
            result.to_excel(writer, sheet_name='stock_tracking', index=False)
            result2.to_excel(writer, sheet_name='c_bond_tracking', index=False)
            result1.to_excel(writer, sheet_name='future_tracking', index=False)
            df3.to_excel(writer, sheet_name='option_tracking', index=False)
            result3.to_excel(writer, sheet_name='bond_tracking', index=False)


def GYZY_N01(file_name):
    # 输入为 高益振英一号文件名
    # 输出为 解析后的四级估值表 分为股票 可转债 期货 期权 国债五个模块
    folder_path = glv.get('folder_path')
    for i in file_name:
        current_file = i
        date = current_file[-12:-4]
        product_name = current_file[0:6]
        a = os.path.join(folder_path, current_file)
        a = a.replace('\\', '\\\\')
        if '.xls' in a:
            df = pd.read_excel(a, header=None)                                           #sheet1 stock_tracking
            df.reset_index(drop=True, inplace=True)
            new_columns = df.iloc[4]
            df = df.iloc[8:]
            df.columns = new_columns
            df.set_index('科目代码', inplace=True)
            df1 = df.dropna(subset=['停牌信息'])
            df1 = df1[df1['科目名称'].str.len() <= 5]
            df1 = df1.drop(df1[df1['科目名称'].str.contains('T2')].index)
            df2 = df1[df1['科目名称'].str.contains('转')]
            stock_df = df1[~df1.index.isin(df2.index)]
            stock_index = stock_df.index.tolist()
            stock_code_list = []
            for i in stock_index:
                stock_code = i[-9:-3]
                stock_code_list.append(stock_code)
            day = pd.Series([date] * len(stock_df), index=stock_df.index)
            stock_name = stock_df['科目名称'].tolist()
            stock_amount = stock_df['数量'].tolist()
            stock_close = stock_df['行情'].tolist()
            stock_value = stock_df['市值-本币'].tolist()
            result = pd.DataFrame()
            try:
                result.insert(0, '日期', day)
                result.insert(1, '产品名称', product_name)
                result.insert(2, '股票代码', stock_code_list)
                result.insert(3, '股票名称', stock_name)
                result.insert(4, '数量', stock_amount)
                result.insert(5, '市价', stock_close)
                result.insert(6, '市值', stock_value)
            except:
                print(day+product_name+"产品sheet1出现变动,请更正")

            df = pd.read_excel(a, header=None)                                            #sheet2 future_tracking
            df.reset_index(drop=True, inplace=True)
            new_columns = df.iloc[4]
            df = df.iloc[8:]
            df.columns = new_columns
            df1 = df.dropna(subset=['停牌信息'])
            df1 = df1[df1['科目名称'].str.len() >= 6]
            df1 = df1.drop(df1[df1['科目名称'].str.contains('国债')].index)
            df2_1 = df1[df1['科目名称'].str.contains('期货')]
            df2_2 = df1[df1['科目名称'].str.contains('I')]
            future_df = pd.concat([df2_1, df2_2], axis=0)
            code = future_df['科目代码'].tolist()
            future_name = future_df['科目名称'].tolist()
            future_amount = future_df['数量'].tolist()
            unit_cost = future_df['单位成本'].tolist()
            cost = future_df['成本-本币'].tolist()
            pure_relative_cost = future_df['成本占比'].tolist()
            market_price = future_df['行情'].tolist()
            future_value = future_df['市值-本币'].tolist()
            pure_relative_future_value = future_df['市值占比'].tolist()
            valuation_appreciation = [cost[m]-future_value[m] for m in range(0,len(cost))]
            trade_info = future_df['停牌信息'].tolist()
            type1 = []
            for i in code:
                if '3102' in i:
                    s = '衍生工具'
                    type1.append(s)
            type2 = []
            for j in code:
                if '3102.01' in j:
                    d = '中金所_投机_买方_股指'
                    type2.append(d)
                if '3102.03' in j:
                    d = '中金所_投机_卖方_股指'
                    type2.append(d)
            direction = []
            for k in future_value:
                if k > 0:
                    b = 'long'
                    direction.append(b)
                if k < 0:
                    b = 'short'
                    direction.append(b)
            day = pd.Series([date] * len(future_df), index=future_df.index)
            product_name = current_file[:6]
            result1 = pd.DataFrame()
            try:
                result1.insert(0, '日期', day)
                result1.insert(1, '产品名称', product_name)
                result1.insert(2, '种类', type1)
                result1.insert(3, '种类名称', type2)
                result1.insert(4, '代码', code)
                result1.insert(5, '方向', direction)
                result1.insert(6, '科目名称', future_name)
                result1.insert(7, '数量', future_amount)
                result1.insert(8, '单位成本', unit_cost)
                result1.insert(9, '成本', cost)
                result1.insert(10, '成本占净值%', pure_relative_cost)
                result1.insert(11, '市价', market_price)
                result1.insert(12, '市值', future_value)
                result1.insert(13, '市值占净值%', pure_relative_future_value)
                result1.insert(14, '估值增值', valuation_appreciation)
                result1.insert(15, '停牌信息', trade_info)
            except:
                print(day+product_name+"产品sheet2出现变动,请更正")

            df = pd.read_excel(a, header=None)                                                  # sheet3 c_bond_tracking
            df.reset_index(drop=True, inplace=True)
            new_columns = df.iloc[4]
            df = df.iloc[8:]
            df.columns = new_columns
            df.set_index('科目代码', inplace=True)
            df1 = df.dropna(subset=['停牌信息'])
            df1 = df1[df1['科目名称'].str.len() <= 5]
            df2 = df1[df1['科目名称'].str.contains('转')]
            cbond_df = df2
            cbond_index = cbond_df.index.tolist()
            cbond_code_list = []
            for i in cbond_index:
                cbond_code = i[-9:-3]
                cbond_code_list.append(cbond_code)
            day = pd.Series([date] * len(cbond_df), index=cbond_df.index)
            cbond_name = cbond_df['科目名称'].tolist()
            cbond_amount = cbond_df['数量'].tolist()
            cbond_close = cbond_df['成本-本币'].tolist()
            cbond_value = cbond_df['市值-本币'].tolist()
            result2 = pd.DataFrame()
            try:
                result2.insert(0, '日期', day)
                result2.insert(1, '产品名称', product_name)
                result2.insert(2, '债券代码', cbond_code_list)
                result2.insert(3, '债券名称', cbond_name)
                result2.insert(4, '数量', cbond_amount)
                result2.insert(5, '市价', cbond_close)
                result2.insert(6, '市值', cbond_value)
            except:
                print(day + product_name + "产品sheet3出现变动,请更正")

            df = pd.read_excel(a, header=None)
            df.reset_index(drop=True, inplace=True)
            new_columns = df.iloc[4]
            df = df.iloc[8:]
            df.columns = new_columns
            df1 = df.dropna(subset=['停牌信息'])                                             #sheet4 option_tracking
            df1 = df1[df1['科目名称'].str.len() >= 6]
            df3 = df1[df1['科目名称'].str.contains('-')]
            df3.loc[:, '科目名称'] = df3.loc[:, '科目名称'].apply(option_name_transfer)
            df.set_index('科目代码', inplace=True)
            day = pd.Series([date] * len(df3), index=df3.index)
            product_name = current_file[:6]
            code = df3['科目代码'].tolist()
            type1 = []
            for i in code:
                if '3102' in i:
                    s = '衍生工具'
                    type1.append(s)
                else:
                    s = ''
                    type1.append(s)
            option_value = df3['市值-本币'].tolist()
            direction = []
            for j in option_value:
                if j > 0:
                    z = 'long'
                    direction.append(z)
                if j < 0:
                    z = 'short'
                    direction.append(z)
            try:
                df3['方向'] = direction
                df3['日期'] = day
                df3['产品名称'] = product_name
                df3['种类'] = type1
                df3['代码'] = code
                df3.rename(columns={'成本-本币': '成本'}, inplace=True)
                df3.rename(columns={'市值-本币': '市值'}, inplace=True)
                df3.rename(columns={'估值增值-本币': '估值增值'}, inplace=True)
                df3.rename(columns={'成本占比': '成本占净值%'}, inplace=True)
                df3.rename(columns={'市值占比': '市值占净值%'}, inplace=True)
                df3.rename(columns={'行情': '市价'}, inplace=True)
                df3 = df3[
                    ['日期', '产品名称', '种类', '代码', '科目名称', '方向', '数量', '单位成本', '成本', '成本占净值%',
                     '市价',
                     '市值',
                     '市值占净值%',
                     '估值增值', '停牌信息']]
            except:
                print(day+product_name+"产品sheet4出现变动,请更正")

            df = pd.read_excel(a, header=None)                                                  # sheet5 bond_tracking
            df.reset_index(drop=True, inplace=True)
            new_columns = df.iloc[4]
            df = df.iloc[8:]
            df.columns = new_columns
            df.set_index('科目代码', inplace=True)
            df1 = df.dropna(subset=['停牌信息'])
            df1 = df1[df1['科目名称'].str.len() <= 5]
            df2 = df1[df1['科目名称'].str.contains('T2')]
            bond_df = df2
            bond_index = bond_df.index.tolist()
            bond_code_list = []
            for i in bond_index:
                bond_code = i[-9:-3]
                bond_code_list.append(bond_code)
            day = pd.Series([date] * len(bond_df), index=bond_df.index)
            bond_name = bond_df['科目名称'].tolist()
            bond_amount = bond_df['数量'].tolist()
            bond_close = bond_df['行情'].tolist()
            bond_value = bond_df['市值-本币'].tolist()
            result3 = pd.DataFrame()
            try:
                result3.insert(0, '日期', day)
                result3.insert(1, '产品名称', product_name)
                result3.insert(2, '债券代码', bond_code_list)
                result3.insert(3, '债券名称', bond_name)
                result3.insert(4, '数量', bond_amount)
                result3.insert(5, '市价', bond_close)
                result3.insert(6, '市值', bond_value)
            except:
                print(day + product_name + "产品sheet5出现变动,请更正")

            result_path = glv.get('outputpath_GYZY_N01')
            result_path_final = os.path.join(result_path, date + '_高益振英一号指增产品跟踪.xlsx')
            gt.create_file_directory(result_path_final)
            with pd.ExcelWriter(result_path_final, engine='openpyxl') as writer:
                result.to_excel(writer, sheet_name='stock_tracking', index=False)
                result2.to_excel(writer, sheet_name='c_bond_tracking', index=False)
                result1.to_excel(writer, sheet_name='future_tracking', index=False)
                df3.to_excel(writer, sheet_name='option_tracking', index=False)
                result3.to_excel(writer, sheet_name='bond_tracking', index=False)


def RenRui_N01(file_name):
    # 输入为 仁睿价值精选1号产品文件名
    # 输出为 解析后的四级估值表 分为股票 可转债 期货 期权 国债五个模块
    folder_path = glv.get('folder_path')
    for i in file_name:
        current_file = i
        date = current_file[-17:-7]
        date = date[0:4] + date[5:7] + date[8:]
        product_name = current_file[0:6]
        a = os.path.join(folder_path, current_file)
        a = a.replace('\\', '\\\\')
        if '.xls' in a:
            df = pd.read_excel(a, header=None)                                             #sheet1 stock_tracking
            df.reset_index(drop=True, inplace=True)
            new_columns = df.iloc[3]
            df = df.iloc[4:]
            df.columns = new_columns
            df.set_index('科目代码', inplace=True)
            df1 = df.dropna(subset=['停牌信息'])
            df1 = df1[df1['科目名称'].str.len() <= 5]
            df1 = df1.drop(df1[df1['科目名称'].str.contains('T2')].index)
            df2 = df1[df1['科目名称'].str.contains('转')]
            stock_df = df1[~df1.index.isin(df2.index)]
            stock_index = stock_df.index.tolist()
            stock_code_list = []
            for i in stock_index:
                stock_code = i[-6:]
                stock_code_list.append(stock_code)
            day = pd.Series([date] * len(stock_df), index=stock_df.index)
            stock_name = stock_df['科目名称'].tolist()
            stock_amount = stock_df['数量'].tolist()
            stock_close = stock_df['市价'].tolist()
            stock_value = stock_df['市值'].tolist()
            result = pd.DataFrame()
            try:
                result.insert(0, '日期', day)
                result.insert(1, '产品名称', product_name)
                result.insert(2, '股票代码', stock_code_list)
                result.insert(3, '股票名称', stock_name)
                result.insert(4, '数量', stock_amount)
                result.insert(5, '市价', stock_close)
                result.insert(6, '市值', stock_value)
            except:
                print(day+product_name+"产品sheet1出现变动,请更正")

            df = pd.read_excel(a, header=None)                                               #sheet2 future_tracking
            df.reset_index(drop=True, inplace=True)
            new_columns = df.iloc[3]
            df = df.iloc[4:]
            df.columns = new_columns
            df1 = df.dropna(subset=['停牌信息'])
            df1 = df1[df1['科目名称'].str.len() >= 6]
            df1 = df1.drop(df1[df1['科目名称'].str.contains('国债')].index)
            df2_1 = df1[df1['科目名称'].str.contains('期货')]
            df2_2 = df1[df1['科目名称'].str.contains('I')]
            future_df = pd.concat([df2_1, df2_2], axis=0)
            code = future_df['科目代码'].tolist()
            future_name = future_df['科目代码'].tolist()
            for i in range(len(future_name)):
                current_element = future_name[i]
                scode = current_element[-6:]
                future_name[i] = scode
            future_amount = future_df['数量'].tolist()
            unit_cost = future_df['单位成本'].tolist()
            cost = future_df['成本'].tolist()
            pure_relative_cost = future_df['成本占净值%'].tolist()
            market_price = future_df['市价'].tolist()
            future_value = future_df['市值'].tolist()
            pure_relative_future_value = future_df['市值占净值%'].tolist()
            valuation_appreciation = future_df['估值增值'].tolist()
            trade_info = future_df['停牌信息'].tolist()
            type1 = []
            for i in code:
                if '3102' in i:
                    s = '衍生工具'
                    type1.append(s)
            type2 = []
            for j in code:
                if '310201' in j:
                    d = '中金所_投机_买方_股指'
                    type2.append(d)
                if '310203' in j:
                    d = '中金所_多头'
                    type2.append(d)
                if '310204' in j:
                    d = '中金所_空头'
                    type2.append(d)
            direction = []
            for k in future_value:
                if k > 0:
                    b = 'long'
                    direction.append(b)
                if k < 0:
                    b = 'short'
                    direction.append(b)

            product_name = current_file[:6]
            day = pd.Series([date] * len(future_df), index=future_df.index)
            result1 = pd.DataFrame()
            try:
                result1.insert(0, '日期', day)
                result1.insert(1, '产品名称', product_name)
                result1.insert(2, '种类', type1)
                result1.insert(3, '种类名称', type2)
                result1.insert(4, '代码', code)
                result1.insert(5, '方向', direction)
                result1.insert(6, '科目名称', future_name)
                result1.insert(7, '数量', future_amount)
                result1.insert(8, '单位成本', unit_cost)
                result1.insert(9, '成本', cost)
                result1.insert(10, '成本占净值%', pure_relative_cost)
                result1.insert(11, '市价', market_price)
                result1.insert(12, '市值', future_value)
                result1.insert(13, '市值占净值%', pure_relative_future_value)
                result1.insert(14, '估值增值', valuation_appreciation)
                result1.insert(15, '停牌信息', trade_info)
            except:
                print(day+product_name+"产品sheet2出现变动,请更正")

            df = pd.read_excel(a, header=None)                                                  # sheet3 c_bond_tracking
            df.reset_index(drop=True, inplace=True)
            new_columns = df.iloc[3]
            df = df.iloc[4:]
            df.columns = new_columns
            df.set_index('科目代码', inplace=True)

            df1 = df.dropna(subset=['停牌信息'])
            df1 = df1[df1['科目名称'].str.len() <= 5]
            df2 = df1[df1['科目名称'].str.contains('转')]
            cbond_df = df2
            cbond_index = cbond_df.index.tolist()
            cbond_code_list = []
            for i in cbond_index:
                cbond_code = i[-6:]
                cbond_code_list.append(cbond_code)
            day = pd.Series([date] * len(cbond_df), index=cbond_df.index)
            cbond_name = cbond_df['科目名称'].tolist()
            cbond_amount = cbond_df['数量'].tolist()
            cbond_close = cbond_df['市价'].tolist()
            cbond_value = cbond_df['市值'].tolist()
            result2 = pd.DataFrame()
            try:
                result2.insert(0, '日期', day)
                result2.insert(1, '产品名称', product_name)
                result2.insert(2, '债券代码', cbond_code_list)
                result2.insert(3, '债券名称', cbond_name)
                result2.insert(4, '数量', cbond_amount)
                result2.insert(5, '市价', cbond_close)
                result2.insert(6, '市值', cbond_value)
            except:
                print(day + product_name + "产品sheet3出现变动,请更正")

            df = pd.read_excel(a, header=None)
            df.reset_index(drop=True, inplace=True)
            new_columns = df.iloc[3]
            df = df.iloc[4:]
            df.columns = new_columns
            df1 = df.dropna(subset=['停牌信息'])                                             #sheet4 option_tracking
            df1 = df1[df1['科目名称'].str.len() >= 6]
            df3 = df1[df1['科目名称'].str.contains('-')]
            df3.loc[:, '科目名称'] = df3.loc[:, '科目名称'].apply(option_name_transfer)
            day = pd.Series([date] * len(df3), index=df3.index)
            product_name = current_file[:6]
            code = df3['科目代码'].tolist()
            type1 = []
            for i in code:
                if '3102' in i:
                    s = '衍生工具'
                    type1.append(s)
            direction = []
            option_value = df3['市值'].tolist()
            for j in option_value:
                if j > 0:
                    z = 'long'
                    direction.append(z)
                if j < 0:
                    z = 'short'
                    direction.append(z)
            try:
                df3['方向'] = direction
                df3['日期'] = day
                df3['产品名称'] = product_name
                df3['种类'] = type1
                df3['代码'] = code
                df3 = df3[
                    ['日期', '产品名称', '种类', '代码', '科目名称','方向', '数量', '单位成本', '成本', '成本占净值%', '市价', '市值',
                     '市值占净值%',
                     '估值增值', '停牌信息']]
            except:
                print(day+product_name+"产品sheet4出现变动,请更正")

            df = pd.read_excel(a, header=None)                                                  # sheet5 bond_tracking
            df.reset_index(drop=True, inplace=True)
            new_columns = df.iloc[3]
            df = df.iloc[4:]
            df.columns = new_columns
            df.set_index('科目代码', inplace=True)
            df1 = df.dropna(subset=['停牌信息'])
            df1 = df1[df1['科目名称'].str.len() >= 6]
            df2 = df1[df1['科目名称'].str.contains('国债')]
            bond_df = df2
            bond_index = bond_df.index.tolist()
            bond_code_list = []
            for i in bond_index:
                bond_code = i[-6:]
                bond_code_list.append(bond_code)
            day = pd.Series([date] * len(bond_df), index=bond_df.index)
            bond_name = bond_df['科目名称'].tolist()
            bond_amount = bond_df['数量'].tolist()
            bond_close = bond_df['市价'].tolist()
            bond_value = bond_df['市值'].tolist()
            result3 = pd.DataFrame()
            try:
                result3.insert(0, '日期', day)
                result3.insert(1, '产品名称', product_name)
                result3.insert(2, '债券代码', bond_code_list)
                result3.insert(3, '债券名称', bond_name)
                result3.insert(4, '数量', bond_amount)
                result3.insert(5, '市价', bond_close)
                result3.insert(6, '市值', bond_value)
            except:
                print(day + product_name + "产品sheet5出现变动,请更正")

            result_path = glv.get('outputpath_RenRui_N01')
            result_path_final = os.path.join(result_path, date + '_仁睿价值精选1号产品跟踪.xlsx')
            gt.create_file_directory(result_path_final)
            with pd.ExcelWriter(result_path_final, engine='openpyxl') as writer:
                result.to_excel(writer, sheet_name='stock_tracking', index=False)
                result2.to_excel(writer, sheet_name='c_bond_tracking', index=False)
                result1.to_excel(writer, sheet_name='future_tracking', index=False)
                df3.to_excel(writer, sheet_name='option_tracking', index=False)
                result3.to_excel(writer, sheet_name='bond_tracking', index=False)




def main(product_code, time):
    inputpath_chart_folder_path = glv.get('folder_path')
    filelist = os.listdir(inputpath_chart_folder_path)
    # 主函数 自动运行七个产品检索及输出
    if 'SSS044' in product_code:
        return RR500(extract(product_code,time,filelist))
    elif 'SNY426' in product_code:
        return RRJX(extract(product_code,time,filelist))
    elif 'SZJ339' in product_code:
        return SF500_N08(extract(product_code,time,filelist))
    elif 'SGS958' in product_code:
        return XYHY_N01(extract(product_code,time,filelist))
    elif 'SVX619' in product_code:
        return SF1000_N01(extract(product_code,time,filelist))
    elif 'SVU353' in product_code:
        return GYZY_N01(extract(product_code,time,filelist))
    elif 'SLA626' in product_code:
        return RenRui_N01(extract(product_code,time,filelist))
    else:
        print("No file extracted")


def history_main():
    # 输入为 上一级文件config文件夹下history_main.xlsx
    # 输出为 输入的excel时间段内产品估值表解析
    inputpath = os.path.join(os.path.abspath('../..'), "config\history_main.xlsx")
    df= pd.read_excel(inputpath,sheet_name='History',index_col=0)
    df.reset_index(drop=False, inplace=True)

    for index, row in df.iterrows():
        start_date = pd.to_datetime(row['start_date'],format='%Y%m%d', errors='ignore')
        end_date = pd.to_datetime(row['end_date'],format='%Y%m%d', errors='ignore')
        start_date = gt.strdate_transfer(start_date)
        end_date = gt.strdate_transfer(end_date)
        list = gt.working_days_list(start_date, end_date)
        for i in list:
        # time 为 i
            i = gt.intdate_transfer(i)
            main(row['product_code'], i)

def auto_main1():
    # 判断五类产品最新日期更新 如果更新失败输出为产品日期存在数据缺失
    time = gt.last_workday()
    list=['SSS044','SNY426','SGS958','SZJ339','SVX619','SVU353','SLA626']
    for i in list:
        main(i,time)

if __name__ == '__main__':
    history_main()
    # auto_main1()
    pass

