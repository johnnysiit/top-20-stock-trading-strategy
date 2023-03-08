import matplotlib.pyplot as plt
import os
import pandas as pd
#pyxl is also required. If you don't have it, please install it by using pip3 install pyxl


#导入数据
def read_file():
    close_price_table = pd.read_excel('沪深300成分股收盘价.xlsx', index_col=0)
    #weight_table = pd.read_excel('沪深300成分股.xlsx')
    close_price_table = close_price_table.replace(0, pd.np.nan)
    return close_price_table

#生成因子/涨跌幅表格
def calculate_return(close_price_table):
    return_table = close_price_table.pct_change()
    return_table.to_csv('涨跌幅表格.csv')
    return return_table

#分成五级组合
def portfolio_divider(return_table):
    total_week_dict = {}
    for i in range(0, len(return_table)):
        weekly_table = return_table.iloc[i]
        weekly_table = weekly_table.dropna()
        weekly_table = weekly_table.sort_values(ascending=False)
        week_dict = {}
        for level in range(0,5):
            week_dict[(level+1)] = list()
            for k in range(level*int(len(weekly_table)/5),(level+1)*int(len(weekly_table)/5)):
                week_dict[(level+1)].append(weekly_table.index[k])
        total_week_dict[return_table.index[i]] = week_dict
    return total_week_dict

#输出每周每个组合的详细持仓列表
def weekly_holding(total_week_holding):
    df = pd.DataFrame(total_week_holding)
    df = df.T
    df = df.stack()
    df = df.reset_index()
    df.columns = ['date', 'level', 'holding']
    df = df.set_index(['date', 'level'])
    df = df['holding'].apply(pd.Series)
    df.to_csv('五个组合每周详细持仓.csv')

#输出每周每个组合的平均收益率
def weekly_portf_return(return_table, holding_dict):
    weekly_porf_return_dict = {}
    for date in holding_dict:
        try:
            nextdate = return_table.index[return_table.index.get_loc(date)+1]
            for level in holding_dict[date]:
                return_percentage_list = []
                for stock in holding_dict[date][level]:
                    return_percentage_list.append(return_table[stock][nextdate])
                try:
                    weekly_porf_return_dict[(date, level)] = sum(return_percentage_list)/len(return_percentage_list)
                except:
                    weekly_porf_return_dict[(date, level)] = 0
        except:
            pass

    #生成表格
    df = pd.DataFrame(weekly_porf_return_dict, index=[0])
    df = df.T
    df = df.reset_index()
    df.columns = ['date', 'level', 'return']
    df = df.set_index(['date', 'level'])

    return_sheet = pd.DataFrame()
    return_sheet['date'] = (df.loc[df.index.get_level_values('level') == 1]).index.get_level_values('date')
    return_sheet = return_sheet.set_index('date')
    for i in range(1,6):
        df_level = df.loc[df.index.get_level_values('level') == i]
        df_level = df_level.droplevel('level')
        for j in range(0, len(df_level)):
            return_sheet.loc[df_level.index[j], str("Level"+str(i))] = df_level.iloc[j]['return']
    return_sheet.to_csv('五个组合每周收益率.csv')

    return df

#计算净值曲线
def value_test(df):
    #分级计算价值并输出表格
    initia_value = 1
    value_sheet = pd.DataFrame()
    value_sheet['date'] = (df.loc[df.index.get_level_values('level') == 1]).index.get_level_values('date')
    value_sheet = value_sheet.set_index('date')
    for i in range(1,6):
        value = initia_value
        df_level = df.loc[df.index.get_level_values('level') == i]
        df_level = df_level.droplevel('level')
        for j in range(0, len(df_level)):
            value = value * (1 + df_level.iloc[j]['return'])
            value_sheet.loc[df_level.index[j], str("Level"+str(i))] = value
    value_sheet.to_csv('净值曲线表格.csv')



    #生成净值曲线图
    plt.clf()
    for i in range(1,6):
        plt.plot(value_sheet.index, value_sheet['Level'+str(i)])
    plt.legend(['Level1', 'Level2', 'Level3', 'Level4', 'Level5'])
    plt.title('Value Curve')
    plt.xlabel('Date')
    plt.ylabel('Value')
    plt.savefig('净值曲线图.png')



if __name__ == '__main__':
    close_price_table = read_file()
    return_table = calculate_return(close_price_table)
    portf_divided = portfolio_divider(return_table)
    weekly_holding(portf_divided)
    weekly_por_return = weekly_portf_return(return_table, portf_divided)
    value_test(weekly_por_return)