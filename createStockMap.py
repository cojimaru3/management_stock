import requests
import pandas as pd
import lxml.html
import sys
import time
from time import strftime

SCREENING = True
ENABLE_DIVIDEND_YIELD = True
TARGET_DIVIDEND_YIELD = 3.75

ENABLE_PAYOUT_RATIO = True
TARGET_PAYOUT_RATIO = 50

ENABLE_PBR = True
TARGET_PBR = 1.5

ENABLE_CAPTIAL_ADEQUACY_RATIO = True
TARGET_CAPTIAL_ADEQUACY_RATIO = 50

ENABLE_OPERATING_PROFIT_RATIO = True
TARGET_OPERATING_PROFIT_RATIO = 10

ENABLE_CONTINUOUS_DIVIDEND_INCREASE = True
TARGET_CONTINUOUS_DIVIDEND_INCREASE = 3

"""
# Purpose
Convert source variable contained in str variable to destination variable.
"""
def util_replace(text, source, destination):
    # if text has source
    if text.find(source) > 0:
        # remove unnecessary ','
        tmp = text.replace(',','')
        return float(tmp.replace(source, destination))
    else:
        # Replace non-valued data with None.
        return None

"""
# Purpose
Confirm that it matches the following conditions.
"""

def util_screening(dividend_yield,dividend_payout_ratio,pbr,capital_adequacy_ratio,operating_profit_ratio,continuous_dividend_increase):
    if ENABLE_DIVIDEND_YIELD:
        # Dividend_yield has a high priority and does not allow None.
        if dividend_yield is not None and dividend_yield >= TARGET_DIVIDEND_YIELD: 
            pass
        else:
            return False
    if ENABLE_PAYOUT_RATIO:
        # Since dividend_payout_ratio has a low priority, allow None.
        if dividend_payout_ratio is None or dividend_payout_ratio <= TARGET_PAYOUT_RATIO:
            pass
        else:
            return False
    if ENABLE_PBR:
        # Since pbr has a low priority, allow None.
        if pbr is None or pbr <= TARGET_PBR:
            pass
        else:
            return False
    if ENABLE_CAPTIAL_ADEQUACY_RATIO:
        if capital_adequacy_ratio is None or capital_adequacy_ratio >= TARGET_CAPTIAL_ADEQUACY_RATIO:
            pass
        else:
            return False
    if ENABLE_OPERATING_PROFIT_RATIO:
        # Since operating_profit_ratio has a low priority, allow None.
        if operating_profit_ratio is None or operating_profit_ratio >= TARGET_OPERATING_PROFIT_RATIO:
            pass
        else:
            return False
    if ENABLE_CONTINUOUS_DIVIDEND_INCREASE:
        # Since continuous_dividend_increase has a high priority, does not allow None.
        if continuous_dividend_increase is not None and continuous_dividend_increase >= TARGET_CONTINUOUS_DIVIDEND_INCREASE:
            pass
        else:
            return False
    return True

"""
# Purpose
Parse dom_tree and get text data.
"""
def parse_dom_tree(dom_tree, xpath, source, destination):
    raw_data = dom_tree.xpath(xpath)
    data = util_replace(raw_data[0].text,source,destination)
    return data

"""
# Purpose
Confirm that the string can be converted to float type.
"""
def isfloat(text):
    try:
        float(text)
    except ValueError:
        return False
    else:
        return True

"""
# Purpose
Convert string type to float type. 
"""
def convert_string_float(text):
    if isfloat(text):
        return float(text)
    else:
        return None

if __name__ == '__main__':
    # Obtain the code and stock information listed on the First Section of the Tokyo Stock Exchange from https://www.jpx.co.jp/. 
    source_path = 'https://www.jpx.co.jp/markets/statistics-equities/misc/tvdivq0000001vg2-att/data_j.xls'
    data_frame = pd.read_excel(source_path,usecols=['コード','銘柄名','市場・商品区分','17業種区分'],sheet_name='Sheet1')
    stock_data_frame = data_frame[data_frame['市場・商品区分']=='市場第一部（内国株）']

    merged_data_path = 'fy-merged-sheet.csv' 
    merged_data_frame = pd.read_csv(merged_data_path,encoding='shift-jis',usecols=['コード','年度','自己資本比率','売上高','営業利益','一株配当','配当性向','連続増配','減配なし'])
    stock_merged_data_frame = merged_data_frame.set_index('コード')
    
    output_list = []
    for index,row in stock_data_frame.iterrows():
        stock_code = row[0]
        stock_name = str(row[1])
        stock_industry_type = str(row[3])
        if stock_code in stock_merged_data_frame.index:
            merged_data_list = stock_merged_data_frame.loc[int(stock_code)]
        else:
            continue
        chart_url = "https://minkabu.jp/stock/" + str(stock_code) + "/chart"
        
        try:
            chart_html = requests.get(chart_url)
            chart_html.raise_for_status()
            chart_dom_tree = lxml.html.fromstring(chart_html.content)

            print(stock_code)
            # 前日終値
            stock_previous_closing_place = parse_dom_tree(chart_dom_tree,'//*[@id="contents"]/div[3]/div[1]/div/div/div[2]/div/div[1]//tr[1]/td[1]','円','')            
            #print('stock_previous_closing_place : ' + str(stock_previous_closing_place))
                
            # 高値
            stock_high_place = parse_dom_tree(chart_dom_tree,'//*[@id="contents"]/div[3]/div[1]/div/div/div[2]/div/div[1]//tr[3]/td[1]','円','')
            #print('stock_high_place : ' + str(stock_high_place))

            # 安値
            stock_low_place = parse_dom_tree(chart_dom_tree,'//*[@id="contents"]/div[3]/div[1]/div/div/div[2]/div/div[1]//tr[4]/td[1]','円','')
            #print('stock_low_place : ' + str(stock_low_place))

            # PER
            stock_per = parse_dom_tree(chart_dom_tree,'//*[@id="contents"]/div[3]/div[1]/div/div/div[2]/div/div[2]//tr[1]/td[1]','倍','')
            #print('stock_per : ' + str(stock_per))

            # PBR
            stock_pbr = parse_dom_tree(chart_dom_tree,'//*[@id="contents"]/div[3]/div[1]/div/div/div[2]/div/div[2]//tr[2]/td[1]','倍','')
            #print('stock_pbr : ' + str(stock_pbr))

            # 年初来高値
            stock_high_place_per_year = parse_dom_tree(chart_dom_tree,'//*[@id="contents"]/div[3]/div[1]/div/div/div[2]/div/div[2]//tr[6]/td[1]','円','')
            #print('stock_high_place_per_year : ' + str(stock_high_place_per_year))

            # 年初来安値
            stock_low_place_per_year = parse_dom_tree(chart_dom_tree,'//*[@id="contents"]/div[3]/div[1]/div/div/div[2]/div/div[2]//tr[7]/td[1]','円','')
            #print('stock_low_place_per_year : ' + str(stock_low_place_per_year))

            # 配当利回り
            stock_dividend_yield = parse_dom_tree(chart_dom_tree,'//*[@id="contents"]/div[3]/div[1]/div/div/div[2]/div/div[2]//tr[3]/td[1]','%','')
            #print('stock_dividend_yield : ' + str(stock_dividend_yield))

            # 年度
            stock_term = merged_data_list['年度']

            # 自己資本比率
            stock_capital_adequacy_ratio = merged_data_list['自己資本比率']

            # 売上高営業利益率 営業利益 / 売上高
            stock_operating_profit = convert_string_float(merged_data_list['営業利益'])
            stock_sales = convert_string_float(merged_data_list['売上高'])
            if stock_operating_profit is None or stock_sales is None:
                stock_operating_profit_ratio = None
            else:
                stock_operating_profit_ratio = (stock_operating_profit / stock_sales) * 100

            # 一株配当
            stock_dividend_per_share = convert_string_float(merged_data_list['一株配当'])

            # 配当性向
            stock_dividend_payout_ratio = convert_string_float(merged_data_list['配当性向'])

            # 連続増配
            stock_continuous_dividend_increase = merged_data_list['連続増配']

            # 減配なし
            stock_dividend_reduction = merged_data_list['減配なし']

            if(SCREENING == True):
                if(util_screening(stock_dividend_yield,stock_dividend_payout_ratio,\
                stock_pbr,stock_capital_adequacy_ratio,\
                stock_operating_profit_ratio,stock_continuous_dividend_increase)):
                    list = [stock_code,\
                            stock_name,\
                            stock_previous_closing_place,\
                            stock_high_place,\
                            stock_low_place,\
                            stock_high_place_per_year,\
                            stock_low_place_per_year,\
                            stock_per,\
                            stock_pbr,\
                            stock_dividend_yield,\
                            stock_term,\
                            stock_sales,\
                            stock_operating_profit,\
                            stock_operating_profit_ratio,\
                            stock_capital_adequacy_ratio,\
                            stock_dividend_per_share,\
                            stock_dividend_payout_ratio,\
                            stock_continuous_dividend_increase,\
                            stock_dividend_reduction,\
                            stock_industry_type]
                else:
                    continue
            else:
                list = [stock_code,\
                        stock_name,\
                        stock_previous_closing_place,\
                        stock_high_place,\
                        stock_low_place,\
                        stock_high_place_per_year,\
                        stock_low_place_per_year,\
                        stock_per,\
                        stock_pbr,\
                        stock_dividend_yield,\
                        stock_term,\
                        stock_sales,\
                        stock_operating_profit,\
                        stock_operating_profit_ratio,\
                        stock_capital_adequacy_ratio,\
                        stock_dividend_per_share,\
                        stock_dividend_payout_ratio,\
                        stock_continuous_dividend_increase,\
                        stock_dividend_reduction,\
                        stock_industry_type]                
            output_list.append(list)
        except requests.exceptions.RequestException as e:
            print(e)
    columns = ['コード',\
                '銘柄',\
                '前日終値(円)',\
                '高値(円)',\
                '安値(円)',\
                '年初来高値(円)',\
                '年初来安値(円)',\
                'PER(倍)',\
                'PBR(倍)',\
                '配当利回り(%)',\
                '年度',\
                '売上高',\
                '営業利益',\
                '売上高営業利益率',\
                '自己資本率',\
                '一株配当',\
                '配当性向(%)',\
                '連続増配',\
                '減配なし',\
                '業種']
    output_data_frame = pd.DataFrame(data=output_list,columns=columns)
    sorted_output_data_frame = output_data_frame.sort_values('配当利回り(%)',ascending=False)
    localtime = strftime("%y%m%d_%H%M%S",time.localtime())
    output_file_name = './stock_' + localtime + '.xlsx'
    sorted_output_data_frame.to_excel(output_file_name,index=None,engine='openpyxl')
    