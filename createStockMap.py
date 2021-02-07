import requests
import pandas as pd
import lxml.html
import sys
import time
from time import strftime

"""
Convert source variable contained in str variable to destination variable.
"""
def util_replace(str, source, destination):
    text = str[0].text
    # if text has source
    if text.find(source) > 0:
        # remove unnecessary ','
        tmp = text.replace(',','')
        return float(tmp.replace(source, destination))
    else:
        # Replace non-valued data with None.
        return None

"""
Confirm that it matches the following conditions.

conditions:
Dividend yield is 3.75 or higher.
Dividend payout ratio is 50% or less.
PBR is 1.5 times or less.
PER is pass.
"""
def util_screening(dividend_yield,dividend_payout_ratio,per,pbr):
    if (dividend_yield is not None and dividend_yield >= 3.75) \
     and (dividend_payout_ratio is None or dividend_payout_ratio <= 50) \
     and (pbr is None or pbr <= 1.5):
        return True
    else:
        return False

if __name__ == '__main__':
    # Obtain the code and stock information listed on the First Section of the Tokyo Stock Exchange from https://www.jpx.co.jp/. 
    source_path = 'https://www.jpx.co.jp/markets/statistics-equities/misc/tvdivq0000001vg2-att/data_j.xls'
    data_frame = pd.read_excel(source_path,usecols=['コード','銘柄名','市場・商品区分','17業種区分'],sheet_name='Sheet1')
    stock_data_frame = data_frame[data_frame['市場・商品区分']=='市場第一部（内国株）']

    output_list = []
    for index,row in stock_data_frame.iterrows():
        stock_code = row[0]
        stock_name = str(row[1])
        stock_industry_type = str(row[3])      
        url = "https://minkabu.jp/stock/" + str(stock_code) + "/chart"
        try:
            html = requests.get(url)
            html.raise_for_status()

            dom_tree = lxml.html.fromstring(html.content)
            print(stock_code)
            # 前日終値
            previous_closing_place = dom_tree.xpath('//*[@id="contents"]/div[3]/div[1]/div/div/div[2]/div/div[1]//tr[1]/td[1]')
            stock_previous_closing_place = util_replace(previous_closing_place,'円','')
            #print('stock_previous_closing_place : ' + str(stock_previous_closing_place))
                
            # 高値
            high_place = dom_tree.xpath('//*[@id="contents"]/div[3]/div[1]/div/div/div[2]/div/div[1]//tr[3]/td[1]')
            stock_high_place = util_replace(high_place,'円','')
            #print('stock_high_place : ' + str(stock_high_place))

            # 安値
            low_place = dom_tree.xpath('//*[@id="contents"]/div[3]/div[1]/div/div/div[2]/div/div[1]//tr[4]/td[1]')
            stock_low_place = util_replace(low_place,'円','')
            #print('stock_low_place : ' + str(stock_low_place))

            # PER
            per = dom_tree.xpath('//*[@id="contents"]/div[3]/div[1]/div/div/div[2]/div/div[2]//tr[1]/td[1]')
            stock_per = util_replace(per, '倍','')
            #print('stock_per : ' + str(stock_per))

            # PBR
            pbr = dom_tree.xpath('//*[@id="contents"]/div[3]/div[1]/div/div/div[2]/div/div[2]//tr[2]/td[1]')
            stock_pbr = util_replace(pbr, '倍','')
            #print('stock_pbr : ' + str(stock_pbr))

            # 配当利回り
            dividend_yield = dom_tree.xpath('//*[@id="contents"]/div[3]/div[1]/div/div/div[2]/div/div[2]//tr[3]/td[1]')
            stock_dividend_yield = util_replace(dividend_yield,'%','')
            #print('stock_dividend_yield : ' + str(stock_dividend_yield))

            # 配当性向
            dividend_payout_ratio = dom_tree.xpath('//*[@id="contents"]/div[3]/div[1]/div/div/div[2]/div/div[2]//tr[4]/td[1]')
            stock_dividend_payout_ratio = util_replace(dividend_payout_ratio,'%','')
            #print('stock_dividend_payout_ratio : ' + str(stock_dividend_payout_ratio))

            # 年初来高値
            high_place_per_year = dom_tree.xpath('//*[@id="contents"]/div[3]/div[1]/div/div/div[2]/div/div[2]//tr[6]/td[1]')
            stock_high_place_per_year = util_replace(high_place_per_year,'円','')
            #print('stock_high_place_per_year : ' + str(stock_high_place_per_year))

            # 年初来安値
            low_place_per_year = dom_tree.xpath('//*[@id="contents"]/div[3]/div[1]/div/div/div[2]/div/div[2]//tr[7]/td[1]')
            stock_low_place_per_year = util_replace(low_place_per_year,'円','')
            #print('stock_low_place_per_year : ' + str(stock_low_place_per_year))

            if(util_screening(stock_dividend_yield,stock_dividend_payout_ratio,stock_per,stock_pbr)):
                list = [stock_code,stock_name,stock_previous_closing_place,stock_high_place,\
                        stock_low_place,stock_high_place_per_year,stock_low_place_per_year,\
                        stock_dividend_yield,stock_dividend_payout_ratio,stock_per,stock_pbr,stock_industry_type]
            else:
                continue
            output_list.append(list)
        except requests.exceptions.RequestException as e:
            print(e)

    columns = ['コード','銘柄','前日終値(円)','高値(円)','安値(円)','年初来高値(円)','年初来安値(円)','配当利回り(%)','配当性向(%)','PER(倍)','PBR(倍)','業種']
    output_data_frame = pd.DataFrame(data=output_list,columns=columns)
    sorted_output_data_frame = output_data_frame.sort_values('配当利回り(%)',ascending=False)
    localtime = strftime("%y%m%d_%H%M%S",time.localtime())
    output_file_name = './stock_' + localtime + '.xlsx'
    sorted_output_data_frame.to_excel(output_file_name,index=None,engine='openpyxl')