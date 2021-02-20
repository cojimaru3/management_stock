#!/usr/bin/env python
# -*- coding: utf-8 -*-

import requests
import pandas as pd
import lxml.html
import sys
import time
from time import strftime
from utility import util_replace,parse_dom_tree,isfloat,convert_string_float,util_createdir,createBook
from screening import SCREENING,util_screening
from sendSlack import SlackManager

if __name__ == '__main__':
    # Obtain the code and stock information listed on the First Section of the Tokyo Stock Exchange from https://www.jpx.co.jp/. 
    source_path = 'https://www.jpx.co.jp/markets/statistics-equities/misc/tvdivq0000001vg2-att/data_j.xls'
    data_frame = pd.read_excel(source_path,usecols=['コード','銘柄名','市場・商品区分','17業種区分'],sheet_name='Sheet1')
    stock_data_frame = data_frame[data_frame['市場・商品区分']=='市場第一部（内国株）']

    merged_data_path = './data/fy-merged-sheet.csv' 
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
                
            # 高値
            stock_high_place = parse_dom_tree(chart_dom_tree,'//*[@id="contents"]/div[3]/div[1]/div/div/div[2]/div/div[1]//tr[3]/td[1]','円','')

            # 安値
            stock_low_place = parse_dom_tree(chart_dom_tree,'//*[@id="contents"]/div[3]/div[1]/div/div/div[2]/div/div[1]//tr[4]/td[1]','円','')

            # PER
            stock_per = parse_dom_tree(chart_dom_tree,'//*[@id="contents"]/div[3]/div[1]/div/div/div[2]/div/div[2]//tr[1]/td[1]','倍','')

            # PBR
            stock_pbr = parse_dom_tree(chart_dom_tree,'//*[@id="contents"]/div[3]/div[1]/div/div/div[2]/div/div[2]//tr[2]/td[1]','倍','')

            # 年初来高値
            stock_high_place_per_year = parse_dom_tree(chart_dom_tree,'//*[@id="contents"]/div[3]/div[1]/div/div/div[2]/div/div[2]//tr[6]/td[1]','円','')

            # 年初来安値
            stock_low_place_per_year = parse_dom_tree(chart_dom_tree,'//*[@id="contents"]/div[3]/div[1]/div/div/div[2]/div/div[2]//tr[7]/td[1]','円','')

            # 配当利回り
            stock_dividend_yield = parse_dom_tree(chart_dom_tree,'//*[@id="contents"]/div[3]/div[1]/div/div/div[2]/div/div[2]//tr[3]/td[1]','%','')

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
                stock_operating_profit_ratio = round((stock_operating_profit / stock_sales) * 100,2)

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
                    stock_list = [stock_code,\
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
                stock_list = [stock_code,\
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
            output_list.append(stock_list)
        except requests.exceptions.RequestException as e:
            print(e)
    columns = ['コード',\
                '銘柄',\
                '前日終値',\
                '高値',\
                '安値',\
                '年初来高値',\
                '年初来安値',\
                'PER(倍)',\
                'PBR(倍)',\
                '配当利回り(%)',\
                '年度',\
                '売上高',\
                '営業利益',\
                '営業利益率',\
                '自己資本率',\
                '一株配当',\
                '配当性向(%)',\
                '連続増配',\
                '減配なし',\
                '業種']

    delete_columns = ['コード','年初来高値','年初来安値','年度','売上高','営業利益']
    localtime = strftime("%y%m%d_%H%M%S",time.localtime())

    output_data_frame = pd.DataFrame(data=output_list,columns=columns)
    sorted_output_data_frame = output_data_frame.sort_values('配当利回り(%)',ascending=False)
    mobile_data_frame = sorted_output_data_frame.drop(delete_columns,axis='columns')

    # Book creation
    output_file_name = 'stock_' + localtime + '.xlsx'
    output_dir_path = './stock'
    util_createdir(output_dir_path)
    output_file_path = output_dir_path + '/' + output_file_name
    createBook(sorted_output_data_frame,output_file_path,'C2')

    # Book creation for mobile
    mobile_file_name = 'mobile_stock_' + localtime + '.xlsx'
    mobile_dir_path = './mobile_stock'
    util_createdir(mobile_dir_path)
    mobile_file_path = mobile_dir_path + '/' + mobile_file_name
    createBook(mobile_data_frame,mobile_file_path,'B2')

    slack = SlackManager()
    slack.upload_file(mobile_file_path,mobile_file_name)