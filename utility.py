#!/usr/bin/env python
# -*- coding: utf-8 -*-

from openpyxl import utils
import pandas as pd
import os

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
"""
# Purpose
Create the directory if the directory is not existed.
"""
def util_createdir(dir_name):
    if os.path.isdir(dir_name):
        pass
    else:
        os.mkdir(dir_name)

"""
# Purpose
Automatically adjust the column width of Book and save it. 
"""
def createBook(data_frame,file_name,freeze_panes=None):
    writer = pd.ExcelWriter(file_name,engine='openpyxl')
    data_frame.to_excel(writer,index=None)

    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    for idx,column in enumerate(worksheet.columns):
        max_length = 0
        for cell in column:
            new_value_length = len(str(cell.value))
            if(new_value_length > max_length):
                max_length = new_value_length
        worksheet.column_dimensions[utils.get_column_letter(idx + 1)].width = (max_length + 2) * 1.3
    
    if freeze_panes is None:
        pass
    else:
        worksheet.freeze_panes = freeze_panes
    writer.save()