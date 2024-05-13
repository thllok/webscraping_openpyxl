from selenium import webdriver
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.by import By
import pandas as pd
import time
import numpy as np
from datetime import datetime, timedelta
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import Alignment
from math import floor
from copy import copy
from openpyxl.styles import *
import os
from selenium.webdriver.chrome.service import Service

import sys
def findUserName():
    path = os.path.expanduser('~')
    pMax = len(path)
    pMin = path.find('Users') + 6
    userName = path[pMin:pMax]
    return userName


def getdata(i):
    service = Service()
    options = webdriver.ChromeOptions()
    driver = webdriver.Chrome(service=service, options=options)
    driver.get("https://sanctionssearch.ofac.treas.gov/")
    select1 = Select(driver.find_element(By.ID, "ctl00_MainContent_ddlCountry"))
    select1.select_by_visible_text(i)
    select2 = Select(driver.find_element(By.ID, "ctl00_MainContent_ddlType"))
    select2.select_by_visible_text("Entity")
    element = driver.find_element(By.NAME, "ctl00$MainContent$btnSearch")
    element.click()
    saving_file = driver.find_element(By.NAME, "ctl00$MainContent$ImageButton1")
    saving_file.click()
    data = driver.find_element(By.ID, "gvSearchResults").text
    time.sleep(5)

def jaro_distance(s1, s2):
    if (s1 == s2):
        return 1.0
    len1 = len(s1)
    len2 = len(s2)
    max_dist = floor(max(len1, len2) / 2) - 1
    match = 0
    hash_s1 = [0] * len(s1)
    hash_s2 = [0] * len(s2)
    for i in range(len1):
        for j in range(max(0, i - max_dist), min(len2, i + max_dist + 1)):
            if (s1[i] == s2[j] and hash_s2[j] == 0):
                hash_s1[i] = 1
                hash_s2[j] = 1
                match += 1
                break
    if (match == 0):
        return 0.0
    t = 0
    point = 0
    for i in range(len1):
        if (hash_s1[i]):
            while (hash_s2[point] == 0):
                point += 1
            if (s1[i] != s2[point]):
                t += 1
            point += 1
    t = t // 2
    return (match / len1 + match / len2 + (match - t) / match) / 3.0


def copy_sheet(source_sheet, target_sheet):
    copy_cells(source_sheet, target_sheet)  # copy all the cel values and styles
    copy_sheet_attributes(source_sheet, target_sheet)

def copy_sheet_attributes(source_sheet, target_sheet):
    target_sheet.sheet_format = copy(source_sheet.sheet_format)
    target_sheet.sheet_properties = copy(source_sheet.sheet_properties)
    target_sheet.merged_cells = copy(source_sheet.merged_cells)
    target_sheet.page_margins = copy(source_sheet.page_margins)
    target_sheet.freeze_panes = copy(source_sheet.freeze_panes)

    # set row dimensions
    # So you cannot copy the row_dimensions attribute. Does not work (because of meta data in the attribute I think). So we copy every row's row_dimensions. That seems to work.
    for rn in range(len(source_sheet.row_dimensions)):
        target_sheet.row_dimensions[rn] = copy(source_sheet.row_dimensions[rn])

    if source_sheet.sheet_format.defaultColWidth is None:
        print('Unable to copy default column wide')
    else:
        target_sheet.sheet_format.defaultColWidth = copy(source_sheet.sheet_format.defaultColWidth)

    # set specific column width and hidden property
    # we cannot copy the entire column_dimensions attribute so we copy selected attributes
    for key, value in source_sheet.column_dimensions.items():
        target_sheet.column_dimensions[key].min = copy(source_sheet.column_dimensions[
                                                           key].min)  # Excel actually groups multiple columns under 1 key. Use the min max attribute to also group the columns in the targetSheet
        target_sheet.column_dimensions[key].max = copy(source_sheet.column_dimensions[
                                                           key].max)  # https://stackoverflow.com/questions/36417278/openpyxl-can-not-read-consecutive-hidden-columns discussed the issue. Note that this is also the case for the width, not onl;y the hidden property
        target_sheet.column_dimensions[key].width = copy(
            source_sheet.column_dimensions[key].width)  # set width for every column
        target_sheet.column_dimensions[key].hidden = copy(source_sheet.column_dimensions[key].hidden)


def copy_cells(source_sheet, target_sheet):
    for (row, col), source_cell in source_sheet._cells.items():
        target_cell = target_sheet.cell(column=col, row=row)

        target_cell._value = source_cell._value
        target_cell.data_type = source_cell.data_type

        if source_cell.has_style:
            target_cell.font = copy(source_cell.font)
            target_cell.border = copy(source_cell.border)
            target_cell.fill = copy(source_cell.fill)
            target_cell.number_format = copy(source_cell.number_format)
            target_cell.protection = copy(source_cell.protection)
            target_cell.alignment = copy(source_cell.alignment)

        if source_cell.hyperlink:
            target_cell._hyperlink = copy(source_cell.hyperlink)

        if source_cell.comment:
            target_cell.comment = copy(source_cell.comment)
def cleaning_file():
    for i in range(size):
        if i == 0:
            os.remove(download_address + "Search_Results.xls")
        else:
            os.remove(download_address + "Search_Results " + "(" + str(i) + ")" + ".xls")


username = findUserName()
download_address = 'C:/Users/' + username + '/Downloads/'
country = ["Taiwan", "United States", "Hong Kong", "China"]
columns = ['Name', "Address", "Type", "Program(s)", "List", "Score"]
today = datetime.today()
#come here if you forgot and change the below date
#today=datetime(year,month,day)
first = today.replace(day=1)
last_month = first - timedelta(days=1)
lastMonth = last_month.strftime("%h")
lastMonth_Year = last_month.strftime("%Y")
old_file = "01. Firm_OFAC Sanction List Check_" + str((last_month.replace(day=1)-timedelta(days=1)).strftime("%h")) + " " + str((last_month.replace(day=1)-timedelta(days=1)).strftime("%Y")) + ".xlsx"

for i in country:
    getdata(i)
size = len(country)
dataset = pd.DataFrame(columns=columns)
for i in range(size):
    if i == 0:
        place = download_address + "Search_Results.xls"
        data = list(pd.read_html(place))
        data = pd.DataFrame(np.reshape(np.array(data), (np.array(data).shape[1], np.array(data).shape[2])),
                            columns=columns)
        data["Country"] = country[i]
        dataset = pd.concat([dataset, data])
    else:
        place = download_address + "Search_Results " + "(" + str(i) + ")" + ".xls"
        data = list(pd.read_html(place))
        data = pd.DataFrame(np.reshape(np.array(data), (np.array(data).shape[1], np.array(data).shape[2])),
                            columns=columns)
        data["Country"] = country[i]
        dataset = pd.concat([dataset, data])
