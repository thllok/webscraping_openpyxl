import sys
from win32com.client import Dispatch
import win32com.client as win32
from openpyxl.formatting.rule import DataBarRule
import pandas as pd
from openpyxl import load_workbook,Workbook
from datetime import datetime
from openpyxl.styles import *
from datetime import datetime, timedelta
from openpyxl.utils.dataframe import dataframe_to_rows
from dateutil.relativedelta import relativedelta
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os
import shutil
def findUserName():
    path = os.path.expanduser('~')
    pMax = len(path)
    pMin = path.find('Users')+6
    userName = path[pMin:pMax]
    return userName
def web_scrap(url, x_path,renamed_file):
    download_dir = 'C:/Users/' + findUserName() + '/Downloads/'
    file_name = "SEMI_holdings"
    options = webdriver.ChromeOptions()
    options.add_experimental_option("useAutomationExtension", False)
    options.add_experimental_option('excludeSwitches', ['enable-automation', 'enable-logging'])
    s = Service('C:/Users/' + findUserName() + '/Documents/chromedriver.exe')
    # driver = webdriver.Chrome(service=s,options=options)
    driver = webdriver.Chrome()
    driver.get(url)
    driver.maximize_window()
    time.sleep(5)
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="onetrust-accept-btn-handler"]'))).click()
    # WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="fundHeaderDocLinks"]/li[4]/a'))).click()
    btn = driver.find_element(By.XPATH, x_path)
    driver.execute_script("arguments[0].click();", btn)
    print("Downloading index...")
    downloaded = 0
    timer = 0
    while not downloaded:
        #if(filename[1:] in os.listdir(download_dir)):
        if file_name+".csv" in os.listdir(download_dir):
            downloaded = 1
            break
        time.sleep(1)
        timer += 1
        print(str(timer) + " second(s) have passed")
    file_date= pd.read_csv(r"C:\Users\elvistsui\Downloads\SEMI_holdings.csv", on_bad_lines='skip', header=None).iloc[0, 1]
    file_date=datetime.strptime(file_date,"%d/%b/%Y")
    shutil.move(download_dir + file_name+".csv", renamed_dir +"\\"+file_name+"_"+file_date.strftime("%Y%m%d")+".csv")
    wb_tickertostrnexcel(file_date,renamed_dir +"\\"+file_name+"_"+file_date.strftime("%Y%m%d"))
    driver.close()
    return renamed_dir +"\\"+file_name+"_"+file_date.strftime("%Y%m%d")+".xlsx", file_date
def wb_tickertostrnexcel(file_date,file):
    benchmark= pd.read_csv(file+".csv",on_bad_lines='skip',header =2 ).dropna(axis=0)
    benchmark["Ticker"]=[str(i) for i in benchmark["Ticker"]]
    new_file = Workbook()
    ws = new_file[new_file.sheetnames[0]]
    ws.title = "SEMI_holdings"+file_date.strftime("%m%d")
    ws["A1"].value = "Fund Holdings as of"
    ws["B1"].value = file_date.strftime("%d-%b-%y")
    rows = dataframe_to_rows(benchmark, index=False, header=True)
    for r_idx, row in enumerate(rows, 3):
        for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)
    new_file.save(file+".xlsx")

def formatting():
    global font_bold,font_notbold,border,purpleFill,yellowFill,orangeFill,no_fill,no_border,GrayFill,GreenFill,rule_databar
    no_fill = PatternFill(fill_type=None)
    side = Side(border_style=None)
    no_border = borders.Border(
        left=side,
        right=side,
        top=side,
        bottom=side,
    )
    font_bold=Font(name="Arial",size=11,bold=True,color="000000")
    font_notbold=Font(name="Arial",size=11,bold=False,color="000000")
    thin=Side(border_style="thin", color="000000")
    double=Side(border_style="double", color="000000")
    border=Border(top=thin, left=thin, right=thin, bottom=thin)

    yellow="FFFF99"  ##YELLOW
    yellowFill=PatternFill(start_color=yellow,
                             end_color=yellow,
                             fill_type='solid')

    purple='9999FF'
    purpleFill=PatternFill(start_color=purple,
                             end_color=purple,
                             fill_type='solid')

    orange="FFC125"  ##orange
    OrangeFill=PatternFill(start_color=orange,
                             end_color=orange,
                             fill_type='solid')
    gray="E7E6E6"  ##gray
    GrayFill=PatternFill(start_color=gray,
                             end_color=gray,
                             fill_type='solid')
    green="B8C2AD"  ##green

    GreenFill=PatternFill(start_color=green,
                             end_color=green,
                             fill_type='solid')
    rule_databar = DataBarRule(start_type="percentile", end_type="percentile", color="63C384", showValue="None", minLength=None,maxLength=None)
    rule_databar = DataBarRule(start_type="percentile",start_value=10, end_type="percentile",
        end_value="90", color="FF638EC6", showValue="None", minLength=None,maxLength=None)


def last_friday():
    today = datetime.today()
    offset = (today.weekday() - 4) % 7
    last_friday_date = today - timedelta(days=offset)
    return last_friday_date
def check_if_updated(date):
    lastfriday= last_friday()
    if date!= lastfriday:
        print("The date shown in the file is "+date.strftime("%d/%b/%Y"))
        print("Please run the code again later today")
        time.sleep(10)
        return 0
    else:
        return 1

def portfolio_weighting(ws):
    ws["A2"].value = "as of "+last_friday().strftime("%m%d%Y")
    ws["A2"].font=font_bold
    ws["A2"].alignment=Alignment(vertical="bottom")
    ws["I2"].value = (datetime(last_friday().year,last_friday().month,1)-timedelta(1)).strftime("%m/%d/%Y")
    ws["J2"].value = (datetime(last_friday().year,1,1)-timedelta(1)).strftime("%m/%d/%Y")
    ws["I4"].value = "MTD ("+last_friday().strftime("%b")+") Performance"
    ws["J4"].value = "YTD "+str(last_friday().year)+" Performance"
    #holding_report = pd.read_excel("Holdings - "+last_friday().strftime("%m%d")+"xlsx",header=9).iloc[1:-2,1::]
    holding_report = pd.read_excel(r"C:\Users\elvistsui\PycharmProjects\attribution_summary\Holdings - 0329.XLSX",header=9).iloc[1:-2,1::]
    Name = holding_report[holding_report.columns[0]]
    ticker = [i +" Equity" for i in holding_report["Ticker"]]
    weighting = holding_report["% Wgt"]
    for row in ws['A5:R100']:
        for cell in row:
            cell.value = None
    for i, (x,y,z) in enumerate(zip(Name,ticker,weighting), start=5):
        ws.cell(row=i, column=1).value = x
        ws.cell(row=i, column=2).value = y
        ws.cell(row=i, column=3).value = "=IFERROR(INDEX('Benchmark Weighting'!E:E,MATCH('Portfolio Weighting'!Q"+str(i)+",'Benchmark Weighting'!A:A,0)),"+chr(34)+"Taiwan"+chr(34)+")"
        ws.cell(row=i, column=4).value = "=IFERROR(INDEX('Portfolio Theme & GICS'!B:B,MATCH(B"+str(i)+",'Portfolio Theme & GICS'!A:A,0)),"+chr(34)+chr(34)+")"
        ws.cell(row=i, column=5).value = "=IFERROR(INDEX('Portfolio Theme & GICS'!F:F,MATCH(B"+str(i)+",'Portfolio Theme & GICS'!E:E,0)),"+chr(34)+chr(34)+")"
        ws.cell(row=i, column=6).value = z/100
        ws.cell(row=i, column=7).value = "=IFERROR(INDEX('Benchmark Weighting'!F:F,MATCH('Portfolio Weighting'!Q"+str(i)+",'Benchmark Weighting'!A:A,0)),0)"
        ws.cell(row=i, column=8).value = "=F"+str(i)+"-G"+str(i)
        ws.cell(row=i, column=9).value = "=FDS(R"+str(i)+","+chr(34)+"P_TOTAL_RETURNC("+chr(34)+"&$I$2&"+chr(34)+","+chr(34)+"&$I$3&"+chr(34)+")"+chr(34)+")/100"
        ws.cell(row=i, column=10).value ="=FDS(R"+str(i)+","+chr(34)+"P_TOTAL_RETURNC("+chr(34)+"&$J$2&"+chr(34)+","+chr(34)+"&$K$3&"+chr(34)+")"+chr(34)+")/100"
        ws.cell(row=i, column=11).value ="=FDS(R"+str(i)+","+chr(34)+"P_TOTAL_RETURNC("+chr(34)+"&$K$2&"+chr(34)+","+chr(34)+"&$K$3&"+chr(34)+")"+chr(34)+")/100"
        ws.cell(row=i, column=12).value = "=I"+str(i)+"-$V$6"
        ws.cell(row=i, column=13).value = "=J"+str(i)+"-$W$6"
        ws.cell(row=i, column=14).value = "=K"+str(i)+"-$X$6"
        ws.cell(row=i, column=15).value = "=IF(AND(H"+str(i)+">0,M"+str(i)+"<-5%),"+chr(34)+"Underperformed Bet"+chr(34)+","+chr(34)+chr(34)+")"
        ws.cell(row=i, column=16).value = "=ABS(H"+str(i)+")"
        ws.cell(row=i, column=17).value = "=LEFT(B"+str(i)+",FIND("+chr(34)+" "+chr(34)+",B"+str(i)+")-1)"
        ws.cell(row=i, column=18).value = "=LEFT(B"+str(i)+",FIND("+chr(34)+" "+chr(34)+",B"+str(i)+")-1)&IF(RIGHT(B"+str(i)+",9)="+chr(34)+"TT Equity"+chr(34)+","+chr(34)+"-TW"+chr(34)+",IF(RIGHT(B"+str(i)+",9)="+chr(34)+"CH Equity"+chr(34)+","+chr(34)+"-CN"+chr(34)+",IF(RIGHT(B"+str(i)+",9)="+chr(34)+"GR Equity"+chr(34)+","+chr(34)+"-DE"+chr(34)+",IF(RIGHT(B"+str(i)+",9)="+chr(34)+"NA Equity"+chr(34)+","+chr(34)+"-NL"+chr(34)+","+chr(34)+"-"+chr(34)+"&MID(B"+str(i)+",FIND("+chr(34)+" "+chr(34)+",B"+str(i)+")+1,2)))))"
    ws["O"+str(i+2)].value = "Other ACWI IMI Weighting"
    ws["O" + str(i + 3)].value = "Active Share"

    ws["P"+str(i+2)].value = "=1-SUM(G5:G"+str(i)+")"
    ws["P" + str(i + 3)].value = "=SUM(P5:P"+str(i)+")/2"
    ws["O"+str(i+2)].value = "Other ACWI IMI Weighting"
    ws["O" + str(i + 3)].value = "Active Share"

    ws["P"+str(i+2)].font=font_notbold
    ws["P" + str(i + 3)].font=font_notbold

    return ws
def benchmark_weighting(ws,file_name):

    for row in ws['A5:G1000']:
        for cell in row:
            cell.value = None
    for row in ws['N5:O1000']:
        for cell in row:
            cell.value = None
    ws["A1"].value = "iShares MSCI Global Semiconductors ETF (proxy of MSCI ACWI IMI Semi Index)"
    ws["A2"].value = last_friday().strftime("%m%d%Y")
    for i in range(2):
        ws["A"+str(i+1)].font=font_bold
        ws["A" + str(i + 1)].alignment = Alignment(vertical="center",horizontal="left")


    header = ['Ticker',
     'Name',
     'GICS Sector',
     'GICS Sub Industry',
     'Location',
     'Index Weight (%)',
     'Weighting in Portfolio',
     'Overweight/UnderWeight',
     'MTD Performance',
     'YTD Performance',
     'MTD Excess Return',
     'YTD Excess Return',
     'Missed Out Names\n(Unweighted & YTD outperformed >5%)']
    benchmark = pd.read_excel(file_name,header =2 ).dropna(axis=0)
    print(benchmark)
    print(benchmark.columns)
    x = benchmark[["Ticker","Name","Sector","Location","Weight (%)"]]
    rows = dataframe_to_rows(x, index=False, header=False)

    for r_idx, row in enumerate(rows, 5):
        for c_idx, value in enumerate(row, 1):
            if c_idx==4:
                ws.cell(row=r_idx, column=c_idx+1, value=value)
            elif c_idx==5:
                ws.cell(row=r_idx, column=c_idx + 1, value=value/100)
            else:
                ws.cell(row=r_idx, column=c_idx, value=value)

    for i in range(len(x)):
        if  ws["E"+str(5+i)].value == "United States":
            ticker = "-US"
        elif ws["E"+str(5+i)].value == "Taiwan":
            ticker = "-TW"
        elif ws["E"+str(5+i)].value == "Japan":
            ticker = "-JP"
        elif ws["E"+str(5+i)].value == "Germany":
            ticker = "-DE"
        elif ws["E"+str(5+i)].value == "Netherlands":
            ticker = "-NL"
        elif ws["E"+str(5+i)].value == "China":
            ticker = "-CH"
        elif ws["E"+str(5+i)].value == "France":
            ticker = "-FR"
        elif ws["E"+str(5+i)].value == "Korea (South)":
            ticker = "-KR"
        elif ws["E"+str(5+i)].value == "Hong Kong":
            ticker = "-HK"
        else:
            pass

        ticker = ws["A"+str(5+i)].value+ticker
        ws["D"+str(5+i)].value = "=INDEX('Benchmark Sub-industry'!C:C,MATCH('Benchmark Weighting'!N"+str(5+i)+",'Benchmark Sub-industry'!A:A,0))"
        ws["G"+str(5+i)].value = "=IFERROR(INDEX('Portfolio Weighting'!F:F,MATCH('Benchmark Weighting'!A"+str(5+i)+",'Portfolio Weighting'!Q:Q,0)),0)"
        ws["N"+str(5+i)].value = "=IF(E"+str(5+i)+"=$Q$5,A"+str(5+i)+"&"+chr(34)+" US"+chr(34)+",IF(E"+str(5+i)+"=$Q$6,A"+str(5+i)+"&"+chr(34)+" TT"+chr(34)+",IF(E"+str(5+i)+"=$Q$7,A"+str(5+i)+"&"+chr(34)+" NA"+chr(34)+",IF(E"+str(5+i)+"=$Q$8,A"+str(5+i)+"&"+chr(34)+" JP"+chr(34)+",IF(E"+str(5+i)+"=$Q$9,A"+str(5+i)+"&"+chr(34)+" GR"+chr(34)+",IF(E"+str(5+i)+"=$Q$10,A"+str(5+i)+"&"+chr(34)+" KS"+chr(34)+",IF(E"+str(5+i)+"=$Q$11,A"+str(5+i)+"&"+chr(34)+" CH"+chr(34)+",A"+str(5+i)+")))))))"
        if i<30:
            ws["H" + str(5 + i)].value="=G"+str(5+i)+"-F"+str(5+i)
            ws["I" + str(5 + i)].value ="=FDS("+chr(34)+ticker+chr(34)+","+chr(34)+"P_TOTAL_RETURNC('"+(last_friday().replace(day=1)-timedelta(1)).strftime("%m/%d/%Y")+"','"+last_friday().strftime("%m/%d/%Y")+"')"+chr(34)+") / 100"
            ws["J" + str(5 + i)].value ="=FDS("+chr(34)+ticker+chr(34)+","+chr(34)+"P_TOTAL_RETURNC('"+(last_friday().replace(day=1,month=1)-timedelta(1)).strftime("%m/%d/%Y")+"','"+last_friday().strftime("%m/%d/%Y")+"')"+chr(34)+") / 100"
            ws["K" + str(5 + i)].value ="=I"+str(5+i)+"-'Portfolio Weighting'!V$6"
            ws["L" + str(5 + i)].value ="=J"+str(5+i)+"-'Portfolio Weighting'!W$6"
            ws["M" + str(5 + i)].value ="=IF(AND(H"+str(5+i)+"<0,L"+str(5+i)+">5%),"+chr(34)+"Missed"+chr(34)+","+chr(34)+chr(34)+")"
            ws["O" + str(5 + i)].value = str(i + 1)
        for j in range(5):
            ws.cell(row=5+i, column=j+1).font=font_notbold
            ws.cell(row=5 + i, column=j + 1).alignment = Alignment(vertical="center",horizontal="left")
        for j in range(5):
            ws.cell(row=5+i,column=j+6).alignment = Alignment(vertical="center",horizontal="center")
            ws.cell(row=5 + i, column=j + 6).font=font_notbold
            ws.cell(row=5 + i, column=j + 6).number_format="0.0%"

    return ws
def run_macro(macro_wb, macro_name, result_wb):
    xl = Dispatch("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    wb_macro = xl.Workbooks.Open(macro_wb)
    wb = xl.Workbooks.Open(result_wb)
    xl.Application.Run(macro_name)
    wb.Close(SaveChanges=True)
    wb_macro.Close()
    xl.Application.Quit
    del xl
if __name__ == "__main__":
    formatting()
    xls="/html/body/div[1]/div[2]/div/div/div/div/div/div[1]/div/div[2]/header[2]/div[1]/div[2]/ul/li[4]/a"
    csv="/html/body/div[1]/div[2]/div/div/div/div/div/div[17]/div/div/div/div[2]/a"
    x_path = '/html/body/div[1]/div[2]/div/div/div/div/div/div[13]/div/div/div/div[2]/a[1]'
    x_url = 'https://www.ishares.com/uk/professional/en/products/319084/ishares-msci-global-semiconductors-ucits-etf?switchLocale=y&siteEntryPassthrough=true'
    renamed_dir = os.getcwd()
    file_name, file_date= web_scrap(x_url, csv, renamed_dir)
    #correct_date=check_if_updated(file_date)
    correct_date=1
    if correct_date:
        #wb=load_workbook("F:\Elvis Tsui\10. Global Semi\SVLO Semi Paper Portfolio_"+last_friday().strftime("%Y%m%d")+".xlsx")
        wb=load_workbook(r"C:\Users\elvistsui\PycharmProjects\attribution_summary\SVLO Semi Paper Portfolio_20240331.xlsx")
        ws=wb["Portfolio Weighting"]
        ws=portfolio_weighting(ws)
        ws=wb["Benchmark Weighting"]
        ws=benchmark_weighting(ws,file_name)
        wb.save("testing.xlsx")
   # run_macro("formatting_macro.xlsm","databar_format", "testing.xlsx")
    print("testing")