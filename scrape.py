#!/usr/bin/env python
# coding: utf-8

import sys
import xlsxwriter
import datetime

from stock import Stock
from stock import Webpage

import os
import numpy as np
import warnings
warnings.filterwarnings("ignore")

now = datetime.datetime.now()


def search(term: str, df, index = 1):
    result = df[df[0].str.contains('(?i)' + term)][index].values[0]
    if type(result) == 'str':
        result.replace(",", "")
    return float(result)


def rename_excel(my_stock, excel_name):
    """rename excel sheet with npv and last price for easy viewing"""

    operating_cf = search("Cash From Operating Activities", my_stock.cash_flow)
    shares_outstanding = search("Shares Outstanding", my_stock.overview)/1000000
    last_price = search("Last Price", my_stock.overview)

    cash_flow = []
    for i in range(1, 11):
        operating_cf = operating_cf * (1 + my_stock.growth_rate)
        cash_flow.append(operating_cf)

    values = cash_flow
    rate = my_stock.discount_rate

    npv = (values / (1+rate)**np.arange(1, len(values)+1)).sum(axis=0) / shares_outstanding
    print(f"NPV per Share: {npv}")
    print(f"Last Price: {last_price}")

    os.rename(excel_name, my_stock.stock_cd + "-" + str(round(npv, 2)) + "-" + str(last_price) + ".xlsx") 


def analyse(company_name):
    my_stock = Stock(company_name)

    excel_name = my_stock.stock_cd + ".xlsx"
    sheet_name = "Sheet1"

    # writer = pd.ExcelWriter(excel_name, engine='xlsxwriter') writer.save()
    workbook = xlsxwriter.Workbook(excel_name)
    worksheet = workbook.add_worksheet(sheet_name)

    # format excel
    worksheet.set_column('A:A', 20)
    worksheet.set_column('A:I', 10)

    currency_format = workbook.add_format({'num_format': '$#,##0.00'})
    percentage_format = workbook.add_format({
        'num_format': '0.0%',
        'bg_color': '#dae8ec',
        'border': 1})
    colored_format = workbook.add_format({
        'num_format': '$#,##0.00',
        'bg_color': '#dae8ec',
        'border': 1
    })


    # Stock and write to excel

    # Required data for npv calculation, table 1
    # --------------------------------------------------------
    table1 = {
        "Name of Stock": my_stock.stock_cd.replace("-", " ").title(),
        "Operating Cash Flow": search("Cash From Operating Activities", my_stock.cash_flow),
        "Total Debt": search("Total Long Term Debt", my_stock.balance_sheet),
        "Cash & Equivalent": search("Cash & Equivalent", my_stock.balance_sheet),
        "Growth Rate": 0,
        "No. of Shares Outstanding": search("Shares Outstanding", my_stock.overview)/1000000,
        "Discount Rate": 0
    }

    worksheet.write_column('A1', table1.keys())
    worksheet.write_column('B1', table1.values(), colored_format)

    # rewrite in percentage format
    worksheet.write('B5', my_stock.growth_rate, percentage_format)
    worksheet.write('B7', my_stock.discount_rate, percentage_format)



    # Ten-year cash flow calculations, bottom table
    # --------------------------------------------------------
    start_row = 14

    # headers
    worksheet.write_column(start_row, 0, ["Year", "Cash Flow", "Discount Rate", "Discounted Value"])
    worksheet.write_row(start_row, 1, list(range(now.year, now.year + 10, 1)))

    # calculation formulas
    cash_flow = ["=B2*(1+B5)"]
    cash_flow.extend(["=" + chr(ord('B') + i) + str(start_row+2) + "*(1+$B$5)" for i in range(10)])
    # +1, +2
    cf_row = start_row + 1
    for i in range(10):
        worksheet.write_formula(cf_row, i+1, cash_flow[i], currency_format)
        
    # +2, +3
    dr_row = start_row + 2
    discount_rate = ["=1/(1 + $B$7)^" + str(i) for i in range(1, 11)]
    for i in range(10):
        worksheet.write_formula(dr_row, i+1, discount_rate[i])

    # +3, +4
    dv_row = start_row + 3
    discounted_value = ["=PRODUCT("+chr(ord('B')+i)+str(cf_row+1)+":"+chr(ord('B')+i)+str(dr_row+1)+")" for i in range(10)]
    for i in range(10):
        worksheet.write_formula(dv_row, i+1, discounted_value[i], currency_format)


    # NPV and intrinsic values calculations, table 2
    # --------------------------------------------------------
    worksheet.write_column('D2', ["PV of 10 yr Cash Flows", "Intrinsic Value per Share", 
                                  "- Debt per Share", "+ Cash per share", "net Cash per Share"])
    worksheet.write_column('E2', [f"=SUM(B{dv_row+1}:K{dv_row+1})", "=E2/B6", "=B3/B6", "=B4/B6", "=E3-E4+E5"], colored_format)


    # Stock overview, table 3
    # --------------------------------------------------------
    df = my_stock.overview.reset_index(drop=True)
    index = [0, 5, 6, 7, 8, 9, 11, 15]
    worksheet.write_column('G2', df.iloc[index, 0])
    worksheet.write_column('H2', df.iloc[index, 1])


    # Jot down links from simply wall st and infront analytics
    # --------------------------------------------------------
    worksheet.write_column('A22', my_stocks.urls)

    workbook.close()
    # Shift + Ctrl + F9

    rename_excel(my_stock, excel_name)


def main():
    # print('Number of arguments: {}'.format(len(sys.argv[1:])))
    # print('Argument(s) passed: {}'.format(str(sys.argv[1:])))

    companies = sys.argv[1:]
    list(map(lambda x: analyse(x),companies))
  
if __name__ == "__main__":
    main()