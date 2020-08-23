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
    result = str(result)
    result = result.replace(",", "")
    if '%' in result:
        result = result.replace("%", "")
        result = float(result) / 100
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

    excel_name = "stocks/" + my_stock.stock_cd + ".xlsx"
    sheet_name = "Sheet1"

    # colours
    blue = '#98C4D1'
    yellow = '#FEC240'
    red = '#DE4B43'

    # writer = pd.ExcelWriter(excel_name, engine='xlsxwriter') writer.save()
    workbook = xlsxwriter.Workbook(excel_name)
    worksheet = workbook.add_worksheet(sheet_name)

    # format excel
    worksheet.set_row(0, 40)
    worksheet.set_column('A:A', 20)
    worksheet.set_column('A:I', 10)

    title_format = workbook.add_format({
        'bold': True,
        'font_color': blue,
        'font_size': 16
        })
    currency_format = workbook.add_format({
        'num_format': '$#,##0.00',
        'border': 1
        })
    percentage_format = workbook.add_format({
        'num_format': '0.0%',
        'bg_color': blue,
        'border': 1
        })
    colored_format = workbook.add_format({
        'bg_color': blue,
        'border': 1
        })
    colored_currency_format = workbook.add_format({
        'num_format': '$#,##0.00',
        'bg_color': blue,
        'border': 1
        })
    border_format = workbook.add_format({
        'border': 1
        })


    # Stock and write to excel

    # Required data for npv calculation, table 1
    # --------------------------------------------------------
    table01 = (0, 0)
    table1 = {
        "Name of Stock": my_stock.stock_cd.replace("-", " ").title(),
        "Operating Cash Flow": search("Cash From Operating Activities", my_stock.cash_flow),
        "Total Debt": search("Total Long Term Debt", my_stock.balance_sheet),
        "Cash & Equivalent": search("Cash & Equivalent", my_stock.balance_sheet),
        "Growth Rate": 0,
        "No. of Shares Outstanding": search("Shares Outstanding", my_stock.overview) / 1000000,
        "Discount Rate": 0
    }

    worksheet.write_column('A1', table1.keys(), border_format)
    worksheet.write_column('B1', table1.values(), colored_currency_format)

    # rewrite in title and percentage format
    worksheet.write('B1', my_stock.stock_cd.replace("-", " ").title(), title_format)
    worksheet.write('B5', my_stock.growth_rate, percentage_format)
    worksheet.write('B7', my_stock.discount_rate, percentage_format)


    # Ten-year cash flow calculations, bottom table
    # --------------------------------------------------------
    table11 = (11, 0)
    calc_row = table11[0]

    # headers
    worksheet.write_column(calc_row, 0, ["Year", "Cash Flow", "Discount Rate", "Discounted Value"], border_format)
    worksheet.write_row(calc_row, 1, list(range(now.year, now.year + 10, 1)), border_format)

    # calculation formulas
    cash_flow = ["=B2*(1+B5)"]
    cash_flow.extend(["=" + chr(ord('B') + i) + str(calc_row+2) + "*(1+$B$5)" for i in range(10)])
    # +1, +2
    cf_row = calc_row + 1
    for i in range(10):
        worksheet.write_formula(cf_row, i+1, cash_flow[i], currency_format)
        
    # +2, +3
    dr_row = calc_row + 2
    discount_rate = ["=1/(1 + $B$7)^" + str(i) for i in range(1, 11)]
    for i in range(10):
        worksheet.write_formula(dr_row, i+1, discount_rate[i], border_format)

    # +3, +4
    dv_row = calc_row + 3
    discounted_value = ["=PRODUCT("+chr(ord('B')+i)+str(cf_row+1)+":"+chr(ord('B')+i)+str(dr_row+1)+")" for i in range(10)]
    for i in range(10):
        worksheet.write_formula(dv_row, i+1, discounted_value[i], currency_format)


    # NPV and intrinsic values calculations, table 2
    # --------------------------------------------------------
    # table02 = ()
    worksheet.write_column('D2', ["PV of 10 yr Cash Flows", "Intrinsic Value per Share", 
                                  "- Debt per Share", "+ Cash per share", "net Cash per Share"], border_format)
    worksheet.write_column('E2', [f"=SUM(B{dv_row+1}:K{dv_row+1})", "=E2/B6", "=B3/B6", "=B4/B6", "=E3-E4+E5"], colored_currency_format)


    # Stock overview, table 3
    # --------------------------------------------------------
    # table03 = ()
    df = my_stock.overview.reset_index(drop=True)
    index = [0, 5, 6, 7, 8, 9, 11, 15]
    worksheet.write_column('G2', df.iloc[index, 0], border_format)
    worksheet.write_column('H2', df.iloc[index, 1], colored_format)


    # Jot down links from simply wall st and infront analytics
    # --------------------------------------------------------
    row = table11[0] + 5
    worksheet.write_column(row, 0, my_stock.urls)


    # Overview by i3investor
    # --------------------------------------------------------
    i3summary = my_stock.i3summary
    i3business_performance = my_stock.i3business_performance

    # i3investor table, table 4
    # table04 = ()
    i3summary_column = 9
    worksheet.set_column(i3summary_column, i3summary_column+1, 20)  # Width of column B set to 30.

    worksheet.write_column(1, i3summary_column, i3summary[0], border_format)
    worksheet.write_column(1, i3summary_column+1, i3summary[1], colored_currency_format)

    # summary tables
    start_row = 23
    start_column = 0

    for key in i3business_performance:
        worksheet.write(start_row, start_column, key)
        
        cur_df = i3business_performance[key]
        for col in cur_df.columns:
            cur_col = []
            cur_col.append(col.replace("Unnamed: 0", ""))
            cur_col.extend(cur_df[col])
            worksheet.write_column(start_row+1, start_column, cur_col, border_format)
            start_column += 1
        start_column += 1


    # Ratios by investing, table 5
    # --------------------------------------------------------
    # table05
    start_row = 1
    start_column = 12
    ratios_header = my_stock.ratios.head(6)
    ratios_header = ratios_header.rename({0: '', 1: 'Company', 2: 'Industry'}, axis=1)
    for col in ratios_header.columns:
        cur_col = []
        cur_col.append(col)
        cur_col.extend(ratios_header[col])
        worksheet.write_column(start_row, start_column, cur_col, border_format)
        start_column += 1

    total_assets = search("Total Assets", my_stock.balance_sheet)
    total_liabilities = search("Total Liabilities", my_stock.balance_sheet)
    current_shares_outstanding = search("Total Common Shares Outstanding", my_stock.balance_sheet)
    total_equity = search("Total Equity", my_stock.balance_sheet)

    net_assets = total_assets - total_liabilities
    net_asset_value = net_assets / current_shares_outstanding
    net_asset_value = round(net_asset_value, 2)

    table5 = {
        "EPS": search("Basic EPS ANN", my_stock.ratios),
        "EPS(MRQ) vs Qtr. 1 Yr. Ago MRQ": search("EPS\(MRQ\) vs Qtr. 1 Yr. Ago MRQ", my_stock.ratios),
        "EPS(TTM) vs TTM 1 Yr. Ago TTM": search("EPS\(TTM\) vs TTM 1 Yr. Ago TTM", my_stock.ratios),
        "5 Year EPS Growth 5YA": search("5 Year EPS Growth 5YA", my_stock.ratios),
            
        "Return on Equity TTM": search("Return on Equity TTM", my_stock.ratios),
        "Return on Equity 5YA": search("Return on Equity 5YA", my_stock.ratios),
            
        "Price to Earnings Ratio": search("P/E Ratio TTM", my_stock.ratios),
            
        "Dividend per Share": search("Dividend Yield ANN", my_stock.ratios),
        "Dividend Yield 5 Year Avg. 5YA": search("Dividend Yield 5 Year Avg. 5YA", my_stock.ratios),
        "Dividend Growth Rate ANN": search("Dividend Growth Rate ANN", my_stock.ratios),
            
        "Net Asset per Share": net_asset_value,
        "Price to Book": search("Price to Book MRQ", my_stock.ratios),
        "LT Debt to Equity": search("LT Debt to Equity", my_stock.ratios)
    }

    # Continuation, table 5
    start_row = len(ratios_header) + 2
    start_column = 12
    worksheet.write_column(start_row, start_column, table5.keys(), border_format)
    worksheet.write_column(start_row, start_column+1, table5.values(), colored_format)

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