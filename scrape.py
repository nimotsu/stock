#!/usr/bin/env python
# coding: utf-8

import sys
import requests
import pandas as pd
import re
from bs4 import BeautifulSoup
import xlsxwriter
import datetime

import os
import numpy as np
import warnings
warnings.filterwarnings("ignore")

now = datetime.datetime.now()


def url2html(url, headers=None, params=None, data=None):
    headers={'User-Agent': 'Mozilla/5.0'}
    try:
        req = requests.get(url, headers=headers, params=params, data=data)
    except:
        req = requests.get(url, headers=headers, params=params, data=data, verify=False)
    html = req.text
    return html
    
def search(term: str, df, index = 1):
    return float(df[df[0].str.contains('(?i)' + term)][index].values[0].replace(",", ""))


# Handle all urls and htmls
class Webpage:
    def __init__(self, html):
        self.html = html
        self.soup = BeautifulSoup(self.html, 'html.parser')
        try:
            self.tables = pd.read_html(self.html)
        except:
            self.tables = None
    
    @classmethod
    def from_url(cls, url, headers=None, params=None, data=None):
        """constructor with url"""

        html = url2html(url, headers)
        return cls(html)
    
    def get_span(self, tag: str, class_name: list):
        """return df from columns not in <table>"""
        
        def get_tag(tag, class_name):
            tags = self.soup.find_all(tag, {'class': class_name})
            text = [i.get_text() for i in tags if i.get_text() != '']
            return text
        
        attrib = get_tag(tag, class_name[0])
        data = get_tag(tag, class_name[1])

        ls = list(zip(attrib, data))
        df = pd.DataFrame(ls)
        return df
        

# Handle all methods related to stock
class Stock:
    def __init__(self, stock_cd):
        self.stock_cd = stock_cd
        print(f"Stock Cd: {self.stock_cd}")
        self.overviewp = Webpage.from_url(f"https://www.investing.com/equities/{stock_cd}")
        self.stock_id = self.get_id()
        print(f"Stock Id: {self.stock_id}")
        self.ratiosp = Webpage.from_url(f"https://www.investing.com/equities/{stock_cd}-ratios")
        self.cash_flowp = Webpage.from_url(f"https://www.investing.com/instruments/Financials/changereporttypeajax?action=change_report_type&pair_ID={self.stock_id}&report_type=CAS&period_type=Annual")
        self.balance_sheetp = Webpage.from_url(f"https://www.investing.com/instruments/Financials/changereporttypeajax?action=change_report_type&pair_ID={self.stock_id}&report_type=BAL&period_type=Annual")
        # self.income_statementp = Webpage.from_url(f"https://www.investing.com/equities/{stock_cd}-income-statement")
        # self.earningsp = Webpage.from_url(f"https://www.investing.com/equities/{stock_cd}-earnings")
        # self.financialp = Webpage.from_url(f"https://www.investing.com/equities/{stock_cd}-financial-summary")
    
    def get_growth_rate(self):
        """scrape growth rate from simply wall st"""

        stock_cd = self.stock_cd.replace("-", " ")
        params = (
            ('x-algolia-agent', 'Algolia for JavaScript (4.2.0); Browser (lite)'),
            ('x-algolia-api-key', 'be7c37718f927d0137a88a11b69ae419'),
            ('x-algolia-application-id', '17IQHZWXZW'),
        )
        data = f'{{"query":"{stock_cd} klse","highlightPostTag":" ","highlightPreTag":" ","restrictHighlightAndSnippetArrays":true}}'
        try:
            response = requests.post('https://17iqhzwxzw-dsn.algolia.net/1/indexes/companies/query', params=params, data=data)
            stock_url = response.json()['hits'][0]['url']
            url = "https://simplywall.st" + stock_url
        except:
            return None
        
        html = url2html(url)
        soup = BeautifulSoup(html, 'html.parser')
        growth = soup.find('p', {'data-cy-id': 'key-metric-value-forecasted-annual-earnings-growth'}).get_text().replace('%', '')
        self.growth_rate = float(growth)/100
        print(f"Growth Rate: {self.growth_rate}")
        return [self.growth_rate, url]
    
    def get_beta(self):
        """scrape beta from infrontanalytics.com"""

        url = f"https://www.infrontanalytics.com/fe-EN/33123FM/{self.stock_cd}-/Beta"
        html = url2html(url)
        m = re.search(r"shows a Beta of ([+-]?\d+\.\d+).", "shows a Beta of 1.56.")
        beta = m.groups()[0]
        return float(beta), url
    
    def get_discount_rate(self):
        """convert beta to discount rate for dcf model"""

        beta, url = self.get_beta()
                
        dr = {0.8: 5, 
        1: 6, 
        1.1: 6.8, 
        1.2: 7, 
        1.3: 7.9, 
        1.4: 8, 
        1.5: 8.9}
        for key in dr:
            if beta < key:
                discount_rate = dr[key]
            else:
                discount_rate = 9
        self.discount_rate = discount_rate/100
        print(f"Discount Rate: {self.discount_rate}")
        return [self.discount_rate, url]
        
    def get_id(self):
        m = re.search('data-pair-id="(\d+)"', self.overviewp.html)
        stock_id = m.groups()[0]
        return stock_id
    
    def overview(self):
        soup = self.overviewp.soup
        last_price = soup.find('span', {'id':'last_last'}).get_text()
        ls = ['Last Price', last_price]
        df = pd.DataFrame([ls])
        
        overview = self.overviewp.get_span('span', ['float_lang_base_1', 'float_lang_base_2'])
        return pd.concat([df, overview])
    
    def ratios(self):
        tables = self.ratiosp.tables
        numbers = range(1, 9)
        ratios = pd.concat(tables[i] for i in numbers)
        return ratios
    
    def cash_flow(self):
        df = self.cash_flowp.tables[0]
        cash_flow = df[~df[1].str.contains("a|e|i|o|u")]
        return cash_flow
    
    def balance_sheet(self):
        df = self.balance_sheetp.tables[0]
        balance_sheet = df[~df[1].str.contains("a|e|i|o|u")]
        return balance_sheet
    
    def earnings(self):
        s = requests.Session()
        url = f"https://www.investing.com/equities/{self.stock_cd}-earnings"
        headers={ "User-Agent": "Mozilla/5.0"}
        r = s.get(url, headers={ "User-Agent": "Mozilla/5.0"})
        
        # get more history - to work on
        '''
        more_history = "https://www.investing.com/equities/morehistory"
        headers = {
            'User-Agent': 'Mozilla/5.0',
            'X-Requested-With': 'XMLHttpRequest',
            'Referer': url,
        }
        data = {"pairID" : "41688", "last_timestamp": "2019-0-02"}
        r = s.post(more_history, headers=headers, cookies=r.cookies, data=data)
        r.json()['historyRows']
        '''
        return r.text
    
    def financial_summary(self):
        def get_summary(html):
            webpage = Webpage(html)
            soup = webpage.soup

            title = soup.find('h3').text
            df = webpage.get_span('span', ['float_lang_base_1', 'float_lang_base_2'])
            table = pd.read_html(str(soup))[0]
            return [title, table, df] # pd.concat([table, df], axis=0, ignore_index=True)
        
        financial_summary = f"https://www.investing.com/instruments/Financials/changesummaryreporttypeajax?action=change_report_type&pid={stock_id}&financial_id={stock_id}&ratios_id={stock_id}&period_type="
        annual = financial_summary + "Annual"
        # interim = financial_summary + "Interim"
        
        df = pd.DataFrame()
        soup = Webpage.from_url(annual).soup
        sections = soup.find_all('div', "companySummaryIncomeStatement")
        result = []
        for i in sections:
            result.append(get_summary(str(i)))
        return result
            
'''
10% for public companies
15% for private companies that are scaling predictably (say above $10m in ARR, and growing greater than 40% year on year)
20% for private companies that have not yet reached scale and predictable growth
'''

def rename_excel(my_stock, excel_name):
    """rename excel sheet with npv and last price for easy viewing"""

    operating_cf = search("Cash From Operating Activities", my_stock.cash_flow())
    shares_outstanding = search("Shares Outstanding", my_stock.overview())/1000000
    last_price = search("Last Price", my_stock.overview())

    cash_flow = []
    for i in range(1, 11):
        operating_cf = operating_cf * (1 + my_stock.growth_rate)
        cash_flow.append(operating_cf)

    values = cash_flow
    rate = my_stock.discount_rate

    npv = (values / (1+rate)**np.arange(1, len(values)+1)).sum(axis=0) / shares_outstanding
    print(f"NPV per share: {npv}")

    os.rename(excel_name, my_stock.stock_cd + "-" + str(round(npv, 2)) + "-" + str(last_price) + ".xlsx") 


def analyse(company_name):
    my_stock = Stock(company_name)

    excel_name = my_stock.stock_cd + ".xlsx"
    sheet_name = "Sheet1"

    # writer = pd.ExcelWriter(excel_name, engine='xlsxwriter') writer.save()
    workbook = xlsxwriter.Workbook(excel_name)
    worksheet = workbook.add_worksheet(sheet_name)

    # format excel
    worksheet.set_column('A:K', 10)
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
    table1 = {
        "Name of Stock": my_stock.stock_cd.replace("-", " ").title(),
        "Operating Cash Flow": search("Cash From Operating Activities", my_stock.cash_flow()),
        "Total Debt": search("Total Long Term Debt", my_stock.balance_sheet()),
        "Cash & Equivalent": search("Cash & Equivalent", my_stock.balance_sheet()),
        "Growth Rate": 0,
        "No. of Shares Outstanding": search("Shares Outstanding", my_stock.overview())/1000000,
        "Discount Rate": 0
    }


    # Required data, table 1
    worksheet.write_column('A1', table1.keys())
    worksheet.write_column('B1', table1.values(), colored_format)

    # rewrite in percentage
    worksheet.write('B5', my_stock.get_growth_rate()[0], percentage_format)
    worksheet.write('B7', my_stock.get_discount_rate()[0], percentage_format)



    # Ten-year cash flow calculations, bottom table
    start_row = 11

    # headers
    worksheet.write_column(start_row, 0, ["Year", "Cash Flow", "Discount Rate", "Discounted Value"])
    worksheet.write_row(start_row, 1, list(range(now.year, now.year + 10, 1)))

    # calculations
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


    # Intrinsic values, table 2
    worksheet.write_column('D2', ["PV of 10 yr Cash Flows", "Intrinsic Value per Share", 
                                  "- Debt per Share", "+ Cash per share", "net Cash per Share"])
    worksheet.write_column('E2', [f"=SUM(B{dv_row+1}:K{dv_row+1})", "=E2/B6", "=B3/B6", "=B4/B6", "=E3-E4+E5"], colored_format)


    # Stock overview, table 3
    df = my_stock.overview().reset_index(drop=True)
    index = [0, 5, 6, 7, 8, 9, 11, 15]
    worksheet.write_column('G2', df.iloc[index, 0])
    worksheet.write_column('H2', df.iloc[index, 1])

    # report
    worksheet.write_column('A18', [my_stock.get_growth_rate()[1], my_stock.get_beta()[1]])

    workbook.close()
    # Shift + Ctrl + F9

    rename_excel(my_stock, excel_name)


def main():

    # print('Number of arguments: {}'.format(len(sys.argv[1:])))
    # print('Argument(s) passed: {}'.format(str(sys.argv[1:])))

    companies = sys.argv[1:]

    list(map(lambda x: analyse(x),companies))








  
main()