import requests
import pandas as pd
import re
from bs4 import BeautifulSoup

def url2html(url, headers=None, params=None, data=None):
    headers={'User-Agent': 'Mozilla/5.0'}
    try:
        req = requests.get(url, headers=headers, params=params, data=data)
    except:
        req = requests.get(url, headers=headers, params=params, data=data, verify=False)
    html = req.text
    return html

    
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
        self.urls = []
        self.stock_cd = stock_cd
        print(f"Stock Cd: {self.stock_cd}")

        self.overview, self.stock_id = self.scrape_overview()
        print(f"Stock Id: {self.stock_id}")

        self.growth_rate, simplywallst_url = self.scrape_growth_rate()
        self.beta, infrontanalytics_url = self.scrape_beta()
        self.discount_rate = self.scrape_discount_rate()
        self.urls.append(simplywallst_url).append(infrontanalytics_url)

        self.ratios = self.scrape_ratios()
        self.cash_flow = self.scrape_cash_flow()
        self.balance_sheet = self.scrape_balance_sheet()
        # self.income_statementp = Webpage.from_url(f"https://www.investing.com/equities/{stock_cd}-income-statement")
        # self.earningsp = Webpage.from_url(f"https://www.investing.com/equities/{stock_cd}-earnings")
        # self.financialp = Webpage.from_url(f"https://www.investing.com/equities/{stock_cd}-financial-summary")
    
    """
    Simply Wall St
    """
    def scrape_growth_rate(self):
        """scrape growth rate from simply wall st"""

        # search the link for stock
        stock_cd = self.stock_cd.replace("-", " ")
        params = (
            ('x-algolia-agent', 'Algolia for JavaScript (4.2.0); Browser (lite)'),
            ('x-algolia-api-key', 'be7c37718f927d0137a88a11b69ae419'),
            ('x-algolia-application-id', '17IQHZWXZW'),
        )
        data = f'{{"query":"{stock_cd} klse","highlightPostTag":" ","highlightPreTag":" ","restrictHighlightAndSnippetArrays":true}}'
        try:
            response = requests.post('https://17iqhzwxzw-dsn.algolia.net/1/indexes/companies/query', params=params, data=data)

            # generate link
            stock_url = response.json()['hits'][0]['url']
            url = "https://simplywall.st" + stock_url
        except:
            return None
        
        html = url2html(url)
        soup = BeautifulSoup(html, 'html.parser')
        growth = soup.find('p', {'data-cy-id': 'key-metric-value-forecasted-annual-earnings-growth'}).get_text().replace('%', '')
        self.growth_rate = float(growth)/100
        print(f"Growth Rate: {self.growth_rate}")
        return self.growth_rate, url
    
    """
    Infront Analytics
    """
    def scrape_beta(self):
        """scrape beta from infrontanalytics.com"""

        # search the link for stock
        params = (
            ('keyname', self.stock_cd.replace("-", " ")),
        )
        response = requests.get('https://www.infrontanalytics.com/Eurofin/autocomplete', params=params, verify=False)
        result = response.json()[0]

        # generate stock url
        name = result['name'].replace(" ", "-").replace(".", "") + "-"
        code = result['isin']
        url = f"https://www.infrontanalytics.com/fe-en/{code}/{name}/beta"

        # get beta
        html = url2html(url)
        m = re.search(r"shows a Beta of ([+-]?\d+\.\d+).", html)
        beta = m.groups()[0]
        print(f"Beta: {beta}")
        return float(beta), url
    
    def scrape_discount_rate(self):
        """convert beta to discount rate for dcf model"""
        
        discount_rate = 0
        dr = {0.8: 5, 
        1: 6, 
        1.1: 6.8, 
        1.2: 7, 
        1.3: 7.9, 
        1.4: 8, 
        1.5: 8.9}
        for key in dr:
            if self.beta < key:
                discount_rate = dr[key]
            else:
                discount_rate = 9
        discount_rate = round(discount_rate/100, 2)
        print(f"Discount Rate: {discount_rate}")
        return discount_rate

    """
    i3investor
    """
    def srape_isummary(self):
        headers = {
            'User-Agent': 'Mozilla',
        }

        params = (
            ('qt', 'lscomn'),
            ('qp', 'nestle'),
        )

        response = requests.get('https://klse.i3investor.com/cmservlet.jsp', headers=headers, params=params)
        query = response.text.split(":")[0]
        params = (
    ('sa', 'ss'),
    ('q', query),
)

response = requests.get('https://klse.i3investor.com/quoteservlet.jsp', headers=headers, params=params)
html = response.text


        

    """
    investing
    """
    def scrape_overview(self):
        stock_cd = self.stock_cd
        def scrape_id(overviewp):
            m = re.search('data-pair-id="(\d+)"', overviewp.html)
            stock_id = m.groups()[0]
            return stock_id

        url = f"https://www.investing.com/equities/{stock_cd}"
        overviewp = Webpage.from_url(url)
        soup = overviewp.soup
        last_price = soup.find('span', {'id':'last_last'}).get_text()
        ls = ['Last Price', last_price]
        df = pd.DataFrame([ls])
        
        overview = overviewp.get_span('span', ['float_lang_base_1', 'float_lang_base_2'])
        stock_id = scrape_id(overviewp)

        return pd.concat([df, overview]), stock_id, url
    
    def scrape_ratios(self):
        stock_cd = self.stock_cd
        ratiosp = Webpage.from_url(f"https://www.investing.com/equities/{stock_cd}-ratios")
        tables = ratiosp.tables
        numbers = range(1, 9)
        ratios = pd.concat(tables[i] for i in numbers)
        return ratios
    
    def scrape_cash_flow(self):
        stock_id = self.stock_id
        cash_flowp = Webpage.from_url(f"https://www.investing.com/instruments/Financials/changereporttypeajax?action=change_report_type&pair_ID={self.stock_id}&report_type=CAS&period_type=Annual")
        df = cash_flowp.tables[0]
        cash_flow = df[~df[1].str.contains("a|e|i|o|u")]
        return cash_flow
    
    def scrape_balance_sheet(self):
        stock_id = self.stock_id
        balance_sheetp = Webpage.from_url(f"https://www.investing.com/instruments/Financials/changereporttypeajax?action=change_report_type&pair_ID={self.stock_id}&report_type=BAL&period_type=Annual")
        df = balance_sheetp.tables[0]
        balance_sheet = df[~df[1].str.contains("a|e|i|o|u")]
        return balance_sheet
    
    def scrape_earnings(self):
        stock_cd = self.stock_cd
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
    
    def scrape_financial_summary(self):
        def get_summary(html):
            webpage = Webpage(html)
            soup = webpage.soup

            title = soup.find('h3').text
            df = webpage.get_span('span', ['float_lang_base_1', 'float_lang_base_2'])
            table = pd.read_html(str(soup))[0]
            return [title, table, df] # pd.concat([table, df], axis=0, ignore_index=True)
        
        stock_id = self.stock_id
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