from openpyxl import Workbook
from scrapy.http import FormRequest
import csv
import os.path
import scrapy
import glob
from decouple import config


class ItemsSpider(scrapy.Spider):
    name = 'items'
    allowed_domains = ['cs.trade']
    start_urls = ['http://cs.trade/']

    def start_requests(self):
        frmdata = {}
        url = "https://cs.trade/loadBotInventory?order_by=price_desc&bot=all&_=1648140077424"
        headers = {
            'Connection': 'keep-alive',
            'sec-ch-ua': '" Not A;Brand";v="99", "Chromium";v="98", "Google Chrome";v="98"',
            'Accept': 'application/json, text/javascript, */*; q=0.01',
            'sec-ch-ua-mobile': '?0',
            'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/98.0.4758.102 Safari/537.36',
            'sec-ch-ua-platform': '"Linux"',
            'Origin': 'https://cs.trade',
            'Sec-Fetch-Site': 'same-site',
            'Sec-Fetch-Mode': 'cors',
            'Sec-Fetch-Dest': 'empty',
            'Referer': 'https://cs.trade/',
            'Accept-Language': 'en-US,en;q=0.9,vi;q=0.8',
            'cookie': config("COOKIE")
        }

        yield FormRequest(url, callback=self.parse, formdata=frmdata, headers=headers)

    def parse(self, response):
        items = response.json()["inventory"]
        for item in items:
            if item["app_id"] == self.appId:
                yield {
                    "name": item["market_hash_name"],
                    "price": item["price"],
                }

    def close(self, reason):
        csv_file = max(glob.iglob('*csv'), key=os.path.getctime)

        wb = Workbook()
        ws = wb.active

        with open(csv_file, 'r') as f:
            for row in csv.reader(f):
                ws.append(row)

        wb.save(csv_file.replace('.csv', '') + '.xlsx')
