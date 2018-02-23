# -*- coding: utf-8 -*-
from xml.sax.saxutils import unescape
import pyexcel as pe
from itertools import islice
import os
import time
import re
from os.path import basename
import progressbar
import scrapy
from scrapy.crawler import CrawlerProcess
from scrapy.http.request import Request
from scrapy.spidermiddlewares.httperror import HttpError
from twisted.internet.error import DNSLookupError
from twisted.internet.error import TimeoutError

bar = progressbar.ProgressBar()
overall_bar = progressbar.ProgressBar()

# Master TBL
records = pe.get_sheet(file_name="tbl_master-in-ECM-test-1.30.18.xlsx")

list_of_pids = []
for pid in records.column[2]:
    if type(pid) == int:
        list_of_pids.append(str(int(pid)))
    else:
        list_of_pids.append(pid)

# List of locales that are present in the TBL Master in the same order as they are there excluding en-us
list_of_locales = ["cs-cz","da-dk","de-de","de-at","de-ch","en","en-ca","en-gb","en-id","en-ie","en-in","en-my","en-ph","en-sg","en-th","en-tt","en-vn","es-mx","es-ar","es-bo","es-cl","es-co","es-cr","es-do","es-ec","es-gt","es-pe","es-sv","es-uy","es-ve","th-th","tr-tr","vi-vn","zh-cn","cn","zh-tw","en-au","es-es","fi-fi","fr-fr","fr","fr-be","fr-ca","fr-ch","id-id","it-it","ja-jp","ko-kr","nl-nl","nl-be","no-no","pl-pl","pt-br","pt-pt","ru-ru","sv-se"]
list_of_row_locales = []
row_urls = []

def ecm_to_row_locale(locale):
    if locale == 'en-gb':
        new_locale = 'uken'
    elif locale == 'zh-tw':
        new_locale = 'twen'
    elif locale == ''
    else:
        last = locale[-2:]
        first = locale[0:2]
        new_locale = last + first
    fr - m1fr
    en - 3 possibilities
    print(new_locale)

    return new_locale

def row_url_builder(locale, pid):
    locale = ecm_to_row_locale(locale)
    url = f'http://www.fluke.com/fluke/{locale}/digital-multimeters/wireless-testers/Fluke-279-FC.htm?PID={pid}'
    return url

def get_pid(row):
    if not row[2]:
        pid = 'unknown'
    else:
        pid = str(int(row[2]))
    return pid

for locale in list_of_locales:
    list_of_row_locales.append(ecm_to_row_locale(locale))

for row in bar(islice(records, 1, None)):
    pid = get_pid(row)
    for x, status in enumerate(islice(row, 14, 70)):
        locale = records[0, x+14]
        if status == 'X' and pid != 'unknown':
            row_urls.append(row_url_builder(locale, pid))
        elif status == 'X' and pid == 'unknown':
            row[x+14] = 1
        else:
            pass
    
pbar = progressbar.ProgressBar(max_value=len(row_urls))

# for url in row_urls:
#   print(url)

PAGES = [*row_urls]

pbar.start()

count = 0

class FindX(scrapy.Spider):
    name = "Find X"
    handle_httpstatus_list = [404, 504, 503] 
    start_urls = PAGES
    # start_urls = ['http://www.fluke.com/fluke/czcs/digital-multimeters/wireless-testers/Fluke-279-FC.htm?PID=54672']
    # start_urls = ['http://www.fluke.com/fluke/twen/digital-multimeters/wireless-testers/Fluke-279-FC.htm?PID=56120']
    def parse(self, response):
        global count
        count += 1
        pbar.update(count)
        # print('\nURL:',response.url)

        referring_url = str(response.meta.get('redirect_urls', [response.url])[0])
        # print('re',referring_url)
        re_url = re.search("/fluke/(.*)/digital-multimeters/wireless-testers/Fluke-279-FC.htm\?PID=(.*)", referring_url)

        locale = re_url.group(1)
        pid = re_url.group(2)

        # print(pid)
        # print(locale)

        # Redirects
        if response.status == 404:
            publish_status = 0
        # Leaving 504 and 503 errors as-is for now
        elif response.status == 504:
            publish_status = 'X'
        elif response.status == 503:
            publish_status = 'X'
        # Has red Discontinued text
        elif response.xpath('//*[@id="tblMainContent"]//font[contains(@color, "red")]').extract_first():
            publish_status = 0
        # The requested url is different than the response url
        elif referring_url != response.url:
            publish_status = 0
        else:
            publish_status = 1

        column_of_locale = list_of_row_locales.index(locale)
        matching_rows = [i for i, x in enumerate(list_of_pids) if x == pid]
       
        for row in matching_rows:
            records[row, column_of_locale+14] = publish_status
        time.sleep(0.1)

process = CrawlerProcess({
    'USER_AGENT': 'Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 5.1)',
    'LOG_LEVEL': 'INFO',
})

process.crawl(FindX)
process.start() # the script will block here until the crawling is finished

pbar.finish()

records.save_as('tbl_master-in-ECM-test-1.30.18.xlsx')