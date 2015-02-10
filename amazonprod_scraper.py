#!/usr/bin/python2.7
# -*- coding: utf-8 -*-
# by: pantuts
# http://pantuts.com
# Dependencies: python2.7, BeautifulSoup4, xlrd, xlwt
# Licence? None. Do what you want. Just a credit is fine.
# Agreement: This script is for educational purposes only. By using this script you agree
# that you alone will be responsible for any act you make. The author will not be liable
# of your actions.

from bs4 import BeautifulSoup
import os.path, re, sys, time
import urllib2
import xlrd
from xlwt import *

def usage():
    print 'python2.7 amazonprod_scraper.py infile.xlsx outfile'


def main():

    if len(sys.argv) < 3: 
        usage()
        sys.exit(0)

    infile = sys.argv[1]
    outfile = sys.argv[2] if '.xls' in sys.argv[2] else sys.argv[2] + '.xls'

    # open sheet
    workbook = xlrd.open_workbook(infile)
    sheet = workbook.sheet_by_index(0)
    urls = sheet.col_values(0)[1:]

    # check if sheet to write exists
    if os.path.isfile(outfile):
        chfile = time.strftime("%d-%b-%Y %H:%M:%S", time.gmtime())
        print 'File exists! Changing output filename to ' + chfile + ' - ' + outfile
        outfile = chfile + ' - ' + outfile
    else:
        outfile = outfile

    # write sheet with labels
    write_book = Workbook()
    wb_sheet = write_book.add_sheet('Amazon Images')
    wb_sheet.write(0, 1, 'Brand')
    wb_sheet.write(0, 2, 'Specification')
    wb_sheet.write(0, 3, 'Model')
    wb_sheet.write(0, 4, 'Price')
    wb_sheet.write(0, 5, 'Delivery')
    wb_sheet.write(0, 6, 'LevelGreen')
    wb_sheet.write(0, 7, 'TopSpecs')
    wb_sheet.write(0, 8, 'Description')
    for i in range(11):
        if i == 0:
            wb_sheet.write(0, 0, 'URLs')
        else:
            if i > 0 and i < 6:
                wb_sheet.write(0, i + 8, 'HiRES')
            elif i > 5 and i < 11:
                wb_sheet.write(0, i + 8, 'MidRES')
    write_book.save(outfile)
    

    for num, url in enumerate(urls):

        try:

            req = urllib2.Request(url)
            req.add_header('User-Agent', 'Mozilla/5.0 (Windows NT 6.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/33.0.1750.146 Safari/537.36')
            opener = urllib2.build_opener()
            print '[+] ' + str(num + 1), 'Scraping: ' + url
            res = opener.open(req)
            response = res.read()
            soup = BeautifulSoup(response, 'html.parser')

        except urllib2.URLError, e:
            print '[-] ' + str(e)
            # write not found url
            with open("NOT FOUND URLs.txt", "a+") as f:
                f.write('Row: ' + str(num + 2) + ' - ' + url + '\n')
            continue
        except Exception, e:
            print '[-] ' + str(e)
            continue

        brand = ''.join([ i.get_text() for i in soup.select('a#brand') ])
        if brand:
            wb_sheet.write(num + 1, 1, brand)
        else:
            wb_sheet.write(num + 1, 1, 'NONE')
        # brand = brand.encode('utf-8')
        
        price = ''.join([ i.get_text() for i in soup.select('span#priceblock_ourprice') ])
        if price:
            price = '£' + price.encode('ascii', 'ignore')
            wb_sheet.write(num + 1, 4, price.decode('utf-8'))
        else:
            sale_price = ''.join([ i.get_text() for i in soup.select('span#priceblock_saleprice') ])
            if sale_price:
                sale_price = '£' + sale_price.encode('ascii', 'ignore')
                wb_sheet.write(num + 1, 4, sale_price.decode('utf-8'))
            else:
                wb_sheet.write(num + 1, 4, 'NONE')

        deliv = [ i.get_text() for i in soup.find_all('b', text=re.compile(r'.*?Delivery.*?')) ]
        # delivery = deliv.encode('utf-8')
        if deliv:
            if deliv[0] != u'Delivery Destinations:':
                wb_sheet.write(num + 1, 5, deliv[0])
            else:
                wb_sheet.write(num + 1, 5, 'NONE')
        else:
            wb_sheet.write(num + 1, 5, 'NONE')
        
        availability = [ i.get_text().strip() for i in soup.select("div#availability span") ]
        if availability:
            wb_sheet.write(num + 1, 6, availability[0])
        else:
            wb_sheet.write(num + 1, 6, 'NONE')
        # availability = availability.encode('utf-8')
        
        prod_desc = re.compile('<div class="bucket" id="productDescription">\s*<h2>Product Description</h2>\s*<div class="content">\s*(.*?)\s*<script type="text/javascript">',re.S|re.I)
        prod_desc = re.findall(prod_desc, urllib2.unquote(response))
        if prod_desc:
            soup2 = BeautifulSoup(prod_desc[0], 'html.parser')
            product_desc = [ i.get_text() for i in soup2.select('div.productDescriptionWrapper') ][0]
        else:
            product_desc = 'NONE'
        # product_desc = product_desc.encode('utf-8')
        wb_sheet.write(num + 1, 8, '<p>' + product_desc.strip() + '</p>')
        
        model = ''
        
        prod_specification = ''
        prod_td1 = [ i.get_text() for i in soup.select("table#technicalSpecifications_sections td.td1") ]
        prod_td2 = [ i.get_text() for i in soup.select("table#technicalSpecifications_sections td.td2") ]
        if prod_td1:
            for n, v in enumerate(prod_td1):
                # print n, v
                if 'Model number' not in prod_td1:
                    if v.lower() == 'part number':
                        model = prod_td2[n]
                if v.lower() == "model number":
                    # add to model
                    model = prod_td2[n]
                prod_specification += '\t<li>' + v.strip() + ': ' + prod_td2[n] + '</li>\n'
        else:
            prod_specification = 'BS4: no parsed data'
            print 'BS4: no parsed data'
        wb_sheet.write(num + 1, 2, '<ul>\n' + prod_specification + '</ul>')
        if model:
            wb_sheet.write(num + 1, 3, model)
        else:
            wb_sheet.write(num + 1, 3, 'NONE')
        
        soup3 = BeautifulSoup(urllib2.unquote(response), 'html.parser')
        top_specification = [ i.get_text() for i in soup3.select('div#feature-bullets li span')]
        if top_specification:
            top_specs = ''
            for sp in top_specification:
                top_specs += '\t<li>' + sp + '</li>\n'
            wb_sheet.write(num + 1, 7, '<ul>\n' + top_specs + '</ul>')
        else:
            wb_sheet.write(num + 1, 7, 'NONE')
        # top_specification = top_specification.encode('utf-8')

        # regex for hiRes: re.findall(r'hiRes":"http.*?.jpg', s)
        # regex for main(midRes): re.findall(r'main":{\S.*?}', s)
        # regex for midRes: re.findall(r'http:.*?.jpg', s[0])[1] using for loop
        # once hiRes is taken, it means there will always be a midres in main
        
        # large list to be used when hiRes is null
        # we set list to remove duplicates
        lge = re.compile(r'large":"http.*?.jpg')
        hres = re.compile(r'hiRes":"http.*?.jpg')
        mmidres = re.compile(r'main":{\S.*?}')
        mdres = re.compile(r'http:.*?.jpg')

        large = re.findall(lge, response)
        hires = list(set(re.findall(hres, response)))
        mainmidres = re.findall(mmidres, response)
        midres = []
        if mainmidres:
            for mres in mainmidres:
                tmp_mres = re.findall(mdres, mres)
                if len(tmp_mres) < 2:
                    midres.append(tmp_mres[0])
                else:
                    midres.append(tmp_mres[1])
            # set midres
        midres = list(set(midres))

        # if resolutions are detected, write to file
        if hires:
            wb_sheet.write(num + 1, 0, url)
            # if hiRes is more than expected links
            if len(hires) > 5:
                hires = hires[:5]
            for i in range(len(hires)):
                wb_sheet.write(num + 1, i + 9, hires[i].replace('hiRes":"', ''))
            if midres:
                for i in range(len(midres)):
                    wb_sheet.write(num + 1, i + 14, midres[i])
            else:
                print 'No midRes found'
            print '>>> done'
        else:
            if large:
                wb_sheet.write(num + 1, 0, url)
                print 'No hiRes found: Using large instead'
                wb_sheet.write(num + 1, 9, large[0].replace('large":"', ''))
                if midres:
                    for i in range(len(midres)):
                        wb_sheet.write(num + 1, i + 14, midres[i])
                else:
                    print 'No midRes found'
                print '>>> done'
            else:
                print 'REGEX: nothing found, skipping.'
                continue

        # save
        write_book.save(outfile)

if __name__ == '__main__':
    try:
        main()
    except KeyboardInterrupt:
        print '\nKeyboard Interrupt!'