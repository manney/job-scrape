#! /usr/bin/python
# -*- coding: utf-8 -*-

from lxml import html
import requests
import csv
import sys
from openpyxl import Workbook
import datetime


def main(filename):
    # Get the jobs from the page(s) and put them in to jobDict
    jobDict = {}
    page = requests.get("http://www.thingamajob.com/Job-Search-Results.aspx?SORTF=&SORTD=D&PAGE_NO=1&searchType=4&state=26&statename=Missouri")
    tree = html.fromstring(page.content)

    navigation = tree.xpath('//span[@id="lblCurrentPageTop"]/text()')
    navSplit = navigation[0].split(' ')
    totalJobs = int(navSplit[-1])
    jobsPerPage = int(navSplit[-3])
    totalPages = (totalJobs / jobsPerPage) + 1

    pn = 1
    while pn <= totalPages:
        if pn != 1:
            page = requests.get("http://www.thingamajob.com/Job-Search-Results.aspx?SORTF=&SORTD=D&PAGE_NO=" + str(pn) + "&searchType=4&state=26&statename=Missouri")
            tree = html.fromstring(page.content)

        oldJobListings = tree.xpath('//tr[@class="job1"]/. | //tr[@class="job2"]/.')
        newJobListings = tree.xpath('//tr[@class="job1New"]/. | //tr[@class="job2New"]/.')

        if newJobListings != []:
            for x in newJobListings:
                company = x.xpath('./td[1]/span/text()')
                title = x.xpath('./td[2]/strong/a/text()')
                link = x.xpath('./td[2]/strong/a/@href')
                location = x.xpath('./td[3]/text()')
                posted = x.xpath('./td[4]/text()')
                code = link[0][-7:]
                jobDict[code] = {'company': company[0],
                                 'title': title[0].strip(),
                                 'link': link[0],
                                 'location': location[0].strip(),
                                 'posted': posted[0].strip(),
                                 'new': True}

        if oldJobListings != []:
            for x in oldJobListings:
                company = x.xpath('./td[1]/span/text()')
                title = x.xpath('./td[2]/strong/a/text()')
                link = x.xpath('./td[2]/strong/a/@href')
                location = x.xpath('./td[3]/text()')
                posted = x.xpath('./td[4]/text()')
                code = link[0][-7:]
                jobDict[code] = {'company': company[0],
                                 'title': title[0].strip(),
                                 'link': link[0],
                                 'location': location[0].strip(),
                                 'posted': posted[0].strip(),
                                 'new': False}
        pn = pn + 1

    # Dump jobDict in to a CSV file
    if filename[-3:] == "csv":
        with open(filename, 'wb') as csvfile:
            writer = csv.writer(csvfile, dialect='excel')
            writer.writerow(['Company', 'New', 'Job Title', 'Link',
                             'Location', 'Posting Date'])
            for x in jobDict:
                rowToWrite = []
                rowToWrite.append(jobDict[x]['company'])
                rowToWrite.append(jobDict[x]['new'])
                rowToWrite.append(jobDict[x]['title'])
                rowToWrite.append(jobDict[x]['link'])
                rowToWrite.append(' '.join(jobDict[x]['location'].split()))
                rowToWrite.append(jobDict[x]['posted'])
                writer.writerow(rowToWrite)
    elif filename[-4:] == "xlsx":
        wb = Workbook()
        ws = wb.active
        ws.append(['Company', 'New', 'Job Title', 'Link', 'Location',
                   'Posting Date'])
        for x in jobDict:
            rowToWrite = []
            rowToWrite.append(jobDict[x]['company'])
            if jobDict[x]['new']:
                rowToWrite.append('true')
            else:
                rowToWrite.append('false')
            rowToWrite.append(jobDict[x]['title'])
            rowToWrite.append(jobDict[x]['link'])
            rowToWrite.append(' '.join(jobDict[x]['location'].split()))
            jobDates = jobDict[x]['posted'].split('/')
            rowToWrite.append(datetime.date(int(jobDates[2]),
                                            int(jobDates[0]),
                                            int(jobDates[1])))
            ws.append(rowToWrite)
        wb.save(filename)
    else:
        print "Unknown file type.  Must be either a CSV or XLSX."
        sys.exit()

if __name__ == '__main__':
    if len(sys.argv) > 1:
        main(sys.argv[1])
    else:
        print "job-scrape.py filename"
        print "  filename is the name of the CSV or Excel file(s) to be created."
        sys.exit()
