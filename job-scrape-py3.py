#! /usr/bin/python3
# -*- coding: utf-8 -*-


from lxml import html
import requests
import csv
from openpyxl import Workbook
import datetime
import sys


def main(filename):
    jobDict = {}    # Main dictionary to hold the data we need to export
    pn = 1          # Page number were scraping
    totalPages = 1  # Total number of pages to scrape (will change in loop)

    # Main loop
    while pn <= totalPages:
        # Try to get the page, error out if we can't connect for some reason
        try:
            page = requests.get("http://www.thingamajob.com/Job-Search-" +
                                "Results.aspx?SORTF=&SORTD=D&PAGE_NO=" +
                                str(pn) + "&searchType=4&state=26&statename" +
                                "=Missouri")
        except:
            print("Unexpected error: ", sys.exc_info()[0], file=sys.stderr)
            sys.exit(1)
        tree = html.fromstring(page.content)

        if pn == 1:
            # Get the total number of jobs available, number of jobs per page
            # and total pages based off of the previous two variables
            nav = tree.xpath('//span[@id="lblCurrentPageTop"]/text()')
            navSplit = nav[0].split(' ')
            totalJobs = int(navSplit[-1])
            jobsPerPage = int(navSplit[-3])
            totalPages = (totalJobs / jobsPerPage) + 1

        # Grab all of the new and old jobs on the current page
        oldJobListings = tree.xpath('//tr[@class="job1"]/. | //tr[@class=' +
                                    '"job2"]/.')
        newJobListings = tree.xpath('//tr[@class="job1New"]/. | //tr[@class' +
                                    '="job2New"]/.')

        # Break down the information we've gathered
        jobDict = updateJobDict(jobDict, oldJobListings, False)
        jobDict = updateJobDict(jobDict, newJobListings, True)

        pn += 1

    # Save to either CSV or to XLSX
    tries = 0
    while tries < 2:
        # Try creating a CSV
        if filename[-3:] == "csv":
            if not createCSV(jobDict, filename):
                if tries == 0:
                    print("Cannot create CSV. Creating a XLSX instead.",
                          file=sys.stderr)
                    filename = filename[:-3] + ".xlsx"
                    tries += 1
                    continue
                else:
                    print("Cannot create CSV.  Aborting.", file=sys.stderr)
                    tries += 1
                    continue
            else:
                break

        # Try creating a XLSX
        if filename[-4:] == "xlsx":
            if not createXLSX(jobDict, filename):
                if tries == 0:
                    print("Cannot create XLSX. Creating a CSV instead.",
                          file=sys.stderr)
                    filename = filename[:-4] + ".csv"
                    tries += 1
                    continue
                else:
                    print("Cannot create XLSX.  Aborting.", file=sys.stderr)
                    tries += 1
                    continue
            else:
                break


def updateJobDict(d, l, new=False):
    # Iterate through l and update d accordingly
    if l != []:
        for x in l:
            company = x.xpath('./td[1]/span/text()')
            title = x.xpath('./td[2]/strong/a/text()')
            link = x.xpath('./td[2]/strong/a/@href')
            location = x.xpath('./td[3]/text()')
            posted = x.xpath('./td[4]/text()')
            code = link[0][-7:]
            d[code] = {'company': company[0],
                       'title': title[0].strip(),
                       'link': link[0],
                       'location': location[0].strip(),
                       'posted': posted[0].strip(),
                       'new': new}
    # This will return an unchanged d if there was nothing in l
    return d


def createCSV(d, filename):
    # Dump jobDict in to a CSV file; Each entry in jobDict is a new row
    try:
        with open(filename, 'wt') as csvfile:
            writer = csv.writer(csvfile, dialect='excel')
            writer.writerow(['Company', 'New', 'Job Title', 'Link',
                             'Location', 'Posting Date'])
            for x in d:
                rowToWrite = []
                rowToWrite.append(d[x]['company'])
                rowToWrite.append(d[x]['new'])
                rowToWrite.append(d[x]['title'])
                rowToWrite.append(d[x]['link'])
                rowToWrite.append(' '.join(d[x]['location'].split()))
                rowToWrite.append(d[x]['posted'])
                writer.writerow(rowToWrite)
    except:
        return False

    # Everything seems peachy
    return True


def createXLSX(d, filename):
    # Create an XLSX from jobDict; Each entry in jobDict is a new row
    try:
        wb = Workbook()
        ws = wb.active
        ws.append(['Company', 'New', 'Job Title', 'Link', 'Location',
                   'Posting Date'])
        for x in d:
            rowToWrite = []
            rowToWrite.append(d[x]['company'])
            if d[x]['new']:
                rowToWrite.append('true')
            else:
                rowToWrite.append('false')
            rowToWrite.append(d[x]['title'])
            rowToWrite.append(d[x]['link'])
            rowToWrite.append(' '.join(d[x]['location'].split()))
            jobDates = d[x]['posted'].split('/')
            rowToWrite.append(datetime.date(int(jobDates[2]),
                                            int(jobDates[0]),
                                            int(jobDates[1])))
            ws.append(rowToWrite)
        wb.save(filename)
    except:
        return False

    # Everything seems peachy
    return True


if __name__ == '__main__':
    if len(sys.argv) > 1:
        # Give the filename a quick check to see if it's CSV or XLSX
        if sys.argv[1][-3:] == 'csv' or sys.argv[1][-4:] == 'xlsx':
            main(sys.argv[1])
        else:
            sys.exit("Unknown file type.  Must be either a CSV or XLSX.")
    else:
        print("job-scrape-py3.py filename", file=sys.stderr)
        print("    'filename' is the name of the CSV or XLSX file to be " +
              "created.", file=sys.stderr)
        sys.exit(1)
