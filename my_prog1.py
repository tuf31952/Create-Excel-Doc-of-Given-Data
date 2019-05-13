#!/usr/bin/env python

# for the random number list I used a random number generator then stored each number into a list
# for the names column I used an online random number generator then saved the html file and scrapped each of the names into a new line in a file
# for the websites list I used a site called https://majestic.com/reports/majestic-million which has a list of the top million websites as well as an option to export a file to excel I then edited the data down to 100,000 and stored it to an array
# for the addresses data I used a similar random US address generator I found online and exported that to a file with each address on a new line

# also the assignment says to output in seperate files but Dr. Picone sent an email that he would like everything in one workbook

import xlwt 
from xlwt import Workbook 
import xlsxwriter
import sys
import os
import random
import numpy as np

def main():

    # declare variables for storage
    names = []
    websites = []
    addresses = []

    # generate random numbers for first column
    randnums = np.random.randint(1,100001,100000)

    # open rach file and store data into lists
    with open('names.txt') as name_file:
            for lines in name_file:
                names.append(lines)

    with open('websites.txt') as websites_file:
            for lines in websites_file:
                websites.append(lines)

    with open('addresses.txt') as addresses_file:
            for lines in addresses_file:
                addresses.append(lines)

    # create workbook for output file
    workbook = xlsxwriter.Workbook('output.xlsx')

    # store each required list into a different sheet
    sheet1 = workbook.add_worksheet('List of 3')
    sheet2 = workbook.add_worksheet('List of 100')
    sheet3 = workbook.add_worksheet('List of 1000')
    sheet4 = workbook.add_worksheet('List of 100000')

    # set x to 0 to make index reset
    x = 0

    # add each entry from data files into columns one one each row
    for i in randnums:
        sheet1.write(x, 0, i) 
        x += 1
        if (x>2):
            break

    x = 0

    for i in names:
        sheet1.write(x, 2, i) 
        x += 1
        if (x>2):
            break

    x = 0

    for i in websites:
        sheet1.write(x, 1, i) 
        x += 1
        if (x>2):
            break

    x = 0

    for i in addresses:
        sheet1.write(x, 3, i) 
        x += 1
        if (x>2):
            break

    x = 0

    for i in randnums:
        sheet2.write(x, 0, i) 
        x += 1
        if (x>99):
            break

    x = 0

    for i in names:
        sheet2.write(x, 2, i) 
        x += 1
        if (x>99):
            break

    x = 0
    
    for i in websites:
        sheet2.write(x, 1, i) 
        x += 1
        if (x>99):
            break

    x = 0

    for i in addresses:
        sheet2.write(x, 3, i) 
        x += 1
        if (x>99):
            break

    x = 0

    for i in randnums:
        sheet3.write(x, 0, i) 
        x += 1
        if (x>999):
            break

    x = 0

    for i in names:
        sheet3.write(x, 2, i) 
        x += 1
        if (x>999):
            break

    x = 0
    
    for i in websites:
        sheet3.write(x, 1, i) 
        x += 1
        if (x>999):
            break

    x = 0

    for i in addresses:
        sheet3.write(x, 3, i) 
        x += 1
        if (x>999):
            break

    x = 0

    for i in randnums:
        sheet4.write(x, 0, i) 
        x += 1
        if (x>99999):
            break

    x = 0

    for i in names:
        sheet4.write(x, 2, i) 
        x += 1
        if (x>99999):
            break

    x = 0
    
    for i in websites:
        sheet4.write(x, 1, i) 
        x += 1
        if (x>99999):
            break

    x = 0

    for i in addresses:
        sheet4.write(x, 3, i) 
        x += 1
        if (x>99999):
            break

    # close the workbook
    workbook.close()

if __name__ == "__main__": 
    main()