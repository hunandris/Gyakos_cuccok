# -*- coding: utf-8 -*-
"""
Created on Tue Jul  5 12:32:00 2022

@author: HAA4BP
"""

import csv
import os
import re

csv_regex = re.compile(r'\S*\.csv\Z')
folder = input("folder:\n")



for root, dirs, files in os.walk(folder, topdown = False):
        for file in files:
            if csv_regex.search(file):
                print(file)
                path = os.path.join(root,file)
                with open (file, 'w', newline = '') as filtered_csvfile:
                    filewriter = csv.writer(filtered_csvfile, delimiter = ';')
                    with open (path, newline = '') as csvfile:
                        csvreader = csv.reader(csvfile, delimiter = ';' )
                        next(csvreader)
                        for row in csvreader:
                            cells = []
                            if row[0] == "TEST":
                                if row[4] == '7':
                                    continue
                                else:
                                    for column in range(len(row)):
                                        cells.append(row[column])
                                    
                            else:
                                for column in range(len(row)):
                                    cells.append(row[column])

                            filewriter.writerow(cells)
