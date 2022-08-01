import csv
import os
import re

csv_regex = re.compile(r'\S*\.csv\Z')
folder = input("folder:\n")
search = input("search:\n")
filename = input("combined file name:\n")

columns_list = ["Sequenznummer", "Einheit in [N]", "Bruchcode",
                "Messkommentar", "Datum", "Uhrzeit"]
column_no_dict = {}

with open(rf'{folder}\{filename}.csv', 'w', newline='') as combined_csvfile, open (rf'{folder}\{filename}_fmode.csv', 'w', newline='') as combined_csvfile_fmode:
    filewriter = csv.writer(combined_csvfile, delimiter = ';')
    filewriter_fmode = csv.writer(combined_csvfile_fmode, delimiter = ';')

    for root, dirs, files in os.walk(folder, topdown = False):
        
            for file in files:
                
                if csv_regex.search(file) and search in file:
                    path = os.path.join(root,file)
                    break
            with open (path, newline = '') as csvfile:
                csvreader = csv.reader(csvfile, delimiter = ';' )
                for index, row in enumerate(csvreader):
                    cells =[]

                    for column in range(len(row)):
                        for item in columns_list:
                            if item == row[column].strip():
                                first_row_no = index+1
                                column_no_dict[item] = column
                                cells.append(row[column].strip())
                    if cells:
                        filewriter.writerow(cells)
                                
            for file in files:                   
                
                if csv_regex.search(file) and search in file:
                    print(file)
                    path = os.path.join(root,file)
                    fmode_dict_local = {}
                    fmode_dict_local_total = 0
                    fmode_list_local = []
                    
                    with open (path, newline = '') as csvfile:
                        
                        csvreader = csv.reader(csvfile, delimiter = ';' )
                        
                        for row in csvreader:
                            sample = row[0]
                            filewriter_fmode.writerow([sample])
                            break
                        
                        
                        for i in range (first_row_no):
                            next(csvreader)

                        for index, row in enumerate(csvreader):
                            cells = []
                            
                            for column in range(len(row)):
                                cells.append(row[column].strip())
                                if column == column_no_dict["Bruchcode"]:
                                    fmode_list_local.append(row[column])
                                
                            if cells:
                                if index == 0:
                                    cells.append(sample)
                                    
                                filewriter.writerow(cells)
                                
                    for item in fmode_list_local:
                        if item not in fmode_dict_local:
                            fmode_dict_local[item] = 1
                            fmode_dict_local_total +=1
                        else:
                            fmode_dict_local[item] += 1
                            fmode_dict_local_total +=1
                    for keys,values in fmode_dict_local.items():
                        fmode_cells = [keys, values/fmode_dict_local_total*100]
                        filewriter_fmode.writerow(fmode_cells)
