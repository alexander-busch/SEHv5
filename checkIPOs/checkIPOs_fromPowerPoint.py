# -*- coding: utf-8 -*-
"""
Created on Wed Apr  5 16:57:27 2023

@author: AlexanderB
"""

import os
from pptx import Presentation
from tabulate import tabulate
from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter

# Directory path
path = r"C:\Users\alexanderb\INCOSE\SEHv5 German Translation - SEHv5\_Übersetzungsfiles\Abbildungen\zz_old"

# pptx file containing the IPO diagrams
file = r'SEHB_V5_All_Figures_and_Tables2022-11-12.pptx'

#Debugging only
#file = r'test.pptx'

#search string to identify slides that contain an IPO diagram
search_string = "IPO diagram"

# Define the headers for the table
headers = ['IPO diagram name',
           'Typical Inputs', 
           'Controls',
           'Activities',
           'Typical Outputs',
           'Enablers']

# Define an empty list to store the table data
IPO_data = []

# Function used for sorting results such that it conforms to structure of headers
def extract_subarray(array, key):
    for subarray in array:
        if subarray[0] == key:
            return subarray[1]
    return None

# For the case of multiple .pptx files, not needed here
#for filename in os.listdir(folder_path):
#   if filename.endswith(".pptx"):

# Get presentation and loop all slides
presentation = Presentation(os.path.join(path,file))
for slide in presentation.slides:
    
    # Check if slides has notes and get notes text
    if slide.has_notes_slide:
        notes_slide = slide.notes_slide
        notes_text = notes_slide.notes_text_frame.text
        
        # Check if this is an IPO diagram
        if search_string in notes_text:
            
            # Extract IPO name; Split the string into two parts based on the delimiter and extract the remaining part that follows the delimiter
            delimiter = "IPO diagram for "
            parts = notes_text.split(delimiter, 1)
            IPO_name = parts[1].strip()
            print(IPO_name)
            
            # Loop all shapes on that slide
            category_array=[]            
            for shape in slide.shapes:
                
                # Check if the shape is a group
                if shape.shape_type == 6:
    
                    # Loop through all shapes in the group
                    result_array = []
                    for group_shape in shape.shapes:
                        
                        # Check if this shape has a text frame
                        if group_shape.has_text_frame:
                            text_frame = group_shape.text_frame
                            
                            # Loop all paragraphs in this text frame and write to result_array
                            for paragraph in text_frame.paragraphs:
                                result_array.append(paragraph.text)
                                
                                # Loop elements of rsult_array, identify the header and separate content
                                content_array = []
                                for entry in result_array:
                                    if not entry == '':
                                        if entry in headers:
                                            header = entry
                                            print('Category {}' .format(header))
                                        else:
                                            content_array.append(entry)
                                            print('Content {}' .format(entry))
                
                    # Build array for this IPO diagram
                    category_array.append([header, content_array])
                   
            # Sort category_array such that it conforms to the structure of headers
            category_array_sorted=[]
            for header in headers:
                category_array_sorted.append(extract_subarray(category_array, header))
            
            #Debugging only   
            #category_array_sorted = [None, ['a','a'], ['b','b', 'b','b'],['c','c'],['d'],['d','d']]
            
            # Write all data to table
            #IPO_data.append([IPO_name, category_array_sorted[1], category_array_sorted[2], category_array_sorted[3], category_array_sorted[4], category_array_sorted[5]])
            IPO_data.append([IPO_name, str(category_array_sorted[1]), str(category_array_sorted[2]), str(category_array_sorted[3]), str(category_array_sorted[4]), str(category_array_sorted[5])])
            
# Print tabulated data
print(tabulate(IPO_data, headers=headers))


# Write to Excel

# Create a new Workbook object and select the active worksheet
workbook = Workbook()
worksheet = workbook.active

# Add the data and column headers from the Table object to the worksheet
# Start appending rows from row 2
worksheet.append(headers)
for row in IPO_data:
    worksheet.append(row)


# Get the number of rows and columns in the worksheet
num_rows = worksheet.max_row
num_cols = worksheet.max_column

# Create an Excel cell reference using the number of rows and columns
start_cell = 'A2'
end_cell = get_column_letter(num_cols) + str(num_rows)

# Create a new Table object and set its style
table = Table(displayName="IPO", ref=':'.join(['A1',get_column_letter(num_cols) + str(num_rows)]))
style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False, showLastColumn=False, showRowStripes=True, showColumnStripes=True)
table.tableStyleInfo = style

# Add the Table object to the worksheet and save the workbook to a file
worksheet.add_table(table)
workbook.save(os.path.join(path,"IPO.xlsx"))

