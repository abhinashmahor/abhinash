# -*- coding: utf-8 -*-
"""
Created on Thu Nov 11 12:30:22 2021

@author: Lenovo
"""

from file2 import Search_functions
from file1 import Make_grid
import openpyxl


'''
Enter Top left coords here
'''
top_left = (19.00093234571652, 72.83252167275216)
'''
Enter top right coordinates here
'''
top_right =(19.00003964872778, 72.861146205218)
'''
Enter bottom right coordinates here
'''
bottom_right = (18.973134761476267, 72.86067413700916)
'''
e
Enter horizontal here
'''
horizontal_distance = 3
'''
Enter vertical distance here
'''
vertical_distance = 3
'''
Enter Terms you want to search here
'''
search_list = ['hospital','nursing home']


'''
Enter Name of the city
'''
cn = 'Mumbai'


'''
Input file name you want it to name for example: 'bhopal_hospitals.xlsx'
'''
file_name ='Mumbai_city.xlsx'

'''
Enter path of the file you want to save the file in
'''
file_name = f"C:\\Users\\Lenovo\\Downloads\\SCRAPER\\places_list_program\\{file_name}.xlsx"





Make_grid(top_left,top_right, bottom_right,horizontal_distance,vertical_distance).method1()
wb= openpyxl.Workbook()
sheet2 = wb.active




try:
    p = Search_functions.read_cord_slow(cn,search_list,sheet2)
except:
    wb.save(file_name)
else:
    wb.save(file_name)