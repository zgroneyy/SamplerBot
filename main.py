#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# TODO: Selection of the sheet on selected XLSX file will be added.
#        Selection of the rows that keeps headings will be added.        
#       Select document -> select sheet -> select # of row that keeps headings
#    Print result on the program window!!!!
#    File selection default path will be determined as Desktop. 
#    Show headings in a table (Create Entry e, e.grid(row=i, column=j)
#    Ask before exiting the program
#    
"""
Created on Thu Nov 16 22:48:56 2023

@author: pro
"""

from home_gui import create_gui
from methods import create_test_file, random_sampler, stratified_sampler, systematic_sampler, cluster_sampler

if __name__ == "__main__":
    
    create_test_file("sample.xlsx",3,250)
    create_gui()
