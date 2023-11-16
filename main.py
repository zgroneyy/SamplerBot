#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Thu Nov 16 22:48:56 2023

@author: pro
"""

from home_gui import create_gui
from methods import create_test_file, random_sampler, stratified_sampler, systematic_sampler, cluster_sampler

if __name__ == "__main__":
    
    create_test_file("sample.xlsx",3,250)
    create_gui()