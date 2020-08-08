# -*- coding: utf-8 -*-
#! /usr/bin/env python3

"""Writing down pivots and resonances for each participant in <data> folder, measuring acoustic reflex. Outputs in corresponding pivot files.

::Input:: ```<data>``` | data folder path
::Output:: ``excel.xlsx``, ``total.xlsx``| output path

.. Examples::

    Default using::

        $ reflex.py

"""

import os
import pandas as pd
import xlsxwriter
from openpyxl import load_workbook
import re
from scipy.optimize import fsolve,root,brentq
import numpy as np
import matplotlib.pyplot as plt
from scipy.interpolate import BPoly, CubicSpline
import warnings
import xlrd
from tqdm import tqdm

from utils import write_xls, find_resonance 
import argparse


def main():   

    parser = argparse.ArgumentParser(usage='"Parses input and output paths for pivots and resonances for each participant in <data> folder')

    parser.add_argument('-i', '--input', help='Path to <data> folder', required=False, default='../data/')

    parser.add_argument('-o', '--output_pivot',
                        help='Output file name [default: ../excel/total.xlsx from current directory]',
                        required=False, default='../excel/total.xlsx')

    parser.add_argument('-a', '--output_alfa', help='Output resonance file name [default: ../excel/alfa.xlsx from current directory]',
                        required=False, default='../excel/alfa.xlsx')
    
    args = parser.parse_args()
    lst_files = os.listdir(args.input)

    if not os.path.exists('../excel'):
        os.makedirs('../excel')
        
    print ('Writing pivot table...')
    write_xls(file_path = args.output_pivot, files_dir = args.input, \
              lst_files = lst_files)

    print ('Calculating the resonances...')

    find_resonance(args.output_pivot, args.output_alfa)
    print('Done !')
    

if __name__ == "__main__":
    main()
