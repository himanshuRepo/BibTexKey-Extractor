#!/usr/bin/env python

"""
Code for generating the bibtex key from Google Scholar for the list of papers, whose names are stored
    in the excel sheet.

Command to Run: python main.py

Author: Himanshu Mittal (himanshu.mittal224@gmail.com)
Referred: https://github.com/venthur/gscholar
"""

import optparse
import sys
import os
import xlsxwriter
import pandas as pd
import gscholar as gs

def main():

    # Path to the excel sheet containing the list of paper title in the second colum, heading as 'Name'.
    pathToFile="PaperList.xlsx"

    xl = pd.ExcelFile(pathToFile)
    df = xl.parse("Sheet1")
    bt = []
    f=df['PaperName']
    for i in range(f.size):
        a=f[i]
        x = a.replace(u'\xa0', ' ')
        args=x.encode('ascii','ignore')
        # args ="Detection of skin cancer by classification of Raman spectra"
        biblist = gs.query(args)
        print(biblist[0])
        k=biblist[0]
        k1 = k.replace(u'\n ', ' ')
        x1=k1.encode('ascii','ignore')
        bt.append(x1)
    df1 = pd.DataFrame({'bibtex': bt})
    f=pd.concat([df, df1], axis=1)

    # Create a Pandas Excel writer using XlsxWriter as the engine.
    writer = pd.ExcelWriter('bibtexFile.xlsx', engine='xlsxwriter')

    # Convert the dataframe to an XlsxWriter Excel object.
    f.to_excel(writer, sheet_name='Sheet1')

    # Get the xlsxwriter objects from the dataframe writer object.
    workbook  = writer.book
    worksheet = writer.sheets['Sheet1']



if __name__ == "__main__":
    main()