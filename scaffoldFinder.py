from Adjustment import *
from Excel1 import *
from Excel2 import *
import argparse
import re

parse = argparse.ArgumentParser()
parse.add_argument('-targets', '--TARGET_LIST', required=True, type=str)
parse.add_argument('-excel1', '--EXCEL_FORM_1', required=True, type=str)
parse.add_argument('-excel2', '--EXCEL_FORM_2', required=True, type=str)
parse.add_argument('-out', '--OUTPUT_PATH', required=True, type=str)
parse.add_argument('-log2FoldChange_value', '--LOG2FOLDCHANGE', required=False, default=1.0, type=float)
parse.add_argument('-pvalue_value', '--PVALUE', required=False, default=0.05, type=float)
args = parse.parse_args()

targetList = args.TARGET_LIST
Excel1Path = args.EXCEL_FORM_1
Excel2Path = args.EXCEL_FORM_2
outputPath = args.OUTPUT_PATH
log2FoldChangeValue = args.LOG2FOLDCHANGE
pvalueValue = args.PVALUE

Annotations = {}
Definition = {}

# Recive targets.
targets = []
tarPattern = re.compile(r'K\d*')
with open(targetList) as tar:
    for line in tar:
        if tarPattern.search(line) != None:
            targets.append(line.replace('\n', ''))

# Process Excel 1.
Excel1 = excel1Process(Excel1Path, targets, Annotations, Definition)
Excel1.processExcel1()

# Process Excel 2.
Excel2 = excel2Process(Excel2Path, Annotations, Definition, log2FoldChangeValue, pvalueValue)
Finalwb, finalSheet = Excel2.processExcel2()

# Adjustment.
print('\nGenerating output...\n')
adjust = Adjustment()
adjust.adjustment(Finalwb, finalSheet, outputPath)

print('\nDone.')
