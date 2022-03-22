import openpyxl
import os

class excel2Process:
    def __init__(self, Excel2Path, Annotations, Definition, log2FoldChangeValue, pvalueValue):
        self.Excel2Path = Excel2Path
        self.Annotations = Annotations
        self.Definition = Definition
        self.log2FoldChangeValue = log2FoldChangeValue
        self.pvalueValue = pvalueValue

    def searchAnnotation(self, wb2Sheet):
        # Start from A2 to Ax.
        x = 'A'
        y = 2

        # Create final Excel.
        Finalwb = openpyxl.Workbook()
        finalSheet = Finalwb.active
        finalSheet['A1'] = 'Annotation'
        finalSheet['B1'] = 'Entry'
        finalSheet['C1'] = 'log2FoldChange'
        finalSheet['D1'] = 'Definition'
        finalSheet['E1'] = 'group'
        finaly = 2

        # Search Annotation in Excel 2.
        while True:
            location = x + str(y)
            currentCell = wb2Sheet[location]
            cellValue = currentCell.value

            # Search target.
            if cellValue in self.Annotations.keys():

                log2FoldChange = wb2Sheet[f'C{str(y)}']
                pvalue = wb2Sheet[f'F{str(y)}']

                if abs(log2FoldChange.value) >= self.log2FoldChangeValue and pvalue.value < self.pvalueValue:
                    print(f'Find Scaffold - {cellValue}')
                    finalSheet[f'A{finaly}'] = cellValue
                    finalSheet[f'B{finaly}'] = self.Annotations[cellValue]
                    finalSheet[f'C{finaly}'] = log2FoldChange.value
                    finalSheet[f'D{finaly}'] = self.Definition[cellValue]
                    finalSheet[f'E{finaly}'] = wb2Sheet[f'H{str(y)}'].value
                    finaly += 1
            
            elif cellValue == None:
                break

            y += 1
        
        return Finalwb, finalSheet

    def processExcel2(self):
        # Process Excel 2.
        print(f'Processing Excel - {os.path.basename(self.Excel2Path)}')
        wb2 = openpyxl.load_workbook(self.Excel2Path)
        wb2Sheet = wb2.active

        Finalwb, finalSheet = self.searchAnnotation(wb2Sheet)

        # Finish.
        wb2.save(os.path.basename(self.Excel2Path))

        print('\nFinish processing Excel 2.\n')

        return Finalwb, finalSheet
        