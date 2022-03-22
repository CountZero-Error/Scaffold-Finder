import openpyxl
import os, re

class excel1Process:
    def __init__(self, Excel1Path, targets, Annotations, Definition):
        self.Excel1Path = Excel1Path
        self.targets = targets
        self.Annotations = Annotations
        self.Definition = Definition
    
    def check(self, annotation, cellValue, definition):
        # If group(2) length != 4, add 0 till length = 4.
        changeAnnotation = re.compile(r'(evm.model.Scaffold)(\d*)(.\d*)')
        patternResult = changeAnnotation.search(annotation.value)
        if patternResult != None:
            length = len(patternResult.group(2))
            if len != 4:
                newAnnotation = f'{patternResult.group(1)}{str(0)*(4 - length)}{patternResult.group(2)}{patternResult.group(3)}'
                self.Annotations[newAnnotation] = cellValue
                self.Definition[newAnnotation] = definition.value
            else:
                self.Annotations[annotation.value] = cellValue
                self.Definition[annotation.value] = definition.value
    
    def searchEntry(self, wb1Sheet):
        # Start from cell B2 to Bx.
        x = 'B'
        y = 2

        # Search Entry in Excel 1.
        while True:
            location = x + str(y)
            currentCell = wb1Sheet[location]
            cellValue = currentCell.value

            # Search target.
            if cellValue in self.targets:
                print(f'Find Entry - {cellValue}')
                annotation = wb1Sheet[f'A{str(y)}']
                definition = wb1Sheet[f'D{str(y)}']

                # If group(2) length != 4, add 0 till length = 4.
                self.check(annotation, cellValue, definition)

            elif cellValue == None:
                break

            y += 1

    def processExcel1(self):
        # Process Excel 1.
        print(f'Processing Excel - {os.path.basename(self.Excel1Path)}')
        wb1 = openpyxl.load_workbook(self.Excel1Path)
        wb1Sheet = wb1.active

        # Search Entry in Excel 1.
        self.searchEntry(wb1Sheet)
                    
        print('\nFinish processing Excel 1.\n')
        wb1.save(os.path.basename(self.Excel1Path))
