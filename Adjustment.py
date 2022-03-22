import os


class Adjustment:
    def adjustment(self, Finalwb, finalSheet, outputPath):
        # Adjustment.
        ALength = len(finalSheet['A1'].value)
        BLength = len(finalSheet['B1'].value)
        CLength = len(finalSheet['C1'].value)
        DLength = len(finalSheet['D1'].value)
        ELength = len(finalSheet['E1'].value)
        i = 2

        while True:
            if finalSheet[f'A{str(i)}'].value == None:
                break
            else:
                if ALength <= len(str(finalSheet[f'A{str(i)}'].value)):
                    ALength = len(str(finalSheet[f'A{str(i)}'].value))

                if BLength <= len(str(finalSheet[f'B{str(i)}'].value)):
                    BLength = len(str(finalSheet[f'B{str(i)}'].value))

                if CLength <= len(str(finalSheet[f'C{str(i)}'].value)):
                    CLength = len(str(finalSheet[f'C{str(i)}'].value))

                if DLength <= len(str(finalSheet[f'D{str(i)}'].value)):
                    DLength = len(str(finalSheet[f'D{str(i)}'].value))

                if ELength <= len(str(finalSheet[f'E{str(i)}'].value)):
                    ELength = len(str(finalSheet[f'E{str(i)}'].value))
            
            i += 1

        finalSheet.column_dimensions['A'].width = ALength + 5
        finalSheet.column_dimensions['B'].width = BLength + 5
        finalSheet.column_dimensions['C'].width = CLength + 5
        finalSheet.column_dimensions['D'].width = DLength + 5
        finalSheet.column_dimensions['E'].width = ELength + 5

        fileNumber = 1
        output = os.path.join(outputPath, 'Result.xlsx')
        while True:
            exist = os.path.exists(output)
            if exist:
                output = os.path.join(outputPath, f'Result{fileNumber}.xlsx')
                fileNumber += 1
            else:
                Finalwb.save(output)
                break
