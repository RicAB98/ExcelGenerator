import os
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell
import re
import json
from JsonHelper import JsonHelper

class Generator:
    def __init__(self, fileName, numberPeriods, numberStudents, sections, finalSection):
        if not os.path.exists(f'FicheirosGerados'):
            os.makedirs('FicheirosGerados')
            
        self.workbook = xlsxwriter.Workbook(f'FicheirosGerados/{fileName}.xlsx')
        self.fileName = fileName
        self.numberPeriods = numberPeriods
        self.numberStudents = numberStudents
        self.sections = sections
        self.finalSection = finalSection

        self.GenerateFormats() 

    def GenerateFormats(self):
        self.fullBorderFormat = self.workbook.add_format()
        self.fullBorderFormat.set_border_color('#000000')
        self.fullBorderFormat.set_border(1)
        self.fullBorderFormat.set_center_across()
        self.fullBorderFormat.set_locked(False)

        self.sideBorderFormat = self.workbook.add_format()
        self.sideBorderFormat.set_left_color('#000000')
        self.sideBorderFormat.set_left(1)
        self.sideBorderFormat.set_right_color('#000000')
        self.sideBorderFormat.set_right(1)
        self.sideBorderFormat.set_center_across()
        self.sideBorderFormat.set_locked(False)

        self.lastCellBorderFormat = self.workbook.add_format()
        self.lastCellBorderFormat.set_left_color('#000000')
        self.lastCellBorderFormat.set_left(1)
        self.lastCellBorderFormat.set_right_color('#000000')
        self.lastCellBorderFormat.set_right(1)
        self.lastCellBorderFormat.set_bottom_color('#000000')
        self.lastCellBorderFormat.set_bottom(1)
        self.lastCellBorderFormat.set_center_across()
        self.lastCellBorderFormat.set_locked(False)

        self.bottomBorderFormat = self.workbook.add_format()
        self.bottomBorderFormat.set_bottom_color('#000000')
        self.bottomBorderFormat.set_bottom(1)
        self.bottomBorderFormat.set_locked(False)

        self.mediaFormat = self.workbook.add_format()
        self.mediaFormat.set_border_color('#000000')
        self.mediaFormat.set_border(1)
        self.mediaFormat.set_bg_color('#D3D3D3')
        self.mediaFormat.set_locked(True)
        self.mediaFormat.set_center_across()

    def Generate(self):
        for i in range(1,self.numberPeriods+1):
            self.currentRow = 1
            self.currentColumn = 1
            self.localAverageCells = []

            worksheet = self.workbook.add_worksheet(f'{i}º Período')
            worksheet.protect()

            self.CreateNumbersColumn(worksheet)
            self.CreateStudentsColumn(worksheet)

            for s in self.sections:
                self.CreateSection(s, worksheet)

            self.CreateFinalSection(worksheet, i)
        self.workbook.close()

    def GenerateJson(self):
        properties = {
            'fileName': self.fileName,
            'numberPeriods': self.numberPeriods,
            'numberStudents': self.numberStudents,
            'sections': self.sections,
            'finalSection': self.finalSection,
        }

        dictionary = JsonHelper.ConvertTo(properties)

        with open(f'FicheirosGerados/{self.fileName}.json', 'w') as f:
            json.dump(dictionary, f)

    def CreateNumbersColumn(self, worksheet):
        worksheet.write(self.currentRow + 1, self.currentColumn, 'Nº', self.fullBorderFormat)

        for n in range(1, self.numberStudents + 1):
            worksheet.write(self.currentRow + 1 + n, self.currentColumn, n, self.fullBorderFormat)

    def CreateStudentsColumn(self, worksheet):
        self.currentColumn += 1

        worksheet.write(self.currentRow + 1, self.currentColumn, 'Aluno', self.fullBorderFormat)
        worksheet.set_column(self.currentColumn, self.currentColumn, 35)

        for n in range(1, self.numberStudents + 1):
            if n == self.numberStudents: #Add bottom border to last cell of column
                worksheet.write(self.currentRow + 1 + n, self.currentColumn, ' ', self.lastCellBorderFormat)
                continue

            worksheet.write(self.currentRow + 1 + n, self.currentColumn, ' ', self.sideBorderFormat)

    def CreateSection(self, section, worksheet):
        self.currentColumn += 1

        colors = section.ColorToHex()

        #Header formatting
        cellFormat = self.workbook.add_format()
        cellFormat.set_border_color('#000000')
        cellFormat.set_border(1)
        cellFormat.set_bg_color(str(colors["Header"]))
        cellFormat.set_font_color(colors["Text"])
        cellFormat.set_center_across()
        cellFormat.set_locked(False)

        noBorderFormat = self.workbook.add_format()
        noBorderFormat.set_border_color('#000000')
        noBorderFormat.set_border(1)
        noBorderFormat.set_center_across()
        noBorderFormat.set_locked(False)

        #Merge cells
        worksheet.merge_range(self.currentRow, self.currentColumn, self.currentRow, self.currentColumn + section.subSections - 1, section.name, cellFormat) #Section cell
        worksheet.merge_range(self.currentRow + 1, self.currentColumn, self.currentRow + 1, self.currentColumn + section.subSections - 1,'', cellFormat) #SubSections cell

        for n in range(0, section.subSections):
            for m in range(0, self.numberStudents - 1):
                worksheet.write(self.currentRow + m + 2, self.currentColumn + n, ' ', noBorderFormat) #Add no locked to normal cells
            worksheet.write(self.currentRow + self.numberStudents + 1, self.currentColumn + n, ' ', noBorderFormat) #Add bottom border to last cell of column

        self.currentColumn += section.subSections
        
        worksheet.write(self.currentRow, self.currentColumn, f'{section.percentage}%', cellFormat) #Percentage cell
        worksheet.write(self.currentRow + 1, self.currentColumn, 'Média', cellFormat) #Average cell

        #Add average formula to last column of component
        for n in range(0, self.numberStudents):
            worksheet.write_formula(self.currentRow + 2 + n, 
            self.currentColumn, 
            f'=IFERROR(AVERAGE({xl_rowcol_to_cell(self.currentRow + 2 + n, self.currentColumn - section.subSections)}:{xl_rowcol_to_cell(self.currentRow + 2 + n, self.currentColumn - 1)}),0)', self.mediaFormat)
        
        averageCell = xl_rowcol_to_cell(self.currentRow + 1, self.currentColumn)

        #Store section cells to dictionary
        self.localAverageCells.append(
            {'AverageColumn': averageCell[:re.search(r"\d", averageCell).start()], 
            'PercentageCell': xl_rowcol_to_cell(self.currentRow, self.currentColumn)
            })

    def CreateFinalSection(self, worksheet, period):
        self.currentColumn += 1

        colors = self.finalSection.ColorToHex()

        cellFormat1 = self.workbook.add_format()
        cellFormat1.set_border_color('#000000')
        cellFormat1.set_border(1)
        cellFormat1.set_bg_color(str(colors["Header"]))
        cellFormat1.set_font_color(str(colors["Text"]))
        cellFormat1.set_center_across()

        formulaString = '='

        for c in self.localAverageCells:
            formulaString += f'{c["AverageColumn"]}ROWNUMBER * {c["PercentageCell"]}'

            if c != self.localAverageCells[-1]:
                formulaString += ' + '

        for n in range(0, self.numberStudents + 1):
            worksheet.write(self.currentRow + 1 + n, self.currentColumn, formulaString.replace('ROWNUMBER', str(self.currentRow + 2 + n)), self.mediaFormat)
            
            averageFormulaString = '=AVERAGE('

            for m in range(1, period + 1):
                averageFormulaString += f'\'{m}º Período\'!{xlsxwriter.utility.xl_col_to_name(self.currentColumn)}{self.currentRow + 2 + n},'

            averageFormulaString = averageFormulaString[:-1] + ')'

            worksheet.write(self.currentRow + 1 + n, self.currentColumn + 1, averageFormulaString, self.mediaFormat)
            worksheet.write(self.currentRow + 1 + n, self.currentColumn + 2, ' ', self.fullBorderFormat)
            worksheet.write(self.currentRow + 1 + n, self.currentColumn + 3, ' ', self.fullBorderFormat)

        worksheet.merge_range(self.currentRow, self.currentColumn, self.currentRow + 1, self.currentColumn, f'Média {period}ºP', cellFormat1)
        worksheet.set_column(self.currentColumn, self.currentColumn, len(f'Média {period}ºP'))
        self.currentColumn += 1

        worksheet.merge_range(self.currentRow, self.currentColumn, self.currentRow + 1, self.currentColumn, 'Média final', cellFormat1)
        worksheet.set_column(self.currentColumn, self.currentColumn, len('Média final'))
        self.currentColumn += 1

        worksheet.merge_range(self.currentRow, self.currentColumn, self.currentRow + 1, self.currentColumn, 'Autoavaliação', cellFormat1)
        worksheet.set_column(self.currentColumn, self.currentColumn, len('Autoavaliação'))
        self.currentColumn += 1

        worksheet.merge_range(self.currentRow, self.currentColumn, self.currentRow + 1, self.currentColumn, 'Nível a atribuir', cellFormat1)
        worksheet.set_column(self.currentColumn, self.currentColumn, len('Nível a atribuir'))