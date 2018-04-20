import xlrd
from xlutils.copy import copy
 
class ParsingExcel:
    def __init__(self, name, change_text):
        self.name = name
        self.change_text = change_text
        self.index = None
        self.book = None
        self.copy = None
        self.__getWorkbook()
        self.__addWorkbook()
    
    def __getWorkbook(self):
        self.book = xlrd.open_workbook(self.name)
        
    def __addWorkbook(self):
        self.copy = copy(self.book)
    # Get sheet
    def getpmEventFormat(self):
        return self.book.sheet_by_name('pmEventFormat')
    
    def getpmEvents(self):
        return self.book.sheet_by_name('pmEvents')
    
    def getLocalEvents(self):
        return self.book.sheet_by_name('LocalEvents')
    
    def getLocalEventFormat(self):
        return self.book.sheet_by_name('LocalEventFormat')
    
    def getEventParams(self):
        return self.book.sheet_by_name('EventParams')
    
    def getDeprecationReasonList(self, sheet):
        col_index = 0
#         print sheet.name
        for cell_name in sheet.row_values(0): # col_list
            if cell_name == u"Deprecation? (N/D/'blank')":
#                 print cell_name
                deprecation_list = sheet.col_values(col_index)
#                 print len(deprecation_list)
            elif cell_name == u'Reason for change':
#                 print cell_name
                self.index = col_index
#                 change_list = sheet.col_values(col_index)
#                 print change_list
            else: pass
            col_index += 1
        return deprecation_list
    
    def reasonforChange(self, sheet, dep):
        getsheet = self.copy.get_sheet(sheet.name)
        row_index = 3
        for col_value in dep:
            if col_value == u'D' or col_value == u'N': 
                getsheet.write(row_index, self.index, u'%s' % self.change_text)
            else: pass
            row_index += 1

    def save(self):
        self.copy.save(self.name)

if __name__ == '__main__':
    excel = 'CAH1091864_27_R46A_PA2.xls'
#     excel = 'test.xlsx'
    parser = ParsingExcel(excel, 'test')
    eventparams_sheet = parser.getEventParams()
    dep = parser.getDeprecationReasonList(eventparams_sheet)
    parser.reasonforChange(eventparams_sheet, dep)
    parser.save()