import win32com.client

class HandleExcel:

    def __init__(self, filepath):
        self.xlsx_app = win32com.client.Dispatch('Excel.Application')
        self.filepath = filepath
        self.xlsx_book = self.xlsx_app.Workbooks.Open(filepath, ReadOnly=False)

    def get_nrows(self, sheet):
        # 获取总行数
        sht = self.xlsx_book.Worksheets(sheet)
        return sht.UsedRange.Rows.Count
    
    def get_ncols(self, sheet):
        # 获取总列数
        sht = self.xlsx_book.Worksheets(sheet)
        return sht.UsedRange.Columns.Count

    def get_col_list(self, sheet, row, col):
        # 获取指定单元格及其以下的该列数据列表，包含该单元格
        sht = self.xlsx_book.Worksheets(sheet)
        col_list = []
        nrows = self.get_nrows(sheet)
        for i in range(row, nrows+1):
            col_list.append(sht.Cells(i,col).value)
        return col_list
    
    def get_row_list(self, sheet, row, col):
        sht = self.xlsx_book.Worksheets(sheet)
        row_list = []
        ncols = self.get_ncols(sheet)
        for i in range(col, ncols+1):
            row_list.append(sht.Cells(row, i).value)
        return row_list
    
    def getCell(self, sheet, row, col): 
        sht = self.xlsx_book.Worksheets(sheet) 
        return sht.Cells(row, col).Value

    def setCell(self, sheet, row, col, value): 
        sht = self.xlsx_book.Worksheets(sheet)
        sht.Cells(row, col).Value = value

    def save(self, newfilename=None):
        if newfilename: 
            self.xlsx_book.SaveAs(newfilename) 
        else: 
            self.xlsx_book.Save()

    def close(self): 
        self.xlsx_book.Close(SaveChanges=1)