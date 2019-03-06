import xlrd

class MyXlsx:
    
    def __init__(self,file):
        self.ExcelFile=xlrd.open_workbook(file)
        self.ExcelSheetNames=self.ExcelFile.sheet_names()
        self.ExcelSheets = self.ExcelFile.sheets()
        self.SheetDatas = []
        for sheet in self.ExcelSheets:
            
            rows = sheet.nrows
            cols = sheet.ncols
            data = {}
            data["row"]=rows
            data["col"]=cols
            data["merged"]=sheet.merged_cells
            data["name"]=sheet.name
            data["data"]=[]
            
            for i in range(rows):
                data["data"].append(sheet.row(i))

            self.SheetDatas.append(data)
            
     

if __name__ == '__main__':
    a= MyXlsx("asdf.xlsx")
    print(a.ExcelSheets[0].name)
