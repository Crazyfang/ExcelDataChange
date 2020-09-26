import xlrd
import xlwt


class ChangeClass:
    def __init__(self, file_path):
        self.file_path = file_path
        self.value = []

    def read_data_from_file(self):
        excel = xlrd.open_workbook(self.file_path)
        ws = excel.sheet_by_index(2)

        print(ws.nrows)
        print(ws.ncols)

        for row_index in range(0, ws.nrows, 8):
            index = ws.cell(rowx=row_index, colx=0).value
            for col_index in range(2, ws.ncols):
                if ws.cell(rowx=row_index+1, colx=col_index).value.strip() != '':
                    self.value.append([index,
                                       ws.cell(rowx=row_index+1, colx=col_index).value,
                                       ws.cell(rowx=row_index+2, colx=col_index).value,
                                       ws.cell(rowx=row_index+3, colx=col_index).value,
                                       ws.cell(rowx=row_index+4, colx=col_index).value,
                                       ws.cell(rowx=row_index+5, colx=col_index).value,
                                       ws.cell(rowx=row_index+6, colx=col_index).value,
                                       ws.cell(rowx=row_index+7, colx=col_index).value])

        print(self.value)
        print(len(self.value))

    def write_data_to_file(self):
        workbook = xlwt.Workbook(encoding='utf-8')
        data_sheet = workbook.add_sheet('胜光')

        for i in range(0, len(self.value)):
            for j in range(0, 8):
                data_sheet.write(i, j, self.value[i][j])

        workbook.save('胜光.xls')


if __name__ == '__main__':
    test = ChangeClass('C:\\Users\\Administrator\\Desktop\\敬德村户口簿.xls')
    test.read_data_from_file()
    test.write_data_to_file()
