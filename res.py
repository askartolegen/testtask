import os
import xlsxwriter


class IterFiles:
    def __init__(self, path):
        self._path = path
        self.__res = []

    def __get_iter_files(self, rootdir):
        try:
            for lists in os.listdir(rootdir):
                path = os.path.join(rootdir, lists)
                self.__res.append(path)
                if os.path.isdir(path):
                    self.__get_iter_files(path)
        except WindowsError:
            print('Some directory does not exist')
        finally:
            return self.__res

    def __call__(self, *args, **kwargs):
        return self.__get_iter_files(self._path)


class Writer:
    def __init__(self, source, col_names):
        self.row = 0
        self.col = 0
        self.workbook = xlsxwriter.Workbook(source)
        self.worksheet = self.workbook.add_worksheet()
        if len(col_names.split(' | ')) != 4:
            raise ValueError('Столбцов должы быть 4')
        for elem in col_names.split(' | '):
            self.worksheet.write(self.row, self.col, elem)
            self.col += 1
        self.row += 1
        self.col = 0

    def __call__(self, lst):
        for elem in lst:
            self.worksheet.write(self.row, self.col, str(self.row))
            temp = elem.split('\\')
            self.worksheet.write(self.row, self.col+1, temp[-2])
            last = temp[-1]
            spliter = last.split('.')
            if last.startswith('.') or len(spliter) == 1:
                self.worksheet.write(self.row, self.col + 2, last)
            else:
                self.worksheet.write(self.row, self.col + 2, spliter[0])
                self.worksheet.write(self.row, self.col + 3, spliter[1])
            self.row += 1
        self.workbook.close()


a = IterFiles(os.getcwd())
writer = Writer('result.xlsx', 'Номер строки | Папка в которой лежит файл | название файла | расширение файла')
writer(a())
