# import os
# import xlsxwriter
#
# workbook = xlsxwriter.Workbook('result.xlsx')
# worksheet = workbook.add_worksheet()
#
# worksheet.write('A1', 'Номер строки')
# worksheet.write('B1', 'Папка в которой лежит файл')
# worksheet.write('C1', 'Название файла')
# worksheet.write('D1', 'Расширение файла')
#
# row = 1
# col = 0
# for dirpath, dirnames, filenames in os.walk(r"..\testtask"):
#     dirpath = dirpath[3:]
#
#     for dirname in dirnames:
#         worksheet.write(row, col, row)
#         worksheet.write(row, col + 1, dirpath)
#         worksheet.write(row, col + 2, dirname)
#         row += 1
#     for filename in filenames:
#         temp = filename.split('.')
#         if filename[0] == '.' or len(temp) == 1:
#             worksheet.write(row, col, row)
#             worksheet.write(row, col + 1, dirpath)
#             worksheet.write(row, col + 2, filename)
#         else:
#             worksheet.write(row, col, "Файл")
#             worksheet.write(row, col + 1, dirpath)
#             worksheet.write(row, col + 2, temp[0])
#             worksheet.write(row, col + 3, temp[1])
#         row += 1
#
# workbook.close()
