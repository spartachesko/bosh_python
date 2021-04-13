import openpyxl
import os

target_path = 'C:/Users/borisov/PycharmProjects/bosh_python/Excel/automate_online-materials/example.xlsx'
path_normalized = os.path.normpath(target_path)
wb = openpyxl.load_workbook(path_normalized)
type(wb)
