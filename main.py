from properties import exel_file_name
from weather import *
from openpyxl import load_workbook

province = input('원하는 동네 이름을 입력하세요')

if __name__ == "__main__":
    request_and_print(province)
        