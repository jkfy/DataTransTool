import pandas as pd
from openpyxl import load_workbook

# 模板和导出表的对应关系,只导表中存在的项目, key 新TC表的项目（列），value，旧TC表的项目（列）
_transDict = {
    '采购类型': '物料类型',
    '名称1': '名称',
    '部件编码': '部件编号',
    '类型编码': '代号（型号）',
    'ERP编码': 'ERP编码',
    '制造商': '备注 (厂家)',
    '重量': '重量(kg)',
    '材料': '材质',
}

_tc_data_new = pd.DataFrame(columns=['是否合法', '部件编码', '名称1', 'ERP编码', '类型编码', '制造商', '重量', '材料', '采购类型'])
_tc_data_old = pd.DataFrame(columns=['序号', '物料类型', '申请人', '标准管理员', '名称', '代号（型号）', '规格', '材质', '部件编号', 'ERP编码', '备注 (厂家)', '重量(kg)'])
_row_init = 6 # 模板中序号下面一行的行号
_column_init = 2 # 模板中序号所在的列号
_column_count = 12 # 模板中需要读取的列的数量，从序号开始计数


def fill_old_data(worksheet, row : int):
    global _tc_data_old
    column = _column_init
    data_temp = []
    for index in range(_column_count):
        data_temp.append(worksheet.cell(row, column + index).value)
    print(data_temp)

    _tc_data_old.loc[len(_tc_data_old.index)] = data_temp

def read_data_from_old_excel(file_path : str):
    wb = load_workbook(file_path, data_only=True)
    ws = wb.active
    
    row = _row_init
    column = _column_init
    while ws.cell(row = row, column= column).value is not None:
        fill_old_data(ws, row)
        row = row + 1

def save_data_to_new_excel(file_path):
    for key, value in _transDict.items():
        _tc_data_new[key] = _tc_data_old[value]
    print(_tc_data_old.head(5))
    _tc_data_new.to_excel(file_path, index=False)

def main():
    file_path_read = "./data/物料编码申请单数据总模板.xlsx"
    file_path_save = "./data/旧TC申请表.xlsx"
    read_data_from_old_excel(file_path_read)
    save_data_to_new_excel(file_path_save)

if __name__ == "__main__":
    main()