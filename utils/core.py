import os
from docx import Document
from tqdm import tqdm
from openpyxl import Workbook
from utils.common import get_output_file


def get_template(template_file, omit_doc_name=False):
    '''
    Get the template of the docx file
    '''
    # word中的表格如果有使用合并单元格，python-docx读取会重复，使用text_map进行去重
    # 如果单元格内容不为空，就是要提取的位置，记录第几个表格，行，列，作为模板
    column_name = [] if omit_doc_name else ['']  # excel列名
    template, text_map = {}, {}
    try:
        dfile = Document(template_file)
    except:
        print('打开模板文件失败!')
        exit()

    tables = dfile.tables
    for tid, table in enumerate(tables):
        for rid, row in enumerate(table.rows):
            for cid, cell in enumerate(row.cells):
                text = cell.text.strip()
                if text != '' and text_map.get(text) == None:
                    key = f'{tid},{rid},{cid}'
                    template[key] = text
                    text_map[text] = key
                    column_name.append(text)
    return template, column_name


def word2excel(template_file, docx_dir, output_dir='output', omit_doc_name=False):
    template, column_name = get_template(template_file, omit_doc_name)
    wb = Workbook()
    ws = wb.active
    ws.append(column_name)

    for docx_file in tqdm(os.listdir(docx_dir)):
        file = os.path.join(docx_dir, docx_file)

        try:
            dfile = Document(file)
        except:
            print(f'打开文件 {docx_file} 失败!如果该文件在其他软件中打开，请关闭后再试!')
            exit()

        data = [] if omit_doc_name else [docx_file]
        tables = dfile.tables
        for tid, table in enumerate(tables):
            for rid, row in enumerate(table.rows):
                for cid, cell in enumerate(row.cells):
                    key = f'{tid},{rid},{cid}'
                    if template.get(key) != None:
                        data.append(cell.text.strip())
        ws.append(data)

    output_file = get_output_file(output_dir)
    wb.save(output_file)
