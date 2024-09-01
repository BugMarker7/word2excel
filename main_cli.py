import os
import argparse
from utils.core import word2excel


if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        prog='word2excel',
        description='从Word文档中批量提取数据，并生成Excel文件',
        epilog='详细介绍及后续更新请参考：https://github.com/BugMarker7/word2excel'
    )
    parser.add_argument('-t', '--template_file',
                        type=str, help='模板Word文件路径')
    parser.add_argument('-d', '--docx_dir',
                        type=str, help='需要转换的Word文档所在目录')
    parser.add_argument('-o', '--output_dir', default='output',
                        type=str, help='保存生成excel文档的目录，默认在当前文件夹的"output"目录下.')
    parser.add_argument('--omit-doc-name', action='store_true',
                        help='Excel文件名是否包含一列Word文件名，默认为False')
    parser.add_argument('-v', '--version', action='version',
                        version='%(prog)s v0.1.0')
    args = parser.parse_args()

    assert os.path.isfile(
        args.template_file), f'模板文件{args.template_file}不存在或者不是一个文件!'
    assert os.path.isdir(args.docx_dir), 'Word文档所在目录不存在!'

    word2excel(template_file=args.template_file,
               docx_dir=args.docx_dir,
               output_dir=args.output_dir,
               omit_doc_name=args.omit_doc_name
               )
