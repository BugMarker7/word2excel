# word2excel

[![GitHub](https://img.shields.io/badge/GitHub-BugMarker7-blue.svg)](https://github.com/BugMarker7/word2excel)

## 简介

`word2excel` 是一个用于从相同模板Word文档中批量提取数据，并生成Excel文件的工具。

## 功能特点

- 使用模板文件来定义数据提取规则，便于较好地通用性。
- 可同时提取Word文档中多个表格以及文本段落中的信息。
- 可在输出的Excel文件中添加Word文档的文件名作为标识。

## 安装

确保您的环境中已安装Python以及必要的依赖包。可以使用pip安装所需的依赖：

```bash
pip install -r requirements.txt
```

## 使用方法

通过命令行运行脚本，提供必要的参数：

```bash
python word2excel.py -t <template_path> -d <docx_directory> [-o <output_directory>] [--omit-doc-name]
```

其中：

- `-t` 或 `--template_file` 指定模板Word文件的路径。
- `-d` 或 `--docx_dir` 指定包含待处理Word文档的目录。
- `-o` 或 `--output_dir` 指定输出Excel文件的目录，默认为`output`。
- `--omit-doc-name` 如果使用该选项，则生成的Excel文件中不会包含Word文档的文件名列。
- `-v` 或 `--version` 显示程序版本信息。

## 示例

假设我们有以下文件结构：

```
project/
├── main_cli.py
├── template.docx
└── docx_files/
    ├── file1.docx
    ├── file2.docx
    └── ...
```

要将`docx_files`目录下的所有Word文档转换为Excel文件，可以执行：

```bash
python main_cli.py -t template.docx -d docx_files
```

## 参数说明

- `template_file`: 必须是一个有效的Word文档路径，用于定义数据提取的规则。
- `docx_dir`: 必须是一个存在的目录路径，包含需要处理的所有Word文档。
- `output_dir`: 输出目录，默认创建在当前工作目录下的`output`文件夹。
- `omit_doc_name`: 控制是否在Excel文件中加入源Word文档的名字，默认为False。

## 模板文件
模板文件是一个Word文档，用于定义数据提取的规则。根据内容的不同，分为文本段落和表格两种类型。

### 文本段落模板定义
文本段落模板的定义方式如下：
将要提取的文本内容用 {{ }} 包裹起来，**双括号中间填写Excel表列名**。例如，有以下内容：

```
姓名:    年龄:
```
其对应的模板定义为：
```
姓名:{{姓名}}    年龄:{{年龄}}
```

### 表格模板定义
表格模板的定义方式如下：
1. 清除表格全部内容（可以点击表格左上角选中表格，然后按快捷键`del`删除）

2. 将Excel表列名写入**值所在的单元格**，例如，有以下表格：

   | 姓名 |      | 年龄 |      |
   | ---- | ---- | ---- | ---- |

   那么模板定义为：

   |      | 姓名 |      | 年龄 |
   | ---- | ---- | ---- | ---- |


## 注意事项

- 请确保模板文件正确配置，以便程序能够正确解析Word文档中的数据。
- 脚本仅支持`.docx`格式的Word文档。
- **模板中表名不能重复**。

## 致谢

感谢我的朋友周周提出了这个需求，让我有完成这个代码的动力。感谢本项目使用到的依赖和库。如果您在使用过程中遇到任何问题或有任何改进建议，请随时提交Issue或Pull Request。
