# -*- coding: utf-8 -*-
"""docx 转 pdf 模块：调用本机 Word（COM / pythoncom 底层）把 docx 转成 pdf。

依赖：
- 电脑已安装 Microsoft Word
- .venv 已安装 pywin32（pip install pywin32）

用法（在 01 里）：
    from docx_to_pdf import convert_docx_to_pdf
    pdf_path = convert_docx_to_pdf(report_path)
"""

import os
import pythoncom
import win32com.client


def convert_docx_to_pdf(docx_path, pdf_path=None):
    """把 docx 转成 pdf，返回 pdf 的完整路径。

    pdf_path 为空时，自动生成与 docx 同名、后缀为 .pdf 的文件（同目录）。
    """
    docx_path = os.path.abspath(docx_path)
    if not os.path.exists(docx_path):
        raise FileNotFoundError(f"找不到 docx 文件：{docx_path}")

    if pdf_path is None:
        pdf_path = os.path.splitext(docx_path)[0] + ".pdf"
    pdf_path = os.path.abspath(pdf_path)

    # Streamlit 在后台线程运行，必须先在当前线程初始化 COM，否则报
    # "尚未调用 CoInitialize"（重复调用是安全的）
    pythoncom.CoInitialize()
    word = None
    try:
        # 启动 Word（后台不可见）
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False

        doc = word.Documents.Open(docx_path)
        try:
            # FileFormat=17 即 wdFormatPDF，把文档另存为 PDF
            doc.SaveAs(pdf_path, FileFormat=17)
        finally:
            doc.Close(False)  # False = 不保存修改
    finally:
        if word is not None:
            word.Quit()  # 退出 Word，避免后台进程残留
        pythoncom.CoUninitialize()  # 释放当前线程的 COM

    return pdf_path


if __name__ == "__main__":
    # 自己单独测试时，把下面改成你文件夹里真实存在的 docx 路径
    test_docx = r"C:\Users\王张含\Desktop\daima\prompt\学习规划_20260925_180859.docx"
    print("生成的 PDF：", convert_docx_to_pdf(test_docx))
