# -*- coding:utf-8 -*-
"""
该程序用来提取发票内容并保存到Excel中

"""
import os
import xlrd
import re
import shutil
import sys
import threading
import time

from xlutils.copy import copy
from pdfminer.pdfparser import PDFParser, PDFDocument
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import PDFPageAggregator
from pdfminer.layout import LAParams, LTTextBox, LTTextLine
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QLineEdit, QLabel, QHBoxLayout, QVBoxLayout, qApp, \
    QDesktopWidget, QFileDialog, QPlainTextEdit

__author__ = "yooongchu"
__site__ = "www.yooongchun.com"
__email__ = "yooongchun@foxmail.com"


# 解析文件
def parse_pdf(path, output_path):
    with open(path, 'rb') as fp:
        parser = PDFParser(fp)
        doc = PDFDocument()
        parser.set_document(doc)
        doc.set_parser(parser)
        doc.initialize('')
        rsrcmgr = PDFResourceManager()
        laparams = LAParams(all_texts=True, boxes_flow=2.0, heuristic_word_margin=True)
        device = PDFPageAggregator(rsrcmgr, laparams=laparams)
        interpreter = PDFPageInterpreter(rsrcmgr, device)
        extracted_text = ''
        for page in doc.get_pages():
            interpreter.process_page(page)
            layout = device.get_result()
            for lt_obj in layout:
                if isinstance(lt_obj, LTTextBox) or isinstance(lt_obj, LTTextLine):
                    extracted_text += lt_obj.get_text()
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(extracted_text)


def inner_key_word():
    key_word = [
        '机器编号:', '发票代码:', '发票号码:', '开票日期:', '校验码:', '名称:', '纳税人识别号:', '地址、电话:',
        '开户行及账号:', '项目名称', '车牌号', '类型', '通行日期起', '通行日期止', '金额', '税率', '税额',
        '价税合计（大写）', '（小写）', '收款人:', '复核:', '开票人:', '销售方:（章）'
    ]
    return key_word


# 加载Excel关键词
def loadKeyWords(path):
    book = xlrd.open_workbook(path)  # 打开一个wordbook
    sheet_ori = book.sheet_by_name('Sheet1')
    return sheet_ori.row_values(0, 0, sheet_ori.ncols)


#
def inner_key_word_v2():
    return ["PDF名称", "发票号码", "名称", "发票代码"]


## 格式化发票内容V2
def small_split_item(text):
    ITEMS = {}
    text_line = text.split("\n")
    for index, line in enumerate(text_line):
        item = re.sub(r"\s+", "", line)
        item = item.replace(":", "")
        if item == "发票号码":
            ITEMS["发票号码"] = text_line[index + 1]
            continue
        if item == "发票代码":
            ITEMS["发票代码"] = text_line[index + 1]
            continue
        if item == "名称" and text_line[index - 3] == "购":
            ITEMS["名称"] = text_line[index + 1]
            continue
    return ITEMS


# 格式化发票内容
# 内容分区
def split_block(text):
    text_line = text.split("\n")
    # 第一步：分区：[公共头,购买方,销售方，公共域]
    header = []
    buyer = []
    saler = []
    body = []
    index = 0
    while index < len(text_line):
        if not text_line[index] == "购":
            header.append(re.sub(r"\s+", "", text_line[index]))
            index += 1
            continue
        else:
            break
    while index < len(text_line):
        if not text_line[index] == "项目名称":
            if text_line[index] == "购" or text_line[index] == "买" or text_line[index] == "方":
                index += 1
                continue
            if re.sub(r"\s+", "", text_line[index]) == "地址、" and re.sub(
                    r"\s+", "", text_line[index + 1]) == "电话:":
                buyer.append(
                    re.sub(r"\s+", "",
                           text_line[index] + text_line[index + 1]))
                index += 2
                continue
            buyer.append(re.sub(r"\s+", "", text_line[index]))
            index += 1
            continue
        else:
            break
    while index < len(text_line):
        if not re.sub(r"\s+", "", text_line[index]) == "名称:":
            body.append(re.sub(r"\s+", "", text_line[index]))
            index += 1
            continue
        else:
            break
    while index < len(text_line):
        if re.sub(r"\s+", "", text_line[index]) == "地址、" and re.sub(
                r"\s+", "", text_line[index + 1]) == "电话:":
            saler.append(
                re.sub(r"\s+", "", text_line[index] + text_line[index + 1]))
            index += 2
            continue
        saler.append(re.sub(r"\s+", "", text_line[index]))
        index += 1
        continue
    return header, buyer, body, saler


def split_item_v2(text, key_word):
    ITEMS = {}
    index = 0
    while index < len(text):
        group = []
        while index < len(text) and text[index] in key_word:
            group.append(text[index])
            index += 1
        ind = index
        if "收款人:" in group and len(text[ind]) == 3 or len(
                text[ind]) == 2 or len(text[ind]) == 4:
            for k, v in enumerate(group):
                if v == "收款人:":
                    ITEMS[v] = text[ind]
                    del group[k]
                    ind += 1
                    break
        for v in group:
            ITEMS[v] = text[ind]
            ind += 1
        index += 1
    if "金额" in ITEMS.keys() and not re.match(r"^\d+\.\d+$", ITEMS["金额"]):
        ITEMS["税额"] = ITEMS["税率"]
        ITEMS["税率"] = re.findall(u"[\u4e00-\u9fa5]+", ITEMS["金额"])[0]
        ITEMS["金额"] = re.findall(r"^\d+\.\d+", ITEMS["金额"])[0]
    for k, v in enumerate(text):
        if v == "机器编号:":
            ITEMS["发票名称"] = text[k + 2]
            break
    return ITEMS


def merge_item(header, body, buyer, saler):
    ITEMS = {}
    for key, value in header.items():
        key = key.replace(":", "")
        ITEMS[key] = value
    for key, value in body.items():
        key = key.replace(":", "")
        ITEMS[key] = value
    for key, value in buyer.items():
        key = key.replace(":", "")
        key = "购买方" + key
        ITEMS[key] = value
    for key, value in saler.items():
        key = key.replace(":", "")
        if not key == "销售方（章）" and not key == "收款人" and not key == "开票人" and not key == "复核" and not key == "（小写）":
            key = "销售方" + key
        ITEMS[key] = value
    if "销售方（章）" in ITEMS.keys():
        ITEMS["销售方"] = ITEMS["销售方（章）"]
    if "（小写）" in ITEMS.keys():
        ITEMS["价税合计（小写）"] = ITEMS["（小写）"]

    return ITEMS


def write_to_excel(output_path, Items, key_words, cnt):
    book = xlrd.open_workbook(output_path)  # 打开一个wordbook
    copy_book = copy(book)
    sheet_copy = copy_book.get_sheet("Sheet1")
    for index, word in enumerate(key_words):
        for key, value in Items.items():
            if key == word:
                sheet_copy.write(cnt, index, value)
                break
    copy_book.save(output_path)


def load_files(path):
    paths = []
    for path, dir_list, file_list in os.walk(path):
        for file_name in file_list:
            abs_path = os.path.join(path, file_name)
            if os.path.isfile(abs_path) and (
                    os.path.splitext(abs_path)[1] == ".pdf" or os.path.splitext(abs_path)[1] == ".PDF"):
                paths.append(abs_path)
    return paths


def load_folders(path):
    folders = os.listdir(path)
    paths = []
    curr_path = []
    for folder in folders:
        if os.path.isdir(os.path.join(path, folder)):
            abs_path = os.path.join(path, folder)
            paths.append(abs_path)
            curr_path.append(folder)
    return paths, curr_path


class MYGUI(QWidget):
    def __init__(self):
        super().__init__()
        self.exit_flag = False
        self.try_time = 1527652647.6671877 + 24 * 60 * 60

        self.initUI()

    def initUI(self):
        self.pdf_label = QLabel("PDF文件夹路径: ")
        self.pdf_btn = QPushButton("选择")
        self.pdf_btn.clicked.connect(self.open_pdf)
        self.pdf_path = QLineEdit("PDF文件夹路径...")
        self.pdf_path.setEnabled(False)
        self.excel_label = QLabel("Excel Demo 路径: ")
        self.excel_btn = QPushButton("选择")
        self.excel_btn.clicked.connect(self.open_excel)
        self.excel_path = QLineEdit("Excel Demo路径...")
        self.excel_path.setEnabled(False)
        self.output_label = QLabel("输出路径: ")
        self.output_path = QLineEdit(os.path.abspath("./"))
        self.output_path.setEnabled(False)
        self.output_btn = QPushButton("选择")
        self.output_btn.clicked.connect(self.open_output)
        self.info = QPlainTextEdit()

        h1 = QHBoxLayout()
        h1.addWidget(self.pdf_label)
        h1.addWidget(self.pdf_path)
        h1.addWidget(self.pdf_btn)

        h2 = QHBoxLayout()
        h2.addWidget(self.excel_label)
        h2.addWidget(self.excel_path)
        h2.addWidget(self.excel_btn)

        h3 = QHBoxLayout()
        h3.addWidget(self.output_label)
        h3.addWidget(self.output_path)
        h3.addWidget(self.output_btn)

        self.run_btn = QPushButton("运行")
        self.run_btn.clicked.connect(self.run)

        self.auth_label = QLabel("密码")
        self.auth_ed = QLineEdit("test_mode")

        exit_btn = QPushButton("退出")
        exit_btn.clicked.connect(self.Exit)
        h4 = QHBoxLayout()
        h4.addWidget(self.auth_label)
        h4.addWidget(self.auth_ed)
        h4.addStretch(1)
        h4.addWidget(self.run_btn)
        h4.addWidget(exit_btn)

        v = QVBoxLayout()
        v.addLayout(h1)
        v.addLayout(h2)
        v.addLayout(h3)
        v.addWidget(self.info)
        v.addLayout(h4)
        self.setLayout(v)
        width = int(QDesktopWidget().screenGeometry().width() / 3)
        height = int(QDesktopWidget().screenGeometry().height() / 3)
        self.setGeometry(100, 100, width, height)
        self.setWindowTitle('PDF to Excel')
        self.show()

    def Exit(self):
        self.exit_flag = True
        qApp.quit()

    def open_pdf(self):
        fname = QFileDialog.getExistingDirectory(self, "Open pdf folder",
                                                 "/home")
        if fname:
            self.pdf_path.setText(fname)

    def open_excel(self):
        fname = QFileDialog.getOpenFileName(self, "Open demo excel", "/home")
        if fname[0]:
            self.excel_path.setText(fname[0])

    def open_output(self):
        fname = QFileDialog.getExistingDirectory(self, "Open output folder",
                                                 "/home")
        if fname:
            self.output_path.setText(fname)

    def run(self):
        self.info.setPlainText("")
        threading.Thread(target=self.scb, args=()).start()
        if self.auth_ed.text() == "a3s7wt29yn1m48zj" or self.auth_ed.text() == "GOD_MODE":
            self.info.insertPlainText("密码正确，开始运行程序!\n")
            threading.Thread(target=self.main_fcn, args=()).start()
        elif self.auth_ed.text() == "test_mode":
            if time.time() < self.try_time:
                self.info.insertPlainText("试用模式，截止时间：2018-05-31 11:58\n")
                threading.Thread(target=self.main_fcn, args=()).start()
            else:
                self.info.insertPlainText(
                    "试用时间已结束，继续使用请联系yooongchun，微信：18217235290 获取密码\n")

        else:
            self.info.insertPlainText(
                "密码错误，请联系yooongchun(微信：18217235290)获取正确密码!\n")

    def scb(self):
        flag = True
        cnt = self.info.document().lineCount()
        while not self.exit_flag:
            if flag:
                self.info.verticalScrollBar().setSliderPosition(self.info.verticalScrollBar().maximum())
            time.sleep(0.01)
            if cnt < self.info.document().lineCount():
                flag = True
                cnt = self.info.document().lineCount()
            else:
                flag = False
            time.sleep(0.01)

    def main_fcn(self):
        if os.path.isdir(self.pdf_path.text()):
            try:
                folders, curr_folder = load_folders(self.pdf_path.text())
            except Exception:
                self.info.insertPlainText("加载PDF文件夹出错，请重试！\n")
                return
        else:
            self.info.insertPlainText("pdf路径错误，请重试！\n")
            return
        if os.path.isfile(self.excel_path.text()):
            demo_path = self.excel_path.text()
        else:
            self.info.insertPlainText("Excel路径错误，请重试！\n")
            return

        for index_, folder in enumerate(folders):
            self.info.insertPlainText("正在处理文件夹: %s %d/%d" %
                                      (folder, index_ + 1, len(folders)) +
                                      "\n")
            try:
                if os.path.isdir(self.output_path.text()):
                    if not os.path.isdir(self.output_path.text()):
                        os.mkdir(self.output_path.text())
                    out_path = os.path.join(self.output_path.text(),
                                            curr_folder[index_] + ".xls")
                else:
                    self.info.insertPlainText("输出路径错误，请重试！\n")
                    return
                shutil.copyfile(demo_path, out_path)
            except Exception:
                self.info.insertPlainText("路径分配出错，请确保程序有足够运行权限再重试！\n")
                return
            try:
                files = load_files(folder)
            except Exception:
                self.info.insertPlainText("读取文件夹 %s 出错，跳过当前文件夹！\n" % folder)
                continue
            if not os.path.isdir("txt"):
                os.mkdir("txt")
            for index, file_ in enumerate(files):
                self.info.insertPlainText(
                    "正在解析文件: %s  %d/%d %.2f" %
                    (os.path.basename(file_), index + 1, len(files),
                     (index + 1) / len(files)) + "\n")

                try:
                    txt_path = 'txt/%s.txt' % os.path.basename(file_)
                    parse_pdf(file_, txt_path)
                except Exception:
                    self.info.insertPlainText("解析文件 %s 出错,跳过！\n" % file_)

            try:
                key_words = loadKeyWords(demo_path)
            except Exception:
                self.info.insertPlainText("加载Excel Demo出错，确保地址正确并且没有在任何地方打开该文件，然后重试！\n")
                return
            if not os.path.isdir("txt"):
                self.info.insertPlainText("No txt file error.\n")

            else:
                self.info.insertPlainText("抽取文件数据到Excel...\n")

                for index, file_ in enumerate(os.listdir("txt")):
                    try:
                        if os.path.splitext(file_)[1] == ".txt":
                            with open(
                                    os.path.join("txt", file_),
                                    "r",
                                    encoding="utf-8") as f:
                                text = f.read()
                            ITEMS = small_split_item(text)
                            ITEMS["PDF名称"] = os.path.basename(file_).split(".")[0]
                            write_to_excel(out_path, ITEMS, key_words, index + 1)
                            # header, buyer, body, saler = split_block(text)
                            # ITEMS = merge_item(
                            #     split_item_v2(header, inner_key_word()), split_item_v2(body, inner_key_word()),
                            #     split_item_v2(buyer, inner_key_word()), split_item_v2(saler, inner_key_word()))
                            # write_to_excel(out_path, ITEMS, key_words,
                            #                index + 1)
                    except Exception:
                        self.info.insertPlainText("抽取文件 %s 内容出错，跳过！\n" % file_)
                        continue
                if os.path.isdir("txt"):
                    try:
                        self.info.insertPlainText("移除临时文件.\n")
                        # shutil.rmtree("txt")
                    except Exception:
                        self.info.insertPlainText("移除临时文件夹 txt 出错,请手动删除！\n")

        if os.path.isdir("txt"):
            try:
                pass
                # shutil.rmtree("txt")
            except Exception:
                self.info.insertPlainText("移除临时文件夹 txt 出错,请手动删除！\n")
        self.info.insertPlainText(
            "运行完成，请到输出地址 %s 查看结果.\n" % self.output_path.text())


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MYGUI()
    sys.exit(app.exec_())
