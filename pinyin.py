import tkinter as tk
from pypinyin import pinyin, Style
import re
import docx
from docx import Document
from docx.shared import Inches
import os

class PinyinConverter:
    def __init__(self, parent):
        self.parent = parent
        
        window = parent
        window.title("拼音练习")

        # Create the input text box
        self.input_entry = tk.Text(window, height=50, width=200)
        self.input_entry.grid(row=0, column=0, columnspan=3)

        # Create the select menu
        window.grid_columnconfigure(0)
        self.label = tk.Label(window, text="列数：")
        self.label.grid(row=1, column=0, sticky="w", padx=10)
        self.select_menu = tk.StringVar(window)
        self.select_menu.set("4")  # Set the default selected value
        options = ["3", "4", "5", "6"]
        window.grid_columnconfigure(1)
        self.select_dropdown = tk.OptionMenu(window, self.select_menu, *options).grid(row=1, column=1, sticky="w", padx=0)
        #self.select_dropdown.grid(row=1, column=0, sticky="e")

        # Create the '转换' button
        window.grid_columnconfigure(2, weight=30)
        self.convert_button = tk.Button(window, text="转换", command=self.convert, height=3, width=10)
        self.convert_button.grid(row=1, column=2, sticky="w", padx=100, pady=10)

        # Start the main event loop
        window.mainloop()

    '''
    Convert the input text into pinyin, retured as a list of strings
    '''
    def convert(self):
        # split input_text into chinese words
        input_text = self.input_entry.get("1.0", tk.END)
        chinese_words = re.findall(r'[\u4e00-\u9fff]+', input_text)
        # convert chinese words into pinyin
        pinyin_words = []
        for word in chinese_words:
            p = pinyin(word, style=Style.TONE, heteronym=True)
            pinyin_words.append(' '.join([i[0] for i in p]))
            self.pinyin_words = pinyin_words

        # open template.docx 
        document = Document('template.docx')
        # save as example.docx
        document.save('example.docx')
        # get the first table in the document
        table = document.tables[0]
        # get width of table
        width = 0
        for cell in table.rows[0].cells:
            width += cell.width
        # set number of rows and columns in the table
        ncol = int(self.select_menu.get())
        nrow = len(pinyin_words) // ncol + 1
        # Add a table to the document
        # Set the number of rows and columns
        actual_nrow = nrow * 3
        for i in range(1, ncol):
            table.add_column(width)
        for i in range(1, actual_nrow):
            table.add_row()
        # add pinyin words to the table
        for i in range(nrow):
            for j in range(ncol):
                if i * ncol + j < len(pinyin_words):
                    table.cell(i*3, j).text = pinyin_words[i * ncol + j]
                    table.cell(i*3 + 1, j).text = '_' * len(pinyin_words[i * ncol + j])
        # Save the document
        document.save('generated.docx')
        # Open the document using the default application
        os.system('start generated.docx')

# 创建主窗口
root = tk.Tk()

# 创建转换器实例
converter = PinyinConverter(root)

# 运行主循环
root.mainloop()
