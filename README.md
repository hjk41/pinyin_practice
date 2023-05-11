# pinyin_practice
看拼音写词练习生成器

输入一堆中文词，自动生成一个看拼音写词的word文件，省去家长自己抄拼音的麻烦

## 使用方法
1. 安装依赖包：
```bash
pip install pypinyin pydocx
```
安装word或者wps，并关联.docx文件，即双击一个.docx文件能够用word或者wps打开。

2. 下载本项目
```bash
git clone git@github.com:hjk41/pinyin_practice.git
```

3. 运行
```bash
python3 pinyin.py
```

本程序是一个GUI程序，用户只需要在文本框中输入所有想要进行练习的中文词，词之间使用空格分隔，然后按“转换”按钮，即可自动生成一个拼音练习。练习以template.docx为模板，每行有N个拼音（N可以在界面进行选择），生成后的文件会自动打开，用户可以选择打印或者对该文件进行手动修改和另存。
