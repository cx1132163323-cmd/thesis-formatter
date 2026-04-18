# 硕士论文格式助手

上传 Word 原稿，自动整理论文格式，几秒钟完成，下载即用。

## 功能

- A4 页面与版心边距
- 正文 Times New Roman 12pt，1.25 倍行距
- 一至四级标题字体与对齐方式
- 图题 / 表题样式与自动编号
- 三线表边框（1.5 pt / 0.75 pt）
- 自动目录 / 图表目录字段
- 奇偶页眉（论文题目 / 章节名）
- 居中页码页脚

## 在线使用

> 部署后在此填写链接

## 本地运行

```bash
git clone https://github.com/cx1132163323-cmd/thesis-formatter.git
cd thesis-formatter
pip install -r requirements.txt
streamlit run app.py
```

浏览器会自动打开，上传 `.docx` 文件即可。

## 使用说明

1. 上传 Word 原稿（`.docx` 格式）
2. 可选填写论文标题，留空则自动识别
3. 点击「开始格式化」
4. 下载处理后的文件
5. 用 Word 打开，按 `Ctrl + A` 再按 `F9`，刷新目录页码

## 依赖

- Python 3.8+
- [python-docx](https://python-docx.readthedocs.io/)
- [Streamlit](https://streamlit.io/)
