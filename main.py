"""
Convert Excel files to PDF
"""
import pandas
import glob
from fpdf import FPDF
from pathlib import Path


# 将多个文件路径读入list:filepaths
filepaths = glob.glob("excels/*.xlsx")
# print(filepaths)

for filepath in filepaths:

    # 拿到excel文件名
    filename = Path(filepath).stem
    # 拿到文件名中的文件编号和文件日期
    excel_number, excel_date = filename.split("-")

    # 设置生成的pdf文件样式
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    # 添加pdf页面
    pdf.add_page()

    # 设置页眉
    pdf.set_font(family="Times", size=20, style="B")
    pdf.cell(w=0, h=10, txt=f"Excel number: {excel_number}", ln=1)
    pdf.cell(w=0, h=10, txt=f"Date: {excel_date}", ln=1)

    # 读入excel文件，convert excel to dataframe
    content = pandas.read_excel(filepath, sheet_name="Sheet 1")
    # content.columns 可以拿到第一行所有名称，是pandas 的Index类，转换为list
    columns = list(content.columns)
    # 去掉名称中间的下划线
    columns = [item.replace("_", ' ').title() for item in columns]

    # 设置table第一行
    pdf.set_font(family="Times", size=10, style="B")
    pdf.cell(w=30, h=8, txt=columns[0], border=1)
    pdf.cell(w=50, h=8, txt=columns[1], border=1)
    pdf.cell(w=40, h=8, txt=columns[2], border=1)
    pdf.cell(w=30, h=8, txt=columns[3], border=1)
    pdf.cell(w=30, h=8, txt=columns[4], border=1, ln=1)

    # 设置table剩余行
    for index, row in content.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.cell(w=30, h=8, txt=str(row['product_id']), border=1)
        pdf.cell(w=50, h=8, txt=str(row['product_name']), border=1)
        pdf.cell(w=40, h=8, txt=str(row['amount_purchased']), border=1)
        pdf.cell(w=30, h=8, txt=str(row['price_per_unit']), border=1)
        pdf.cell(w=30, h=8, txt=str(row['total_price']), border=1, ln=1)

    # 设置table最后一行总计
    total = content['total_price'].sum()
    pdf.cell(w=30, h=8, border=1)
    pdf.cell(w=50, h=8, border=1)
    pdf.cell(w=40, h=8, border=1)
    pdf.cell(w=30, h=8, border=1)
    pdf.cell(w=30, h=8, txt=str(total), border=1, ln=1)

    # 设置总计说明行
    pdf.set_font(family="Times", size=15, style="B")
    pdf.cell(w=30, h=8, txt=f"The total price is {total}.", ln=1)

    # 设置图标
    pdf.set_font(family="Times", size=15, style="B")
    pdf.cell(w=30, h=8, txt=f"Made by ZYR")
    pdf.image("images.jpg", w=5, h=8)

    # 生成pdf
    pdf.output(f"pdfs/{filename}.pdf")






