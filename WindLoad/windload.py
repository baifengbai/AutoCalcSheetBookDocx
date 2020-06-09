from docx import Document
from docx.shared import Inches
from docx.shared import Pt
from docx.shared import Mm
from docx.shared import RGBColor
from docx.oxml.ns import qn
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import urllib.parse
import urllib.request
from PIL import Image
from time import strftime
from time import localtime
from datetime import datetime
from docx.enum.text import WD_LINE_SPACING
from math import ceil
from os import mkdir
from os.path import exists

print('======================================')
print('塔机抗风稳定校核计算书word生成器')
print('20200609 by 徐明')
print('适用工况：塔机随风自由旋转')
print('======================================')
print('计算书开始生成……')

path = 'images'
if not exists(path):
    mkdir(path)

# 插入公式图片函数，参数1公式字符串，参数2公式图片的名称，返回值为公式图片的原始英寸宽度
def add_image(latex, pngname):
    math = urllib.parse.quote(latex)
    query_url = 'http://latex.xuming.science/latex-image.php?math=' + math
    try:
        chart = urllib.request.urlopen(query_url)
        f = open(f'{path}/{pngname}.png', 'wb')
        f.write(chart.read())
        f.close()
        img = Image.open(f'{path}/{pngname}.png')
        return img.size[0] / 96
    except:
        print('无法连接公式服务器，请检查网络连接或向软件作者提交问题')
        return 0


print('文档格式初始化……')
# 创建文档对象
calc_book = Document()
# 设置正文字体
calc_book.styles['Normal'].font.name = 'Italic'
calc_book.styles['Normal'].element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
calc_book.styles['Normal'].font.size = Pt(12)
calc_book.styles['Normal'].font.color.rgb = RGBColor(0x00, 0x00, 0x00)
calc_book.styles['Normal'].paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
# 首行缩进，字符宽度等于字符高度，12pt=4.23mm, 1pt=0.3527mm
calc_book.styles['Normal'].paragraph_format.first_line_indent = Mm(8.46)
calc_book.styles['Normal'].paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
calc_book.styles['Normal'].paragraph_format.space_before = Mm(0)
calc_book.styles['Normal'].paragraph_format.space_after = Mm(0)
# 增加一个居中显示的样式
calc_book.styles['No Spacing'].paragraph_format.first_line_indent = Mm(0)
calc_book.styles['No Spacing'].paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
calc_book.styles['No Spacing'].paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
calc_book.styles['No Spacing'].paragraph_format.space_before = Mm(0)
calc_book.styles['No Spacing'].paragraph_format.space_after = Mm(0)
# 增加一个右对齐的样式
calc_book.styles['Quote'].paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
calc_book.styles['Quote'].font.name = 'Italic'
calc_book.styles['Quote'].element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
calc_book.styles['Quote'].font.size = Pt(12)
calc_book.styles['Quote'].font.italic = False
calc_book.styles['Quote'].font.color.rgb = RGBColor(0x00, 0x00, 0x00)
calc_book.styles['Quote'].paragraph_format.space_before = Mm(0)
calc_book.styles['Quote'].paragraph_format.space_after = Mm(0)
calc_book.styles['Quote'].paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
calc_book.styles['Quote'].paragraph_format.first_line_indent = Mm(0)
# 设置标题字体
calc_book.styles['Heading 1'].font.name = 'Times New Roman'
calc_book.styles['Heading 1'].element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
calc_book.styles['Heading 1'].element.rPr.rFonts.set(qn('w:asciiTheme'), 'Times New Roman')
calc_book.styles['Heading 1'].element.rPr.rFonts.set(qn('w:eastAsiaTheme'), '微软雅黑')
calc_book.styles['Heading 1'].element.rPr.rFonts.set(qn('w:hAnsiTheme'), 'Times New Roman')
calc_book.styles['Heading 1'].element.rPr.rFonts.set(qn('w:cstheme'), 'Times New Roman')
calc_book.styles['Heading 1'].font.size = Pt(22)
calc_book.styles['Heading 1'].font.color.rgb = RGBColor(0x00, 0x00, 0x00)
calc_book.styles['Heading 1'].paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
calc_book.styles['Heading 1'].paragraph_format.first_line_indent = Mm(0)
calc_book.styles['Heading 1'].paragraph_format.keep_with_next = True
calc_book.styles['Heading 1'].paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
calc_book.styles['Heading 1'].paragraph_format.space_before = Mm(0)
calc_book.styles['Heading 1'].paragraph_format.space_after = Mm(0)

'''
计算书文档属性设置
'''
print('文档属性初始化……')
# 文档标题
calc_book.core_properties.title = '计算书'
# 文档主题
calc_book.core_properties.subject = '计算书'
# 文档作者
calc_book.core_properties.author = 'xuming'
# 文档类别
calc_book.core_properties.category = 'calculation sheet'
# 文档注释
calc_book.core_properties.comments = 'Designed By Xuming. All Rights Reserved.'
# 文档创建时间
calc_book.core_properties.created = datetime.utcnow()
# 文档修改时间
calc_book.core_properties.modified = datetime.utcnow()

'''
计算书正文开始
'''

'''
参数定义
'''
# 非工作10m高处基本风压
pn = 1000
# 塔机最高处的计算高度
ht = 220
# 风压高度变化系数, 保留两位小数
kh = round(((((ht/10)**0.14)+0.4)/1.4)**2, 2)




'''
文档生成
'''