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
from math import cos
from math import sin
from math import radians

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
calc_book.styles['Normal'].font.name = 'Times New Roman'
calc_book.styles['Normal'].element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
calc_book.styles['Normal'].font.size = Pt(12)
calc_book.styles['Normal'].font.color.rgb = RGBColor(0x00, 0x00, 0x00)
calc_book.styles['Normal'].paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
# 首行缩进，字符宽度等于字符高度，12pt=4.23mm, 1pt=0.3527mm
font_height = 4.23
calc_book.styles['Normal'].paragraph_format.first_line_indent = Mm(0)
calc_book.styles['Normal'].paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
calc_book.styles['Normal'].paragraph_format.space_before = Mm(0)
calc_book.styles['Normal'].paragraph_format.space_after = Mm(0)
# 增加一个居中显示的样式
calc_book.styles['No Spacing'].paragraph_format.first_line_indent = Mm(0)
calc_book.styles['No Spacing'].paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
calc_book.styles['No Spacing'].paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
calc_book.styles['No Spacing'].paragraph_format.space_before = Mm(0)
calc_book.styles['No Spacing'].paragraph_format.space_after = Mm(0)
# 增加一个左对齐的样式
calc_book.styles['Quote'].paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
calc_book.styles['Quote'].font.name = 'Times New Roman'
calc_book.styles['Quote'].element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
calc_book.styles['Quote'].font.size = Pt(10)
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
# 设置标题2字体
calc_book.styles['Heading 2'].font.name = 'Times New Roman'
calc_book.styles['Heading 2'].element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
calc_book.styles['Heading 2'].element.rPr.rFonts.set(qn('w:asciiTheme'), 'Times New Roman')
calc_book.styles['Heading 2'].element.rPr.rFonts.set(qn('w:eastAsiaTheme'), '微软雅黑')
calc_book.styles['Heading 2'].element.rPr.rFonts.set(qn('w:hAnsiTheme'), 'Times New Roman')
calc_book.styles['Heading 2'].element.rPr.rFonts.set(qn('w:cstheme'), 'Times New Roman')
calc_book.styles['Heading 2'].font.size = Pt(18)
calc_book.styles['Heading 2'].font.color.rgb = RGBColor(0x00, 0x00, 0x00)
calc_book.styles['Heading 2'].paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
calc_book.styles['Heading 2'].paragraph_format.first_line_indent = Mm(0)
calc_book.styles['Heading 2'].paragraph_format.keep_with_next = True
calc_book.styles['Heading 2'].paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
calc_book.styles['Heading 2'].paragraph_format.space_before = Mm(0)
calc_book.styles['Heading 2'].paragraph_format.space_after = Mm(0)
# 设置标题3字体
calc_book.styles['Heading 3'].font.name = 'Times New Roman'
calc_book.styles['Heading 3'].element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
calc_book.styles['Heading 3'].element.rPr.rFonts.set(qn('w:asciiTheme'), 'Times New Roman')
calc_book.styles['Heading 3'].element.rPr.rFonts.set(qn('w:eastAsiaTheme'), '微软雅黑')
calc_book.styles['Heading 3'].element.rPr.rFonts.set(qn('w:hAnsiTheme'), 'Times New Roman')
calc_book.styles['Heading 3'].element.rPr.rFonts.set(qn('w:cstheme'), 'Times New Roman')
calc_book.styles['Heading 3'].font.size = Pt(14)
calc_book.styles['Heading 3'].font.color.rgb = RGBColor(0x00, 0x00, 0x00)
calc_book.styles['Heading 3'].paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
calc_book.styles['Heading 3'].paragraph_format.first_line_indent = Mm(0)
calc_book.styles['Heading 3'].paragraph_format.keep_with_next = True
calc_book.styles['Heading 3'].paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
calc_book.styles['Heading 3'].paragraph_format.space_before = Mm(0)
calc_book.styles['Heading 3'].paragraph_format.space_after = Mm(0)
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
# 计算书名称
jobname = '大疆ZSL1150抗风计算书'
# 塔机型号
tower_model = 'ZSL1150'
# 非工作10m高处基本风压
# pn = 1000
# 塔机最高处的计算高度
# ht = 220
# 风压高度变化系数, 保留两位小数
# kh = round(((((ht / 10) ** 0.14) + 0.4) / 1.4) ** 2, 2)

# 抗台风计算风压指定值
# 深圳防台风技术规程5.1.3中最大风压1870Pa
pnh = 1870
# 吊臂长度
beam_len = 58.5
# 非工作吊臂仰角
beam_ang = 60
# 塔机总力矩
tower_m = 1488
# 塔机后倾平衡力矩
tower_back = - ceil(tower_m * 0.45)
# 吊臂0度力矩
beam_0 = 702
# 吊臂非工作状态前倾力矩
beam_m = ceil(beam_0 * cos(radians(beam_ang)))
# 钩头重量
hook_mass = 1.3
# 钩头产生的前倾力矩
hook_m = ceil(hook_mass * beam_len * cos(radians(beam_ang)))
# 非工作状态下，上部结构不平衡
m0 = tower_back + beam_m + hook_m

# 吊臂
beam_d1 = 133  # 主弦直径mm
beam_d2 = 70  # 腹杆直径mm
beam_b = 2.25  # 主弦中心距m
beam_fg_len = 1.27 * beam_len * 2  # 腹杆长度总和
beam_a = beam_len * 2 * beam_d1 * 0.001 + beam_fg_len * beam_d2 * 0.001  # 特征面积
beam_phi = beam_a / (beam_len * (beam_b + beam_d1 * 0.001) * sin(radians(beam_ang)))   # 充实率
wind_v = (pnh / 0.625)**0.5   # 通过风压反推风速
re = 0.667 * wind_v * beam_d1 * 0.001  # 单位10^5

'''
文档生成
'''
calc_book.add_heading('2.1 无风状态塔机上部结构不平衡力矩', level=2)
calc_book.add_paragraph(f'吊臂长度：{beam_len}m', style='Normal')
calc_book.add_paragraph(f'非工作状态吊臂仰角：{beam_ang}°', style='Normal')
calc_book.add_paragraph(f'塔机后倾平衡力矩：{tower_back}t.m', style='Normal')
calc_book.add_paragraph(f'吊臂{beam_ang}°产生的前倾力矩：{beam_m}t.m', style='Normal')
calc_book.add_paragraph(f'钩头产生的前倾力矩：{hook_m}t.m', style='Normal')
calc_book.add_paragraph(f'非工作状态下塔机上部结构总不平衡力矩：M={tower_back}t.m+{beam_m}t.m+{hook_m}t.m={m0}t.m', style='Normal')
calc_book.add_paragraph('*为便于计算，规定由回转中心朝吊臂方向弯矩为正，由回转中心朝配重块方向弯矩为负。', style='Quote')

calc_book.add_heading('2.2 非工作状态风载荷', level=2)
mathtemp = r'F_{WN} = p_n(h) \times C \times A'
width = add_image(mathtemp, 'fwn')
calc_book.add_paragraph('', style='No Spacing').add_run('').add_picture(f'{path}/fwn.png', width=Inches(width))
calc_book.add_paragraph('式中：', style='Normal')

para1 = calc_book.add_paragraph('', style='Normal')
mathtemp = r'F_{WN}'
add_image(mathtemp, 'fwn1')
para1.add_run('').add_picture(f'{path}/fwn1.png', height=Mm(font_height))  # Inches(width))
para1.add_run('——非工作状态垂直作用在所指构件纵轴线上的风载荷，单位为牛顿')
mathtemp = r'(N)'
add_image(mathtemp, 'N')
para1.add_run('').add_picture(f'{path}/N.png', height=Mm(font_height))  # Inches(width))

para1 = calc_book.add_paragraph('', style='Normal')
mathtemp = r'p_n(h)'
add_image(mathtemp, 'pnh')
para1.add_run('').add_picture(f'{path}/pnh.png', height=Mm(font_height))  # Inches(width))
para1.add_run('——高度h处的非工作状态计算风压，单位为牛顿每平方米')
mathtemp = r'(N/m^2)'
add_image(mathtemp, 'Nm2')
para1.add_run('').add_picture(f'{path}/Nm2.png', height=Mm(font_height))  # Inches(width))

para1 = calc_book.add_paragraph('', style='Normal')
mathtemp = r'C'
add_image(mathtemp, 'C')
para1.add_run('').add_picture(f'{path}/C.png', height=Mm(font_height))
para1.add_run('——所指构件的空气动力系数，与构件的特征面积A一起使用')

para1 = calc_book.add_paragraph('', style='Normal')
mathtemp = r'A'
add_image(mathtemp, 'A')
para1.add_run('').add_picture(f'{path}/A.png', height=Mm(font_height))
para1.add_run('——所指构件的特征面积，单位为平方米')
mathtemp = r'(m^2)'
add_image(mathtemp, 'm2')
para1.add_run('').add_picture(f'{path}/m2.png', height=Mm(font_height))  # Inches(width))

calc_book.add_heading('2.3 非工作状态计算风压', level=2)
calc_book.add_paragraph(f'{tower_model}塔机非工作状态计算风压取深圳防台风技术规程5.1.3中最大风压1870Pa。', style='Normal')
calc_book.add_paragraph('因此，非工作状态计算风压：', style='Normal')
mathtemp = r'p_n(h) = 1870 Pa'
width = add_image(mathtemp, 'pnh1870')
calc_book.add_paragraph('', style='No Spacing').add_run('').add_picture(f'{path}/pnh1870.png', width=Inches(width))

calc_book.add_heading('2.4 吊臂风载荷计算', level=2)






'''
计算书结束，输出docx文档
'''
filename = f'{jobname}' + strftime("%Y-%m-%d-%H%M%S", localtime())
calc_book.save(f'{filename}.docx')
print(f'计算书生成结束，保存在程序目录下，文件名为{filename}.docx')
