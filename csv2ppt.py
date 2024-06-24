import pandas as pd
from pptx import Presentation
import csv
import copy
import six
import sys


# 读取 CSV 文件
csv_file = 'data.csv'  # 请替换为你的 CSV 文件路径
df = pd.read_csv(csv_file)

input_file = csv.DictReader(open(csv_file))

# 加载 PowerPoint 文件
ppt_template = 'template.pptx'  # 请替换为你的 PowerPoint 模板文件路径
presentation = Presentation(ppt_template)

# 假设我们只处理第一张幻灯片
slide = presentation.slides[0]

# 定义一个函数，将数据填入幻灯片
def fill_slide(slide, data):

    slide_title = find_shape_by_name(slide.shapes,'index')
    add_text(slide_title,data['序号'])
    slide_title = find_shape_by_name(slide.shapes,'hiragana')
    add_text(slide_title,data['日文音标'])
    slide_title = find_shape_by_name(slide.shapes,'kanji')
    add_text(slide_title,data['日文单词'])
         
         
# 复制幻灯片
def duplicate_slide(pres, index):
    template = pres.slides[index]
    try:
        blank_slide_layout = pres.slide_layouts[5]
    except:
        blank_slide_layout = pres.slide_layouts[len(pres.slide_layouts)]

    copied_slide = pres.slides.add_slide(blank_slide_layout)

    for shp in template.shapes:
        el = shp.element
        newel = copy.deepcopy(el)
        copied_slide.shapes._spTree.insert_element_before(newel, 'p:extLst')

    for _, value in six.iteritems(template.part.rels):
        # Make sure we don't copy a notesSlide relation as that won't exist
        if "notesSlide" not in value.reltype:
            copied_slide.part.rels._add_relationship(
                value.reltype,
                value._target,
                value.rId
            )

    return copied_slide
    
def find_shape_by_name(shapes, name):
    for shape in shapes:
        if shape.name == name:
            return shape
    return None

def add_text(shape, text, alignment=None):
    
    if alignment:
        shape.vertical_anchor = alignment

    tf = shape.text_frame
    tf.clear()
    run = tf.paragraphs[0].add_run()
    run.text = text if text else ''
    
    
# 遍历 CSV 文件中的每条记录，创建对应的幻灯片
for row in input_file:
    # 从模板创建一张新的幻灯片
    new_slide = duplicate_slide(presentation, 0)
    # 填充文本
    fill_slide(new_slide, row)
    
    

# 保存修改后的幻灯片
output_ppt = 'output.pptx'  # 输出文件路径
presentation.save(output_ppt)

print(f'幻灯片已保存为 {output_ppt}')

