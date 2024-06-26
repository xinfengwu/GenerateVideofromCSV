import os
import gtts
from pptx import Presentation
import csv
import copy
from pptx.util import Pt
import ppt2img

import aspose.slides as slides
import aspose.pydrawing as drawing


# 创建文件夹
def create_folder(folder_name):
    # 构建文件夹路径
    folder_path = os.path.join(os.getcwd(), folder_name)
    
    # 检查目录是否存在
    if not os.path.exists(folder_path):
        # 如果不存在，创建目录
        os.makedirs(folder_path)
        print(f"文件夹 '{folder_path}' 已创建！")
    else:
        print(f"文件夹 '{folder_path}' 已存在！")
        empty_folder(folder_path)
        print(f"文件夹 '{folder_path}' 已清空！")
    return folder_path

# 删除空文件
def empty_folder(folder_path):
    # 遍历文件夹中的所有文件
    for file_name in os.listdir(folder_path):
        file_path = os.path.join(folder_path, file_name)

        # 如果是空文件，则删除
        if os.path.isfile(file_path) and os.path.getsize(file_path) == 0: # 空文件
            os.remove(file_path)
        # 如果是文件夹，则递归调用 empty_folder 函数清空该文件夹
        elif os.path.isdir(file_path):
            empty_folder(file_path)

# 按行生成语音
def text_to_mp3(line,output_folder_path):
    # print(gtts.lang.tts_langs()) 输出支持的语言
    try:
        tts = gtts.gTTS(line.strip(), lang='ja')  ##  request google to get synthesis
        output_file_path = output_folder_path +'/'+ line.strip()+'.mp3'
        if not os.path.exists(output_file_path):
            tts.save(output_file_path)  ##  save audio
    except Exception as e:
        print(f"转换 '{line.strip()}' 出现错误：{e}")

# 定义一个函数，将数据填入幻灯片
def fill_slide(slide, data):

    slide_title = find_shape_by_name(slide.shapes,'index')
    add_text(slide_title,data['index'])
    slide_title = find_shape_by_name(slide.shapes,'hiragana')
    add_text(slide_title,data['phonetic symbol'])
    slide_title = find_shape_by_name(slide.shapes,'kanji')
    add_text(slide_title,data['word'])

# 复制幻灯片
def duplicate_slide(prs, index):
    slide_to_copy= prs.slides[index]
    # Create a new slide object
    new_slide = prs.slides.add_slide(slide_to_copy.slide_layout)
    # Copy content from the original slide to the new slide
    for shape in slide_to_copy.shapes:
        el = shape.element
        newel = copy.deepcopy(el)
        new_slide.shapes._spTree.insert_element_before(newel, 'p:extLst')

    return new_slide

# 删除幻灯片
def delete_slide(presentation, slide_index):
    xml_slides = presentation.slides._sldIdLst  # Access the slide list XML elements
    if slide_index < 0 or slide_index >= len(xml_slides):
        print(f"Slide index {slide_index} out of range.")
        return
    slide_id = xml_slides[slide_index].rId
    presentation.slides._sldIdLst.remove(xml_slides[slide_index])  # Remove slide from XML elements
    presentation.part.drop_rel(slide_id)  # Remove slide from related parts


# 查找文本框
def find_shape_by_name(shapes, name):
    for shape in shapes:
        if shape.name == name:
            return shape
    return None


# 修改文字
def add_text(shape, text, alignment=None):

    if alignment:
        shape.vertical_anchor = alignment

    tf = shape.text_frame
    tf.clear()
    run = tf.paragraphs[0].add_run()
    run.text = text if text else ''
    font=run.font
    font.name = 'Calibri'
    font.size = Pt(58)
    font.bold = True

# 将PPTX导出为图片
def ppt_to_jpg(ppt,img_path):
    pres = slides.Presentation(ppt)
    for sld in pres.slides:
        bmp = sld.get_thumbnail(1, 1)
        bmp.save(img_path +'/'+ "Slide_{num}.jpg".format(num=str(sld.slide_number)), drawing.imaging.ImageFormat.jpeg)

    print(f'幻灯片已导出为图片 ')

