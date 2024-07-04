import os
import gtts
from pptx import Presentation
import csv
import copy
import subprocess
import tempfile
from pdf2image import convert_from_path
from pdf2image.exceptions import (
    PDFInfoNotInstalledError,
    PDFPageCountError,
    PDFSyntaxError
)
import subprocess
from PIL import Image
from pptx.util import Pt
import time

# 读取csv文件

def read_csv_data(file_path):
    data = []
    with open(file_path, newline='') as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            data.append(row)
    return data


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
def text_to_mp3(row,output_folder_path):
    # print(gtts.lang.tts_langs()) 输出支持的语言
    try:
        tts = gtts.gTTS(row['kanji'].strip(), lang='ja')  ##  request google to get synthesis
        output_file_path = output_folder_path +'/'+ row['index'].strip()+'.mp3'
        if not os.path.exists(output_file_path):
            tts.save(output_file_path)  ##  save audio
    except Exception as e:
        
        # line = row['kanji']
        print(f"转换 '{row['kanji'].strip()}' 出现错误：{e}")
    
# 定义一个函数，将数据填入幻灯片
def fill_slide(slide, row):
    slide_title = find_shape_by_name(slide.shapes,'index')
    add_text(slide_title,row['index'])
    slide_title = find_shape_by_name(slide.shapes,'hiragana')
    add_text(slide_title,row['hiragana'])
    slide_title = find_shape_by_name(slide.shapes,'kanji')
    add_text(slide_title,row['kanji'])


# csv to ppt
def csv_to_ppt(data, ppt_template, output_ppt):
    prs = Presentation(ppt_template)
    new_prs = Presentation()

    # 遍历 CSV 文件中的每条记录，创建对应的幻灯片
    for row in data:
        # 从模板创建一张新的幻灯片
        new_slide = duplicate_slide(prs, 0)
        # 填充文本
        fill_slide(new_slide, row)

    # 删除模版幻灯片
    delete_slide(prs,0)
    # 保存修改后的幻灯片
    prs.save(output_ppt)
    print(f'幻灯片已保存为 {output_ppt}')



# 复制幻灯片
def duplicate_slide(prs, index):
    slide_to_copy = prs.slides[index]
    # Create a new slide object
    new_slide = prs.slides.add_slide(slide_to_copy.slide_layout)
    # Copy content from the original slide to the new slide
    for shape in slide_to_copy.shapes:
        el = shape.element
        newel = copy.deepcopy(el)
        new_slide.shapes._spTree.insert_element_before(newel, 'p:extLst')

    return new_slide

# 删除幻灯片
def delete_slide(prs, slide_index):
    xml_slides = prs.slides._sldIdLst  # Access the slide list XML elements
    if slide_index < 0 or slide_index >= len(xml_slides):
        print(f"Slide index {slide_index} out of range.")
        return
    slide_id = xml_slides[slide_index].rId
    prs.slides._sldIdLst.remove(xml_slides[slide_index])  # Remove slide from XML elements
    prs.part.drop_rel(slide_id)  # Remove slide from related parts


# 查找文本框
def find_shape_by_name(shapes, name):
    for shape in shapes:
        # print(shape.shape_id)
        if shape.name == name:
            return shape

# 修改文字
def add_text(shape, text,):
    tf = shape.text_frame
    tf.clear()
    run = tf.paragraphs[0].add_run()
    run.text = text if text else ''
    font=run.font
    font.name = 'Calibri'
    font.size = Pt(58)
    font.bold = True
    
# 将ppt转化成pdf
def ppt_to_pdf(ppt_file):
    # 将ppt转化成pdf
    cmdLine = "soffice --headless --convert-to pdf " + ppt_file
    subprocess.call(cmdLine, shell=True)
    name, ext = os.path.splitext(ppt_file)
    pdf_file = name + ".pdf"
    print(f'幻灯片已导出为PDF ')
    return pdf_file
    
    
# 将pdf转化成jpg
def pdf_to_img(pdf_file,imgs_folder):
    pages = convert_from_path(pdf_file, first_page=1)
    for img in pages:
        img_path = imgs_folder + '/' + str(pages.index(img)+1) + ".jpg"
        img.save(img_path , quality=100)
    print(f'PDF已导出为图片 ')


def filter_files_by_extension(folder, extension):
    # 获取指定文件夹中特定后缀名的文件列表
    file_list = [file for file in os.listdir(folder) if file.endswith(extension)]
    return file_list

def compare_mp3_and_img_files(mp3_folder, img_folder):
    # 筛选出 mp3_folder 中所有的 mp3 文件和 img_folder 中所有的 img 文件
    mp3_files = filter_files_by_extension(mp3_folder, '.mp3')
    img_files = filter_files_by_extension(img_folder, '.jpg')
    
    # 去掉后缀,将文件名列表转换为集合，以便进行快速比较
    mp3_set = {file.split('.')[0] for file in mp3_files}
    img_set = {file.split('.')[0] for file in img_files}
    
    # 找出两个集合中共同的文件名
    common_files = mp3_set.intersection(img_set)
    
    # 找出只在 mp3 文件夹中的文件名
    unique_to_mp3 = mp3_set - img_set
    
    # 找出只在 img 文件夹中的文件名
    unique_to_img = img_set - mp3_set
    # 打印结果
    print(f'共同的文件名: {common_files}')
    print(f'只在 mp3 文件夹中的文件名: {unique_to_mp3}')
    print(f'只在 img 文件夹中的文件名: {unique_to_img}')

    return common_files, unique_to_mp3, unique_to_img


# img + mp3 ---> video
def img_mp3_to_mp4(mp3_folder, img_folder, mp4_folder):
    # 执行比较
    common_files, unique_to_mp3, unique_to_img = compare_mp3_and_img_files(mp3_folder, img_folder)
    
    for name in common_files:
        mp3_file = mp3_folder + '/' + name + '.mp3'
        img_file = img_folder + '/' + name + '.jpg'
        mp4_file = mp4_folder + '/' + name + '.mp4'
        ffmpeg_cmd = "ffmpeg \
        -i " + mp3_file + \
        " -loop 1 -i " + img_file + \
        " -vcodec libx264 -pix_fmt yuv420p -shortest -y " \
        " -vf 'scale=2204:1240' " \
        +  mp4_file
        # print(ffmpeg_cmd)
        subprocess.call(ffmpeg_cmd, shell=True)


# 将多个短视频合成为一个视频
def concatenate_mp4_files(folder_path, output_file):
    # 获取文件夹中所有的 MP4 文件,
    files = []
    for file in os.listdir(folder_path):
        if file.endswith('.mp4'):
            # print(file)
            files.append(os.path.join(folder_path, file))
    # 并使用 .sort() 方法进行排序（直接修改原始列表）
    files.sort()
    # 生成输入文件列表
    input_files = '|'.join(files)
    # print(input_files)
    
    # 构建 FFmpeg 命令
    ffmpeg_cmd = "sudo ffmpeg -i concat:" +input_files+ " -c copy -bsf:a aac_adtstoasc -movflags +faststart " + output_file
    # print(ffmpeg_cmd)
    # 执行 FFmpeg 命令
    subprocess.call(ffmpeg_cmd, shell=True)


