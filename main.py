#! /usr/bin/env python3

from myutil import *


# 主函数
def main():
    # 构建项目文件夹路径
    root_folder_name = "30天背完小学英语单词"
    level2_folder_name = "Lesson1"
    
    root_folder_path = os.path.join(os.getcwd(), root_folder_name) 
    level2_folder_path = os.path.join(os.getcwd(), root_folder_name, level2_folder_name) 
    
    root_folder = create_folder(root_folder_path)
    leve2_folder = create_folder(level2_folder_path)
    
    output_mp4 = os.path.join(root_folder, "Lesson1.mp4")
# 读取 CSV 文件
    data = read_csv_data(os.path.join(root_folder, "day1.csv"))
    
    
# csv to ppt
    print("csv to ppt")
    # 生成正文ppt
    body_ppt_template = os.path.join(root_folder, "body_template.pptx")
    body_output_ppt = os.path.join(root_folder, "body.pptx")
    # 新幻灯片背景图片
    bg_img_path = ""
    new_Presentation = create_ppt_with_csv(data, body_ppt_template, bg_img_path, body_output_ppt)
    
    new_slide_width = new_Presentation.slide_width
    new_slide_height = new_Presentation.slide_height


    # 封面封底ppt模板
    covers_ppt_template = os.path.join(root_folder, "covers_template.pptx")
    

    
# pppt---> pdf ---> img
    print("ppt to img")
    # 创建文件夹pdf
    pdf_folder = create_folder(os.path.join(leve2_folder, "pdf"))
    body_pdf = os.path.join(pdf_folder, "body.pdf")
    # 转换正文ppt---> pdf ---> img
    # ppt_to_pdf_by_soffice(body_output_ppt, pdf_folder)
    ppt_to_pdf_by_unoconv(body_output_ppt, body_pdf)
    # 创建文件夹body_img
    body_img_folder = create_folder(os.path.join(leve2_folder, "body_img"))
    pdf_to_img(body_pdf, body_img_folder)
    
    # 创建文件夹covers_img
    covers_img_folder = create_folder(os.path.join(leve2_folder, "covers_img"))
    # 转换封面封底ppt---> pdf ---> img
    covers_pdf = os.path.join(pdf_folder, "covers_template.pdf")
    ppt_to_pdf_by_soffice(covers_ppt_template, pdf_folder, covers_pdf)
    pdf_to_img(covers_pdf,covers_img_folder)
    
    
# csv to mp3
    print("csv to mp3")
    # 创建文件夹body_mp3文件夹
    body_mp3_folder = create_folder(os.path.join(leve2_folder, "body_mp3")) 
    for row in data:
        keyword = row['kanji'].strip()
        file_name = row['index'].strip() + ".mp3"
        output_file = os.path.join(body_mp3_folder, file_name)
        text_to_mp3(keyword, 'en', output_file)
        
    
# img + mp3 to video
    print("img + mp3 to video")
    # 转换body img---> video
    # 创建文件夹body_mp4
    body_mp4_folder = create_folder(os.path.join(leve2_folder, "body_mp4"))
    # print("Starting mp4")
    # 执行比较
    common_files, unique_to_mp3, unique_to_img = compare_mp3_and_img_files(body_mp3_folder, body_img_folder)
    # print(common_files)
    
    resolution = set_video_resolution(new_slide_width, new_slide_height)
    
    for name in common_files:
        mp3_file = os.path.join(body_mp3_folder, name + '.mp3')
        img_file = os.path.join(body_img_folder, name + '.jpg')
        mp4_file = os.path.join(body_mp4_folder, name + '.mp4')
        img_mp3_to_mp4(mp3_file, img_file, resolution, mp4_file)
    
    # 转换封面封底img---> video
    # 创建文件夹covers_mp4
    covers_mp4_folder = create_folder(os.path.join(leve2_folder, "covers_mp4"))
    
    cover_img = os.path.join(covers_img_folder, "1.jpg")
    cover_mp4 = os.path.join(covers_mp4_folder, "1.mp4")
    image_to_video(cover_img, resolution, 3, cover_mp4)
    
    back_cover_img = os.path.join(covers_img_folder, "2.jpg")
    back_cover_mp4 = os.path.join(covers_mp4_folder, "2.mp4")
    image_to_video(back_cover_img, resolution, 3, back_cover_mp4)
        
        
# video1 + video2 + ... to videos
    print("video1 + video2 + ... to videos")
    body_mp4_files = create_mp4_filelist(body_mp4_folder)
    # 将列表里的每个视频文件重复3次
    repeated_mp4_files = repeat_elements(body_mp4_files, 3)
    # print(repeated_mp4_files)

    # 在每个列表的视频文件后插入2秒静音视频间隔
    silence_file = os.path.join(root_folder, "silence.mp4")
    gap_image_path = os.path.join(level1_folder_name, "gap.jpg")
    image_to_video(gap_image_path, resolution, 2, silence_file)
    result_files = add_elements_with_silence(repeated_mp4_files, silence_file)
    
    # 将封面、封底分别插入列表首尾位置
    # 在列表开头插入元素
    result_files.insert(0, cover_mp4)
    # 在列表末尾插入元素
    result_files.append(back_cover_mp4)
    
    # 将列表写入txt文件
    input_txt = os.path.join(root_folder, "input.txt")
    with open(input_txt, 'w') as file:
        for item in result_files:
            file.write(f"file {item}\n")
    # print(result_files)
    # print("文件写入完成")

    # 合成 封面+正文+封底 视频
    concatenate_mp4_files(input_txt, output_mp4)
    
    # 删除临时文件及文件夹
    # delete_folder(body_mp3_folder)
    delete_folder(body_mp4_folder)
    # delete_folder(body_img_folder)
    delete_folder(covers_img_folder)
    delete_folder(covers_mp4_folder)
    delete_folder(pdf_folder)
    
    delete_file(body_output_ppt)
    delete_file(input_txt)
    delete_file(silence_file)
    

# if __name__ == "__main__":
main()
