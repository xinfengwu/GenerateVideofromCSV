#! /usr/bin/env python3

from myutil import *


# 主函数
def main():
    # 构建项目文件夹路径
    folder_path = os.path.join(os.getcwd(), "nihon", "Lesson1") 
    root_folder = create_folder(folder_path)
# 读取 CSV 文件
    data = read_csv_data(os.path.join(root_folder, "data.csv"))
    
# csv to mp3
    # 创建文件夹body_mp3文件夹
    body_mp3_folder = create_folder(os.path.join(root_folder, "body_mp3")) 
    for row in data:
        keyword = row['kanji'].strip()
        file_name = row['index'].strip() + ".mp3"
        output_file = os.path.join(body_mp3_folder, file_name)
        text_to_mp3(keyword,output_file)
    
    
# csv to ppt
    print("csv to ppt")
    # 生成正文ppt
    body_ppt_template = os.path.join(root_folder, "body_template.pptx")
    body_output_ppt= os.path.join(root_folder, "body.pptx")
    create_ppt_with_csv(data, body_ppt_template, body_output_ppt)
    
    # 生成封面封底ppt
    covers_ppt_template = os.path.join(root_folder, "covers_template.pptx")
    # covers_output_ppt = "covers.pptx"
    covers_output_ppt = os.path.join(root_folder, "covers_template.pptx")
    # print(covers_output_ppt)
   

# Create a 3-second silent MP3 file
    silence_mp3 = create_silent_mp3(3, os.path.join(body_mp3_folder,"0.mp3"))
    # print(silence_mp3)

    
# ppt to img
    # 创建文件夹pdf
    pdf_folder = create_folder(os.path.join(root_folder, "pdf"))
    body_pdf = os.path.join(pdf_folder, "body.pdf")
    # 转换正文ppt---> pdf ---> img
    body_pdf_file = ppt_to_pdf_by_soffice(body_output_ppt, pdf_folder, body_pdf)
    # 创建文件夹body_img
    body_img_folder = create_folder(os.path.join(root_folder, "body_img"))
    pdf_to_img(body_pdf_file, body_img_folder)
    
    # 创建文件夹covers_img
    covers_img_folder = create_folder(os.path.join(root_folder, "covers_img"))
    # 转换封面封底ppt---> pdf ---> img
    covers_pdf = os.path.join(pdf_folder, "covers_template.pdf")
    covers_pdf_file = ppt_to_pdf_by_soffice(covers_output_ppt, pdf_folder, covers_pdf)
    pdf_to_img(covers_pdf_file,covers_img_folder)
    
    
# img + mp3 to video
    # 转换body img---> video
    # 创建文件夹body_mp4
    body_mp4_folder = create_folder(os.path.join(root_folder, "body_mp4"))
    print("Starting mp4")
    # 执行比较
    common_files, unique_to_mp3, unique_to_img = compare_mp3_and_img_files(body_mp3_folder, body_img_folder)
    print(common_files)
    # 视频解析度要为2的倍数
    resolution=[1240,2204]
    for name in common_files:
        mp3_file = os.path.join(body_mp3_folder, name + '.mp3')
        img_file = os.path.join(body_img_folder, name + '.jpg')
        mp4_file = os.path.join(body_mp4_folder, name + '.mp4')
        img_mp3_to_mp4(mp3_file, img_file, resolution, mp4_file)
    
    # 转换封面封底img---> video
    # 创建文件夹covers_mp4
    covers_mp4_folder = create_folder(os.path.join(root_folder, "covers_mp4"))
    
    cover_img = os.path.join(covers_img_folder, "1.jpg")
    cover_mp4 = os.path.join(covers_mp4_folder, "1.mp4")
    image_to_video(cover_img, resolution, cover_mp4)
    
    back_cover_img = os.path.join(covers_img_folder, "2.jpg")
    back_cover_mp4 = os.path.join(covers_mp4_folder, "2.mp4")
    image_to_video(back_cover_img, resolution, back_cover_mp4)
        
        
# video1 + video2 + ... to videos
    output_mp4 = os.path.join(root_folder, "output.mp4")
    
    body_mp4_files = create_mp4_filelist(body_mp4_folder)
    # 将列表里的每个视频文件重复3次
    repeated_mp4_files = repeat_elements(body_mp4_files, 1)
    # print(repeated_mp4_files)

    # 在每个列表的视频文件后插入静音视频间隔
    silence_file = os.path.join(root_folder, "silence.mp4") 
    create_silent_video(3, resolution, silence_file)
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
    print("文件写入完成")

    # 合成 封面+正文+封底 视频
    concatenate_mp4_files(input_txt, output_mp4)

# if __name__ == "__main__":
main()
