#! /usr/bin/env python3

from myutils import *


# 主函数
def main():
    # 构建项目文件夹路径
    # 新中日标准日本语上册
    root_folder_name = "Case1-日语语料库-单词短语句型" 
    root_folder_path = os.path.join(os.getcwd(), root_folder_name)
    root_folder = create_folder(root_folder_path)
        
    section_csv_folder_name = "section_csv"
    section_csv_folder_path = os.path.join(root_folder, section_csv_folder_name)
    lessons_data = read_csv_data(os.path.join(root_folder, "lessons.csv"))
    
    ppt_template = os.path.join(os.getcwd(), "portrait_templates.pptx")
    # 打开源PPTX文件
    src_prs = Presentation(ppt_template)   
    
    section_mp4_folder_name = "section_mp4"
    section_mp4_folder_path = os.path.join(root_folder, section_mp4_folder_name)
    section_mp4_folder = create_folder(section_mp4_folder_path) 
    
    
    for lesson in lessons_data:
        section_folder_name = lesson['lesson_name']
        section_folder_path = os.path.join(os.getcwd(), root_folder_name, section_folder_name)
        section_folder = create_folder(section_folder_path)
        
        word_data = read_csv_data(os.path.join(section_csv_folder_path, section_folder_name+".csv"))
        section_mp4 = os.path.join(root_folder, section_mp4_folder, section_folder_name+".mp4")
      
        # 设置朗读语言
        lang = 'ja' # en
          
    # csv to ppt
        print("csv to ppt")

        body_output_ppt = os.path.join(section_folder, "body.pptx")
        body_bg_img_path = ""
        new_body_presentation = create_ppt_with_csv(word_data, src_prs, 1, body_bg_img_path, body_output_ppt)
        
        new_slide_width = new_body_presentation.slide_width
        new_slide_height = new_body_presentation.slide_height

        covers_output_ppt = os.path.join(section_folder, "covers.pptx")
        # covers_bg_img_path = os.path.join(root_folder, "cover.jpg")
        covers_bg_img_path = ""
        new_covers_presentation = create_ppt_with_csv(lessons_data, src_prs, 2, covers_bg_img_path, covers_output_ppt)
        

    # ppt---> pdf ---> img
        print("ppt to img")
        pdf_folder = create_folder(os.path.join(section_folder, "pdf"))
        body_pdf = os.path.join(pdf_folder, "body.pdf")
        covers_pdf = os.path.join(pdf_folder, "covers.pdf")
        
        # ppt_to_pdf_by_unoconv(body_output_ppt, body_pdf)
        ppt_to_pdf_by_soffice(body_output_ppt, pdf_folder, body_pdf)
        ppt_to_pdf_by_soffice(covers_output_ppt, pdf_folder, covers_pdf)

        body_img_folder = create_folder(os.path.join(section_folder, "body_img"))
        covers_img_folder = create_folder(os.path.join(section_folder, "covers_img"))

        pdf_to_img(body_pdf, body_img_folder)
        pdf_to_img(covers_pdf, covers_img_folder)
        
        
    # csv to mp3
        print("csv to mp3")
        body_mp3_folder = create_folder(os.path.join(section_folder, "body_mp3")) 
        for row in word_data:
            keyword = row['平假名注音'].strip()
            # keyword = row['日本語'].strip()
            file_name = row['序号'].strip() + ".mp3"
            output_file = os.path.join(body_mp3_folder, file_name)
            text_to_mp3(keyword, lang, output_file)
            
        
    # img + mp3 to video
        print("img + mp3 to video")
        body_mp4_folder = create_folder(os.path.join(section_folder, "body_mp4"))

        common_files, unique_to_mp3, unique_to_img = compare_mp3_and_img_files(body_mp3_folder, body_img_folder)
        resolution = set_video_resolution(new_slide_width, new_slide_height)
        
        for name in common_files:
            mp3_file = os.path.join(body_mp3_folder, name + '.mp3')
            img_file = os.path.join(body_img_folder, name + '.jpg')
            mp4_file = os.path.join(body_mp4_folder, name + '.mp4')
            img_mp3_to_mp4(mp3_file, img_file, resolution, mp4_file)
        
        # 转换封面封底img---> video
        covers_mp4_folder = create_folder(os.path.join(section_folder, "covers_mp4"))
        #index = section_folder_name.split("_")[1]
        cover_img = os.path.join(covers_img_folder, lesson['index']+".jpg")
        cover_mp4 = os.path.join(covers_mp4_folder, lesson['index']+".mp4")
        image_to_video(cover_img, resolution, 5, cover_mp4)
        
        back_cover_img = os.path.join(root_folder, "back_cover.jpg")
        back_cover_mp4 = os.path.join(covers_mp4_folder, "back_cover.mp4")
        image_to_video(back_cover_img, resolution, 3, back_cover_mp4)
            
            
    # video1 + video2 + ... to videos
        print("video1 + video2 + ... to videos")
        body_mp4_files = sort_filelist(body_mp4_folder, ".mp4")
        # 将列表里的每个视频文件重复2次
        repeated_mp4_files = repeat_elements(body_mp4_files, 2)
        # print(repeated_mp4_files)

        # 在每个列表的视频文件后插入1秒静音视频间隔
        silence_file = os.path.join(root_folder, "silence.mp4")
        gap_image_path = os.path.join(root_folder, "gap.jpg")
        if not os.path.exists(silence_file):
            image_to_video(gap_image_path, resolution, 1, silence_file)
        result_files = add_elements_with_silence(repeated_mp4_files, silence_file)
        
        # 将封面、封底分别插入列表首尾位置
        # 在列表开头插入封面
        result_files.insert(0, cover_mp4)
        # 在列表末尾插入封底
        result_files.append(back_cover_mp4)
        
        # 将列表写入txt文件
        input_txt = os.path.join(section_folder, "input.txt")
        with open(input_txt, 'w') as file:
            for item in result_files:
                file.write(f"file {item}\n")
        # print(result_files)
        # print("文件写入完成")

        # 合成 封面+正文+封底 视频
        concatenate_mp4_files(input_txt, section_mp4)
        
        # 删除临时文件及文件夹
        #delete_file(body_output_ppt)
        #delete_file(covers_output_ppt)
        #delete_file(input_txt)
        
        # delete_folder(section_mp4_folder)
        #delete_folder(section_folder)
        
        print("Done: ", lesson['lesson_name'])
        
    
# if __name__ == "__main__":
main()
