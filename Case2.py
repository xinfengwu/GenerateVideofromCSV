#! /usr/bin/env python3

from myutils import *


# 主函数
def main():
    # 构建项目文件夹路径
    root_folder_name = "Case2-新版中日标准日本语初级_句型_磨耳朵专项训练" 
    root_folder_path = os.path.join(os.getcwd(), root_folder_name)
    root_folder = create_folder(root_folder_path)
        
    lessons_csv_folder_name = "Lessons_csv"
    lessons_csv_folder_path = os.path.join(root_folder, lessons_csv_folder_name)
    lessons_data = read_csv_data(os.path.join(root_folder, "lessons_data.csv"))
    
    ppt_template = os.path.join(os.getcwd(), "portrait_templates.pptx")
    # 打开源PPTX文件
    src_prs = Presentation(ppt_template)   
    
    lessons_mp4_folder_name = "Lessons_mp4"
    lessons_mp4_folder_path = os.path.join(root_folder, lessons_mp4_folder_name)
    lessons_mp4_folder = create_folder(lessons_mp4_folder_path)
    
    # 创建静音mp3
    silent_mp3 = os.path.join(root_folder, "silent.mp3")
    create_silent_mp3(1, silent_mp3)
    
    for lesson in lessons_data:
        lesson_folder_name = lesson['lesson_name']
        lesson_folder_path = os.path.join(os.getcwd(), root_folder_name, lesson_folder_name)
        lesson_folder = create_folder(lesson_folder_path)
        pdf_folder = create_folder(os.path.join(lesson_folder, "pdf"))
        body_img_folder = create_folder(os.path.join(lesson_folder, "body_img"))
        covers_img_folder = create_folder(os.path.join(lesson_folder, "covers_img"))
        
        body_data = read_csv_data(os.path.join(lessons_csv_folder_path, lesson_folder_name+".csv"))
        
        lesson_mp4_filename = lesson_folder_name + '-' + root_folder_name.split('-')[1]+".mp4"
        lesson_mp4 = os.path.join(root_folder, lessons_mp4_folder, lesson_mp4_filename)
      
        # Choose language and region
        lang = 'ja' # Language code (e.g., 'en', 'fr', 'jp')
        
          
    # 封面封底 csv to video
        print("封面封底 csv to video")
        # 根据模板创建新的封面封底pptx
        covers_output_ppt = os.path.join(lesson_folder, "covers.pptx")
        covers_bg_img_path = os.path.join(root_folder, "cover.jpg")
        
        # cover_prs = Presentation()
        # 设置新PPTX的方向和源PPTX一致
        # cover_prs.slide_width = src_prs.slide_width
        # cover_prs.slide_height = src_prs.slide_height
        # duplicate_slide(src_prs, 4, cover_prs) # 复制模板中第4张幻灯片做封面slide 不用填数据 单纯的复制
        # cover_prs.save(covers_output_ppt)
        
        cover_bg_img_path = ""
        new_cover_presentation = create_ppt_with_csv(lessons_data, src_prs, 4, cover_bg_img_path, covers_output_ppt) # 复制模板中第4张幻灯片做封面slide
        
        # 转换封面pptx ---> pdf ---> img
        covers_pdf = os.path.join(pdf_folder, "covers.pdf")
        ppt_to_pdf_by_soffice(covers_output_ppt, pdf_folder, covers_pdf)
        pdf_to_img(covers_pdf, covers_img_folder)
        # 转换封面img ---> video
        covers_mp4_folder = create_folder(os.path.join(lesson_folder, "covers_mp4"))
        new_slide_width = new_cover_presentation.slide_width
        new_slide_height = new_cover_presentation.slide_height
        resolution = set_video_resolution(new_slide_width, new_slide_height)
        
        cover_img = os.path.join(covers_img_folder, lesson['index']+".jpg")
        cover_mp4 = os.path.join(covers_mp4_folder, lesson['index']+".mp4")
        image_to_video(cover_img, resolution, 5, cover_mp4)
        
        back_cover_img = os.path.join(root_folder, "back_cover.jpg")
        back_cover_mp4 = os.path.join(covers_mp4_folder, "back_cover.mp4")
        image_to_video(back_cover_img, resolution, 3, back_cover_mp4)
        
    # 主体 csv to img
        print("主体 csv to video")
        # 根据模板创建新的主体pptx
        body_output_ppt = os.path.join(lesson_folder, "body.pptx")
        body_bg_img_path = ""
        new_body_presentation = create_ppt_with_csv(body_data, src_prs, 5, body_bg_img_path, body_output_ppt) # 复制模板中第5张幻灯片做主体slide
        # 转换主体pptx ---> pdf  ---> img
        body_pdf = os.path.join(pdf_folder, "body.pdf")
        ppt_to_pdf_by_unoconv(body_output_ppt, body_pdf)
        pdf_to_img(body_pdf, body_img_folder)
             
        # csv to mp3
        print("csv to mp3")
        body_mp3_folder = create_folder(os.path.join(lesson_folder, "body_mp3"))
        
        for row in body_data:     
            keyword = row['問句'].strip()
            file_name_question = row['序号'].strip() + "_question.mp3"
            output_file_question = os.path.join(body_mp3_folder, file_name_question)
            text_to_mp3(keyword, lang, output_file_question)
            
            keyword = row['答句'].strip()
            file_name_answer = row['序号'].strip() + "_answer.mp3"
            output_file_answer = os.path.join(body_mp3_folder, file_name_answer)
            text_to_mp3(keyword, lang, output_file_answer)
            
            mp3_list = []
            mp3_list.append(output_file_question)
            # 问句和答句之间有1秒的停顿
            mp3_list.append(silent_mp3)
            mp3_list.append(output_file_answer)
            
            # 将列表写入txt文件
            input_mp3_txt = os.path.join(lesson_folder, "input_mp3.txt")
            file_name = row['序号'].strip() + ".mp3"
            output_file = os.path.join(body_mp3_folder, file_name)
            with open(input_mp3_txt, 'w') as file:
                for item in mp3_list:
                    file.write(f"file {item}\n")
            # 合并多个mp3
            concatenate_mp3_files(input_mp3_txt, output_file)
            
            #delete_file(output_file_question)
            #delete_file(output_file_answer)
            
                 
        # img + mp3 to video
        print("img + mp3 to video")
        body_mp4_folder = create_folder(os.path.join(lesson_folder, "body_mp4"))

        common_files, unique_to_mp3, unique_to_img = compare_mp3_and_img_files(body_mp3_folder, body_img_folder)

        
        for name in common_files:
            mp3_file = os.path.join(body_mp3_folder, name + '.mp3')
            img_file = os.path.join(body_img_folder, name + '.jpg')
            mp4_file = os.path.join(body_mp4_folder, name + '.mp4')
            img_mp3_to_mp4(mp3_file, img_file, resolution, mp4_file)
            
    # video1 + video2 + ... to videos
        print("video1 + video2 + ... to videos")
        body_mp4_files = sort_filelist(body_mp4_folder, ".mp4")
        # 将列表里的每个视频文件重复3次
        repeated_mp4_files = repeat_elements(body_mp4_files, 3)
        # print(repeated_mp4_files)

        # 在每个列表的视频文件后插入2秒静音视频间隔
        silence_file = os.path.join(root_folder, "silence.mp4")
        gap_image_path = os.path.join(root_folder, "gap.jpg")
        if not os.path.exists(silence_file):
            image_to_video(gap_image_path, resolution, 2, silence_file)
        result_files = add_elements_with_silence(repeated_mp4_files, silence_file)
        
        # 将封面、封底分别插入列表首尾位置
        # 在列表开头插入封面
        result_files.insert(0, cover_mp4)
        # 在列表末尾插入封底
        result_files.append(back_cover_mp4)
        
        # 将列表写入txt文件
        input_mp4_txt = os.path.join(lesson_folder, "input_mp4.txt")
        with open(input_mp4_txt, 'w') as file:
            for item in result_files:
                file.write(f"file {item}\n")
        # print(result_files)
        # print("文件写入完成")

        # 合成 封面+正文+封底 视频
        concatenate_mp4_files(input_mp4_txt, lesson_mp4)
        
        # 删除临时文件及文件夹
        #delete_file(body_output_ppt)
        #delete_file(covers_output_ppt)
        #delete_file(input_txt)
        
        #delete_folder(lessons_mp4_folder)
        delete_folder(lesson_folder)
        print("Done: ", lesson['lesson_name'])
        
    
# if __name__ == "__main__":
main()
