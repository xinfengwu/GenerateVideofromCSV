#! /usr/bin/env python3

from myutils import *

# 主函数
def main():

# 输入输出文件及文件夹
    root_folder_name = "Case3-中上级日语影子跟读" 
    root_folder_path = os.path.join(os.getcwd(), root_folder_name)
    root_folder = create_folder(root_folder_path)
    # 输入文件
    covers_data = read_csv_data(os.path.join(root_folder, "case3_covers.csv"))
    ppt_template = os.path.join(os.getcwd(), "portrait_templates.pptx")
    # 输入文件夹  
    mp3_folder_name = "mp3"
    mp3_folder = os.path.join(root_folder, mp3_folder_name)
    
    # 输出文件
    covers_ppt = os.path.join(root_folder, "covers.pptx")
    pdf_file = os.path.join(root_folder, "covers.pdf")
    # 输出文件夹
    img_folder_name = "img"
    img_folder = create_folder(os.path.join(root_folder, img_folder_name))   
    mp4_folder_name = "mp4"
    mp4_folder = create_folder(os.path.join(root_folder, mp4_folder_name))
    
    # 打开源PPTX文件
    src_prs = Presentation(ppt_template)
    
# csv to ppt
    bg_img_path = ""
    new_presentation = create_ppt_with_csv(covers_data, src_prs, 3, bg_img_path, covers_ppt)

# ppt to pdf
    ppt_to_pdf_by_unoconv(covers_ppt, pdf_file)
    
# pdf to img
    pdf_to_img(pdf_file, img_folder)

# 获取mp3文件列表
    mp3_filenames = list(get_filenames_by_extension(mp3_folder, ".mp3"))
    
# 生成视频
    resolution = set_video_resolution(new_presentation.slide_width, new_presentation.slide_height)
    
    for mp3_filename in mp3_filenames:
        mp3_file = os.path.join(mp3_folder, mp3_filename+".mp3")
        img_file = os.path.join(img_folder, mp3_filename+".jpg")
        mp4_file = os.path.join(mp4_folder, "SHADOWING\ 日语影子跟读\ 新シャドーイング\ 日本語を話そう\ 中上級編_"+mp3_filename+".mp4")
        
        img_mp3_to_mp4(mp3_file, img_file, resolution, mp4_file)
        
# 删除文件及文件夹
    delete_file(covers_ppt)
    delete_file(pdf_file)
    
    delete_folder(img_folder)

    print("Done: ", lesson)


# if __name__ == "__main__":
main()
