#! /usr/bin/env python3

from myutil import *


# 主函数
def main():
# 读取 CSV 文件
    data = read_csv_data("data.csv")
        
# csv to mp3
    create_folder("mp3") # 创建文件夹mp3
    mp3_folder = os.path.join(os.getcwd(),"mp3")
    for row in data:
        text_to_mp3(row,mp3_folder)
    
    
# csv to ppt
    print("csv to ppt")
    # 加载 PowerPoint 文件
    ppt_template = 'template.pptx'
    output_ppt= 'output.pptx'
    csv_to_ppt(data, ppt_template, output_ppt)
    
    
# ppt to img
    create_folder("img") # 创建文件夹img
    img_folder = os.path.join(os.getcwd(),"img")
    
    pdf_file = ppt_to_pdf(output_ppt)
    pdf_to_img(pdf_file,img_folder)
    
    
# img + mp3 to video
    create_folder("mp4") # 创建文件夹mp4
    mp4_folder = os.path.join(os.getcwd(),"mp4")
    img_mp3_to_mp4(mp3_folder, img_folder, mp4_folder)
    

# 暂停 5 秒
    time.sleep(5)



# video1 + video2 + ... to videos
    output_mp4 = "output.mp4"
    concatenate_mp4_files(mp4_folder, output_mp4)

# if __name__ == "__main__":
main()
