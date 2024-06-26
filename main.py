#! /usr/bin/env python3

from myutil import *


# 主函数
def main():
# 读取 CSV 文件
    csv_file = 'data.csv'  # 请替换为你的 CSV 文件路径
    input_file = csv.DictReader(open(csv_file))

# csv to mp3
    #create_folder("mp3") # 创建文件夹mp3
    #mp3_path = os.path.join(os.getcwd(),"mp3")
    #for row in input_file:
    #    text_to_mp3(row['word'],mp3_path)

# csv to ppt

    # 加载 PowerPoint 文件
    ppt_template = 'template.pptx'  # 请替换为你的 PowerPoint 模板文件路径
    presentation = Presentation(ppt_template)
    # 遍历 CSV 文件中的每条记录，创建对应的幻灯片
    for data in input_file:
        # 从模板创建一张新的幻灯片
        new_slide = duplicate_slide(presentation, 0)
        # 重命名slide
        # new_slide.shapes.title.text = data['word']
        # 填充文本
        fill_slide(new_slide, data)

    # 删除模版幻灯片
    delete_slide(presentation,0)
    # 保存修改后的幻灯片
    output_ppt = 'output.pptx'  # 输出文件路径
    presentation.save(output_ppt)
    print(f'幻灯片已保存为 {output_ppt}')

# ppt to img
    create_folder("img") # 创建文件夹img
    img_path = os.path.join(os.getcwd(),"img")
    # ppt_to_jpg(output_ppt,img_path)
    for slide_idx, slide in enumerate(presentation.slides):
        ppt2img(output_ppt,slide_idx,img_path,'jpg',"/usr/local/bin/soffice")
# img + mp3 to video

# if __name__ == "__main__":
main()
