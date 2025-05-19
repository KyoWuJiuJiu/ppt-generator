import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Cm
from io import BytesIO
from PIL import Image
import tempfile
import os
import copy

st.title("📊 Excel + 图片生成 PowerPoint")

# 上传PPT模板
ppt_file = st.file_uploader("上传 PPT 模板 (.pptx)", type=["pptx"])
# 上传Excel文件
excel_files = st.file_uploader("上传一个或多个 Excel 文件", type=["xlsx", "xls"], accept_multiple_files=True)
# 上传图片
image_files = st.file_uploader("上传产品图片（可多选）", type=["jpg", "jpeg", "png", "bmp", "gif"], accept_multiple_files=True)

if ppt_file and excel_files:
    if st.button("生成 PowerPoint"):
        with tempfile.TemporaryDirectory() as tmpdir:
            # 保存图片到临时文件夹
            image_folder = os.path.join(tmpdir, "images")
            os.makedirs(image_folder, exist_ok=True)
            for file in image_files:
                with open(os.path.join(image_folder, file.name), "wb") as f:
                    f.write(file.getbuffer())

            # 合并Excel数据
            df_all = pd.concat([pd.read_excel(f) for f in excel_files], ignore_index=True)

            # 加载PPT模板
            ppt = Presentation(ppt_file)
            template_slide = ppt.slides[0]
            new_ppt = Presentation()
            new_ppt.slide_width = ppt.slide_width
            new_ppt.slide_height = ppt.slide_height

            def clone_slide(pres, slide):
                layout = pres.slide_layouts[6]
                new_slide = pres.slides.add_slide(layout)
                for shape in slide.shapes:
                    new_shape = copy.deepcopy(shape.element)
                    new_slide.shapes._spTree.insert_element_before(new_shape, 'p:extLst')
                return new_slide

            def replace_text(slide, data):
                for shape in slide.shapes:
                    if shape.has_text_frame:
                        for para in shape.text_frame.paragraphs:
                            for run in para.runs:
                                text = run.text
                                for col, val in data.items():
                                    if pd.isna(val): val = ""
                                    text = text.replace(f"{{{col.strip()}}}", str(val))
                                run.text = text

            def insert_images(slide, item_number):
                matches = [f for f in os.listdir(image_folder) if f.startswith(str(item_number)) and f.lower().endswith(('.jpg','.jpeg','.png','.bmp','.gif'))]
                if not matches: return
                slide_width = new_ppt.slide_width
                slide_height = new_ppt.slide_height
                pic_height = Cm(18)
                right_margin = Cm(3)

                total_pic_height = pic_height * len(matches)
                space = (slide_height - total_pic_height) / (len(matches)+1) if len(matches) > 1 else (slide_height - total_pic_height)/2
                current_top = space

                for fname in sorted(matches):
                    img_path = os.path.join(image_folder, fname)
                    with Image.open(img_path) as img:
                        w, h = img.size
                        aspect_ratio = w / h if h else 1
                        pic_width = pic_height * aspect_ratio
                    slide.shapes.add_picture(img_path, slide_width - right_margin - pic_width, current_top, height=pic_height)
                    current_top += pic_height + space

            for _, row in df_all.iterrows():
                row_data = {k.strip(): v for k, v in row.to_dict().items()}
                for field in ["Item Width(Inch)", "Item Depth (inch)", "Item Height (inch)"]:
                    val = row_data.get(field, "")
                    try:
                        row_data[field] = round(float(val) * 2.54, 1) if val and not pd.isna(val) else ""
                    except:
                        row_data[field] = ""
                slide = clone_slide(new_ppt, template_slide)
                replace_text(slide, row_data)
                item_no = row_data.get("ITEM#", "")
                insert_images(slide, item_no)

            output = BytesIO()
            new_ppt.save(output)
            st.success("✅ 幻灯片已生成")
            st.download_button("📥 点击下载 PPT", output.getvalue(), file_name="output.pptx")
