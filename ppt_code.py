from pptx import Presentation
from pptx.util import Cm, Inches, Pt
from wand.image import Image
import os

# image path
img_path = "images/image1.jpg"

# create presentation
prs = Presentation()

# choosing slide layout
sl = prs.slide_layouts[8]

# adding slides
slide1 = prs.slides.add_slide(sl)
slide2 = prs.slides.add_slide(sl)
slide3 = prs.slides.add_slide(sl)
slide4 = prs.slides.add_slide(sl)
slide5 = prs.slides.add_slide(sl)

# create directory for composite images
try:
    path = os.path.join("/home/neeraj/Desktop/Development/Python/Create_ppt/", "logo_images")
    os.mkdir(path)

except FileExistsError:
    print("File is already exists")

# image path
logo = 'images/nike_black.png'
image1 = 'images/image1.jpg'
image2 = 'images/image2.jpg'
image3 = 'images/image3.jpg'
image4 = 'images/image4.jpg'
image5 = 'images/image5.jpg'

# set logo in image
images = [image1, image2, image3, image4, image5]
x = 0
for i in images:
    x = x+1
    print("set logo in Image%s is complete" % x)
    with Image(filename=i) as img:
        img.resize(500, 400)
        with Image(filename=logo) as logo_img:
            logo_img.resize(170, 60)
            img.composite(logo_img, left=5, top=5)
            loc = "logo_images/logo_image%s.jpg" % x
        img.save(filename=loc)

# adding Images
slides = [slide1, slide2, slide3, slide4, slide5]
x = 0
for i in slides:
    x = x + 1
    title = i.shapes
    top = Inches(1.5)
    left = Inches(1.5)
    title.add_picture("logo_images/logo_image%s.jpg" % x, left, top)

# adding a title in slide
for i in slides:
    left = Inches(1.4)
    top = Inches(0.5)
    width = height = Inches(7)
    txt = i.shapes.add_textbox(left, top, width, height)
    tf = txt.text_frame
    p = tf.add_paragraph()
    p.text = "This ppt file is created by Python"
    p.font.size = Pt(30)

# saving file
prs.save("my_ppt.ppt")