from pptx import Presentation
from pptx.util import Cm, Inches, Pt
from wand.image import Image
import os

# create presentation
prs = Presentation()

# choosing slide layout
sl = prs.slide_layouts[8]

# adding slides
slides = {}
for n in range(1,6):
    slides["slide%s" %n] = prs.slides.add_slide(sl)

# create directory for composite images
try:
    path = os.path.join("/home/neeraj/Desktop/Development/Python/Create_ppt/", "logo_images")
    os.mkdir(path)

except FileExistsError:
    print("File is already exists")

# image path
logo = 'images/nike_black.png'
images = {}
for m in range(1, 6):
    images["image%s" % m] = 'images/image%s.jpg' % m

# set logo in image
x = 0
for i in images.values():
    x = x+1
    print("set logo in Image%s" % x)
    with Image(filename=i) as img:
        img.resize(500, 400)
        with Image(filename=logo) as logo_img:
            logo_img.resize(170, 60)
            img.composite(logo_img, left=5, top=5)
            loc = "logo_images/logo_image%s.jpg" % x
        img.save(filename=loc)

# adding image and title in slide
x = 0
for i in slides.values():
    # adding image
    x = x+1
    i.shapes.add_picture("logo_images/logo_image%s.jpg" % x, Inches(1.5), Inches(1.5))
    # adding title
    left = Inches(1.4)
    top = Inches(0.5)
    width = height = Inches(7)
    p = i.shapes.add_textbox(left, top, width, height).text_frame.add_paragraph()
    p.text = "This ppt file is created by Python"
    p.font.size = Pt(30)

# saving file
prs.save("my_ppt.pptx")