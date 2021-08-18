from pptx import Presentation
from wand.image import Image
import os

# image path
img_path = "images/image1.jpg"

# create presentation
prs = Presentation()

# choosing slide layout
sl = prs.slide_layouts[8]

# adding a slide
slide1 = prs.slides.add_slide(sl)
slide2 = prs.slides.add_slide(sl)
slide3 = prs.slides.add_slide(sl)
slide4 = prs.slides.add_slide(sl)
slide5 = prs.slides.add_slide(sl)

# image path
logo = 'images/nike_black.png'
image1 = 'images/image1.jpg'
image2 = 'images/image2.jpg'
image3 = 'images/image3.jpg'
image4 = 'images/image4.jpg'
image5 = 'images/image5.jpg'

# create directory for composite images
path = os.path.join("/home/neeraj/Desktop/Development/Python/Create_ppt/", "logo_images")
os.mkdir(path)

# set logo in image
with Image(filename=image1) as main_img:
    main_img.resize(1000, 800)
    with Image(filename=logo) as logo_img:
        logo_img.resize(140, 70)
        main_img.composite(logo_img, left=170, top=130)
    main_img.save(filename="logo_images/logo_image1.jpg")

with Image(filename=image2) as main_img:
    main_img.resize(1000, 800)
    with Image(filename=logo) as logo_img:
        logo_img.resize(140, 70)
        main_img.composite(logo_img, left=170, top=130)
    main_img.save(filename="logo_images/logo_image2.jpg")

with Image(filename=image3) as main_img:
    main_img.resize(1000, 800)
    with Image(filename=logo) as logo_img:
        logo_img.resize(140, 70)
        main_img.composite(logo_img, left=170, top=130)
    main_img.save(filename="logo_images/logo_image3.jpg")

with Image(filename=image4) as main_img:
    main_img.resize(1000, 800)
    with Image(filename=logo) as logo_img:
        logo_img.resize(140, 70)
        main_img.composite(logo_img, left=170, top=130)
    main_img.save(filename="logo_images/logo_image4.jpg")

with Image(filename=image5) as main_img:
    main_img.resize(1000, 800)
    with Image(filename=logo) as logo_img:
        logo_img.resize(140, 70)
        main_img.composite(logo_img, left=170, top=130)
    main_img.save(filename="logo_images/logo_image5.jpg")

# adding Images
slide1.placeholders[1].insert_picture("logo_images/logo_image1.jpg")
slide2.placeholders[1].insert_picture("logo_images/logo_image2.jpg")
slide3.placeholders[1].insert_picture("logo_images/logo_image3.jpg")
slide4.placeholders[1].insert_picture("logo_images/logo_image4.jpg")
slide5.placeholders[1].insert_picture("logo_images/logo_image5.jpg")

# adding a title in slide
slide1.shapes.title.text = "This is a title."
slide2.shapes.title.text = "This is a title."
slide3.shapes.title.text = "This is a title."
slide4.shapes.title.text = "This is a title."
slide5.shapes.title.text = "This is a title."

# saving file
prs.save("my_ppt.ppt")
