import xlsxwriter
from PIL import Image
import os


def get_file_name(path):
    return os.path.basename(path)


def RGB(value):
    digit = list(map(str, range(10))) + list("ABCDEF")
    string = '#'
    for i in value:
        a1 = i // 16
        a2 = i % 16
        string += digit[a1] + digit[a2]
    return string


def convert(book_name, sheet_name, img_path):
    workbook = xlsxwriter.Workbook(book_name)
    worksheet = workbook.add_worksheet(sheet_name)
    imgSrc = Image.open(img_path)
    pixels = imgSrc.load()
    img_w, img_h = imgSrc.size
    task_all = img_w * img_h
    print("Height:{0} Width:{1}".format(img_h, img_w))
    w = 0
    for h in range(img_h):
        for w in range(img_w):
            pixel_color = RGB(pixels[w, h])
            format_dict = {'fg_color': pixel_color}
            my_format = workbook.add_format(format_dict)
            worksheet.write(h, w, " ", my_format)
        print("\rFinished:{0}%".format(round(((h + 1) * (w + 1) * 100) / task_all, 1)), end="")
    print("\n\nSaving File...")
    workbook.close()
    return


if __name__ == "__main__":
    src = "wife.jpg"
    target = ".xlsx"
    sheet_name = get_file_name(src)
    convert(sheet_name + target, sheet_name, src)
