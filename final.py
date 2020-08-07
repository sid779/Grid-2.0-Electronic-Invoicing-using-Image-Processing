from PIL import Image, ImageOps
import subprocess, sys, os, glob
import pytesseract
import re


def split_pdf(filename, density=600):
    prefix = filename[:-4]
    #     Command line
    cmd = "convert -colorspace gray -density  %d %s %s-%%d.png" % (density, filename, prefix)
    subprocess.call([cmd], shell=True)
    images_list = [f for f in glob.glob('%s-?.png' % prefix)]

    return images_list


def pdf_to_images(filename):
    images_path = split_pdf(filename)
    images_path.sort()
    return images_path


PATH = os.getcwd()
print(PATH)
images = pdf_to_images(PATH + "/Sample19.pdf")
print(images)


def extract_data_per_page(path):
    image = Image.open(path)
    text = pytesseract.image_to_string(image, lang='eng', config=r'--dpi 700')

    return text


for i in images:
    text = extract_data_per_page(i)
    print(text)

    # GSTIN no
    pattern = 'GSTIN'
    length = len(text)
    print('Length : ' + str(length))
    for match in re.finditer(pattern, text):
        s = match.start()
        e = match.end()
        print('Found "%s" at %d:%d' % (text[s:e], s, e))
        i = e
        while i < length and text[i] != ':' and text[i]!= '>': i += 1
        i += 1
        if text[i] == '-': i += 1
        if text[i] == ' ': i += 1
        gstin = ''
        while i < length and text[i] != ' ' and text[i] != '\n':
            gstin += text[i]
            i += 1
        print('GSTIN : ' + gstin)

    # PAN
    pattern = 'PAN'
    for match in re.finditer(pattern, text):
        s = match.start()
        e = match.end()
        print('Found "%s" at %d:%d' % (text[s:e], s, e))
        i = e
        while i < length and text[i] != ':' and text[i]!= '>': i += 1
        i += 1
        if text[i] == '-': i += 1
        if text[i] == ' ': i += 1
        pan = ''
        while i < length and text[i] != ' ' and text[i] != '\n':
            pan += text[i]
            i += 1
        if pan and len(pan)==10: print('PAN : ' + pan)

    # CIN
    pattern = 'CIN'
    for match in re.finditer(pattern, text):
        s = match.start()
        e = match.end()
        print('Found "%s" at %d:%d' % (text[s:e], s, e))
        i = e
        while i < length and text[i] != ':' and text[i]!= '>': i += 1
        i += 1
        if text[i] == '-': i += 1
        if text[i] == ' ': i += 1
        cin = ''
        while i < length and text[i] != ' ' and text[i] != '\n':
            cin += text[i]
            i += 1
        if cin: print('CIN : ' + cin)

    # STATE
    pattern = re.compile('STATE',re.IGNORECASE)
    for match in re.finditer(pattern, text):
        s = match.start()
        e = match.end()
        print('Found "%s" at %d:%d' % (text[s:e], s, e))
        i = e
        while i < length and text[i] != ':' and text[i]!= '>': i += 1
        i += 1
        if text[i] == '-': i += 1
        if text[i] == ' ': i += 1
        state = ''
        while i < length and text[i] != ' ' and text[i] != '\n':
            state += text[i]
            i += 1
        if state: print('STATE : ' + state)

    # NAME
    pattern = re.compile('NAME',re.IGNORECASE)
    for match in re.finditer(pattern, text):
        s = match.start()
        e = match.end()
        print('Found "%s" at %d:%d' % (text[s:e], s, e))
        i = e
        while i < length and text[i] != ':' and text[i]!= '>': i += 1
        i += 1
        if text[i] == '-': i += 1
        if text[i] == ' ': i += 1
        name = ''
        while i < length and text[i] != '\n':
            name += text[i]
            i += 1
        if name: print('NAME : ' + name)


    # ADDRESS
    pattern = re.compile('ADDRESS',re.IGNORECASE)
    for match in re.finditer(pattern, text):
        s = match.start()
        e = match.end()
        print('Found "%s" at %d:%d' % (text[s:e], s, e))
        i = e
        while i < length and text[i] != ':' and text[i]!= '>': i += 1
        i += 1
        if text[i] == '-': i += 1
        if text[i] == ' ': i += 1
        addr = ''
        while i < length and text[i] != '\n':
            addr += text[i]
            i += 1
        if addr: print('NAME : ' + addr)