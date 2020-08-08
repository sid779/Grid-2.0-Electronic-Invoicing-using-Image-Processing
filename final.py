from PIL import Image, ImageOps
import subprocess, sys, os, glob
import pytesseract
import re
from datetime import datetime


def split_pdf(filename, density=300):
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
images = pdf_to_images(PATH + "/Sample15.pdf")
print(images)


def extract_data_per_page(path):
    image = Image.open(path)
    text = pytesseract.image_to_string(image, lang='eng', config=r'--dpi 300')

    return text


for i in images:
    text = extract_data_per_page(i)
    print(text)

    # GSTIN no
    pattern = re.compile(r"\bGSTIN\b")
    length = len(text)
    print('Length : ' + str(length))
    for match in re.finditer(pattern, text):
        s = match.start()
        e = match.end()
        print('Found "%s" at %d:%d' % (text[s:e], s, e))
        i = e
        while i < length and text[i] != ':' and text[i] != '>' and text[i] != '.' and text[i] != '=': i += 1
        i += 1
        while text[i] == ' ' or text[i] == '\n': i += 1
        if text[i] == '-': i += 1
        if text[i] == '>': i += 1
        if text[i] == ':': i += 1
        if text[i] == '=': i += 1
        while text[i] == ' ' or text[i] == '\n': i += 1
        gstin = ''
        while i < length and text[i] != ' ' and text[i] != '\n':
            gstin += text[i]
            i += 1
        print('GSTIN : ' + gstin)

    # PAN
    pattern = re.compile(r"\bPAN\b")
    for match in re.finditer(pattern, text):
        s = match.start()
        e = match.end()
        print('Found "%s" at %d:%d' % (text[s:e], s, e))
        i = e
        while i < length and text[i] != ':' and text[i] != '>' and text[i] != '.' and text[i] != '=': i += 1
        i += 1
        while text[i] == ' ' or text[i] == '\n': i += 1
        if text[i] == '-': i += 1
        if text[i] == '>': i += 1
        if text[i] == ':': i += 1
        if text[i] == '=': i += 1
        while text[i] == ' ' or text[i] == '\n': i += 1
        pan = ''
        while i < length and text[i] != ' ' and text[i] != '\n':
            pan += text[i]
            i += 1
        if pan: print('PAN : ' + pan)

    # CIN
    pattern = re.compile(r"\bCIN\b")
    for match in re.finditer(pattern, text):
        s = match.start()
        e = match.end()
        print('Found "%s" at %d:%d' % (text[s:e], s, e))
        i = e
        while i < length and text[i] != ':' and text[i] != '>' and text[i] != '.' and text[i] != '=': i += 1
        i += 1
        while text[i] == ' ' or text[i] == '\n': i += 1
        if text[i] == '-': i += 1
        if text[i] == '>': i += 1
        if text[i] == ':': i += 1
        if text[i] == '=': i += 1
        while text[i] == ' ' or text[i] == '\n': i += 1
        cin = ''
        while i < length and text[i] != ' ' and text[i] != '\n':
            cin += text[i]
            i += 1
        if cin: print('CIN : ' + cin)

    # Currency
    pattern = re.compile(r"\bCURRENCY\b", re.IGNORECASE)
    for match in re.finditer(pattern, text):
        s = match.start()
        e = match.end()
        print('Found "%s" at %d:%d' % (text[s:e], s, e))
        i = e
        while i < length and text[i] != ':' and text[i] != '>' and text[i] != '.' and text[i] != '=': i += 1
        i += 1
        while text[i] == ' ' or text[i] == '\n': i += 1
        if text[i] == '-': i += 1
        if text[i] == '>': i += 1
        if text[i] == ':': i += 1
        if text[i] == '=': i += 1
        while text[i] == ' ' or text[i] == '\n': i += 1
        currency = ''
        while i < length and text[i] != '\n' and text[i] != ' ':
            currency += text[i]
            i += 1
        if currency: print('CURRENCY : ' + currency)

    # Invoice No
    pattern = re.compile(r"\b(INVOICE.NO)|(INVOICE.NUMBER)\b", re.IGNORECASE)
    for match in re.finditer(pattern, text):
        s = match.start()
        e = match.end()
        print('Found "%s" at %d:%d' % (text[s:e], s, e))
        i = e
        while i < length and text[i] != ':' and text[i] != '>' and text[i] != '.' and text[i]!='=': i += 1
        i += 1
        while text[i] == ' ' or text[i] == '\n': i += 1
        if text[i] == '-': i += 1
        if text[i] == '>': i += 1
        if text[i] == ':': i += 1
        if text[i] == '=': i += 1
        while text[i] == ' ' or text[i] == '\n': i += 1
        inv_no = ''
        while i < length and text[i] != '\n' and text[i] != ' ':
            inv_no += text[i]
            i += 1
        if inv_no: print('INVOICE NO : ' + inv_no)

    # ORDER No
    pattern = re.compile(r"\b(ORDER.NO)|(ORDER.NUMBER)\b", re.IGNORECASE)
    for match in re.finditer(pattern, text):
        s = match.start()
        e = match.end()
        print('Found "%s" at %d:%d' % (text[s:e], s, e))
        i = e
        while i < length and text[i] != ':' and text[i] != '>' and text[i] != '.' and text[i] != '=': i += 1
        i += 1
        while text[i] == ' ' or text[i] == '\n': i += 1
        if text[i] == '-': i += 1
        if text[i] == '>': i += 1
        if text[i] == ':': i += 1
        if text[i] == '=': i += 1
        while text[i] == ' ' or text[i] == '\n': i += 1
        order_no = ''
        while i < length and text[i] != '\n' and text[i] != ' ':
            order_no += text[i]
            i += 1
        if order_no: print('ORDER NO : ' + order_no)

    # PO No
    pattern = re.compile(r"\b(PO.NO)|(P\.O\..NO)\b",re.IGNORECASE)
    for match in re.finditer(pattern, text):
        s = match.start()
        e = match.end()
        print('Found "%s" at %d:%d' % (text[s:e], s, e))
        i = e
        while i < length and text[i] != ':' and text[i] != '>' and text[i] != '.' and text[i] != '=': i += 1
        i += 1
        while text[i] == ' ' or text[i] == '\n': i += 1
        if text[i] == '-': i += 1
        if text[i] == '>': i += 1
        if text[i] == ':': i += 1
        if text[i] == '=': i += 1
        while text[i] == ' ' or text[i] == '\n': i += 1
        po = ''
        while i < length and text[i] != '\n' and text[i] != ' ':
            po += text[i]
            i += 1
        if po: print('P.O. No : ' + po)

    # PHONE/CONTACT No
    pattern = re.compile(r"\b(PHONE.NO)|(CONTACT.NO)|(PHONE)|(CONTACT)|(TEL)|(TELE.PHONE)\b", re.IGNORECASE)
    for match in re.finditer(pattern, text):
        s = match.start()
        e = match.end()
        print('Found "%s" at %d:%d' % (text[s:e], s, e))
        i = e
        while i < length and text[i] != ':' and text[i] != '>' and text[i] != '.' and text[i] != '=': i += 1
        i += 1
        while text[i] == ' ' or text[i] == '\n': i += 1
        if text[i] == '-': i += 1
        if text[i] == '>': i += 1
        if text[i] == ':': i += 1
        if text[i] == '=': i += 1
        while text[i] == ' ' or text[i] == '\n': i += 1
        phone = ''
        while i < length and text[i] != '\n' and text[i] != ' ':
            phone += text[i]
            i += 1
        if phone: print(text[s:e] +' : '+ phone)

    # STATE
    pattern = re.compile(r'\bSTATE\b',re.IGNORECASE)
    for match in re.finditer(pattern, text):
        s = match.start()
        e = match.end()
        print('Found "%s" at %d:%d' % (text[s:e], s, e))
        i = e
        while i < length and text[i] != ':' and text[i] != '>' and text[i] != '.' and text[i] != '=': i += 1
        i += 1
        while text[i] == ' ' or text[i] == '\n': i += 1
        if text[i] == '-': i += 1
        if text[i] == '>': i += 1
        if text[i] == ':': i += 1
        if text[i] == '=': i += 1
        while text[i] == ' ' or text[i] == '\n': i += 1
        state = ''
        while i < length and text[i] != '\n':
            state += text[i]
            i += 1
        if state: print('STATE : ' + state)

    # BILLING ADDRESS
    pattern = re.compile(r'BILL((.|\n)*?)INDIA', re.IGNORECASE)
    for match in re.finditer(pattern, text):
        s = match.start()
        e = match.end()
        print('Found "%s" at %d:%d' % (text[s:e], s, e))
        bill_addr = text[s:e]
        bill_addr = re.sub(r':|\)', '', bill_addr)
        bill_addr = re.sub("\n", " ", bill_addr)
        print(bill_addr)

    # SHIPPING ADDRESS
    pattern = re.compile(r'SHIP((.|\n)*?)INDIA', re.IGNORECASE)
    for match in re.finditer(pattern, text):
        s = match.start()
        e = match.end()
        print('Found "%s" at %d:%d' % (text[s:e], s, e))
        ship_addr = text[s:e]
        ship_addr = re.sub(r':|\)','',ship_addr)
        ship_addr = re.sub("\n"," ",ship_addr)
        print(ship_addr)


    # NAME
    pattern = re.compile(r'\bNAME\b',re.IGNORECASE)
    for match in re.finditer(pattern, text):
        s = match.start()
        e = match.end()
        print('Found "%s" at %d:%d' % (text[s:e], s, e))
        i = e
        while i < length and text[i] != ':' and text[i]!= '>' and text[i]!='.': i += 1
        i += 1
        if text[i] == '-': i += 1
        while text[i] == ' ' or text[i] == '\n': i += 1
        name = ''
        while i < length and text[i] != '\n':
            name += text[i]
            i += 1
        if name: print('NAME : ' + name)


    # ADDRESS
    pattern = re.compile(r'\bADDRESS\b',re.IGNORECASE)
    for match in re.finditer(pattern, text):
        s = match.start()
        e = match.end()
        print('Found "%s" at %d:%d' % (text[s:e], s, e))
        i = e
        while i < length and text[i] != ':' and text[i]!= '>' and text[i]!='.' and text[i]!='\n': i += 1
        i += 1
        if text[i] == '-': i += 1
        while text[i] == ' ' or text[i] == '\n': i += 1
        addr = ''
        while i < length and text[i] != '\n':
            addr += text[i]
            i += 1
        if addr: print('ADDRESS : ' + addr)

    # Date
    pattern = re.compile(r"\b(DATE)|(DATED)\b",re.IGNORECASE)
    for match in re.finditer(pattern, text):
        s = match.start()
        e = match.end()
        print('Found "%s" at %d:%d' % (text[s:e], s, e))
        i = e
        while i < length and text[i] != ':' and text[i] != '>' and text[i] != '.' and text[i] != '=' and not (text[i]>='0' and text[i]<='9'): i += 1
        i += 1
        while text[i] == ' ' or text[i] == '\n': i += 1
        if text[i] == '-': i += 1
        if text[i] == '>': i += 1
        if text[i] == ':': i += 1
        if text[i] == '=': i += 1
        while text[i] == ' ' or text[i] == '\n': i += 1
        date = ''
        while i < length and text[i] != '\n':
            date += text[i]
            i += 1

        Date = None
        for fmt in ('%d-%m-%Y', '%d.%m.%Y', '%d/%m/%Y'):
            try:
                Date =  datetime.strptime(date, fmt).date()
            except ValueError:
                pass

        if Date: print('DATE_EXT : ' + str(Date))
        elif date: print('DATE : ' + date)




'''
    # BILLING ADDRESS
    pattern = re.compile(r'\bBILL((.|\n)*)INDIA\b', re.IGNORECASE)
    for match in re.finditer(pattern, text):
        s = match.start()
        e = match.end()
        print('Found "%s" at %d:%d' % (text[s:e], s, e))
        bill_addr = text[s:e]
        print(bill_addr)
'''
