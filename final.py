from PIL import Image, ImageOps
import subprocess, sys, os, glob
import pytesseract
import re
from datetime import datetime
import dateutil.parser
import xlsxwriter


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


# PATH = os.getcwd()
# print(PATH)

def extract_data_per_page(path):
    image = Image.open(path)
    text = pytesseract.image_to_string(image, lang='eng', config=r'--dpi 300 hocr')
    return text


# print(PATH,type(PATH))

def digits_in_a_line(line):
    count = 0
    for i in line:
        if (i >= '0' and i <= '9'): count += 1
    return count


def first_page_content(text, worksheet):
    global row

    length = len(text)
    # print('Length : ' + str(length))
    # Invoice No
    pattern = re.compile(r"\b(INVOICE.NO)|(INVOICE.NUMBER)\b", re.IGNORECASE)
    for match in re.finditer(pattern, text):
        s = match.start()
        e = match.end()
        # print('Found "%s" at %d:%d' % (text[s:e], s, e))
        i = e
        while i < length and text[i] != ':' and text[i] != '>' and text[i] != '.' and text[i] != '=' and (
        not (text[i] >= '0' and text[i] <= '9')): i += 1
        if (not (text[i] >= '0' and text[i] <= '9')): i += 1
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
        # if inv_no: print('INVOICE NO : ' + inv_no)
        if inv_no:
            worksheet.write(row, 0, "Invoice number ")
            worksheet.write(row, 1, inv_no)
            row += 1
            break

    # ORDER No
    pattern = re.compile(r"\b(ORDER.NO)|(ORDER.NUMBER)\b", re.IGNORECASE)
    for match in re.finditer(pattern, text):
        s = match.start()
        e = match.end()
        # print('Found "%s" at %d:%d' % (text[s:e], s, e))
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
        # if order_no: print('ORDER NO : ' + order_no)
        if order_no:
            worksheet.write(row, 0, "Order number ")
            worksheet.write(row, 1, order_no)
            row += 1
            break

    # PO No
    pattern = re.compile(r"\b(PO.NO)|(P\.O\..NO)\b", re.IGNORECASE)
    for match in re.finditer(pattern, text):
        s = match.start()
        e = match.end()
        # print('Found "%s" at %d:%d' % (text[s:e], s, e))
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
        if po:
            worksheet.write(row, 0, "PO number ")
            worksheet.write(row, 1, po)
            row += 1
            break

    # GSTIN no
    pattern = re.compile(r"\b(GSTIN)|(GST)\b")
    gst_list = []
    for match in re.finditer(pattern, text):
        s = match.start()
        e = match.end()
        # print('Found "%s" at %d:%d' % (text[s:e], s, e))
        i = e
        while i < length and text[i] != ':' and text[i] != '>' and text[i] != '.' and text[i] != '=': i += 1
        i += 1
        while text[i] == ' ' or text[i] == '\n': i += 1
        if text[i] == '-': i += 1
        if text[i] == '>': i += 1
        if text[i] == ':': i += 1
        if text[i] == '=': i += 1
        if text[i] == '.': i += 1
        while text[i] == ' ' or text[i] == '\n': i += 1
        gstin = ''
        while i < length and text[i] != ' ' and text[i] != '\n':
            gstin += text[i]
            i += 1
        if gstin not in gst_list and len(gstin) >= 10:
            gst_list.append(gstin)
        # print('GSTIN : ' + gstin)
    if (len(gst_list) > 0):
        worksheet.write(row, 0, "Seller's GST")
        worksheet.write(row, 1, gst_list[0])
        row += 1
        if (len(gst_list) > 1):
            worksheet.write(row, 0, "Buyer's GST")
            worksheet.write(row, 1, gst_list[-1])
            row += 1

    # PAN
    pattern = re.compile(r"\bPAN\b")
    pan_list = []
    for match in re.finditer(pattern, text):
        s = match.start()
        e = match.end()
        # print('Found "%s" at %d:%d' % (text[s:e], s, e))
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
        # if pan: print('PAN : ' + pan)
        if pan not in pan_list: pan_list.append(pan)
    if (len(pan_list) > 0):
        worksheet.write(row, 0, "Seller's PAN no  ")
        worksheet.write(row, 1, pan_list[0])
        row += 1
        if (len(pan_list) > 1):
            worksheet.write(row, 0, "Buyer's PAN no  ")
            worksheet.write(row, 1, pan_list[-1])
            row += 1
    # CIN
    pattern = re.compile(r"\bCIN\b")
    for match in re.finditer(pattern, text):
        s = match.start()
        e = match.end()
        # print('Found "%s" at %d:%d' % (text[s:e], s, e))
        i = e
        while i < length and text[i] != ':' and text[i] != '>' and text[i] != '.' and text[i] != '=': i += 1
        i += 1
        while text[i] == ' ' or text[i] == '\n': i += 1
        if text[i] == '-': i += 1
        if text[i] == '>': i += 1
        if text[i] == ':': i += 1
        if text[i] == '=': i += 1
        while text[i] == ' ' or text[i] == '\n': i += 1
        while (not ((text[i] >= 'a' and text[i] <= 'z') or (text[i] >= 'A' and text[i] <= 'Z') or (
                text[i] >= '0' and text[i] <= '9'))): i += 1
        cin = ''

        while i < length and text[i] != ' ' and text[i] != '\n':
            cin += text[i]
            i += 1
        # if cin: print('CIN : ' + cin)
        if len(cin) >= 4:
            worksheet.write(row, 0, "CIN ")
            worksheet.write(row, 1, cin)
            row += 1
            break
    # Currency
    pattern = re.compile(r"\bCURRENCY\b", re.IGNORECASE)
    for match in re.finditer(pattern, text):
        s = match.start()
        e = match.end()
        # print('Found "%s" at %d:%d' % (text[s:e], s, e))
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
        if len(currency) >= 4:
            worksheet.write(row, 0, "Currency ")
            worksheet.write(row, 1, currency)
            row += 1
            break

    # PHONE/CONTACT No
    pattern = re.compile(r"\b(PHONE.NO)|(CONTACT.NO)|(PHONE)|(CONTACT)|(TEL)|(TELE.PHONE)\b", re.IGNORECASE)
    for match in re.finditer(pattern, text):
        s = match.start()
        e = match.end()
        # print('Found "%s" at %d:%d' % (text[s:e], s, e))
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
        if phone: print(text[s:e] + ' : ' + phone)
        if phone:
            worksheet.write(row, 0, "Phone number ")
            worksheet.write(row, 1, phone)
            row += 1
            break

    # STATE
    pattern = re.compile(r'\bSTATE\b', re.IGNORECASE)
    for match in re.finditer(pattern, text):
        s = match.start()
        e = match.end()
        # print('Found "%s" at %d:%d' % (text[s:e], s, e))
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
        while i < length and text[i] != '\n' and (not (text[i] >= '0' and text[i] <= '9')):
            state += text[i]
            i += 1
        # if state: print('STATE : ' + state)
        if state:
            worksheet.write(row, 0, "State ")
            worksheet.write(row, 1, state)
            row += 1
            break

    # BILLING ADDRESS
    pattern = re.compile(r'BILL((.|\n)*?)INDIA', re.IGNORECASE)
    for match in re.finditer(pattern, text):
        s = match.start()
        e = match.end()
        # print('Found "%s" at %d:%d' % (text[s:e], s, e))
        bill_addr = text[s:e]
        bill_addr = re.sub(r':|\)', '', bill_addr)
        bill_addr = re.sub("\n", " ", bill_addr)
        if bill_addr:
            worksheet.write(row, 0, "Billing Address")
            worksheet.write(row, 1, bill_addr)
            row += 1
            break

    # SHIPPING ADDRESS
    pattern = re.compile(r'SHIP((.|\n)*?)INDIA', re.IGNORECASE)
    for match in re.finditer(pattern, text):
        s = match.start()
        e = match.end()
        # print('Found "%s" at %d:%d' % (text[s:e], s, e))
        ship_addr = text[s:e]
        ship_addr = re.sub(r':|\)', '', ship_addr)
        ship_addr = re.sub("\n", " ", ship_addr)
        # print(ship_addr)
        if ship_addr:
            worksheet.write(row, 0, "Shipping Address")
            worksheet.write(row, 1, ship_addr)
            row += 1
            break

    # NAME
    pattern = re.compile(r'\bNAME\b', re.IGNORECASE)
    for match in re.finditer(pattern, text):
        s = match.start()
        e = match.end()
        print('Found "%s" at %d:%d' % (text[s:e], s, e))
        i = e
        while i < length and text[i] != ':' and text[i] != '>' and text[i] != '.': i += 1
        i += 1
        if text[i] == '-': i += 1
        while text[i] == ' ' or text[i] == '\n': i += 1
        name = ''
        while i < length and text[i] != '\n':
            name += text[i]
            i += 1
        if name: print('NAME : ' + name)

    # ADDRESS
    pattern = re.compile(r'\bADDRESS\b', re.IGNORECASE)
    for match in re.finditer(pattern, text):
        s = match.start()
        e = match.end()
        print('Found "%s" at %d:%d' % (text[s:e], s, e))
        i = e
        while i < length and text[i] != ':' and text[i] != '>' and text[i] != '.' and text[i] != '\n': i += 1
        i += 1
        if text[i] == '-': i += 1
        while text[i] == ' ' or text[i] == '\n': i += 1
        addr = ''
        while i < length and text[i] != '\n':
            addr += text[i]
            i += 1
        if addr: print('ADDRESS : ' + addr)

    # Date
    pattern = re.compile(r"\b(DATE)|(DATED)\b", re.IGNORECASE)
    date_list = []
    for match in re.finditer(pattern, text):
        s = match.start()
        e = match.end()
        print('Found "%s" at %d:%d' % (text[s:e], s, e))
        i = e
        while i < length and text[i] != ':' and text[i] != '>' and text[i] != '.' and text[i] != '=' and not (
                text[i] >= '0' and text[i] <= '9'): i += 1
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

        # print(date,"hello")
        x = re.findall(
            r"^(?:(?:31(\/|-|\.)(?:0?[13578]|1[02]))\1|(?:(?:29|30)(\/|-|\.)(?:0?[13-9]|1[0-2])\2))(?:(?:1[6-9]|[2-9]\d)?\d{2})$|^(?:29(\/|-|\.)0?2\3(?:(?:(?:1[6-9]|[2-9]\d)?(?:0[48]|[2468][048]|[13579][26])|(?:(?:16|[2468][048]|[3579][26])00))))$|^(?:0?[1-9]|1\d|2[0-8])(\/|-|\.)(?:(?:0?[1-9])|(?:1[0-2]))\4(?:(?:1[6-9]|[2-9]\d)?\d{2})$",
            date)
        # print(x,"hii")
        Date = None
        for data in date.split(" "):
            for fmt in ('%d-%m-%Y', '%d.%m.%Y', '%d/%m/%Y'):
                try:
                    DATE = datetime.strptime(data, fmt).date()
                    if DATE: Date = DATE
                    print(DATE)
                except ValueError as e:
                    # print(e)
                    pass

        # Ndate = dateutil.parser.parse(date,fuzzy_with_tokens=True)
        # print(Ndate)

        if Date:
            print('DATE_EXT : ' + str(Date))
        elif date:
            print('DATE : ' + date)
        if Date:
            try:
                Date = datetime.datetime.strptime(Date, '%Y-%m-%d').strftime('%d/%m/%y')
            except:
                pass
        if Date and (str(Date) not in date_list):
            date_list.append(str(Date))
        elif date and (date not in date_list):
            date_list.append(str(date))

    if (len(date_list) == 1):
        worksheet.write(row, 0, "Ordered Date ")
        worksheet.write(row, 1, date_list[0])
        row += 1
    elif (len(date_list) > 1):
        min_date = ""
        max_date = ""
        if (date_list[0] < date_list[len(date_list) - 1]):
            min_date, max_date = date_list[0], date_list[len(date_list) - 1]
        else:
            min_date, max_date = date_list[len(date_list) - 1], date_list[0]
        worksheet.write(row, 0, "Ordered Date ")
        worksheet.write(row, 1, min_date)
        row += 1
        worksheet.write(row, 0, "Delivery Date ")
        worksheet.write(row, 1, min_date)
        row += 1

    # table

    # table_description_keywords = ["material","description","hsn","goods","quantity",
    #         "sgst","cgst","rate","discount","amount","per","gst",
    #         "price off","unit","taxable","code"]
    # table_cut_off = 7
    # table_found = False
    # for line in text.split('\n'):
    #     orig_line = line
    #     if(not table_found):
    #         line = line.lower()
    #         for char in line:
    #             if(not (char>='a' and char<='z')):
    #                 if(char!=' '):
    #                     line = line.replace(char,'')
    #         while("  " in line):
    #             line = line.replace('  ',' ');

    #         #now the text only contains small character
    #         # for processing design according to you
    #         table_key_finder = 0
    #         for key in table_description_keywords:
    #             if(key in line):
    #                 table_key_finder += 1
    #         if(table_key_finder>=table_cut_off):
    #             col = 0
    #             table_found = True
    #             if(("sr" in line) or ("sl" in line)):
    #                 worksheet.write(row,col,"Sl no")
    #                 col+=1
    #             if "material" in line:
    #                 worksheet.write(row,col,"Sl no")
    #                 col+=1
    #             worksheet.write(row,col,"Description of goods")
    #             col+=1
    #             worksheet.write(row,col,"HSN/SAC")
    #             col+=1
    #             worksheet.write(row,col,"Qty")
    #             col+=1
    #             if "unit" in line:
    #                 worksheet.write(row,col,"Unit price")
    #                 col+=1
    #             print(line)
    #     if(table_found):
    #         table_found = False

    # table

    table_description_keywords = ["material", "description", "hsn", "goods", "quantity", "qty", "sgst", "cgst", "rate",
                                  "discount", "amount", "per", "gst",
                                  "price off", "unit", "taxable", "code", "mrp", "item", "total"]
    table_cut_off = 3
    table_found = False
    table_end = False

    row += 1
    print('ROW: ' + str(row))
    for line in text.split('\n'):
        orig_line = line
        counter = 0
        if (not table_found):
            for str_p in table_description_keywords:
                pattern = r'\b(' + str_p + r')\b'
                pattern = re.compile(pattern, re.IGNORECASE)
                match = re.search(pattern, line)
                if (match): counter += 1

            if (counter > table_cut_off):
                print("TABLE FOUND----------------------------------------------------------------")
                worksheet.write(row, 0, "TABLE FOUND----------------------------------------------------------------")
                row += 1
                print(line)
                table_found = True
                line = line.strip()
                col = 0
                for word in line.split():
                    worksheet.write(row, col, word)
                    print('WORD: ' + word)
                    col += 1

        elif not table_end:
            pattern = re.compile(r'in word', re.IGNORECASE)
            match = re.search(pattern, line)
            if (match):
                print("TABLE END--------------------------------------------------------------------")
                worksheet.write(row, 0, "TABLE END----------------------------------------------------------------")
                row += 2
                table_end = True
                line = line.strip()
                print(line)
                worksheet.write(row, col, line)
                break
            print(line)
            line = line.strip()
            col = 0
            for word in line.split():
                worksheet.write(row, col, word)
                col += 1
        else:
            col = 0
            line = line.strip()
            print(line)
            worksheet.write(row, col, line)

        if table_found: row += 1

    print('ROW: ' + str(row))


'''        
    print("Hello from here")
    pattern = (r'\d{4}.\d{2}.\d{2}')
    for match in re.finditer(pattern,text):
    	print("Hello from pattern")
    	s = match.start()
    	e = match.end()
    	print('Found "%s" at %d:%d' % (text[s:e], s, e))

'''

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

row = 0


def main_fun():
    pdf_path = ""
    try:
        pdf_path = sys.argv[1]
    except IndexError:
        print("Enter file path")
        sys.exit(0)
    images = pdf_to_images(pdf_path)
    page_counter = 0

    # print(pdf_path,'\n\n\n\n\n')
    global bold_font

    workbook_path = pdf_path[:-4] + '.xlsx'
    workbook = xlsxwriter.Workbook(workbook_path)
    worksheet = workbook.add_worksheet()
    bold_font = workbook.add_format({'bold': 1})

    for i in images:
        text = extract_data_per_page(i)
        if (page_counter == 0): first_page_content(text, worksheet)
        page_counter += 1
        print(text)
    workbook.close()
    # print(images)


main_fun()
