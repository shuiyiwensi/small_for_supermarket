from PyPDF2 import PdfFileWriter, PdfFileReader
import re
from openpyxl import load_workbook
import os
import csv
import pdfkit
import datetime


def get_fist_page_of_pdf(path):
    inputpdf = PdfFileReader(path)
    os.remove(path)
    output = PdfFileWriter()
    output.addPage(inputpdf.getPage(0))
    with open(path, "wb") as outputStream:
        output.write(outputStream)

def print_html_to_pdf(htmls,file_name):
    config = pdfkit.configuration(wkhtmltopdf=r"D:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe")
    options = {
        'page-size': 'A4',
        'margin-top': '0in',
        'margin-right': '0in',
        'margin-bottom': '0in',
        'margin-left': '0in',
        'encoding': "UTF-8",
        'custom-header': [
            ('Accept-Encoding', 'gzip')
        ],
        #'cookie': [
        #    ('cookie-name1', 'cookie-value1'),
         #   ('cookie-name2', 'cookie-value2'),
        #],
        'outline-depth': 1,
    }
    pdfkit.from_file(htmls, file_name, options=options, configuration=config)

def get_arabic_name(chinese_name, translatelist1):
    if chinese_name in translatelist1:
        arabic_name = translatelist1[chinese_name]
        return arabic_name
    else:
        print(chinese_name)
        print(chinese_name+" 没有阿拉伯语翻译")
        # add_Chinese_name_to_dictionary()
        exit()


def build_dictionary():
    translatelist1 = {}
    with open('dictionary_csv/dictionary.csv', 'r',encoding="utf-16") as csvfile:
        csv_reader = csv.reader(csvfile, delimiter='\t')
        next(csv_reader)
        for row in csv_reader:
            translatelist1[row[0]] = row[1]
    return translatelist1


def build_chinese_name_to_ingredients_dictionary():
    chinese_name_to_ingredients_dictionary = {}
    chinese_name_to_row_number_dictionary = {}
    with open('dictionary_csv/dictionary_name_formula.csv', 'r',encoding="utf-16") as csvfile:
        csv_reader = csv.reader(csvfile, delimiter='\t')
        next(csv_reader)
        for row in csv_reader:
            chinese_name_to_ingredients_dictionary[row[0]] = row[1]
            chinese_name_to_row_number_dictionary[row[0]] = row[6]
    return chinese_name_to_ingredients_dictionary, chinese_name_to_row_number_dictionary

def get_one_blank(fontsize,Chinese_Name,Output_Product_name,Output_Product_Ingredients_name,Our_Product_serving_size,Output_Production_date,Output_Expire_date,Out_Product_place):
    content_in_one_blank='''
    <td align="right">
        <font size="'''+str(fontsize)+'''">
        '''+Chinese_Name+'''<br>
        '''+Output_Product_name+'''<br>
        '''+Output_Product_Ingredients_name+'''<br>
        '''+Our_Product_serving_size+'''<br>
        '''+Output_Production_date+'''<br>
        '''+Output_Expire_date+'''<br>
        '''+Out_Product_place+'''<br>
        </font>
    </td>
    '''
    return content_in_one_blank

def generate_content_in_a_row(content_in_one_blank,number):
    senctence=''''''
    for i in range(number):
        senctence=senctence+content_in_one_blank
    message ='''
    <td>
    '''+senctence+'''
    </td>
    '''
    return message

def generate_content_from_row(one_row_content,number):
    senctence = ''''''
    for i in range(number):
        senctence = senctence +'''<tr>
        ''' + one_row_content + '''
        </tr>
        '''
    return senctence


def generate_html(senctence, margins_size):
    message='''
    <html> <head> <style>
    @page {
    size: 21cm 29.7cm;
    margin: '''+str(margins_size)+'''mm '''+str(margins_size)+'''mm '''+str(margins_size)+'''mm '''+str(margins_size)+'''mm; /* change the margins as you want them to be. */
    }
    table
    {font - family: arial, sans - serif;border - collapse: collapse;width: 100 %;}
    td, th
    {border: 1px solid  # dddddd;text - align: left;padding: 8px;}
    </style> </head>
    <body> <table>
    '''+senctence+'''
    </table> </body> </html>
    '''
    return message


def write2html(fontsize,Ingredients_length_in_one_row, Ingredients_length_in_one_col, Chinese_Name,Ingredients_name_list, product_weight_number, product_weight, product_time, expire_time, Chinese_ProductPlace):
    stringlength = 80
    margins_size = 0.1
    row_num = Ingredients_length_in_one_col
    col_num = Ingredients_length_in_one_row
    translatelist1 = build_dictionary()
    Arabic_Name = get_arabic_name(Chinese_Name, translatelist1)
    Arabic_ProductPlace = get_arabic_name(Chinese_ProductPlace, translatelist1)
    Arabic_product_weight = get_arabic_name(product_weight, translatelist1)
    Arabic_Ingredients_name_list = [get_arabic_name(Ingredients_name_list[n], translatelist1) for n in
                                    range(len(Ingredients_name_list))]

    The_name = 'الاسم: '  # 商品名
    The_Ingredients_name = 'المكونات: '  # 配方
    The_serving_size = 'حجم الحصة: '  # serving size
    Product_date = 'تاريخ الإنتاج: '  # 生产日期
    expire_date = 'تاريخ انتهاء الصلاحية: '  # 过期日期
    expire_date = 'تاريخ الانتهاء:'
    Output_Product_name = The_name + Arabic_Name
    Output_Product_Ingredients_name = The_Ingredients_name
    for n in range(len(Arabic_Ingredients_name_list)):
        Output_Product_Ingredients_name = Output_Product_Ingredients_name + Arabic_Ingredients_name_list[n]
        if n != len(Arabic_Ingredients_name_list) - 1:
            Output_Product_Ingredients_name = Output_Product_Ingredients_name + '، '
    Our_Product_serving_size = The_serving_size + product_weight_number + ' ' + Arabic_product_weight
    Output_Production_date = Product_date + product_time + ' '
    Output_Expire_date = expire_date + expire_time + ' '
    Out_Product_place = Arabic_ProductPlace
    print(Chinese_Name.rjust(stringlength, ' '))
    print(Output_Product_name.ljust(stringlength, ' '))
    print(Output_Product_Ingredients_name.ljust(stringlength, ' '))
    print(Our_Product_serving_size.ljust(stringlength, ' '))
    print(Output_Production_date.ljust(stringlength, ' '))
    print(Output_Expire_date.ljust(stringlength, ' '))
    print(Out_Product_place.ljust(stringlength, ' '))
    content_in_one_blank=get_one_blank(fontsize,Chinese_Name,Output_Product_name,Output_Product_Ingredients_name,Our_Product_serving_size,Output_Production_date,Output_Expire_date,Out_Product_place)
    one_row_content=generate_content_in_a_row(content_in_one_blank, Ingredients_length_in_one_row)
    out_put_content=generate_content_from_row(one_row_content,Ingredients_length_in_one_col)
    html_message=generate_html(out_put_content,margins_size)
    temp_path='temp_word/' + Chinese_Name + '-' + product_weight_number + product_weight + '.html'
    f = open(temp_path, 'w',encoding='utf-8')
    f.write(html_message)
    f.close()
    out_file = os.path.abspath('Out_put/' + Chinese_Name + '-' + product_weight_number + product_weight + '.pdf')
    print_html_to_pdf(temp_path,out_file)
    get_fist_page_of_pdf(out_file)




def add_Ingredients_to_dictionary(Chinese_Name,list):
    #wb = load_workbook(filename='dictionary_csv/dictionary.xlsx')
    #old_sheet = wb.get_sheet_by_name('Sheet2')
    #rows=(Chinese_Name,list,None,None,None,None,9)
    #old_sheet.append(rows)
    #wb.save(filename='dictionary_csv/dictionary.xlsx')

    rows = Chinese_Name+ "\t"+ list+ "\t"+ "\t"+ "\t"+ "\t"+ "\t"+ str(9)+"\n"
    fd = open('dictionary_csv/dictionary_name_formula.csv', 'a',encoding="utf-16")
    fd.write(rows)
    fd.close()


def get_Ingredients_name_list(list,Chinese_Name,Chinese_name_to_Ingredients_dictionary):
    if Chinese_Name in Chinese_name_to_Ingredients_dictionary:
        if list == None or list=='':
            list = Chinese_name_to_Ingredients_dictionary[Chinese_Name]
        Ingredients_name_list = re.split('，|,', list)
    else:
        if list!=None:
            add_Ingredients_to_dictionary(Chinese_Name,list)
            Ingredients_name_list = re.split('，|,', list)
        else:
            print(Chinese_Name+" 需要输入配方")
            exit()
    return Ingredients_name_list


def get_row_number(rownumber,Chinese_name_to_row_number_dictionary):
    if rownumber == None:
        rownumber = Chinese_name_to_row_number_dictionary[Chinese_Name]
    if rownumber == None:
        rownumber = 9
    return rownumber


def get_product_weight_number(list):
    product_weight_number_list = re.split('g|kg|ml|L|l|KG', list)
    return product_weight_number_list[0]


def get_product_weight(product_weight_number,list):
    product_weight_list = re.split(product_weight_number, list)
    return product_weight_list[1]


Chinese_name_to_Ingredients_dictionary, Chinese_name_to_row_number_dictionary = build_chinese_name_to_ingredients_dictionary()
with open('input_excel/input.csv', 'r',encoding="utf-16") as csvfile:
    csv_reader = csv.reader(csvfile, delimiter='\t')
    next(csv_reader)
    print(csv_reader)
    for row in csv_reader:
        if row[0]=='':
            break;
        Chinese_Name =row[0]
        Ingredients_name_list = get_Ingredients_name_list(row[1],Chinese_Name,Chinese_name_to_Ingredients_dictionary)
        product_weight_number = get_product_weight_number(row[2])
        product_weight =  get_product_weight(product_weight_number,row[2])

        product_time_temp=datetime.datetime.strptime(row[3], "%m/%d/%Y").date()
        expire_time_temp=datetime.datetime.strptime(row[4], "%m/%d/%Y").date()
        product_time = product_time_temp.strftime("%Y/%m/%d")
        expire_time =expire_time_temp.strftime("%Y/%m/%d")

        Chinese_ProductPlace = row[5]
        print(Ingredients_name_list)
        print(product_weight_number+product_weight)
        print(product_time)
        print(expire_time)
        print(Chinese_ProductPlace)
        rownumber=get_row_number(int(row[6]),Chinese_name_to_row_number_dictionary)
        write2html(4,5,rownumber, Chinese_Name, Ingredients_name_list, product_weight_number,
                   product_weight, product_time, expire_time, Chinese_ProductPlace)