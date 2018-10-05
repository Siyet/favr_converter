# -*- coding: utf-8 -*-
from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfpage import PDFTextExtractionNotAllowed
from pdfminer.pdfinterp import PDFResourceManager
from pdfminer.pdfinterp import PDFPageInterpreter
from pdfminer.layout import LAParams, LTTextBoxHorizontal, LTFigure
from pdfminer.converter import PDFPageAggregator
import argparse
import comtypes.client
import ConfigParser
import os
import sys
import re
import zipfile
import ctypes
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from PIL import ImageFont
from PIL import Image
from PIL import ImageDraw
# Loading the pyPdf Library
from PyPDF2 import PdfFileWriter, PdfFileReader, PdfFileMerger
from subprocess import call, check_output, Popen, PIPE
import subprocess
import json
import shutil
from errors import error
from lxml import etree

_ = lambda s: s.encode('cp1251')

IS_EXE = os.path.exists(os.path.join(os.path.realpath(os.path.dirname(sys.argv[0])), 'library.zip'))

CONPRINT_PATH = os.path.realpath(os.path.dirname(sys.argv[0]))
COMTYPES_GEN_PATH = os.path.join(CONPRINT_PATH, 'library\\comtypes\\gen') if IS_EXE else 'c:\\Lib\\Python27\\sites-packages\\comtypes\\gen'

HELP_ARG_FUNCTION = u"""Исполняемое действие: convert - конвертировать документ DOCX в PDF/A-1,
print - внедряет в полученный PDF элементы графической визуализации регистрационных данных и ЭП и выводит на печать,
sign_stamps_generate - сгенерировать все штампы ЭП, по списку signatures.json,
regstamp - сгенерировать картинку с регистрационными данными"""

HELP_ARG_INPUT = u"""Исходный файл, при выполнении функции convert должен быть в формате DOCX,
а при выполнении функции print, соответственно, должен иметь формат PDF/A-1"""

FONT = ImageFont.truetype('c:\\Windows\\Fonts\\times.ttf', 10)
FONT_SMALL = ImageFont.truetype('c:\\Windows\\Fonts\\ARIALN.ttf', 9)
FONTBD = ImageFont.truetype('c:\\Windows\\Fonts\\timesbd.ttf', 10)

def get_sign_list():
    slp = os.path.join(CONPRINT_PATH, 'signatures.json')
    if os.path.exists(slp):
        with open(slp) as data_file:
            data = json.load(data_file)
            return data['signatures']
    else:
        error(99)

SIGNATURES_LIST = get_sign_list()


def create_parser():
    parser = argparse.ArgumentParser(add_help=True, version='0.7')
    parser.add_argument('function', help=HELP_ARG_FUNCTION)
    parser.add_argument('input', help=HELP_ARG_INPUT)
    parser.add_argument('-p', '--pdf-out', dest='pdf')
    parser.add_argument('-i', '--ini-out', dest='ini')
    parser.add_argument('-d', '--date', dest='date')
    parser.add_argument('-n', '--number', dest='num')
    parser.add_argument('-s', '--signature', dest='sign')
    # parser.add_argument('-x', '--description-file', dest='xml')

    return parser


def parse_pdf(sourcePDF, outputToINI=True):
    fp = open(sourcePDF, 'rb')
    parser = PDFParser(fp)
    document = PDFDocument(parser)
    if not document.is_extractable:
        raise PDFTextExtractionNotAllowed
    rsrcmgr = PDFResourceManager()
    laparams = LAParams()
    device = PDFPageAggregator(rsrcmgr, laparams=laparams)
    interpreter = PDFPageInterpreter(rsrcmgr, device)
    sign_count = 1
    coord = {}
    if outputToINI:
        ini = namespace.ini if namespace.ini else 'coordinates.ini'
        ini = ini if os.path.isabs(ini) else os.path.join(CONPRINT_PATH, ini)
        if os.path.exists(ini):
            os.remove(ini)
    else:
        ini = {}
    for idx, page in enumerate(PDFPage.create_pages(document)):
        interpreter.process_page(page)
        layout = device.get_result()
        page_params = {'num': idx+1, 'height': page.attrs['MediaBox'][3]}

        sign_count, tmp_coord = parse_obj(layout._objs, page_params, ini, sign_count)
        coord.update(tmp_coord)

    if not outputToINI:
        return coord


def output_coord(section_name, data, output):
    if type(output) != dict:
        if not output.has_section(section_name):
            output.add_section(section_name)
        for key in data.keys():
            output.set(section_name, key, data[key])
    else:
        for key in data.keys():
            if section_name not in output.keys():
                output[section_name] = {}
            output[section_name][key] = data[key]
        return output


def parse_obj(lt_objs, page_params, ini, sign_count):

    flag = type(ini) == str
    if flag:
        Config = ConfigParser.ConfigParser()
        output = Config
    else:
        output = ini
    # loop over the object list
    for obj in lt_objs:
        # if it's a textbox, print text and location
        if isinstance(obj, LTTextBoxHorizontal):
            k = 297.0 / page_params['height']
            if u"¸¸" in obj.get_text():
                pass
                # print obj.get_text()[:-3]
            if u"¸¸" in obj.get_text()[0:2]:
                section_name = 'sign_'+str(sign_count)
                data = {'page_num': page_params['num'],
                        'x': int(k * obj.bbox[0]),
                        'y': int(297 - k * obj.bbox[1] - 5)}
                if flag and output.has_section(section_name):
                    sign_count += 1
                output_coord(section_name, data, output)
            elif u"¸_" in obj.get_text()[0:2]:
                section_name = 'reg_info'
                data = {'page_num': page_params['num'],
                        'x': int(k * obj.bbox[0]),
                        'y': int(297 - k * obj.bbox[1] - 5)}
                output_coord(section_name, data, output)
            elif u"¸¸¸" in obj.get_text()[-6:]:
                section_name = 'sign_'+str(sign_count)
                ogt = re.search(u"[А-Я][А-Яа-я\-]+", obj.get_text())
                if ogt is None:
                    error(90)
                data = {'last_name': ogt.group(0)}
                if flag:
                    data['last_name'] = _(data["last_name"])
                if flag and output.has_section(section_name):
                    sign_count += 1
                output_coord(section_name, data, output)

        # if it's a container, recurse
        elif isinstance(obj, LTFigure):
            parse_obj(obj._objs, page_params, ini, sign_count)
    if flag:
        cfgfile = open(ini, 'a')
        Config.write(cfgfile)
        cfgfile.close()
        ini = {}
    return sign_count, ini


def convert_docx_to_pdf(namespace):
    ifn = namespace.input if os.path.isabs(namespace.input) else os.path.join(CONPRINT_PATH, namespace.input)
    if not os.path.exists(ifn):
        error(93)
    tmpp = os.path.join(CONPRINT_PATH, '~tmp.docx')
    shutil.copy(ifn, tmpp)
    ifn = tmpp
    ofn = namespace.pdf if namespace.pdf else 'output.pdf'
    ofn = ofn if os.path.isabs(ofn) else os.path.join(CONPRINT_PATH, ofn)
    word = comtypes.client.CreateObject('Word.Application')
    doc = word.Documents.Open(ifn)
    try:
        doc.ExportAsFixedFormat(OutputFileName=ofn, ExportFormat=17, IncludeDocProps=True, UseISO19005_1=True)
    except AttributeError:
        if word.Version == '12.0':
            error(91, False)
        else:
            error(92, False, word.Version)
        if IS_EXE:
            with zipfile.ZipFile('library.zip', 'r') as zin:
                with zipfile.ZipFile('_library.zip', 'w') as zout:
                    for item in zin.infolist():
                        if 'comtypes/gen/' not in item.filename:
                            buffer = zin.read(item.filename)
                            zout.writestr(item, buffer)
            os.remove('library.zip')
            os.rename('_library.zip', 'library.zip')
        else:
            os.rmdir(COMTYPES_GEN_PATH)
        sys.exit(9192)
    finally:
        doc.Close()
        word.Quit()

    parse_pdf(ofn)


def insert_stamps_and_print(namespace, coord, spath):

    reg_flag = True if namespace.date and namespace.num else False
    sign_flag = True if namespace.sign else False
    if reg_flag:
        if "reg_info" in coord.keys():
            rtp = os.path.join(CONPRINT_PATH, 'reg_template.jpg')
            if not os.path.exists(rtp):
                error(9601)
            reg_img = Image.open(rtp)
            reg_draw = ImageDraw.Draw(reg_img)
            # 06.05.2016 г.; ТБ-01-21/3345
            num = namespace.num.decode('cp1251')
            num = num[1:-1] if num[:1] == "\'" and num[-1:] == "\'" else num
            text = u"%s г.          %s" % (namespace.date.decode('cp1251'), num)
            reg_draw.text((24, 2), text, (0, 0, 0), font=FONT)
            reg_img.save(os.path.join(CONPRINT_PATH, "reg_stamp.jpg"))
        else:
            error(9501)
    if sign_flag:
        if 'sign' in ' '.join(coord.keys()):
            stp = os.path.join(CONPRINT_PATH, 'sign_template.jpg')
            if not os.path.exists(stp):
                error(9602)
            sign_img = Image.open(stp)
            sign_draw = ImageDraw.Draw(sign_img)
            sign_data = []
            for signature in SIGNATURES_LIST:
                for key in coord.keys():
                    if key[:4] == 'sign' and signature['last_name'] == coord[key]['last_name']:
                        sign_data.append(signature)
                        sign_data[-1]['key'] = key
            if len(sign_data) < 1:
                error(97)
            for sign in sign_data:
                sign_draw.text((59, 50), sign['license_num'], (0, 0, 0), font=FONT)
                sign_draw.text((47, 62), sign['full_name'], (0, 0, 0), font=FONTBD)
                sign_draw.text((67, 74), sign['period_expires'], (0, 0, 0), font=FONT)
                sign_img.save(os.path.join(CONPRINT_PATH, "%s_stamp.jpg"%(sign['key'])))
        else:
            error(9502)
    if reg_flag or sign_flag:
        tpath = os.path.join(CONPRINT_PATH, '_tmp.pdf')
        _canvas = canvas.Canvas(tpath, pagesize=A4)

        cur_page = 1
        for section in coord.keys():
            x = float(coord[section]['x']) * 2.83
            y = (297 - float(coord[section]['y'])) * 2.83
            if section == "reg_info" and reg_flag:
                rsp = os.path.join(CONPRINT_PATH, 'reg_stamp.jpg')
                _canvas.drawImage(rsp, x, y - 16)
                os.remove(rsp)
            elif section[:4] == "sign" and sign_flag:
                while int(coord[section]['page_num']) > cur_page:
                    _canvas.showPage()
                    cur_page += 1
                ssp = os.path.join(CONPRINT_PATH, '%s_stamp.jpg'%(section))
                _canvas.drawImage(ssp, x, y - 91)
                os.remove(ssp)
        pdfSource = PdfFileReader(open(spath, "rb"))
        page_count = pdfSource.getNumPages()
        while page_count >= cur_page:
            _canvas.showPage()
            cur_page += 1
        _canvas.save()

        cur_page = 0
        pdfOutput = PdfFileWriter()
        pdfTmp = PdfFileReader(open(tpath, "rb"))
        while page_count > cur_page:
            page = pdfTmp.getPage(cur_page)
            page.mergePage(pdfSource.getPage(cur_page))
            pdfOutput.addPage(page)
            cur_page += 1
        opath = os.path.join(CONPRINT_PATH, '_2print.pdf')
        outputStream = open(opath, "wb")
        pdfOutput.write(outputStream)
        outputStream.close()
        call(r'C:\Windows\System32\cmd.exe /C %s'%(opath))
    else:
        call(r'C:\Windows\System32\cmd.exe /C %s'%(spath))


def check_container(cpath, result):
    # Перебираем все файлы и папки контейнера
    for item in os.listdir(cpath):
        # Пропускаем все, что не является файлом
        if not os.path.isfile(os.path.join(cpath, item)):
            continue
        try:
            f = open(os.path.join(CONPRINT_PATH, os.path.join(cpath, item)))
            schema_root = etree.XML(f.read())
            f.close()
            result['tmp']['xfp'] = os.path.join(cpath, item)
        except etree.XMLSyntaxError:
            continue

        # # Пропускаем, если не читается как ini-файл
        # conf = ConfigParser.SafeConfigParser()
        # try:
        #     conf.read(os.path.join(cpath, item))
        # except ConfigParser.MissingSectionHeaderError:
        #     continue
        # result['tmp']['ifp'] = os.path.join(cpath, item)  # ini_file_path
        # # Проверяем наличие разделов
        #
        # if _(u"АДРЕСАТЫ") not in conf.sections():
        #     result['Errors']['e87'] = u'В ini-файле контейнера отсутствует раздел АДРЕСАТЫ!'
        # if _(u"ПИСЬМО КП ПС СЗИ") not in conf.sections():
        #     result['Errors']['e84'] = u'В ini-файле контейнера отсутствует раздел ПИСЬМО КП ПС СЗИ!'
        # if _(u"ФАЙЛЫ") not in conf.sections():
        #     result['Errors']['e83'] = u'В ini-файле контейнера отсутствует раздел ФАЙЛЫ!'
        # if len(conf.items(_(u"АДРЕСАТЫ"))) < 1:
        #     result['Errors']['e86'] = u'В ini-файле пустой раздел АДРЕСАТЫ!'
        # if len(conf.items(_(u"ПИСЬМО КП ПС СЗИ"))) < 1:
        #     result['Warnings']['w98'] = u'В ini-файле пустой раздел ПИСЬМО КП ПС СЗИ!'
        # if len(conf.items(_(u"ФАЙЛЫ"))) < 1:
        #     result['Errors']['e82'] = u'В ini-файле пустой раздел ФАЙЛЫ!'
        # if not result['Errors'].get('e87', False) and not result['Errors'].get('e86', False):
        #     e85 = True
        #     for item in conf.items(_(u"АДРЕСАТЫ")):
        #         if item[1] == "A_RVR_S~MEDOGU":
        #             e85 = False
        #     if e85:
        #         result['Errors']['e85'] = u'В контейнере в разделе АДРЕСАТЫ отсутствует Федеральное агенство водных ресурсов!'
        # if not result['Errors'].get('e83', False) and not result['Errors'].get('e82', False):
        #     for item in conf.items(_(u"ФАЙЛЫ")):
        #         if item[1][-4:] == '.xml' and not result['tmp'].get('xfp', False):  # xfp - xml_file_path
        #             result['tmp']['xfp'] = os.path.join(cpath, item[1])
        #         elif item[1][-4:] == '.xml' and result['tmp'].get('xfp', False):
        #             result['Warnings']['w97'] = u"В ini-файле, в разделе ФАЙЛЫ указано более одно файла с раширением xml, за файл описания принят указанный выше всех в списке."
        #         elif item[1][-8:] == '.edc.zip' and not result['tmp'].get('zfp', False):  # zfp - zip_file_path
        #             result['tmp']['zfp'] = os.path.join(cpath, item[1])
        #         elif item[1][-8:] == '.edc.zip' and result['tmp'].get('zfp', False):
        #             result['Warnings']['w96'] = u"В ini-файле, в разделе ФАЙЛЫ указано более одно файла с раширением edc.zip, за архив контейнера принят указанный выше всех в списке."
        #     if not result['tmp'].get('xfp', False):
        #         result['Errors']['e81'] = u"В ini-файле в разделе ФАЙЛЫ отсутствует ссылка на файл описания!"
        #     if not result['tmp'].get('zfp', False):
        #         result['Warning']['w93'] = u"В ini-файле в разделе ФАЙЛЫ отсутствует ссылка на архив с документом!"
        break
    else:
        result['Errors']['e81'] = u'В папке с электронным ссобщением отсутстсвует файл описания электронного сообщения!!'


def xml_validate_and_get_data(xsd_filename, key, result):
    if not os.path.exists(os.path.join(CONPRINT_PATH, xsd_filename)):

        error(79)
    with open(os.path.join(CONPRINT_PATH, xsd_filename)) as f:
        schema_root = etree.XML(f.read())
    schema = etree.XMLSchema(schema_root)
    xml_parser = etree.XMLParser(schema=schema)
    try:
        with open(result['tmp'][key], 'r') as f:
            return etree.fromstring(f.read(), xml_parser)
    except etree.XMLSyntaxError as e:
        if key == 'xfp':
            result['Errors']['e78'] = u"Файл описания сообщения не прошел валидацию xsd-схемой: %s" % e
        elif key == 'pfp':
            result['Errors']['e67'] = u"Файл описания документа не прошел валидацию xsd-схемой: %s" % e
        return None


if __name__ == '__main__':
    parser = create_parser()
    namespace = parser.parse_args()
    if namespace.function == 'convert':
        if namespace.input:
            convert_docx_to_pdf(namespace)
        else:
            error(94)
    elif namespace.function == 'print':
        if not namespace.input:
            error(94)
        spath = namespace.input if os.path.isabs(namespace.input) else os.path.join(CONPRINT_PATH, namespace.input)
        if not os.path.exists(spath):
            error(93)
        coordinates = parse_pdf(spath, False)
        insert_stamps_and_print(namespace, coordinates, spath)
    elif namespace.function == "check":
        if not namespace.input:
            error(94)
        cpath = namespace.input  # container path
        if not os.path.isabs(cpath) or not os.path.isdir(cpath):
            error(89)
        result = {'tmp': {}, 'Warnings': {}, 'Errors': {}}
        check_container(cpath, result)
        # TODO реализовать проверку ЭП контейнера при ее наличии
        xml_data = xml_validate_and_get_data('document.xsd', 'xfp', result)

        if len(result['Errors'].keys()) < 1 or (len(result['Errors'].keys()) == 1 and result['Errors'].get('e85', False)):
            p = xml_data.tag[:-13]  # get and save prefix
            header = xml_data.find(p+'header')
            if not header.attrib[p+'type'] == u'Транспортный контейнер':
                result['Errors']['e77'] = u"Указанный тип сообщения в файле описания:'%s', ожидается 'Транспортный контейнер'"%(header.attrib[p+'type'])
            # if result['tmp'].get('zfp', False) and os.path.join(cpath, xml_data.find(p+'container').find(p+'body').text) != result['tmp']['zfp']:
            #     result['Warnings']['w94'] = u"Наименование архива с документом в ini-файле отличается от наименования, указанного в файле описания сообщения. За основу взят последний."
            #     result['tmp']['zfp'] = xml_data.find(p+'container').find(p+'body').text
            # elif not result['tmp'].get('zfp', False):
            #     result['tmp']['zfp'] = xml_data.find(p+'container').find(p+'body').text
            if not result['tmp'].get('zfp', False):
                result['tmp']['zfp'] = os.path.join(cpath, xml_data.find(p+'container').find(p+'body').text)
            if zipfile.is_zipfile(result['tmp']['zfp']):

                z = zipfile.ZipFile(result['tmp']['zfp'], 'a')
                _zpath = os.path.join(CONPRINT_PATH, '~document.edc')
                if os.path.exists(_zpath):
                    shutil.rmtree(_zpath)
                z.extractall(_zpath)
                container = None
                if os.path.exists(os.path.join(_zpath, 'passport.xml')):
                    result['tmp']['pfp'] = os.path.join(_zpath, 'passport.xml')
                    container = xml_validate_and_get_data('passport.xsd', 'pfp', result)
                else:
                    result['Errors']['e68'] = u"В архиве с документом не найден файл описания passport.xml"
                if len(result["Errors"].keys()) < 1:
                    p = container.tag[:-9]
                    result['tmp']['dfname'] = container.find(p+'document').attrib[p+'localName']
                    result['tmp']['dpath'] = os.path.join(_zpath, result['tmp']['dfname'])
                    signs = []
                    reg_stamps = []
                    ie = 1
                    ie1 = 1
                    for author in container.find(p+'authors').findall(p+'author'):
                        registrationStamp = author.find(p + 'registration').find(p + 'registrationStamp')
                        reg_stamp = (os.path.join(_zpath, registrationStamp.attrib[p+'localName']),
                                     registrationStamp.find(p+'position').find(p+'page').text,
                                     registrationStamp.find(p+'position').find(p+'topLeft').find(p+'x').text,
                                     registrationStamp.find(p+'position').find(p+'topLeft').find(p+'y').text,
                                     registrationStamp.find(p+'position').find(p+'dimension').find(p+'w').text,
                                     registrationStamp.find(p+'position').find(p+'dimension').find(p+'h').text)
                        if os.path.exists(reg_stamp[0]):
                            reg_stamps.append(reg_stamp)
                        else:
                            result["Errors"]['e62'] = u"Не найден файл графической визуализации регистрационных данных, указанный в файле описания. Имя файла: %s" % registrationStamp.attrib[p+'localName']
                        for sign in author.findall(p + 'sign'):
                            lname = re.search(u"[А-Я][А-Яа-я\-]+", sign.find(p+'person').find(p+'name').text)
                            sign_fname = sign.find(p+'documentSignature').attrib[p+'localName']
                            signatureStamp = sign.find(p+'documentSignature').find(p+'signatureStamp')
                            signstamp = (os.path.join(_zpath, signatureStamp.attrib[p+'localName']),
                                         signatureStamp.find(p+'position').find(p+'page').text,
                                         signatureStamp.find(p+'position').find(p+'topLeft').find(p+'x').text,
                                         signatureStamp.find(p+'position').find(p+'topLeft').find(p+'y').text,
                                         signatureStamp.find(p+'position').find(p+'dimension').find(p+'w').text,
                                         signatureStamp.find(p+'position').find(p+'dimension').find(p+'h').text)
                            if os.path.exists(os.path.join(_zpath, sign_fname)) and os.path.exists(signstamp[0]):
                                signs.append((lname, os.path.join(_zpath, sign_fname), signstamp))
                            else:
                                if not os.path.exists(os.path.join(_zpath, sign_fname)):
                                    result['Errors']['e64'+ie] = u"Не найден файл подписи, указанный в файле описания. Имя файла: %s, подписант: %s" % (sign_fname, lname)
                                    ie += 1
                                else:
                                    result['Errors']['e63'+ie1] = u"Не найден файл штампа подписи, указанный в файле описания. Имя файла: %s, подписант: %s" % (sign_fname, lname)
                                    ie1 += 1
                    if os.path.exists(result['tmp']['dpath']) and len(signs) > 0 and len(result["Errors"]) < 1:
                        for sign in signs:
                            args = 'cmd /C "C:\\Lotus\\Notes\\AnswerSigner\\AnswerSigner.exe verifydetached 1 '
                            sof = os.path.join(_zpath, 'stdout')
                            args += str(result['tmp']['dpath']) + ' ' + sign[1] + ' > ' + sof
                            args += ' 2> ' + os.path.join(_zpath, 'stderr') + '"'
                            call(args)
                            try:
                                f = open(sof)
                                if f.read() == "OK\n":
                                    continue
                                e51 = u"ЭП, хранящаяся в файле %s, не является ЭП файла %s" % (str(result['tmp']['dpath']), sign[1])
                                result['Errors']['e51'] = e51 if not result['Errors'].get('e51', False) else str(result['Errors']['e51']) + '; ' + e51
                            except Exception as e:
                                e52 = u"При проверке ЭП, хранящейся в файле %s и относящейся к файлу %s, произошла следующая ошибка: %s" % (str(result['tmp']['dpath']), sign[1], e)

                        tpath = os.path.join(CONPRINT_PATH, '~tmp.pdf')
                        if os.path.exists(tpath):
                            os.remove(tpath)
                        _canvas = canvas.Canvas(tpath, pagesize=A4)
                        cur_page = 1
                        for rstamp in reg_stamps:
                            x = int(rstamp[2]) * 2.84
                            y = (297 - int(rstamp[3])) * 2.84
                            rsp = rstamp[0]
                            _canvas.drawImage(rsp, x, y - int(rstamp[5]) * 2.84, int(rstamp[4]) * 2.84, int(rstamp[5]) * 2.84, mask='auto')
                        for sstamp in signs:
                            while int(sstamp[2][1]) > cur_page:
                                _canvas.showPage()
                                cur_page += 1
                            x = int(sstamp[2][2]) * 2.84
                            y = (297 - int(sstamp[2][3])) * 2.84

                            rsp = sstamp[2][0][:-4] + '.jpg'

                            im = Image.open(sstamp[2][0])
                            bg = Image.new("RGB", im.size, (255,255,255))
                            bg.paste(im, im)
                            bg.save(rsp)
                            _canvas.drawImage(rsp, x, y - int(sstamp[2][5]) * 2.84, int(sstamp[2][4]) * 2.84,
                                              int(sstamp[2][5]) * 2.84)
                        pdfSource = PdfFileReader(open(result['tmp']['dpath'], "rb"))
                        page_count = pdfSource.getNumPages()
                        while page_count >= cur_page:
                            _canvas.showPage()
                            cur_page += 1
                        _canvas.save()
                        cur_page = 0
                        pdfOutput = PdfFileWriter()
                        pdfTmp = PdfFileReader(open(tpath, "rb"))
                        while page_count > cur_page:
                            page = pdfSource.getPage(cur_page)
                            page.mergePage(pdfTmp.getPage(cur_page))
                            pdfOutput.addPage(page)
                            cur_page += 1
                        a = result['tmp']['dpath'][:-4] + '(compiled)' + result['tmp']['dpath'][-4:]
                        if not os.path.exists(a):
                            outputStream = open(a, "wb")
                            pdfOutput.write(outputStream)
                            outputStream.close()
                            z.write(a, os.path.basename(a))
                        else:
                            result['Warnings']['w50'] = u"Проверка контейнера электронного сообщения была запущена повторно!"
                    else:
                        if not os.path.exists(result['tmp']['dpath']):
                            result['Errors']['e66'] = u'В архиве с документом не найден указанный в описании файл самого документа'
                        elif len(signs) < 1:
                            result['Errors']['e65'] = u'Не найдено ни одного файла электронной подписи документа'

                z.close()

                # TODO Найти документы, проверить у них подписи, сгенерить их варианты с наложенными картинками

            else:
                result['Errors']['e69'] = u"Указанный файл архива с документом не является zip-архивом. Путь к файлу:%s"%(result['tmp']['zfp'])
        rconf = ConfigParser.SafeConfigParser()
        for key in result.keys():
            if key == 'tmp' or len(result[key].keys()) < 1:
                continue
            else:
                rconf.add_section(key)
                for subkey in result[key].keys():
                    rconf.set(key, subkey, _(result[key][subkey]))
        cfgfile = open(os.path.join(cpath, 'result.ini'), 'w')
        rconf.write(cfgfile)
        cfgfile.close()

    elif namespace.function == "sign_stamps_generate":
        stp = os.path.join(CONPRINT_PATH, 'sign_template.jpg')
        if not os.path.exists(stp):
            error(9602)
        for sign in SIGNATURES_LIST:
            sign_img = Image.open(stp)
            sign_draw = ImageDraw.Draw(sign_img)
            sign_draw.text((59, 51), sign['license_num'], (0, 0, 0), font=FONT_SMALL)
            sign_draw.text((47, 62), sign['full_name'], (0, 0, 0), font=FONTBD)
            sign_draw.text((67, 74), sign['period_expires'], (0, 0, 0), font=FONT)
            sign_img.save(os.path.join(CONPRINT_PATH, u"%s_stamp.jpg" % (sign['last_name'])))
    elif namespace.function == "regstamp":
        output = namespace.input if namespace.input else 'reg_stamp.jpg'
        output = output if os.path.isabs(output) else os.path.join(CONPRINT_PATH, output)
        rtp = os.path.join(CONPRINT_PATH, 'reg_template.jpg')
        if not os.path.exists(rtp):
            error(9601)
        reg_img = Image.open(rtp)
        reg_draw = ImageDraw.Draw(reg_img)
        # 06.05.2016 г.; ТБ-01-21/3345
        num = namespace.num.decode('cp1251')
        num = num[1:-1] if num[:1] == "\'" and num[-1:] == "\'" else num
        # TODO разделить текст с датой и номером, позицию номера отсчитывать с правой стороны
        text = u"%s г.          %s"%(namespace.date.decode('cp1251'), num)
        reg_draw.text((24, 2), text, (0, 0, 0), font=FONT)
        reg_img.save(output)
    else:
        error(98)
