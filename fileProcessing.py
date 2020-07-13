from docx.shared import Inches, Pt
import docx
import datetime
from pytils import dt
import os
import win32com.client as com
import shutil
import re


def textforlist(textinput):
    textlst = []
    for line in textinput:
        textlst.append(line.strip())
    return textlst


def delete_paragraph(paragraph):
    p = paragraph._element
    p.getparent().remove(p)
    p._p = p._element = None


def get_from_bodylist_azsnum(blist):
    for i in range(len(blist[0].split())):
        if 'АЗС' in blist[0].split()[i]:
            if re.sub("\\D", "", blist[0].split()[i + 1]).isdigit():
                return re.sub("\\D", "", blist[0].split()[i + 1])
            elif re.sub("\\D", "", blist[0].split()[i + 2]).isdigit():
                return re.sub("\\D", "", blist[0].split()[i + 2])


def get_from_bodylist_ssonum(blist):
    for i in range(len(blist[0].split())):
        if 'ССО' in blist[0].split()[i]:
            if re.sub("\\D", "", blist[0].split()[i + 1]).isdigit():
                return re.sub("\\D", "", blist[0].split()[i + 1])
            elif re.sub("\\D", "", blist[0].split()[i + 2]).isdigit():
                return re.sub("\\D", "", blist[0].split()[i + 2])


def del_empty_paragraphs(doc, btext):
    empty_parag = 45 - 3 - len(btext) - 2
    for p in doc.paragraphs:
        if p.text == '' and empty_parag != 0:
            delete_paragraph(p)
            empty_parag -= 1


def get_current_date():
    return str(dt.ru_strftime("%d %B %Y" + ' г.', inflected=True))


def create_docx_file_from_bodylist(blist):
    doc = docx.Document('template.docx')

    bottext = False
    AZSnum = get_from_bodylist_azsnum(blist)
    SSOnum = get_from_bodylist_ssonum(blist)
    ACTnum = get_number_act()
    nowdate = get_current_date()
    numline = 0

    for i in range(len(doc.paragraphs)):
        if i == 0:
            hd = doc.paragraphs[i]
            hd.paragraph_format.space_after = Pt(10)
            doc.paragraphs[i].text = 'АКТ № ' + ACTnum
            doc.paragraphs[i].style = 'Normal'
            doc.paragraphs[i].alignment = 1
            doc.paragraphs[i].runs[0].bold = True
            doc.paragraphs[i].runs[0].font.name = 'Times New Roman'
            doc.paragraphs[i].runs[0].font.size = Pt(14)
            continue
        if i == 1:
            hd2 = doc.paragraphs[i]
            hd2.paragraph_format.space_after = Pt(10)
            doc.paragraphs[i].style = 'Normal'
            doc.paragraphs[i].add_run('АЗС №' + AZSnum + '	ССО №' + SSOnum + '\t\t\t\t\t\t\t\t' + nowdate)
            doc.paragraphs[i].runs[0].font.name = 'Times New Roman'
            doc.paragraphs[i].runs[0].font.size = Pt(12)
            continue
        if i > 1 and not bottext and numline < len(blist):
            p2 = doc.paragraphs[i]
            run2 = p2.add_run('\t' + blist[numline])
            p2.paragraph_format.alignment = 3
            if blist[numline].find('В связи с чем') != -1:
                bottext = True
                p2.paragraph_format.space_before = Pt(0)
                p2.paragraph_format.space_after = Pt(0)
            else:
                p2.paragraph_format.space_before = Pt(0)
                p2.paragraph_format.space_after = Pt(0)
            p2.paragraph_format.line_spacing = Pt(0)
            run2.font.name = 'Times New Roman'
            run2.font.size = Pt(12)
            numline += 1
            continue
        if i > 1 and bottext and numline < len(blist):
            p3 = doc.paragraphs[i]
            doc.paragraphs[i].style = 'List Paragraph'
            run3 = p3.add_run('— ' + blist[numline])
            run3.font.name = 'Times New Roman'
            run3.font.size = Pt(12)
            p3.paragraph_format.left_indent = Inches(0.8)
            p3.paragraph_format.alignment = 0
            p3.paragraph_format.line_spacing = Pt(0)
            p3.paragraph_format.space_before = Pt(0)
            p3.paragraph_format.space_after = Pt(0)
            numline += 1
            continue


def get_path():
    return r'E:\tested\Акты'


def get_number_act():
    filelist = os.listdir(get_path())
    numacts = []
    nextnumact = ''
    for i in range(len(filelist)):
        if '.docx' in filelist[i]:
            numacts.append(int(filelist[i].split('_')[0][3:]))
    numacts.sort(reverse=True)
    for i in range(len(numacts)):
        # print(i, numacts[i])
        if 0 <= (numacts[i] - numacts[i + 1]) <= 2:
            nextnumact = str(numacts[i] + 1)
            break
    return nextnumact

def createdocxnpdffiles(lst):

    doc = docx.Document('template.docx')




    bottext = False
    textlist = lst
    AZSnum = get_from_bodylist_azsnum(lst)
    SSOnum = get_from_bodylist_ssonum(lst)
    ACTnum = get_number_act()

    now = datetime.datetime.now()
    nowdate = str(dt.ru_strftime("%d %B %Y" + ' г.', inflected=True))
    numline = 0

    for i in range(len(doc.paragraphs)):
        if i == 0:
            hd = doc.paragraphs[i]
            hd.paragraph_format.space_after = Pt(10)
            doc.paragraphs[i].text = 'АКТ № ' + ACTnum
            doc.paragraphs[i].style = 'Normal'
            doc.paragraphs[i].alignment = 1
            doc.paragraphs[i].runs[0].bold = True
            doc.paragraphs[i].runs[0].font.name = 'Times New Roman'
            doc.paragraphs[i].runs[0].font.size = Pt(14)
            continue
        if i == 1:
            hd2 = doc.paragraphs[i]
            hd2.paragraph_format.space_after = Pt(10)
            doc.paragraphs[i].style = 'Normal'
            doc.paragraphs[i].add_run('АЗС №' + AZSnum + '	ССО №' + SSOnum + '\t\t\t\t\t\t\t\t' + nowdate)
            doc.paragraphs[i].runs[0].font.name = 'Times New Roman'
            doc.paragraphs[i].runs[0].font.size = Pt(12)
            continue
        if i > 1 and not bottext and numline < len(textlist):
            p2 = doc.paragraphs[i]
            run2 = p2.add_run('\t' + textlist[numline])

            p2.paragraph_format.alignment = 3
            if textlist[numline].find('В связи с чем') != -1:
                bottext = True
                p2.paragraph_format.space_before = Pt(0)
                p2.paragraph_format.space_after = Pt(0)
            else:
                p2.paragraph_format.space_before = Pt(0)
                p2.paragraph_format.space_after = Pt(0)
            p2.paragraph_format.line_spacing = Pt(0)
            run2.font.name = 'Times New Roman'
            run2.font.size = Pt(12)
            numline += 1
            continue
        if i > 1 and bottext and numline < len(textlist):
            p3 = doc.paragraphs[i]
            doc.paragraphs[i].style = 'List Paragraph'
            run3 = p3.add_run('— ' + textlist[numline])
            run3.font.name = 'Times New Roman'
            run3.font.size = Pt(12)
            p3.paragraph_format.left_indent = Inches(0.8)
            p3.paragraph_format.alignment = 0
            p3.paragraph_format.line_spacing = Pt(0)
            p3.paragraph_format.space_before = Pt(0)
            p3.paragraph_format.space_after = Pt(0)
            numline += 1
            continue

    del_empty_paragraphs(doc, textlist)

    filenamedocx = 'local_acts\\' + 'Акт' + ACTnum + '_' + 'АЗС' + AZSnum + '_' + 'ССО' + SSOnum + '.docx'
    filenamepdf = 'local_acts\\' + 'Акт' + ACTnum + '_' + 'АЗС' + AZSnum + '_' + 'ССО' + SSOnum + '.pdf'
    endfiledocx = get_path() + '\\' + 'Акт' + ACTnum + '_' + 'АЗС' + AZSnum + '_' + 'ССО' + SSOnum + '.docx'
    endfilepdf = get_path() + '\\' + 'Акт' + ACTnum + '_' + 'АЗС' + AZSnum + '_' + 'ССО' + SSOnum + '.pdf'

    doc.save(filenamedocx)
    #
    # wdFormatPDF = 17
    #
    # in_file = os.path.abspath(filenamedocx)
    # out_file = os.path.abspath(filenamepdf)
    #
    #
    # word = com.DispatchEx('word.application')
    # doccon = word.Documents.Open(in_file)
    # doccon.SaveAs(out_file, FileFormat=wdFormatPDF)
    # doccon.Close()
    # word.Quit()
    #
    #
    # shutil.copyfile(filenamedocx, endfiledocx)
    # shutil.copyfile(filenamepdf, endfilepdf)
    # return 'Акт' + ACTnum + '_' + 'АЗС' + AZSnum + '_' + 'ССО' + SSOnum + '.pdf'
