from docxtpl import DocxTemplate, RichText
tpl = DocxTemplate('Template1.docx')
rt = RichText()
rt.add('a rich text', style='List Bullet 3')
rt.add(' with ')
rt.add('some italic', italic=True)
rt.add(' and ')
rt.add('some violet', color='#ff00ff')
rt.add(' and ')
rt.add('some striked', strike=True)
rt.add(' and ')
rt.add('some small', size=14)
rt.add(' or ')
rt.add('big', size=60)
rt.add(' text.')
rt.add('\nYou can add an hyperlink, here to ')
rt.add('google', url_id=tpl.build_url_id('http://google.com'))
rt.add('\nEt voilà ! ')
rt.add('\n1st line')
rt.add('\n2nd line')
rt.add('\n3rd line')
rt.add('\n\n<cool>')
rt.add('\nFonts :\n', underline=True)
rt.add('Arial\n', font='Arial')
rt.add('Courier New\n', font='Courier New')
rt.add('Times New Roman\n', font='Times New Roman')
rt.add('\n\nHere some')
rt.add('superscript', superscript=True)
rt.add(' and some')
rt.add('subscript', subscript=True)

rt_embedded = RichText('an example of ', style='List Bullet 3')
rt_embedded.add(rt)

context = {
    'example': rt_embedded,
}
tpl.render(context)
# text = '\tНастоящим подтверждаю, что 8.05.2020 на АЗС №31052 в результате сбоя, пролив топлива АИ-92 объемом 30,12 л.' \
#        ' на сумму 1249,98 руб. был учтен в ССО №58 как пролив объемом 12,21 л. на сумму 506,72 руб. из-за чего' \
#        ' возникло расхождение по счетчикам ТРК. Оплата производилась по банковской карте №****4240 по RRN №' \
#        ' 012923616562.\n\tТакже подтверждаю, что 8.05.2020 на АЗС №30052 в результате сбоя, пролив топлива АИ-92' \
#        ' объемом 12,21 л. на сумму 506 руб. не был учтен в ССО №58. Оплата производилась за наличный расчет.\n\tВ связи' \
#        ' с чем, прошу:\n'
# text2 = '\t\t— Считать пролив топлива АИ-92 по банковской карте №****4240 от 8.05.2020 объемом 30,12 л.' \
#        ' на сумму 1249,98 руб.\n\t\t— Пролив топлива АИ-92 объемом 12,21 л. на сумму 506 руб. считать проливом за' \
#        ' наличный расчет.\n\t\t— Выручку по банковским картам считать равной 135738,51 руб.\n\t\t— Выручку за наличный' \
#        ' расчет считать равной 231838 руб.\n\t\t— Произвести внесение на кассе №1 по АСУ+ККМ на сумму 506 руб.'
#
# rt = RichText()
# rt2 =RichText()
#
# rt.add(text)
# rt2.add(text2)
# #
# context = {'txt': 'АЗС №31051 ССО №516', 'data': rt, 'data1': rt2}
# doc.render(context)


tpl.save('new1.docx')
