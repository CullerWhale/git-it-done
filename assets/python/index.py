import datetime as dt
from docx2pdf import convert
from docxtpl import DocxTemplate, InlineImage


# create a document object
doc = DocxTemplate("./assets/python/photoExhibitionTemplate2.docx")

processedImages = [
    {'images' :  InlineImage(doc, './assets/images/img1.jpg'), 'captions': 'image one'},
    {'images' : InlineImage(doc, './assets/images/img2.jpg'), 'captions': 'image two'},
    {'images' : InlineImage(doc,'./assets/images/img3.jpg'), 'captions': 'image three'}]

todayStr = dt.datetime.now().strftime("%d-%b-%Y")

# create context to pass data to template
context = {
    "reportDtStr": todayStr,
    "processedImages": processedImages
}

# context['result1'] = InlineImage(doc, 'images/img1.jpg')
# context['result2'] = InlineImage(doc, 'images/img2.jpg')

# render context into the document object
doc.render(context)

# save the document object as a word file
reportWordPath = './assets/reports/report_{0}.docx'.format(todayStr)
doc.save(reportWordPath)

# convert the word file as pdf file
convert(reportWordPath, reportWordPath.replace(".docx", ".pdf"))