from docx import Document

def fillindocx(fname, oname, filld):
    document = Document(fname)

    for para in document.paragraphs:
        for i in filld:
            if i in para.text:
                txt = para.text
                otxt = txt.replace(i, fill[i])
                para.text = otxt

    document.save(oname)

gname = "Jack Duggan"
cname = "MERCURY ENGINEERING FRANCE"
cdate = "16/10/2020"
cadde = "MERCURY ENGINEERING, Dublin, Ireland"

fill = {"gname":gname, "cname":cname, "cdate":cdate, "cadde":cadde }

fillindocx(r'siteadmin\wordtemplating\Paris Movement Template.docx', r'siteadmin\wordtemplating\demo.docx', fill)
