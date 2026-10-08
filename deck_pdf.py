import win32com.client, os
src = r"C:\Users\idari\Downloads\Hermes\PertamaxGreen95_PreDefense_Template.pptx"
pdf = r"C:\Users\idari\Downloads\Hermes\deck.pdf"
app = win32com.client.Dispatch("PowerPoint.Application")
pres = app.Presentations.Open(src, WithWindow=False)
pres.SaveAs(pdf, 32)  # ppSaveAsPDF
pres.Close(); app.Quit()
d = open(pdf,"rb").read(8)
print("ok", os.path.getsize(pdf), d)
