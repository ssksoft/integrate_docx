Set word_obj = CreateObject("Word.Application")
word_obj.Visible = True

Dim docxs(2)
docxs(0) = "C:\Users\sasat\Desktop\integrate_docx\sample1.docx"
docxs(1) = "C:\Users\sasat\Desktop\integrate_docx\sample2.docx"

Set src1_obj = word_obj.Documents.Open(docxs(0))
Set src2_obj = word_obj.Documents.Open(docxs(1))

' src1_obj.Close
' word_obj.Quit
Set integrated_docx_obj = word_obj.Documents.Add
Set tmp = word_obj.Selection
integrated_docx_obj.SaveAs2 "C:\Users\sasat\Desktop\integrate_docx\integrate.docx"