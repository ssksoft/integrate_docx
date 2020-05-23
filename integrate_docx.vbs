Set word_obj = CreateObject("Word.Application")
word_obj.Visible = True

cover_filepath = "C:\Users\sasat\Desktop\integrate_docx\cover.docx"
Set cover = word_obj.Documents.Open(cover_filepath)

cover.SaveAs2 "C:\Users\sasat\Desktop\integrate_docx\integrate.docx"
Set integrate = cover

word_obj.Selection.EndKey(6)
word_obj.Selection.InsertBreak(7)
word_obj.Selection.InsertFile("C:\Users\sasat\Desktop\integrate_docx\contents.docx")

integrate.Save
integrate.Close
word_obj.Quit
