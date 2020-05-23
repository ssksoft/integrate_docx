Set word_obj = CreateObject("Word.Application")
word_obj.Visible = True

' Dim docxs(2)
' docxs(0) = "C:\Users\sasat\Desktop\integrate_docx\sample1.docx"
' docxs(1) = "C:\Users\sasat\Desktop\integrate_docx\sample2.docx"

cover_filepath = "C:\Users\sasat\Desktop\integrate_docx\cover.docx"

Set cover = word_obj.Documents.Open(cover_filepath)

cover.SaveAs2 "C:\Users\sasat\Desktop\integrate_docx\integrate.docx"

integrate = cover

' integrate.Selection.InsertBreak=7
word_obj.Selection.EndKey(6)
word_obj.Selection.InsertBreak(7)
' Selection.InsertFile FileName:="sample1.docx", Range:="", _ // ファイル結合
' ConfirmConversions:=False, Link:=False, Attachment:=False