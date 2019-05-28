Option Compare Database
Option Explicit

Public Sub AutoUpdate()
Dim i As Integer

i = Access.Application.CurrentDb.TableDefs.Count


'birgün back end database yaparsam açıldığında sayfaların güncellenmesi için yapılan bir durum.

End Sub


Private Sub metraj()
Dim hakNo As Integer
hakNo = 1
DoCmd.SetParameter "hakNo", hakNo
DoCmd.OpenQuery "SRHAKEDNO", acViewNormal, acReadOnly
End Sub
Function AlreadyOpen(fileName As String, Path As String) As Boolean
Dim exl As Excel.Application
Dim wbook As Excel.Workbook

On Error Resume Next
Set exl = GetObject(, "Excel.Application")
Set wbook = exl.Workbooks(fileName)
If wbook Is Nothing Then
    AlreadyOpen = False
Else
    AlreadyOpen = True
End If

End Function
