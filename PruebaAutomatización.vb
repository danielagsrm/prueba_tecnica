Sub Corona_Scraper()
 
Dim IE As Object
Dim wb As Workbook
Dim tbl As HTMLTable
Dim trCounter As Integer
Dim tdCounter As Integer
Dim objWord
Dim objDoc
trCounter = 1
tdCounter = 1

On Error GoTo EndRoutine
Set Report = ThisWorkbook.Sheets("Reporte")
Set Sheet = ThisWorkbook.Sheets("Hoja1")

Set IE = CreateObject("InternetExplorer.Application")
IE.navigate "https://www.worldometers.info/coronavirus/"
IE.Visible = True
Do While IE.Busy And IE.readyState <> 4: DoEvents: Loop

Report.Range("B2").Value = Date
Report.Cells(5, 1).Value = IE.document.getElementsByClassName("maincounter-number")(0).outerText
Report.Cells(5, 2).Value = IE.document.getElementsByClassName("maincounter-number")(1).outerText
Report.Cells(5, 3).Value = IE.document.getElementsByClassName("maincounter-number")(2).outerText
Report.Cells(10, 1).Value = IE.document.getElementsByClassName("number-table-main")(0).outerText
Report.Cells(13, 1).Value = IE.document.getElementsByClassName("number-table")(0).outerText
Report.Cells(16, 1).Value = IE.document.getElementsByClassName("number-table")(1).outerText
Report.Cells(10, 3).Value = IE.document.getElementsByClassName("number-table-main")(1).outerText
Report.Cells(13, 3).Value = IE.document.getElementsByClassName("number-table")(2).outerText
Report.Cells(16, 3).Value = IE.document.getElementsByClassName("number-table")(3).outerText

'Table scraping
For Each tbl In IE.document.getElementsByTagName("table")
    For Each tr In tbl.getElementsByTagName("tr")
        For Each th In tr.getElementsByTagName("th")
            Sheet.Cells(trCounter, tdCounter).Value = th.innerText
            tdCounter = tdCounter + 1
        Next th
        For Each td In tr.getElementsByTagName("td")
            Sheet.Cells(trCounter, tdCounter).Value = td.innerText
            tdCounter = tdCounter + 1
        Next td
        tdCounter = 1
        trCounter = trCounter + 1
    Next tr
Next tbl

'Table Cleaning
LastRow = Sheet.Range("C" & Rows.Count).End(xlUp).Row
Sheet.Rows("230:" & LastRow).EntireRow.Delete
Sheet.Range("A1:O229").Sort Key1:=Range("A1"), Order1:=xlAscending, Header:=xlYes
Sheet.Rows("222:229").EntireRow.Delete
Report.Range("A1:C17").Copy

'Creates Word Document
Set objWord = CreateObject("Word.Application")
Set objDoc = objWord.Documents.Add
objWord.Visible = True
objWord.Activate
'Pastes Summary - Parrafo 1
objDoc.Paragraphs(objDoc.Paragraphs.Count).Range.Paste
objDoc.Paragraphs.Add
objDoc.Content.InsertAfter Text:="TABLE 1 - SOUTH AMERICA (SUMMARY)"
objDoc.Paragraphs.Add
'Pastes Tabla1
Sheet.Range("A1").AutoFilter Field:=15, Criteria1:="South America"
Sheet.Range("A1:N222").Copy
objDoc.Paragraphs(objDoc.Paragraphs.Count).Range.Paste
objDoc.Paragraphs.Add
'Pastes Tabla2
Sheet.Range("A1").AutoFilter Field:=15, Criteria1:="Europe"
Sheet.Range("A1:N222").Copy
Sheet.AutoFilterMode = False
Sheet.Range("Q1").PasteSpecial Paste:=xlPasteValues
Sheet.Range("Q1").AutoFilter Field:=3, Criteria1:="10", Operator:=xlTop10Items
LastRow = Sheet.Range("Q" & Rows.Count).End(xlUp).Row
Sheet.Range("Q1:AD" & LastRow).Copy
objDoc.Content.InsertAfter Text:="TABLE 2 - EUROPE (TOP 10)"
objDoc.Paragraphs.Add
objDoc.Paragraphs(objDoc.Paragraphs.Count).Range.Paste
objDoc.Paragraphs.Add
'Saving files
objDoc.SaveAs2 ("C:\Users\DGIRALD\Desktop\REPORTES CORONAVIRUS\INFORME CORONAVIRUS.doc")
objDoc.Close
Set wb = Workbooks.Add
Sheet.AutoFilterMode = False
Sheet.Range("Q:AD").ClearContents
Sheet.Copy Before:=wb.Sheets(1)
wb.SaveAs "C:\Users\DGIRALD\Desktop\REPORTES CORONAVIRUS\ANEXO 1_ TABLA CORONAVIRUS.xls"
'Cleaning for next run
Report.Range("B2, A5:C5, A10:C10, A13:C13, A16:C16").ClearContents

EndRoutine:
'Optimize Code
Application.ScreenUpdating = True
Application.EnableEvents = True

'Clear the Clipboard
Application.CutCopyMode = False

End Sub
