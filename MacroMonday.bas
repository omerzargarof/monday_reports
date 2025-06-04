Attribute VB_Name = "Module1"
Sub CleanAndCopyHeaders()

    ' --- עימוד ---

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim header As Boolean
    Dim currentHeader As String
    Dim cellValue As String
    Dim pos As Integer
    Dim userMailResponse As VbMsgBoxResult
    Dim userCCResponse  As VbMsgBoxResult

    ' Set the worksheet to the active sheet
    Set ws = ActiveSheet
    header = False
    
    ' Insert new column before column A
    ws.Columns("A").Insert Shift:=xlToRight
    'ws.Cells(1, 1).Value = "New Column"
    
    ' Find the last row with data in column B (originally column A)
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    i = 1
    
    While i <= lastRow
        ' Check if the cell in column B is empty
        If ws.Cells(i, 2).Value = "" Then
            ws.Rows(i).Delete
            i = i - 1
            currentHeader = ""
            lastRow = lastRow - 1
        ' Check if the cell in column B contains a header
        ElseIf ws.Cells(i, 2).Value = "Started By" Then
            If header = False Then
                header = True
            Else
                ws.Rows(i).Delete
                i = i - 1
                lastRow = lastRow - 1
            End If
        ElseIf ws.Cells(i, 2).Value <> "" And currentHeader = "" Then
            currentHeader = ws.Cells(i, 2).Value
        Else
            ' Copy the current header to the cell in column B
            ws.Cells(i, 2).Value = currentHeader
        End If
        
        ' Split the cell content at "-" and save the first part in the new column (now A)
        cellValue = ws.Cells(i, 2).Value
        pos = InStr(cellValue, "-")
        If pos > 0 Then
            ws.Cells(i, 1).Value = Left(cellValue, pos - 1)
            ws.Cells(i, 2).Value = Mid(cellValue, pos + 1)
        Else
            ws.Cells(i, 1).Value = ""
        End If

        i = i + 1
    Wend
    
    ' Column B (originally A)
    ' Find the last row with data in column B
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    i = 2
    While i <= lastRow
        ' Check if the cell in column B is empty
        If ws.Cells(i, 4).Value = "" Then
            ws.Rows(i).Delete
            i = i - 1
            lastRow = lastRow - 1
        End If
        i = i + 1
    Wend
    
    ' Set Headers
    ws.Cells(2, 1).Value = "Client"
    ws.Cells(2, 2).Value = "Task"
    ws.Cells(2, 3).Value = "Name"
    ws.Columns.AutoFit
    ws.Rows(2).Font.Bold = True
    
    
    ' --- מיון ---

    ' מציאת השורה האחרונה בעמודה
   ' Dim lastRow As Long
    'lastRow = ws.Cells(ws.Rows.Count, "C2").End(xlUp).Row
    
    ' טווח הנתונים למיון
    Dim rng As Range
    Set rng = ws.Range("A2:I" & lastRow) ' שנה את הטווח לפי הצורך

    ' מיון לפי העמודה הראשונה בעדיפות ראשונה
    rng.Sort Key1:=ws.Range("F3"), Order1:=xlAscending, header:=xlYes

    ' מיון לפי העמודה השנייה בעדיפות שנייה
    rng.Sort Key1:=ws.Range("D3"), Order1:=xlAscending, header:=xlYes

    ' מיון לפי העמודה השלישית בעדיפות שלישית
 '   rng.Sort Key1:=ws.Range("A3"), Order1:=xlAscending, header:=xlYes
    
    ' מיון לפי העמודה השלישית בעדיפות רביעית
 '   rng.Sort Key1:=ws.Range("B3"), Order1:=xlAscending, header:=xlYes
    
    
    ' מיון לפי העמודה השלישית בעדיפות חמישית
    rng.Sort Key1:=ws.Range("C3"), Order1:=xlAscending, header:=xlYes
    
    
    
    ' --- צביעה ---

     Dim currentDate As String
    Dim previousDate As String
    Dim colorFlag As Boolean
    colorFlag = False

    For i = 3 To lastRow ' הנחה שהנתונים מתחילים בשורה 3
        currentDate = ws.Cells(i, 4).Value
        If currentDate <> previousDate Then
            colorFlag = Not colorFlag
        End If

        If colorFlag Then
            ws.Rows(i).Interior.Color = RGB(211, 211, 211) ' צבע אפור בהיר
        Else
            ws.Rows(i).Interior.Color = RGB(255, 255, 255) ' צבע לבן
        End If

        previousDate = currentDate
    Next i
    
    'ConvertTimeToDecimal
    Dim timeValue As Date
    Dim cell As Range
    ' Find the last row with data in column A
    lastRow = ws.Cells(ws.Rows.Count, "H").End(xlUp).Row
    
    ' Loop through the cells in column A
    For Each cell In ws.Range("H3:H" & lastRow)
       If Not IsEmpty(cell.Value) And cell.Value <> "" Then
        timeValue = cell.Value
        ' Convert time to decimal format
        decimalTime = Hour(timeValue) + Minute(timeValue) / 60
        cell.Value = decimalTime
        End If
    Next cell
    
    '----סינון תאריכים---
    'Dim ws As Worksheet
    'Dim lastRow As Long
    Dim monthNumber As Integer
    Dim yearNumber As Integer
    Dim currentMonth As Integer
    Dim currentYear As Integer
    Dim dateValue As Date
    
    ' Set the worksheet
   ' Set ws = ThisWorkbook.Sheets("Sheet1") ' שנה את השם לשם הגיליון שלך
    
    ' Get the last row in column 4
    lastRow = ws.Cells(ws.Rows.Count, 4).End(xlUp).Row
    
    ' Get the current year
    currentMonth = Month(Date)
    ' Prompt the user for the month number
    monthNumber = InputBox("עבור איזה חודש נרצה את הדוח? (1-12)", "בחירת חודש", currentMonth)
    
    ' Get the current year
    currentYear = Year(Date)
    ' Prompt the user for the year with the current year as default
    yearNumber = InputBox("אנא הזן שנה (למשל " & currentYear & ")", "בחירת שנה", currentYear)

   If monthNumber <> 0 Then
      ' Loop through the dates in column 4 and hide rows that don't match the selected month and year
      For i = lastRow To 3 Step -1
        dateValue = ws.Cells(i, 4).Value
         If Month(dateValue) <> monthNumber Or Year(dateValue) <> yearNumber Then
            ws.Rows(i).Delete
        End If
    Next i
   End If
   
    
    
    ' --- מיילים ---

' הצגת הודעת שאלה למשתמש
userMailResponse = MsgBox("האם לשלוח את מיילים למופיעים בדוח?", vbYesNo + vbQuestion, "אישור שליחה")
If userMailResponse = vbYes Then
userCCResponse = MsgBox("האם לשלוח עותק לראש צוות?", vbYesNo + vbQuestion, "שליחת לראש צוות כCC")

    Dim OutlookApp As Object
    Dim OutlookMail As Object
    'Dim ws As Worksheet
    'Dim rng As Range
    'Dim cell As Range
    Dim startRow As Long
    Dim endRow As Long
    Dim currentName As String
    'Dim lastRow As Long
     Dim filePath As String
      Dim emailAddress As String
    Dim emailAddressCC As String
       Dim emailWb As Workbook
     ' Set the file path for the email addresses
      Dim headerText As String
      
    filePath = "\\bb8\D\Group Files\WMS\Softwares\Macro_Monday\Contact_List.xlsx"
    
    ' הגדרת גליון העבודה והטווח
    'Set ws = ThisWorkbook.Sheets("Sheet1") ' שנה את שם הגליון לפי הצורך
    'lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row ' מציאת השורה האחרונה בעמודה C
    
    ' יצירת אובייקט Outlook
    Set OutlookApp = CreateObject("Outlook.Application")
    
    ' הגדרת שורת הכותרת
    headerText = "<tr>"
    For Each cell In ws.Range(ws.Cells(2, 1), ws.Cells(2, 8))
        headerText = headerText & "<th>" & cell.Value & "</th>"
    Next cell
    headerText = headerText & "</tr>"
    
     ' נניח שהשמות מתחילים בשורה 3
     startRow = 3
    
    Do While startRow <= lastRow
        currentName = ws.Cells(startRow, "C").Value
        endRow = startRow
        
        ' מציאת הטווח של השם הנוכחי
        Do While endRow <= lastRow And ws.Cells(endRow, "C").Value = currentName
            endRow = endRow + 1
        Loop
        
        ' הגדרת הטווח לשליחה
        Set rng = ws.Range(ws.Cells(startRow, 1), ws.Cells(endRow - 1, 8))
        
        ' בניית גוף המייל כ-HTML
        bodyText = "<html><body><h3>נתוני מאנדיי עבור " & currentName & "</h3><table border='1'>" & headerText
        For Each cell In rng
            If cell.Column = 1 Then
                bodyText = bodyText & "<tr>"
            End If
            bodyText = bodyText & "<td>" & cell.Value & "</td>"
            If cell.Column = rng.Columns.Count Then
                bodyText = bodyText & "</tr>"
            End If
        Next cell
        bodyText = bodyText & "</table></body></html>"
        
         ' Open the workbook with email addresses
    Set emailWb = Workbooks.Open(filePath) ' קובץ מהכתובת filePath
    Set emailWs = emailWb.Sheets(1) ' גיליון מספר 1
        
        ' חיפוש כתובת המייל המתאימה ב-Excel
        emailAddress = ""
        For i = 2 To emailWs.Cells(emailWs.Rows.Count, "A").End(xlUp).Row
            If emailWs.Cells(i, 1).Value = currentName Then ' Assuming names are in column A
                emailAddress = emailWs.Cells(i, 2).Value ' Assuming emails are in column B
                Exit For
            End If
        Next i
      
If userCCResponse = vbYes Then
    emailAddressCC = ""
        For i = 2 To emailWs.Cells(emailWs.Rows.Count, "A").End(xlUp).Row
            If emailWs.Cells(i, 1).Value = "TeamLeader" Then ' Assuming names are in column A
                emailAddressCC = emailWs.Cells(i, 2).Value ' Assuming emails are in column B
                Exit For
            End If
        Next i
    End If

        
        ' יצירת המייל
    If emailAddress <> "" Then
        Set OutlookMail = OutlookApp.CreateItem(0)
        With OutlookMail
            .To = emailAddress ' שנה את כתובת הנמען לפי הצורך
           .CC = emailAddressCC
            .Subject = "נתונים עבור " & currentName
            .HTMLBody = bodyText
            .Send
        End With
    End If
          
        ' עדכון השורה ההתחלתית לשם הבא
        startRow = endRow
    Loop
    
    ' ניקוי אובייקטים
    Set OutlookMail = Nothing
    Set OutlookApp = Nothing
    
    
    ' סגירת קובץ ה-Excel עם כתובות המייל
emailWb.Close SaveChanges:=False
    
  '  Else
  '  MsgBox "שליחת המיילים בוטלה.", vbInformation, "ביטול"
End If
End Sub


