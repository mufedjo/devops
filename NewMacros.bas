Attribute VB_Name = "NewMacros"
Dim a, b, c, RF(8, 2) As Variant

Sub Autofill()
' Autofill Macro
Dim channel1 As Long
If Not Documents.Count = 0 Then ActiveDocument.Close

Set excel = CreateObject("Excel.Application")
channel1 = DDEInitiate(App:="Excel", Topic:="System")
DDEExecute channel:=channel1, Command:="[OPEN(" & Chr(34) _
 & "C:\Users\KMF\Documents\appt.xls" & Chr(34) & ")]"
DDETerminate channel:=channel1
channel1 = DDEInitiate(App:="Excel", Topic:="appt.xls")
rowarray = Split(DDERequest(channel1, "R2"), Chr(9))
Documents.Add Template:="exam", NewTemplate:=False, DocumentType:=0
With Selection
    patient = LCase(InputBox("Enter the patient's Lastname", " PATIENT'S LASTNAME"))
    i = 1
    Do Until rowarray(0) = Chr(9)
        i = i + 1
        Index = "R" + Trim(Str(i))
        Row = DDERequest(channel1, Index)
        If Left(Row, 1) = Chr(9) Then
            MsgBox ("Patient not found in the appointment book")
            Exit Sub
        End If
        rowarray = Split(Row, Chr(9))
        If InStr(LCase(rowarray(0)), patient) Then
            lname = rowarray(0)
            FName = rowarray(1)
            homephone = rowarray(4)
            cellphone = rowarray(6)
            reason = rowarray(8)
            dob = rowarray(20)
            choice = InputBox("Is this the patient?", FName & " " & lname & " born " & dob, lname & " " & FName)
            If Not Len(choice) = 0 Then
                If InStr(choice, " ") > 0 Then If InStr(LCase(choice), LCase(patient)) Then Exit Do
            End If
        End If
    Loop
    Selection.GoTo what:=wdGoToBookmark, Name:="DOS"
        texttotype = Date & ", " & Time
        GoSub DataEnteredCheck
DDETerminateAll
' Look for patient's picture in the photos folder
    Dim fs, workfolder
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set workfolder = fs.GetFolder("C:\Photos\")
    Count = workfolder.Files.Count
    photo = lname & FName
    Do While InStr(photo, " ") > 0
        n = InStr(photo, " ")
        photo = Left(photo, n - 1) & Mid(photo, n + 1)
    Loop
' if a picture is found, insert it on the exam form
    For Each file In workfolder.Files
        Found = IIf(InStr(LCase(file.Name), LCase(photo)), 1, 0)
        If Found = 1 Then
            photo = file.Name
            ActiveDocument.Bookmarks("Photo").Select
            Call NewCanvasPicture(photo)
            Exit For
        End If
    Next
    days = DateDiff("d", dob, Date)
    age = Int(days / 365)
    mo = Int((days Mod 365) / 30)
    da = (days Mod 365) Mod 30
     Selection.GoTo what:=wdGoToBookmark, Name:="Treatment"
        texttotype = "Explanations given and questions answered." + Chr(13) + " Spectacles prescription issued" + Chr(13) + "Follow-up in a year"
        GoSub DataEnteredCheck
    Selection.GoTo what:=wdGoToBookmark, Name:="Diagnosis"
        Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
        texttotype = "Eye examination within normal findings for age"
        GoSub DataEnteredCheck
    Selection.GoTo what:=wdGoToBookmark, Name:="Motility"
        texttotype = "No motility deficit"
        GoSub DataEnteredCheck
    Selection.GoTo what:=wdGoToBookmark, Name:="Align"
        texttotype = "Orthophoria"
        GoSub DataEnteredCheck
    Selection.GoTo what:=wdGoToBookmark, Name:="Medications"
        texttotype = "No significant family history"
        GoSub DataEnteredCheck
    Selection.GoTo what:=wdGoToBookmark, Name:="HP"
        texttotype = "Patient here for " & reason
        GoSub DataEnteredCheck
    Selection.GoTo what:=wdGoToBookmark, Name:="Name"
        texttotype = UCase(Left(FName, 1)) & Right(FName, Len(FName) - 1) & " " & UCase(lname)
        GoSub DataEnteredCheck
    Selection.GoTo what:=wdGoToBookmark, Name:="birthdate"
        texttotype = dob
        GoSub DataEnteredCheck
    Selection.GoTo what:=wdGoToBookmark, Name:="telephone"
        texttotype = IIf(cellphone = "", homephone, cellphone)
        GoSub DataEnteredCheck
    Selection.GoTo what:=wdGoToBookmark, Name:="age"
        texttotype = age & "Y, " & mo & "M, " & da & "D"
        GoSub DataEnteredCheck
Selection.GoTo what:=wdGoToBookmark, Name:="Page"
    texttotype = "One"
    GoSub DataEnteredCheck
GoTo Endroutine
End With
DataEnteredCheck:
        With Selection
            .Font.Name = "Arial"
            .Font.Size = 10
            .Expand unit:=wdCell
            b = Selection.Text
            If Asc(Selection.Text) = 13 Then .TypeText Text:=texttotype
        End With
    Return
Endroutine:
months = Array("jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec")
monthsnumeric = Array("01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
For i = 1 To 12
    If LCase(Left(dob, 3)) = months(i - 1) Then
        m = monthsnumeric(i - 1)
        Exit For
    End If
Next
y = Right(dob, 4)
d = Val(Mid(dob, 5, Len(dob) - 4 - InStr(1, dob, "-")))
d = IIf(d < 10, "0" & Trim(Str(d)), Trim(Str(d)))
birthdate = m + d + y
patient = LCase(lname & FName & "_" & Right(Date, 4) & "-" & Left(Date, InStr(Date, "/") - 1) & "-" & Mid(Date, InStr(Date, "/") + 1, Len(Date) - 5 - InStr(Date, "/")))
' Delete all back-up files
Set workfolder = fs.GetFolder("C:\Users\KMF\Documents\")
fileCount = workfolder.Files.Count
For Each file In workfolder.Files
    If Left(LCase(file.Name), InStr(file.Name, ".") - 1) = patient Then
        RecentFiles.Add (file)
        RecentFiles(1).Open
    End If
Next
savefile = "C:\Users\KMF\Documents\Patients\" + patient & ".docx"
ActiveDocument.SaveAs2 FileName:=savefile, FileFormat:=wdFormatDocumentDefault, Password:="1156"
For Each file In workfolder.Files
    If InStr(1, LCase(file.Name), "backup") Then file.Delete
Next
excel.Quit
End Sub
Sub NewCanvasPicture(photo)
 Dim shpCanvas As Shape
 
 'Add a drawing canvas to the active document
 Set shpCanvas = ActiveDocument.Shapes _
 .AddCanvas(Left:=21, Top:=45, _
 Width:=100, Height:=80)
 
 'Add a graphic to the drawing canvas
 shpCanvas.CanvasItems.AddPicture _
 FileName:="C:\Photos\" + photo, LinkToFile:=False, SaveWithDocument:=True, Top:=0, Left:=0, Width:=100, Height:=80
End Sub
Sub same()
'
' same Macro
'
'
    Selection.GoTo what:=wdGoToBookmark, Name:="Medications"
    Selection.TypeText Text:="Same as most recent exam"
    Selection.GoTo what:=wdGoToBookmark, Name:="Face_R"
    Selection.TypeText Text:="Same as most recent exam"
    Selection.GoTo what:=wdGoToBookmark, Name:="Face_L"
    Selection.TypeText Text:="Same as most recent exam"
    Selection.GoTo what:=wdGoToBookmark, Name:="CorneaR"
    Selection.TypeText Text:="Same as most recent exam"
    Selection.GoTo what:=wdGoToBookmark, Name:="CorneaOU"
    Selection.TypeText Text:="Unchanged"
    Selection.GoTo what:=wdGoToBookmark, Name:="CorneaL"
    Selection.TypeText Text:="Same as most recent exam"
    Selection.GoTo what:=wdGoToBookmark, Name:="Lens_OU"
    Selection.TypeText Text:="Unchanged"
    Selection.GoTo what:=wdGoToBookmark, Name:="Retina_R"
    Selection.TypeText Text:="Same as most recent exam"
        Selection.GoTo what:=wdGoToBookmark, Name:="Vitreous"
    Selection.TypeText Text:="Unchanged"
Selection.GoTo what:=wdGoToBookmark, Name:="Retina_L"
    Selection.TypeText Text:="Same as most recent exam"
    Selection.GoTo what:=wdGoToBookmark, Name:="Macula_R"
    Selection.TypeText Text:="Unchanged"
    Selection.GoTo what:=wdGoToBookmark, Name:="Macula_L"
    Selection.TypeText Text:="Unchanged"
    ActiveWindow.ActivePane.SmallScroll Down:=-6
    Selection.GoTo what:=wdGoToBookmark, Name:="Diagnosis"
    Selection.TypeText Text:="Same as most recent exam"
    Selection.GoTo what:=wdGoToBookmark, Name:="Treatment"
    Selection.TypeText Text:= _
        "Explanations and recommendations given to patient."
    Selection.TypeParagraph
    Selection.TypeText Text:= _
        "Follow-up in 1 year for complete dilated examination."
End Sub
Sub Eyeglasses()
    ' Eyeglasses Macro
    '
'Dim RF(8, 2)
GoTo main
pickalens:
    Dim channel1 As Long, bus As Variant
    DDETerminateAll
    Set excel = CreateObject("Excel.Application")
    channel1 = DDEInitiate(App:="Excel", Topic:="System")
    DDEExecute channel:=channel1, Command:="[OPEN(" & Chr(34) _
     & "C:\Users\KMF\Documents\database.xlsx" & Chr(34) & ")]"
    DDETerminate channel:=channel1
    channel1 = DDEInitiate(App:="Excel", Topic:="Database.xlsx")
    rowarray = Split(DDERequest(channel1, "R2"), Chr(9))
    With Lenspicker
        i = 2
        Do Until rowarray(0) = ""
            .Lenses.AddItem (i)
            For j = 0 To 9
              .Lenses.List(i - 2, j) = rowarray(j)
            Next j
            i = i + 1
            Index = "R" + Trim(Str(i))
            a = DDERequest(channel1, Index)
            rowarray = Split(a, Chr(9))
         Loop
    End With
    Lenspicker.Show
    Return
main:
    Selection.GoTo what:=wdGoToBookmark, Name:="name"
    With Selection
        .Expand unit:=wdLine
        ptname = Left(.Text, Len(.Text) - 2)
        .GoTo what:=wdGoToBookmark, Name:="birthdate"
        .Expand unit:=wdLine
        dob = .Text
       ptage = Left(.Text, Len(.Text) - 2)
        With Selection.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = "PD="
            .Forward = True
            .Format = False
            .MatchCase = False
            .Wrap = wdFindContinue
            .Execute
        End With
        .Expand unit:=wdParagraph
        pd = Left(.Text, Len(.Text) - 2)
        .MoveDown unit:=wdLine
'Copy all the refraction data and store in array RF(x,y) y=0 for right eye, y=1 for left eye, x=0 to 8 for acuity, sphere, cyl, axix, add, PD, K1 and K2, Average
        For i = 0 To 4
            .MoveLeft unit:=wdCell, Count:=1
            .Expand unit:=wdParagraph
            RF(i, 0) = Val(.Text)
            .MoveRight unit:=wdCell, Count:=2
            .Expand unit:=wdParagraph
            RF(i, 1) = Val(.Text)
            .MoveLeft unit:=wdCell, Count:=1
            .Expand unit:=wdParagraph
            RF(i, 2) = Left(.Text, Len(.Text) - 2)
            If RF(i, 2) = "Add" Then
                If RF(i, 1) > RF(i, 0) Then RF(i, 0) = RF(i, 1) Else RF(i, 1) = RF(i, 0)
            End If
            .MoveDown unit:=wdLine, Count:=1
        Next i
        RF(5, 2) = pd
    End With
    ActiveWindow.ActivePane.LargeScroll Down:=-1
    If InputBox("Spectacles Only?", "", "YES") = "YES" Then
        product = 0
        Call Prescription(RF, product, ptname, ptage)
        ActiveDocument.Select
        ActiveDocument.Tables(2).Select
        Selection.Collapse (wdCollapseEnd)
        With Selection
            .ParagraphFormat.LineSpacing = 10
            .ParagraphFormat.LineSpacingRule = wdLineSpaceExactly
            .TypeText vbTab & vbTab
            .Font.Bold = True
            .TypeText IIf(LCase(pd) = "pd", "", pd)
            .Font.Bold = False
            .TypeText vbTab & vbTab
            .TypeText "Good for "
            .Font.Underline = wdUnderlineSingle
            .TypeText "One Year"
            .Expand (wdLine)
            .ParagraphFormat.SpaceBefore = 4
            .GoTo what:=wdGoToBookmark, Name:="Instructions"
            .ParagraphFormat.LeftIndent = MillimetersToPoints(5)
            .Font.Bold = True
            .TypeText ("Instructions for Optician: ")
            .Font.Underline = wdUnderlineNone
            With ActiveDocument
                Text = ""
                If Val(.Tables(2).Cell(2, 2).Range.Text) = 0 And Val(.Tables(2).Cell(2, 2).Range.Text) = Val(.Tables(2).Cell(3, 2).Range.Text) Then Text = "READING GLASSES ONLY"
            End With
            .TypeText Text
            .Font.Bold = False
            .Font.Italic = True
            .TypeText Chr(13) + "Please discuss all parameters with patient to help him/her make an informed choice of frame style and brand, lens material, tint, coating etc...  Check the reported PD before final order."
            With Selection.ParagraphFormat
                .LeftIndent = MillimetersToPoints(10)
                .RightIndent = MillimetersToPoints(10)
                .SpaceBefore = 3
                .LineSpacingRule = wdLineSpaceExactly
                .LineSpacing = 10
                .Alignment = wdAlignParagraphJustify
            End With
            If RF(4, 0) <> 0 Or RF(4, 1) <> 0 Then
                .Font.Underline = wdUnderlineSingle
                .TypeText Chr(13) & " A Progressive Reading Segment is recommended."
                .Font.Underline = wdUnderlineNone
            End If
            .TypeText Chr(13) + "This prescription may be filled after expiration during a second year only if you have tested the visual acuity to be within one line of the visual acuity reported in this prescription."
              .GoTo what:=wdGoToBookmark, Name:="Title"
            .InsertAfter ("Spectacles ")
        End With
        ElseIf InputBox("Contacts Only?", "", "YES") = "YES" Then
'Adjust power for contact lens precription when spectacles power above 3 diopters, aasuming vertex of 13.75 mm and ROUND TO THE NEAREST QUARTER DIOPTER
            csr = RF(1, 0)
            If Abs(csr) > 3 Then
                cornea = Round(4 * (Abs(RF(1, 0)) / RF(1, 0) * 1000 * (1000 * (Abs(RF(1, 0))) ^ -1 - 13.75 * (Abs(RF(1, 0)) / RF(1, 0))) ^ -1), 2)
                RF(1, 0) = IIf(cornea - Int(cornea) < 0.5, Int(cornea), Int(cornea) + 1) / 4
            End If
            ccr = RF(2, 0)
            If Abs(ccr) > 3 Then
                cornea = Round(4 * (Abs(RF(2, 0)) / RF(2, 0) * 1000 * (1000 * (Abs(RF(2, 0))) ^ -1 - 13.75 * (Abs(RF(2, 0)) / RF(2, 0))) ^ -1), 2)
                RF(2, 0) = IIf(cornea - Int(cornea) < 0.5, Int(cornea), Int(cornea) + 1) / 4
            End If
            cornea = Round(RF(1, 0) + RF(2, 0) / 2, 2) * 4
            RF(6, 0) = IIf(cornea - Int(cornea) < 0.5, Int(cornea), Int(cornea) + 1) / 4
            csl = RF(1, 1)
            If Abs(csl) > 3 Then
                cornea = Round(4 * (Abs(RF(1, 1)) / RF(1, 1) * 1000 * (1000 * (Abs(RF(1, 1))) ^ -1 - 13.75 * (Abs(RF(1, 1)) / RF(1, 1))) ^ -1), 2)
                RF(1, 1) = IIf(cornea - Int(cornea) < 0.5, Int(cornea), Int(cornea) + 1) / 4
            End If
            ccl = RF(2, 1)
            If Abs(ccl) > 3 Then
                cornea = Round(4 * (Abs(RF(2, 1)) / RF(2, 1) * 1000 * (1000 * (Abs(RF(2, 1))) ^ -1 - 13.75 * (Abs(RF(2, 1)) / RF(2, 1))) ^ -1), 2)
                RF(2, 1) = IIf(cornea - Int(cornea) < 0.5, Int(cornea), Int(cornea) + 1) / 4
            End If
            cornea = Round(RF(1, 1) + RF(2, 1) / 2, 2) * 4
            RF(6, 1) = IIf(cornea - Int(cornea) < 0.5, Int(cornea), Int(cornea) + 1) / 4
            If InputBox("Sherical fit?", "", "YES") = "YES" Then product = 1.1 Else product = 1
            GoSub pickalens
            Call Prescription(RF, product, ptname, ptage)
            With Selection
                .GoTo what:=wdGoToBookmark, Name:="Instructions"
                .ParagraphFormat.LeftIndent = MillimetersToPoints(5)
                .Font.Underline = wdUnderlineSingle
                .Font.Bold = True
                .TypeText ("Instructions for Optician:")
                .Font.Underline = wdUnderlineNone
                .Font.Bold = False
                .Font.Italic = True
                With Selection.ParagraphFormat
                    .LeftIndent = MillimetersToPoints(5)
                    .RightIndent = MillimetersToPoints(0)
                    .SpaceBefore = 3
                    .LineSpacingRule = wdLineSpaceExactly
                    .LineSpacing = 10
                    .Alignment = wdAlignParagraphJustify
                    .RightIndent = 20
                End With
                If Not RF(2, 0) = 0 Or Not RF(2, 1) = 0 Then
                    .Font.Underline = wdUnderlineNone
                    .TypeText Chr(13) & "The optician is "
                    .Font.Underline = wdUnderlineSingle
                    .TypeText "authorized"
                    .Font.Underline = wdUnderlineNone
                    .TypeText " to dispense lenses "
                    .Font.Underline = wdUnderlineSingle
                    .TypeText "a quarter diopter less cylinder and axis up to 5 degrees less or more"
                    .Font.Underline = wdUnderlineNone
                    .TypeText " than the prescribed axis if the exact cylinder and axis are not available in the recommended lens"
                    Selection.ParagraphFormat.LeftIndent = MillimetersToPoints(10)
                End If
                .TypeText Chr(13) + "Dispensing a different brand with same BC and DIAM is authorized if patient expresses a preferential choice of brand based on previous lens wear experience. Please have him/her sign their choice on this prescription and keep a copy."
                .Font.Underline = wdUnderlineSingle
                .Font.Bold = True
                .TypeText Chr(13) + ("Instructions for Patient:")
                .ParagraphFormat.LeftIndent = MillimetersToPoints(5)
                .Font.Underline = wdUnderlineNone
                .Font.Bold = False
                .TypeText Chr(13) & "Wear your contact lens as instructed during the office visit. Do not exceed the length of time recommended and replace it promptly to minimize the risk of complications. If you develop pain, redness or persistent itching or tearing, return immediately to our office or seek immediate attention from a qualified eye care professional."
                .ParagraphFormat.LeftIndent = MillimetersToPoints(10)
                .GoTo what:=wdGoToBookmark, Name:="Title"
                .InsertAfter ("Contact Lens ")
        End With
    End If
'excel.Quit
End Sub

Private Static Sub Prescription(RF, product, ptname, ptage)
' Procedure called with parameter product coded as follows: 0 for spectacles and 1 for contact lens (1.1 for spherical CL)
    Title = ""
    line1 = ""
    line2 = ""
' Convert to negative cylinder notation
    For i = 0 To 5
            If UCase(RF(i, 2)) = "CYL" Or UCase(RF(i, 2)) = "CYLINDER" Then
                For j = 0 To 1
                    If RF(i, j) > 0 Then
                        RF(i - 1, j) = RF(i - 1, j) + RF(i, j)
                        RF(i + 1, j) = (RF(i + 1, j) + 90) Mod 180
                        If RF(i + 1, j) = 0 Then RF(i + 1, j) = 180
                        RF(i, j) = -1 * RF(i, j)
                    End If
                Next j
                axis1 = RF(3, 0)
                axis2 = RF(3, 1)
            End If
    Next i
    If product >= 1 Then
        For i = 6 To 8
            RF(i, 2) = Choose(i - 5, "Spherical", "BC", "Diam")
        Next i
    End If
' Check and format value for sphere, cylnder, axis and add (i=1 to 4) left and right eyes
    For i = 1 To 8
                If InStr("AXISACUITYBCDIAM", UCase(RF(i, 2))) Then g = Trim(Str(RF(i, 0))) Else g = Trim(Str(RF(i, 0)) & Choose(InStr(StrReverse(Str(RF(i, 0))), ".") + 1, ".00", "", "0"))
                If Mid(g, 2, 1) = "." And Val(g) < 0 Then g = "- 0" & Trim(Mid(g, InStr(g, ".")))
                If Mid(g, 1, 1) = "." And Val(g) > 0 Then g = "0" & Trim(g)
                    line1 = line1 + vbTab + IIf(InStr("AXISACUITYBCDIAM", UCase(RF(i, 2))), IIf(g = "0", "No", g), Choose(Sgn(RF(i, 0)) + 2, "- " & Mid(g, 2), "No", "+ " & g))
                If InStr("AXISACUITYBCDIAM", UCase(RF(i, 2))) Then g = Trim(Str(RF(i, 1))) Else g = Trim(Str(RF(i, 1)) & Choose(InStr(StrReverse(Str(RF(i, 1))), ".") + 1, ".00", "", "0"))
                If Mid(g, 2, 1) = "." And Val(g) < 0 Then g = "- 0" & Trim(Mid(g, InStr(g, ".")))
                If Mid(g, 1, 1) = "." And Val(g) > 0 Then g = "0" & Trim(g)
                    line2 = line2 + vbTab + IIf(InStr("AXISACUITYBCDIAM", UCase(RF(i, 2))), IIf(g = "0", "No", g), Choose(Sgn(RF(i, 1)) + 2, "- " & Mid(g, 2), "No", "+ " & g))
                    Title = Title + vbTab + RF(i, 2)
                    OD = UCase(StrReverse(Left(StrReverse(line1), InStr(1, StrReverse(line1), vbTab))))
                    OS = UCase(StrReverse(Left(StrReverse(line2), InStr(1, StrReverse(line2), vbTab))))
                    OU = UCase(StrReverse(Left(StrReverse(Title), InStr(1, StrReverse(Title), vbTab))))
                    If InStr(OD, "NO") And InStr(OS, "NO") And InStr("BCDIAM", UCase(RF(i, 2))) = 0 Then
                        Title = Left(Title, Len(Title) - Len(OU))
                        line1 = Left(line1, Len(line1) - Len(OD))
                        line2 = Left(line2, Len(line2) - Len(OS))
                    ElseIf UCase(RF(i, 2)) = "SPHERICAL" And product = 1 Or RF(i, 2) = "" Then
                        Title = Left(Title, Len(Title) - Len(OU))
                        line1 = Left(line1, Len(line1) - Len(OD))
                        line2 = Left(line2, Len(line2) - Len(OS))
                    End If
    Next i
    Title = Title + vbTab + RF(0, 2) + vbCr
    line1 = "Right" + line1 + vbTab + Trim(Str(RF(0, 0))) + vbCr
    line2 = "Left" + line2 + vbTab + Trim(Str(RF(0, 1))) + vbCr
' remove extra spaces
    Do While InStr(line1, "  ") > 0
        line1 = Left(line1, InStr(line1, "  ")) & Mid(line1, InStr(line1, "  ") + 2)
    Loop
    Do While InStr(line2, "  ") > 0
        line2 = Left(line2, InStr(line2, "  ")) & Mid(line2, InStr(line2, "  ") + 2)
    Loop
        Documents.Add Template:= _
                "C:\users\kmf\Documents\Templates\Optic.dotx", NewTemplate:=False, _
                DocumentType:=0
    With Selection
     days = DateDiff("d", dob, Date)
    age = Int(days / 365)
    mo = Int((days Mod 365) / 30)
    da = (days Mod 365) Mod 30
       .GoTo what:=wdGoToBookmark, Name:="name"
        .InsertAfter ptname
        .GoTo what:=wdGoToBookmark, Name:="age"
        .InsertAfter ptage
        .GoTo what:=wdGoToBookmark, Name:="date"
        .InsertAfter Date & ", Time: " & Time
        .GoTo what:=wdGoToBookmark, Name:="prescription"
        If product >= 1 Then
            With Selection.PageSetup
                .LeftMargin = MillimetersToPoints(15)
                .RightMargin = MillimetersToPoints(15)
                .SectionStart = wdSectionContinuous
            End With
        End If
        .InsertAfter Title
        .InsertAfter line1
        .InsertAfter line2
        .ConvertToTable Separator:=wdSeparateByTabs, AutoFitBehavior:=wdAutoFitContent
    End With
    With Selection.Tables(1)
        If product = 1.1 Then
            For i = 1 To 3
                .Range.Columns(2).Delete
            Next i
        End If
        Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        .AutoFitBehavior wdAutoFitContent
        .Style = "Table Grid"
'Store table width in variable a by adding all columns widths
        .Cell(2, 1).Range.Bold = True
        .Cell(3, 1).Range.Bold = True
        a = 0
        For i = 1 To .Columns.Count
            a = a + .Columns(i).Width
        Next i
        .AutoFitBehavior wdAutoFitContent
'  Horizontally center the prescrition table
        .Rows.SetLeftIndent LeftIndent:=(350 - a) / 2, RulerStyle:=wdAdjustNone
        .Select
        For j = 2 To 3
            ActiveDocument.Range(.Cell(j, 1).Range.Start, .Cell(j, .Columns.Count - 1).Range.End).Font.Bold = True
        Next j
    End With
End Sub
Sub lensparam(a, b, c)
    For i = 0 To 1
        RF(7, i) = b
        RF(8, i) = c
    Next i
End Sub



