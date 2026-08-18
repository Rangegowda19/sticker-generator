' ============================================================
'  StickerImport - CorelDRAW X8  (FIXED: X8-compatible width scaling)
'  Reads the Python data .txt file, handles any line endings,
'  draws all sticker pages as native editable objects.
' ============================================================
Option Explicit

Public Sub DrawStickers()
    Dim dataFile As String
    dataFile = InputBox("Paste the FULL path to the data .txt file:", "Sticker Import")
    dataFile = Trim(Replace(dataFile, Chr(34), ""))
    If dataFile = "" Then Exit Sub

    If Dir(dataFile) = "" Then
        MsgBox "FILE NOT FOUND:" & vbCrLf & dataFile
        Exit Sub
    End If

    Dim fnum As Integer, whole As String
    fnum = FreeFile
    Open dataFile For Input As #fnum
    whole = Input(LOF(fnum), fnum)
    Close #fnum

    whole = Replace(whole, vbCrLf, vbLf)
    whole = Replace(whole, vbCr, vbLf)
    Dim allLines() As String
    allLines = Split(whole, vbLf)

    Dim pageW As Double, pageH As Double
    Dim fontName As String
    Dim nText As Long, nOval As Long
    fontName = "Arial": pageW = 595.28: pageH = 841.89
    nText = 0: nOval = 0

    Dim doc As Document
    Set doc = Application.CreateDocument
    doc.Unit = cdrPoint
    Dim firstPage As Boolean
    firstPage = True

    Application.Optimization = True

    Dim i As Long, ln As String
    For i = 0 To UBound(allLines)
        ln = Trim(allLines(i))
        If Len(ln) > 0 Then
            Dim p() As String
            p = Split(ln, "|")
            Select Case p(0)
                Case "FONT":  fontName = p(1)
                Case "PAGEW": pageW = CDbl(p(1))
                Case "PAGEH": pageH = CDbl(p(1))
                Case "PAGE"
                    If firstPage Then
                        firstPage = False
                    Else
                        doc.AddPages 1
                        doc.Pages(doc.Pages.Count).Activate
                    End If
                    doc.ActivePage.SetSize pageW, pageH
                Case "TEXT"
                    Dim tx As Double, ty As Double, sz As Double, hsc As Double
                    Dim anchor As String, s As String, k As Integer
                    tx = CDbl(p(1)): ty = CDbl(p(2)): sz = CDbl(p(3))
                    hsc = CDbl(p(4)): anchor = p(5): s = p(6)
                    For k = 7 To UBound(p): s = s & "|" & p(k): Next k
                    DrawTxt tx, ty, sz, hsc, anchor, s, fontName, pageH
                    nText = nText + 1
                Case "OVAL"
                    Dim cx As Double, cy As Double, rx As Double, ry As Double, lw As Double
                    cx = CDbl(p(1)): cy = CDbl(p(2)): rx = CDbl(p(3))
                    ry = CDbl(p(4)): lw = CDbl(p(5))
                    DrawOvl cx, cy, rx, ry, lw, pageH
                    nOval = nOval + 1
            End Select
        End If
    Next i

    Application.Optimization = False
    ActiveWindow.Refresh
    MsgBox "Done." & vbCrLf & "Pages: " & doc.Pages.Count & vbCrLf & _
           "Text drawn: " & nText & vbCrLf & "Ovals drawn: " & nOval
End Sub

Private Sub DrawTxt(x As Double, y As Double, sizePt As Double, _
                    hscale As Double, anchor As String, s As String, _
                    fontName As String, pageH As Double)
    Dim sh As Shape
    Dim cy As Double
    cy = pageH - y
    Set sh = ActiveLayer.CreateArtisticText(x, cy, s)
    sh.Text.Story.Font = fontName
    sh.Text.Story.Size = sizePt
    sh.Text.Story.Bold = True
    sh.Fill.UniformColor.CMYKAssign 0, 0, 0, 100
    sh.Outline.SetNoOutline

    ' X8 has no Story.HorizontalScale -> squeeze the object width instead,
    ' anchored on the correct side so alignment stays right.
    Dim natW As Double
    natW = sh.SizeWidth
    If hscale > 0 And hscale <> 100 Then
        sh.SizeWidth = natW * hscale / 100
    End If

    Dim bw As Double
    bw = sh.SizeWidth
    Select Case anchor
        Case "middle": sh.LeftX = x - bw / 2
        Case "end":    sh.LeftX = x - bw
        Case Else:     sh.LeftX = x
    End Select
End Sub

Private Sub DrawOvl(cx As Double, cy As Double, rx As Double, ry As Double, _
                    lw As Double, pageH As Double)
    Dim yy As Double
    yy = pageH - cy
    Dim sh As Shape
    Set sh = ActiveLayer.CreateEllipse(cx - rx, yy + ry, cx + rx, yy - ry)
    sh.Fill.ApplyNoFill
    sh.Outline.SetProperties lw
    sh.Outline.Color.CMYKAssign 0, 0, 0, 100
End Sub
