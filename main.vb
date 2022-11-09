Option Explicit

Private Const c_sDialogCommand As String = "fDialog"
Const sResourcePrefix As String = "RES_"
Private Const c_sAddinFolder As String = "Analysis"
Private Const c_sXllName As String = "ANALYS32.XLL"

Private Enum RegistrationTerm
    RegistrationAddIn = 1
    RegistrationFunction = 2
End Enum

'Get Culture
Private Function GetATPUICultureTag() As String
    Dim shTemp As Worksheet
    Dim sCulture As String
    Dim sSheetName As String
    
    sCulture = Application.International(xlUICultureTag)
    sSheetName = sResourcePrefix + sCulture
    
    On Error Resume Next
    Set shTemp = ThisWorkbook.Worksheets(sSheetName)
    On Error GoTo 0
    If shTemp Is Nothing Then sCulture = GetFallbackTag(sCulture)
    
    GetATPUICultureTag = sCulture
End Function

'Entry point for RibbonX button click
Sub ShowATPDialog(control As IRibbonControl)
    Dim funcs As Variant
    funcs = Application.RegisteredFunctions
    If (IsNull(funcs)) Then
        'XLL isn't open or didn't register for some reason
        Exit Sub
    End If
    
    Dim sPathSep As String
    sPathSep = Application.PathSeparator
    Dim sXllFullName As String
    sXllFullName = Application.LibraryPath & sPathSep & c_sAddinFolder & sPathSep & c_sXllName
    Dim fFoundCommand As Boolean
    fFoundCommand = False
    Dim iFuncNum As Integer
    For iFuncNum = LBound(funcs) To UBound(funcs)
        If (StrComp(funcs(iFuncNum, RegistrationFunction), c_sDialogCommand, vbTextCompare) = 0) Then
            fFoundCommand = StrComp(funcs(iFuncNum, RegistrationAddIn), sXllFullName, vbTextCompare) = 0
            Exit For
        End If
    Next iFuncNum
    
    If (Not fFoundCommand) Then
        'Dialog command isn't registered or is registered to the wrong XLL
        Exit Sub
    End If
    
    Application.Run (c_sDialogCommand)
End Sub

'Callback for RibbonX button label
Sub GetATPLabel(control As IRibbonControl, ByRef label)
    label = ThisWorkbook.Sheets(sResourcePrefix + GetATPUICultureTag()).Range("RibbonCommand").Value
End Sub

'Callback for screentip
Public Sub GetATPScreenTip(control As IRibbonControl, ByRef label)
    label = ThisWorkbook.Sheets(sResourcePrefix + GetATPUICultureTag()).Range("ScreenTip").Value
End Sub

'Callback for Super Tip
Public Sub GetATPSuperTip(control As IRibbonControl, ByRef label)
    label = ThisWorkbook.Sheets(sResourcePrefix + GetATPUICultureTag()).Range("SuperTip").Value
End Sub

Public Sub GetGroupName(control As IRibbonControl, ByRef label)
    label = ThisWorkbook.Sheets(sResourcePrefix + GetATPUICultureTag()).Range("GroupName").Value
End Sub

'Check for Fallback Languages
Private Function GetFallbackTag(szCulture As String) As String
    'Sorted alphabetically by returned culture tag, then input culture tag
    Select Case (szCulture)
        Case "rm-CH"
            GetFallbackTag = "de-DE"
        Case "ca-ES", "ca-ES-valencia", "eu-ES", "gl-ES"
            GetFallbackTag = "es-ES"
        Case "lb-LU"
            GetFallbackTag = "fr-FR"
        Case "nn-NO"
            GetFallbackTag = "nb-NO"
        Case "be-BY", "ky-KG", "tg-Cyrl-TJ", "tt-RU", "uz-Latn-UZ"
            GetFallbackTag = "ru-RU"
        Case Else
            GetFallbackTag = "en-US"
    End Select
End Function
    
    'Pixel por celda
    Option Explicit

Sub LoadImageIntoExcel()
    
    Me.Activate
    
    Dim strFileName     As String
    
    Dim bmpFileHeader   As BITMAPFILEHEADER
    Dim bmpInfoHeader   As BITMAPINFOHEADER
    Dim ExcelPalette()  As PALETTE
    Dim Palette24       As PALETTE24Bit
    
    Dim I               As Integer
    Dim R As Integer, c As Integer
    Dim dAdjustedWidth  As Double
    Dim Padding         As Byte

   
    AutoSize
    On Error GoTo CloseFile
    strFileName = Application.GetOpenFilename

    Open strFileName For Binary As #1
        
    Get #1, , bmpFileHeader
    Get #1, , bmpInfoHeader


    If bmpInfoHeader.lngWidth Mod 4 > 0 Then
        dAdjustedWidth = (((Int((bmpInfoHeader.lngWidth * bmpInfoHeader.intBitCount) / 32) + 1) * 4#)) / _
                            (bmpInfoHeader.intBitCount / 8#)

        If dAdjustedWidth Mod 4 <> 0 Then dAdjustedWidth = Application.RoundUp(dAdjustedWidth, 0)

    Else
        dAdjustedWidth = bmpInfoHeader.lngWidth
    End If
    
    If bmpInfoHeader.intBitCount <= 8 Then
        ReDim ExcelPalette(0 To 255)
        

        For I = 0 To UBound(ExcelPalette)
            Get #1, , ExcelPalette(I)
        Next I
    
                
        Dim bytPixel As Double
        
        For R = 1 To bmpInfoHeader.lngHeight
            For c = 1 To dAdjustedWidth
                
                If c <= bmpInfoHeader.lngWidth Then
                    Get #1, , bytPixel
                    
                    Me.Cells(bmpInfoHeader.lngHeight + 1 - R, c).Interior.Color = _
                    RGB(ExcelPalette(bytPixel).Red, _
                        ExcelPalette(bytPixel).Green, _
                        ExcelPalette(bytPixel).Blue)
                    DoEvents
                Else
                    Get #1, , Padding
                    Me.Cells(bmpInfoHeader.lngHeight + 1 - R, c).Interior.Color = _
                        RGB(255, 255, 255)
                End If
                


            Next c
        Next R
        
    Else
            
        For R = 1 To bmpInfoHeader.lngHeight
            For c = 1 To dAdjustedWidth
                
                If c <= bmpInfoHeader.lngWidth Then
                    Get #1, , Palette24
                    Me.Cells(bmpInfoHeader.lngHeight + 1 - R, c).Interior.Color = _
                        RGB(Palette24.Red, _
                            Palette24.Green, _
                            Palette24.Blue)
                Else
                    Get #1, , Padding
                    Me.Cells(bmpInfoHeader.lngHeight + 1 - R, c).Interior.Color = _
                        RGB(255, 255, 255)
                End If
                DoEvents
            Next c
        Next R
        
    End If
    

    MsgBox "File loaded - program complete."

CloseFile:
    If Len(Err.Description) > 0 Then MsgBox Err.Description
    Close #1

End Sub

Private Sub AutoSize()

    Dim cels As Range
    Set cels = Me.Cells
    
    cels.ColumnWidth = 1
    cels.RowHeight = cels(1, 1).width
    cels.Clear
    
    Set cels = Nothing
End Sub

                'RGB a blanco y negro
                Option Explicit

Sub LoadImageIntoExcel()
    
    Me.Activate
    
    Dim strFileName     As String
    
    Dim bmpFileHeader   As BITMAPFILEHEADER
    Dim bmpInfoHeader   As BITMAPINFOHEADER
    Dim ExcelPalette()  As PALETTE
    Dim Palette24       As PALETTE24Bit
    
    Dim I               As Integer
    Dim R As Integer, c As Integer
    Dim dAdjustedWidth  As Double
    Dim Padding         As Byte

   
    AutoSize
    On Error GoTo CloseFile
    strFileName = Application.GetOpenFilename

    Open strFileName For Binary As #1
        
    Get #1, , bmpFileHeader
    Get #1, , bmpInfoHeader


    If bmpInfoHeader.lngWidth Mod 4 > 0 Then
        dAdjustedWidth = (((Int((bmpInfoHeader.lngWidth * bmpInfoHeader.intBitCount) / 32) + 1) * 4#)) / _
                            (bmpInfoHeader.intBitCount / 8#)

        If dAdjustedWidth Mod 4 <> 0 Then dAdjustedWidth = Application.RoundUp(dAdjustedWidth, 0)

    Else
        dAdjustedWidth = bmpInfoHeader.lngWidth
    End If
    
    If bmpInfoHeader.intBitCount <= 8 Then
        ReDim ExcelPalette(0 To 255)
        

        For I = 0 To UBound(ExcelPalette)
            Get #1, , ExcelPalette(I)
        Next I
    
                
        Dim bytPixel As Byte
        
        
        For R = 1 To bmpInfoHeader.lngHeight
            For c = 1 To dAdjustedWidth
                
                If c <= bmpInfoHeader.lngWidth Then
                    Get #1, , bytPixel
                    
                    Dim Red, Green, Blue, NTSC As Double
                    'Los valores RGB se convierten a escala de grises mediante la fórmula NTSC: 0.299 · Rojo + 0.587 · Verde + 0.114 · Azul'
                    
                    Red = ExcelPalette(bytPixel).Red
                    Green = ExcelPalette(bytPixel).Green
                    Blue = ExcelPalette(bytPixel).Blue
                    
                    NTSC = 0.299 * Red + 0.587 * Green + 0.114 * Blue

                    Me.Cells(bmpInfoHeader.lngHeight + 1 - R, c).Interior.Color = _
                    RGB(NTSC, NTSC, NTSC)
                    DoEvents
                Else
                    Get #1, , Padding
                    Me.Cells(bmpInfoHeader.lngHeight + 1 - R, c).Interior.Color = _
                        RGB(255, 255, 255)
                End If
                


            Next c
        Next R
        
    Else
            
        For R = 1 To bmpInfoHeader.lngHeight
            For c = 1 To dAdjustedWidth
                
                If c <= bmpInfoHeader.lngWidth Then
                    Get #1, , Palette24
                    
                    'Los valores RGB se convierten a escala de grises mediante la fórmula NTSC: 0.299 · Rojo + 0.587 · Verde + 0.114 · Azul'
                    
                    Red = Palette24.Red
                    Green = Palette24.Green
                    Blue = Palette24.Blue
                    
                    NTSC = 0.299 * Red + 0.587 * Green + 0.114 * Blue
                    Me.Cells(bmpInfoHeader.lngHeight + 1 - R, c).Interior.Color = _
                        RGB(NTSC, NTSC, NTSC)
                Else
                    Get #1, , Padding
                    Me.Cells(bmpInfoHeader.lngHeight + 1 - R, c).Interior.Color = _
                        RGB(255, 255, 255)
                End If
                DoEvents
            Next c
        Next R
        
    End If
    

    MsgBox "File loaded - program complete."

CloseFile:
    If Len(Err.Description) > 0 Then MsgBox Err.Description
    Close #1

End Sub

Private Sub AutoSize()

    Dim cels As Range
    Set cels = Me.Cells
    
    cels.ColumnWidth = 1
    cels.RowHeight = cels(1, 1).width
    cels.Clear
    
    Set cels = Nothing
End Sub

                            'Division de color por canal
                            Option Explicit

Sub LoadImageIntoExcel()
    
    Me.Activate
    
    Dim strFileName     As String
    
    Dim bmpFileHeader   As BITMAPFILEHEADER
    Dim bmpInfoHeader   As BITMAPINFOHEADER
    Dim ExcelPalette()  As PALETTE
    Dim Palette24       As PALETTE24Bit
    
    Dim I               As Integer
    Dim R As Integer, c As Integer
    Dim dAdjustedWidth  As Double
    Dim Padding         As Byte

   
    AutoSize
    On Error GoTo CloseFile
    strFileName = Application.GetOpenFilename

    Open strFileName For Binary As #1
        
    Get #1, , bmpFileHeader
    Get #1, , bmpInfoHeader


    If bmpInfoHeader.lngWidth Mod 4 > 0 Then
        dAdjustedWidth = (((Int((bmpInfoHeader.lngWidth * bmpInfoHeader.intBitCount) / 32) + 1) * 4#)) / _
                            (bmpInfoHeader.intBitCount / 8#)

        If dAdjustedWidth Mod 4 <> 0 Then dAdjustedWidth = Application.RoundUp(dAdjustedWidth, 0)

    Else
        dAdjustedWidth = bmpInfoHeader.lngWidth
    End If
    
    If bmpInfoHeader.intBitCount <= 8 Then
        ReDim ExcelPalette(0 To 255)
        

        For I = 0 To UBound(ExcelPalette)
            Get #1, , ExcelPalette(I)
        Next I
    
                
        Dim bytPixel As Byte
        
        For R = 1 To bmpInfoHeader.lngHeight
            For c = 1 To dAdjustedWidth
                
                If c <= bmpInfoHeader.lngWidth Then
                    Get #1, , bytPixel
                    Dim Rojo, Verde, Azul As Double
                    Dim canal As String
                    canal = Worksheets("Instrucciones").Cells(1, 1).Value
                    
                    If canal = "Red" Then
                        Rojo = ExcelPalette(bytPixel).Red
                        Verde = 0
                        Azul = 0
                    ElseIf canal = "Green" Then
                        Rojo = 0
                        Verde = ExcelPalette(bytPixel).Green
                        Azul = 0
                    ElseIf canal = "Blue" Then
                        Rojo = 0
                        Verde = 0
                        Azul = ExcelPalette(bytPixel).Blue
                    ElseIf canal = "Yellow1" Then
                        Rojo = ExcelPalette(bytPixel).Green
                        Verde = ExcelPalette(bytPixel).Green
                        Azul = 0
                    ElseIf canal = "Yellow2" Then
                        Rojo = ExcelPalette(bytPixel).Blue
                        Verde = ExcelPalette(bytPixel).Blue
                        Azul = 0
                    ElseIf canal = "Purple1" Then
                        Rojo = ExcelPalette(bytPixel).Blue
                        Verde = 0
                        Azul = ExcelPalette(bytPixel).Blue
                    ElseIf canal = "Purple2" Then
                        Rojo = ExcelPalette(bytPixel).Green
                        Verde = 0
                        Azul = ExcelPalette(bytPixel).Green
                    Else
                        Rojo = 255
                        Verde = 255
                        Azul = 255
                    End If
                    
                    Me.Cells(bmpInfoHeader.lngHeight + 1 - R, c).Interior.Color = _
                    RGB(Rojo, Verde, Azul)
                    DoEvents
                Else
                    Get #1, , Padding
                    Me.Cells(bmpInfoHeader.lngHeight + 1 - R, c).Interior.Color = _
                        RGB(255, 255, 255)
                End If
                


            Next c
        Next R
        
    Else
            
        For R = 1 To bmpInfoHeader.lngHeight
            For c = 1 To dAdjustedWidth
                
                If c <= bmpInfoHeader.lngWidth Then
                    Get #1, , Palette24
                    canal = Worksheets("Instrucciones").Cells(1, 1).Value
                    
                    If canal = "Red" Then
                        Rojo = Palette24.Red
                        Verde = 0
                        Azul = 0
                    ElseIf canal = "Green" Then
                        Rojo = 0
                        Verde = Palette24.Green
                        Azul = 0
                    ElseIf canal = "Blue" Then
                        Rojo = 0
                        Verde = 0
                        Azul = Palette24.Blue
                    ElseIf canal = "Yellow1" Then
                        Rojo = Palette24.Green
                        Verde = Palette24.Green
                        Azul = 0
                    ElseIf canal = "Yellow2" Then
                        Rojo = Palette24.Blue
                        Verde = Palette24.Blue
                        Azul = 0
                    ElseIf canal = "Purple1" Then
                        Rojo = Palette24.Blue
                        Verde = 0
                        Azul = Palette24.Blue
                    ElseIf canal = "Purple2" Then
                        Rojo = Palette24.Green
                        Verde = 0
                        Azul = Palette24.Green
                    Else
                        Rojo = 255
                        Verde = 255
                        Azul = 255
                    End If
                        
                    Me.Cells(bmpInfoHeader.lngHeight + 1 - R, c).Interior.Color = _
                        RGB(Rojo, Verde, Azul)
                Else
                    Get #1, , Padding
                    Me.Cells(bmpInfoHeader.lngHeight + 1 - R, c).Interior.Color = _
                        RGB(255, 255, 255)
                End If
                DoEvents
            Next c
        Next R
        
    End If
    

    MsgBox "File loaded - program complete."

CloseFile:
    If Len(Err.Description) > 0 Then MsgBox Err.Description
    Close #1

End Sub

Private Sub AutoSize()

    Dim cels As Range
    Set cels = Me.Cells
    
    cels.ColumnWidth = 1
    cels.RowHeight = cels(1, 1).width
    cels.Clear
    
    Set cels = Nothing
End Sub
                                        'RGB to YUV mapa de calor
                                        Option Explicit

Sub LoadImageIntoExcel()
    
    Me.Activate
    
    Dim strFileName     As String
    
    Dim bmpFileHeader   As BITMAPFILEHEADER
    Dim bmpInfoHeader   As BITMAPINFOHEADER
    Dim ExcelPalette()  As PALETTE
    Dim Palette24       As PALETTE24Bit
    
    Dim I               As Integer
    Dim R As Integer, c As Integer
    Dim dAdjustedWidth  As Double
    Dim Padding         As Byte

   
    AutoSize
    On Error GoTo CloseFile
    strFileName = Application.GetOpenFilename

    Open strFileName For Binary As #1
        
    Get #1, , bmpFileHeader
    Get #1, , bmpInfoHeader


    If bmpInfoHeader.lngWidth Mod 4 > 0 Then
        dAdjustedWidth = (((Int((bmpInfoHeader.lngWidth * bmpInfoHeader.intBitCount) / 32) + 1) * 4#)) / _
                            (bmpInfoHeader.intBitCount / 8#)

        If dAdjustedWidth Mod 4 <> 0 Then dAdjustedWidth = Application.RoundUp(dAdjustedWidth, 0)

    Else
        dAdjustedWidth = bmpInfoHeader.lngWidth
    End If
    
    If bmpInfoHeader.intBitCount <= 8 Then
        ReDim ExcelPalette(0 To 255)
        

        For I = 0 To UBound(ExcelPalette)
            Get #1, , ExcelPalette(I)
        Next I
    
                
        Dim bytPixel As Byte
        
        For R = 1 To bmpInfoHeader.lngHeight
            For c = 1 To dAdjustedWidth
                
                If c <= bmpInfoHeader.lngWidth Then
                    Get #1, , bytPixel
                    
                    Dim Red, Green, Blue, NTSC, Y, U, V As Double
                    
                    Red = ExcelPalette(bytPixel).Red
                    Green = ExcelPalette(bytPixel).Green
                    Blue = ExcelPalette(bytPixel).Blue
                    
                    Y = Red * 0.299 + Green * 0.587 + Blue * 0.114
                    U = Red * -0.168736 + Green * -0.331264 + Blue * 0.5 + 128
                    V = Red * 0.5 + Green * -0.418688 + Blue * -0.081312 + 128
                    
                    Me.Cells(bmpInfoHeader.lngHeight + 1 - R, c).Interior.Color = _
                    RGB(Y, U, V)
                    DoEvents
                Else
                    Get #1, , Padding
                    Me.Cells(bmpInfoHeader.lngHeight + 1 - R, c).Interior.Color = _
                        RGB(255, 255, 255)
                End If
                


            Next c
        Next R
        
    Else
            
        For R = 1 To bmpInfoHeader.lngHeight
            For c = 1 To dAdjustedWidth
                
                If c <= bmpInfoHeader.lngWidth Then
                    Get #1, , Palette24
                    
                    Red = Palette24.Red
                    Green = Palette24.Green
                    Blue = Palette24.Blue
                    
                    Y = Red * 0.299 + Green * 0.587 + Blue * 0.114
                    U = Red * -0.168736 + Green * -0.331264 + Blue * 0.5 + 128
                    V = Red * 0.5 + Green * -0.418688 + Blue * -0.081312 + 128
                    
                    Me.Cells(bmpInfoHeader.lngHeight + 1 - R, c).Interior.Color = _
                        RGB(Y, U, V)
                Else
                    Get #1, , Padding
                    Me.Cells(bmpInfoHeader.lngHeight + 1 - R, c).Interior.Color = _
                        RGB(255, 255, 255)
                End If
                DoEvents
            Next c
        Next R
        
    End If
    

    MsgBox "File loaded - program complete."

CloseFile:
    If Len(Err.Description) > 0 Then MsgBox Err.Description
    Close #1

End Sub

Private Sub AutoSize()

    Dim cels As Range
    Set cels = Me.Cells
    
    cels.ColumnWidth = 1
    cels.RowHeight = cels(1, 1).width
    cels.Clear
    
    Set cels = Nothing
End Sub

                                                    'RGB TO YIQ colores intensos
                                                    Option Explicit

Sub LoadImageIntoExcel()
    
    Me.Activate
    
    Dim strFileName     As String
    
    Dim bmpFileHeader   As BITMAPFILEHEADER
    Dim bmpInfoHeader   As BITMAPINFOHEADER
    Dim ExcelPalette()  As PALETTE
    Dim Palette24       As PALETTE24Bit
    
    Dim I               As Integer
    Dim R As Integer, c As Integer
    Dim dAdjustedWidth  As Double
    Dim Padding         As Byte

   
    AutoSize
    On Error GoTo CloseFile
    strFileName = Application.GetOpenFilename

    Open strFileName For Binary As #1
        
    Get #1, , bmpFileHeader
    Get #1, , bmpInfoHeader


    If bmpInfoHeader.lngWidth Mod 4 > 0 Then
        dAdjustedWidth = (((Int((bmpInfoHeader.lngWidth * bmpInfoHeader.intBitCount) / 32) + 1) * 4#)) / _
                            (bmpInfoHeader.intBitCount / 8#)

        If dAdjustedWidth Mod 4 <> 0 Then dAdjustedWidth = Application.RoundUp(dAdjustedWidth, 0)

    Else
        dAdjustedWidth = bmpInfoHeader.lngWidth
    End If
    
    If bmpInfoHeader.intBitCount <= 8 Then
        ReDim ExcelPalette(0 To 255)
        

        For I = 0 To UBound(ExcelPalette)
            Get #1, , ExcelPalette(I)
        Next I
    
                
        Dim bytPixel As Byte
        
        For R = 1 To bmpInfoHeader.lngHeight
            For c = 1 To dAdjustedWidth
                
                If c <= bmpInfoHeader.lngWidth Then
                    Get #1, , bytPixel
                    
                    Dim Red, Green, Blue, NTSC, Y, Ia, Q As Double
                    
                    Red = ExcelPalette(bytPixel).Red
                    Green = ExcelPalette(bytPixel).Green
                    Blue = ExcelPalette(bytPixel).Blue
                    
                    Y = Red * 0.299 + Green * 0.587 + Blue * 0.114
                    Ia = 0.596 * Red + 0.275 * Green + 0.321 * Blue
                    Q = 0.212 * Red + 0.523 * Green + 0.311 * Blue
                    
                    Me.Cells(bmpInfoHeader.lngHeight + 1 - R, c).Interior.Color = _
                    RGB(Y, Ia, Q)
                    DoEvents
                Else
                    Get #1, , Padding
                    Me.Cells(bmpInfoHeader.lngHeight + 1 - R, c).Interior.Color = _
                        RGB(255, 255, 255)
                End If
                


            Next c
        Next R
        
    Else
            
        For R = 1 To bmpInfoHeader.lngHeight
            For c = 1 To dAdjustedWidth
                
                If c <= bmpInfoHeader.lngWidth Then
                    Get #1, , Palette24
                    
                    Red = Palette24.Red
                    Green = Palette24.Green
                    Blue = Palette24.Blue
                    
                    Y = Red * 0.299 + Green * 0.587 + Blue * 0.114
                    Ia = 0.596 * Red + 0.275 * Green + 0.321 * Blue
                    Q = 0.212 * Red + 0.523 * Green + 0.311 * Blue
                    
                    Me.Cells(bmpInfoHeader.lngHeight + 1 - R, c).Interior.Color = _
                        RGB(Y, Ia, Q)
                Else
                    Get #1, , Padding
                    Me.Cells(bmpInfoHeader.lngHeight + 1 - R, c).Interior.Color = _
                        RGB(255, 255, 255)
                End If
                DoEvents
            Next c
        Next R
        
    End If
    

    MsgBox "File loaded - program complete."

CloseFile:
    If Len(Err.Description) > 0 Then MsgBox Err.Description
    Close #1

End Sub

Private Sub AutoSize()

    Dim cels As Range
    Set cels = Me.Cells
    
    cels.ColumnWidth = 1
    cels.RowHeight = cels(1, 1).width
    cels.Clear
    
    Set cels = Nothing
End Sub

