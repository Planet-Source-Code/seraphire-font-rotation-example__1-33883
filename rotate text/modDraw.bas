Attribute VB_Name = "modDraw"
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File: modDraw (Code)
'' Created on: 4/17/2002
'' Created by: BK
''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
''  Provides drawing routines to rotate text on a canvas
''
''  This is an improved upon version of the source posted by ZATRiX on PSC
''  http://www.planet-source-code.com/vb/scripts/showcode.asp?txtCodeId=25423&lngWId=1
''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Change Notes:
''
''  MM/DD/YY    INITIALS        CHANGE NOTE
''  --------    --------        -----------
''  4/17/2002   BK              Created
''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateFont Lib "gdi32.dll" Alias "CreateFontA" (ByVal nheight As Long, ByVal nWidth As Long, ByVal nEscapement As Long, ByVal nOrientation As Long, ByVal fnWeight As Long, ByVal fdwItalic As Long, ByVal fdwUnderline As Long, ByVal fdwStrikeOut As Long, ByVal fdwCharSet As Long, ByVal fdwOutputPrecision As Long, ByVal fdwClipPrecision As Long, ByVal fdwQuality As Long, ByVal fdwPitchAndFamily As Long, ByVal lpszFace As String) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long

Private Const LOGPIXELSY = 90                    'For GetDeviceCaps - returns the height of a logical pixel
Private Const ANSI_CHARSET = 0                   'Use the default Character set
Private Const CLIP_LH_ANGLES = 16                ' Needed for tilted fonts.
Private Const OUT_TT_PRECIS = 4                  'Tell it to use True Types when Possible
Private Const PROOF_QUALITY = 2                  'Make it as clean as possible.
Private Const DEFAULT_PITCH = 0                  'We want the font to take whatever pitch it defaults to
Private Const FF_DONTCARE = 0                    'Use whatever fontface it is.


Public Enum FontWeight
    FW_DONTCARE = 0
    FW_THIN = 100
    FW_EXTRALIGHT = 200
    FW_ULTRALIGHT = 200
    FW_LIGHT = 300
    FW_NORMAL = 400
    FW_REGULAR = 400
    FW_MEDIUM = 500
    FW_SEMIBOLD = 600
    FW_DEMIBOLD = 600
    FW_BOLD = 700
    FW_EXTRABOLD = 800
    FW_ULTRABOLD = 800
    FW_HEAVY = 900
    FW_BLACK = 900
End Enum


Public Sub DrawRotatedText(ByRef Canvas As Object, _
    ByVal txt As String, _
    ByVal X As Single, ByVal Y As Single, _
    ByVal font_name As String, ByVal size As Long, _
    ByVal Angle As Single, ByVal weight As FontWeight, _
    ByVal Italic As Boolean, ByVal Underline As Boolean, _
    ByVal Strikethrough As Boolean)


    Dim newfont As Long
    Dim oldfont As Long
    Dim nEscapement As Long
    Dim nheight As Long

    'The Angle in CreateFont is in 1/10 of a degree resolution
    'The Angle is also rotated counter-clockwise from "3-o'clock"

    '   0 = 3 o'clock
    '  90 = noon
    ' 180 = 9 o'clock (upside down)
    ' 270 = 6 o'clock (like a book title)

    nEscapement = Angle * 10

    'The height of the call to create font is device dependent.
    'Therefore, use the following formula to convert "Size" to
    'the logical size.
    nheight = -MulDiv(size, GetDeviceCaps(Canvas.hdc, LOGPIXELSY), 72)

    'Create a font resource
    newfont = CreateFont(nheight, 0, nEscapement, nEscapement, weight, 0, 0, 0, ANSI_CHARSET, OUT_TT_PRECIS, CLIP_LH_ANGLES, PROOF_QUALITY, DEFAULT_PITCH Or FF_DONTCARE, "Arial")

    ' Select the new font.
    oldfont = SelectObject(Canvas.hdc, newfont)

    ' Display the text.
    ' Note that the X and Y are in the scale units of the canvas,
    ' so if we are on a form, by default, it will be twips.
    ' No calculation to Pixels is necessary!
    
    ' The pivot point of the box is the top-left corner.  You may
    'need to do some extra calculations to move the pivot point.
    '(x and y would have to be different for each call.)
    Canvas.CurrentX = X
    Canvas.CurrentY = Y
    Canvas.Print txt

    ' Restore the original font.
    newfont = SelectObject(Canvas.hdc, oldfont)

    ' Free font resources (important!)
    DeleteObject newfont
End Sub


