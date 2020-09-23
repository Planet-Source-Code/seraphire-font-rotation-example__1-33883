VERSION 5.00
Begin VB.Form frmSample 
   AutoRedraw      =   -1  'True
   Caption         =   "Font Rotation Example"
   ClientHeight    =   5415
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6075
   LinkTopic       =   "frmSample"
   LockControls    =   -1  'True
   ScaleHeight     =   5415
   ScaleWidth      =   6075
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDraw 
      Caption         =   "Click Me!"
      Height          =   495
      Left            =   75
      TabIndex        =   0
      Top             =   45
      Width           =   1215
   End
End
Attribute VB_Name = "frmSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File: frmSample (Code)
'' Created on: 4/17/2002
'' Created by: BK
''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
''  Example of Font Rotation
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


Private Sub cmdDraw_Click()
    Me.Cls
    'Note the Me.Scale... - this centers the starting
    'position of the text on the form.
    Dim i As Integer
    For i = 0 To 360 Step 10
    Me.ForeColor = QBColor(Int(Rnd() * 16))
    DrawRotatedText Me, "Hello World from Benny!", _
        Me.ScaleWidth / 2, Me.ScaleHeight / 2, _
        "Arial", 12, i, 400, False, False, False
    Next i
End Sub

Private Sub Form_Load()

End Sub
