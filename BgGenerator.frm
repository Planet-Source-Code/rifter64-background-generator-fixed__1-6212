VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "BACKGROUND GENERATOR....BY RIFTER64"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   1200
   ClientWidth     =   15240
   DrawStyle       =   2  'Dot
   LinkTopic       =   "Form1"
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command5 
      BackColor       =   &H000000FF&
      Caption         =   "SPLIT SCREEN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   9840
      Width           =   2775
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H000000FF&
      Caption         =   "TRIANGLES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   9840
      Width           =   2775
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   8895
      Left            =   120
      ScaleHeight     =   8835
      ScaleWidth      =   14955
      TabIndex        =   10
      Top             =   120
      Width           =   15015
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H000000FF&
      Caption         =   "SAVE AS BMP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12600
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   9840
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Caption         =   "FADING IMAGE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   9840
      Width           =   2775
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "BgGenerator.frx":0000
      Left            =   12600
      List            =   "BgGenerator.frx":0019
      TabIndex        =   7
      Text            =   "DRAW STYLE"
      Top             =   9240
      Width           =   2535
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   3840
      TabIndex        =   4
      Text            =   "1"
      Top             =   9240
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2160
      TabIndex        =   3
      Text            =   "1"
      Top             =   9240
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   600
      TabIndex        =   1
      Text            =   "1"
      Top             =   9240
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "CIRCULAR IMAGE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   9840
      Width           =   2775
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "HTTP://WWW.GEOCITIES.COM/JABRONI_DRIVE64"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   120
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   10440
      Width           =   15015
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "GREEN"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3120
      TabIndex        =   6
      Top             =   9240
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "BLUE"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1680
      TabIndex        =   5
      Top             =   9240
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "RED"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   9240
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'THE BACKGROUND GENERATOR WAS CODED AND CREATED BY RYAN HOLLETT, FEB 17 2000
'VISIT HTTP://WWW.GEOCITIES.COM/JABRONI_DRIVE64 FOR OTHER VB PROGRAMS

Private Sub Command1_Click()
' Command button for the circular image
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Then
    MsgBox "A COLOR VALUE MUST BE ENTER FOR ALL THREE, AND MUST BE GREATER THAN 0", , "COLOR ERROR"
    Exit Sub
End If

Dim i As Integer, loops As Integer, rads As Integer

col1 = Text1.Text
col2 = Text2.Text
col3 = Text3.Text

MsgBox "AN IMAGE IS NOW BEING GENERATED WAIT FOR IT TO APPEAR"

Picture1.Cls

If Combo1.ListIndex <= 0 Then
    Picture1.DrawStyle = 0
    Let Picture1.DrawWidth = 2
ElseIf Combo1.ListIndex = 1 Then
    Picture1.DrawStyle = 1
    Picture1.DrawWidth = 1
ElseIf Combo1.ListIndex = 2 Then
    Picture1.DrawStyle = 2
    Picture1.DrawWidth = 1
ElseIf Combo1.ListIndex = 3 Then
    Picture1.DrawStyle = 3
    Picture1.DrawWidth = 1
ElseIf Combo1.ListIndex = 4 Then
    Picture1.DrawStyle = 4
    Picture1.DrawWidth = 1
ElseIf Combo1.ListIndex = 5 Then
    Picture1.DrawStyle = 5
    Picture1.DrawWidth = 1
ElseIf Combo1.ListIndex = 6 Then
    Picture1.DrawStyle = 6
    Picture1.DrawWidth = 1
End If

If Picture1.Height > Picture1.Width Then
    Let loops = Picture1.Height
    Let rads = Picture1.Width
ElseIf Picture1.Width > Picture1.Height Then
    Let loops = Picture1.Width
    Let rads = Picture1.Height
End If

For i = 1 To loops Step 20
Picture1.Circle (Picture1.Width / 2, Picture1.Height / 2), i, RGB(i / col1, i / col2, i / col3), BF
If i >= rads + 2000 Then
    Exit For
End If
Next i

MsgBox "THE IMAGE IS DONE"

End Sub
Private Sub Command2_Click()
' Comand button for the fading image
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Then
    MsgBox "A COLOR VALUE MUST BE ENTER FOR ALL THREE, AND MUST BE GREATER THAN 0", , "COLOR ERROR"
    Exit Sub
End If
Dim i As Integer, loops As Integer, rads As Integer

col1 = Text1.Text
col2 = Text2.Text
col3 = Text3.Text

MsgBox "AN IMAGE IS NOW BEING GENERATED WAIT FOR IT TO APPEAR"

Picture1.Cls

If Combo1.ListIndex <= 0 Then
    Picture1.DrawStyle = 0
    Let Picture1.DrawWidth = 2
ElseIf Combo1.ListIndex = 1 Then
    Picture1.DrawStyle = 1
    Picture1.DrawWidth = 1
ElseIf Combo1.ListIndex = 2 Then
    Picture1.DrawStyle = 2
    Picture1.DrawWidth = 1
ElseIf Combo1.ListIndex = 3 Then
    Picture1.DrawStyle = 3
    Picture1.DrawWidth = 1
ElseIf Combo1.ListIndex = 4 Then
    Picture1.DrawStyle = 4
    Picture1.DrawWidth = 1
ElseIf Combo1.ListIndex = 5 Then
    Picture1.DrawStyle = 5
    Picture1.DrawWidth = 1
ElseIf Combo1.ListIndex = 6 Then
    Picture1.DrawStyle = 6
    Picture1.DrawWidth = 1
End If

If Picture1.Height > Picture1.Width Then
    Let loops = Picture1.Height
    Let rads = Picture1.Width
Else
    Let loops = Picture1.Width
    Let rads = Picture1.Height
End If

For i = 1 To loops Step 20
Picture1.Line (i, Picture1.ScaleHeight)-(i, 0), RGB(i / col1, i / col2, i / col3)
Next i

MsgBox "THE IMAGE IS DONE"

End Sub
Private Sub Command3_Click()
' Command for saving as a files
    filename = InputBox("Give a filename to save:", "Save", "C:\Windows\Desktop\Rifter64.bmp")
    
    If filename = "" Then Exit Sub
    
    SavePicture Form1.Picture1.Image, filename
    
    MsgBox "File saved as:" & vbNewLine & filename, vbInformation, "Saved"
 
End Sub
Private Sub Command4_Click()
' Command button for the trinagle images
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Then
    MsgBox "A COLOR VALUE MUST BE ENTER FOR ALL THREE, AND MUST BE GREATER THAN 0", , "COLOR ERROR"
    Exit Sub
End If

Dim A, C As Long

col1 = Text1.Text
col2 = Text2.Text
col3 = Text3.Text

Picture1.Cls

MsgBox "AN IMAGE IS NOW BEING GENERATED WAIT FOR IT TO APPEAR"

If Combo1.ListIndex <= 0 Then
    Picture1.DrawStyle = 0
    Let Picture1.DrawWidth = 2
ElseIf Combo1.ListIndex = 1 Then
    Picture1.DrawStyle = 1
    Picture1.DrawWidth = 1
ElseIf Combo1.ListIndex = 2 Then
    Picture1.DrawStyle = 2
    Picture1.DrawWidth = 1
ElseIf Combo1.ListIndex = 3 Then
    Picture1.DrawStyle = 3
    Picture1.DrawWidth = 1
ElseIf Combo1.ListIndex = 4 Then
    Picture1.DrawStyle = 4
    Picture1.DrawWidth = 1
ElseIf Combo1.ListIndex = 5 Then
    Picture1.DrawStyle = 5
    Picture1.DrawWidth = 1
ElseIf Combo1.ListIndex = 6 Then
    Picture1.DrawStyle = 6
    Picture1.DrawWidth = 1
End If

For C = 1 To Picture1.Width Step 15
 Picture1.Line (0, A)-(Picture1.Height / 2, Picture1.Width / 2), RGB(C / col1, C / col2, C / col3)
Next C

For C = 1 To Picture1.Width Step 15
 Picture1.Line (A, 0)-(Picture1.Height / 2, Picture1.Width / 2), RGB(C / col1, C / col2, C / col3)
Next C

For C = 1 To Picture1.Width Step 15
 Picture1.Line (Picture1.Width, Picture1.Height - C)-(0, C), RGB(C / col1, C / col2, C / col3)
 
Next C

For C = 1 To Picture1.Width Step 15
 Picture1.Line (Picture1.Width - C, Picture1.Height)-(C, 0), RGB(C / col1, C / col2, C / col3)
Next C

MsgBox "THE IMAGE IS DONE"

End Sub
Private Sub Command5_Click()
' Command for the split screen images
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Then
    MsgBox "A COLOR VALUE MUST BE ENTER FOR ALL THREE, AND MUST BE GREATER THAN 0", , "COLOR ERROR"
    Exit Sub
End If

Dim i As Long

MsgBox "AN IMAGE IS NOW BEING GENERATED WAIT FOR IT TO APPEAR"

If Combo1.ListIndex <= 0 Then
    Picture1.DrawStyle = 0
    Let Picture1.DrawWidth = 2
ElseIf Combo1.ListIndex = 1 Then
    Picture1.DrawStyle = 1
    Picture1.DrawWidth = 1
ElseIf Combo1.ListIndex = 2 Then
    Picture1.DrawStyle = 2
    Picture1.DrawWidth = 1
ElseIf Combo1.ListIndex = 3 Then
    Picture1.DrawStyle = 3
    Picture1.DrawWidth = 1
ElseIf Combo1.ListIndex = 4 Then
    Picture1.DrawStyle = 4
    Picture1.DrawWidth = 1
ElseIf Combo1.ListIndex = 5 Then
    Picture1.DrawStyle = 5
    Picture1.DrawWidth = 1
ElseIf Combo1.ListIndex = 6 Then
    Picture1.DrawStyle = 6
    Picture1.DrawWidth = 1
End If

col1 = Text1.Text
col2 = Text2.Text
col3 = Text3.Text

Picture1.Cls

Picture1.DrawWidth = 4

For i = 1 To Picture1.Width Step 10
Picture1.Line (0, i)-(Picture1.Width, i), RGB(i / col1, i / col2, i / col3)
Next i

For i = 1 To Picture1.Width Step 10
Picture1.Line (Picture1.Height - 6000, i)-(0, i), RGB(i / col3, i / col2, i / col1)
Next i

MsgBox "THE IMAGE IS DONE"

End Sub
Private Sub Form_load()

    Let Combo1.ListIndex = 0

End Sub
Private Sub Text1_gotfocus()
    
    Let Text1.Text = ""

End Sub
Private Sub Text2_gotfocus()
    
    Let Text2.Text = ""

End Sub
Private Sub Text3_gotfocus()
    
    Let Text3.Text = ""

End Sub

