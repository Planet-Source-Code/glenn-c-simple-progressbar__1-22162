VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Simple Progressbar"
   ClientHeight    =   1245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2370
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1245
   ScaleWidth      =   2370
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   75
      Top             =   675
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   315
      Left            =   623
      TabIndex        =   1
      Top             =   780
      Width           =   1125
   End
   Begin VB.PictureBox picProgress 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      DrawMode        =   10  'Mask Pen
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   338
      ScaleHeight     =   225
      ScaleWidth      =   1635
      TabIndex        =   0
      Top             =   270
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-----------------------------------------------
' This is a simple example of how to
' create a progressbar with a caption.
'-----------------------------------------------
' The trick is to set the DrawMode to
' Not Xor Pen and draw the caption
' before drawing the bar.
'-----------------------------------------------

Dim Value As Long         ' current progress value
Dim Interval As Double  ' amount to draw for each percent
Dim cap As String          ' caption

Private Sub Command1_Click()
    Value = 0
    Interval = picProgress.ScaleWidth / 100
    Command1.Enabled = False
    Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    Value = Value + 1
    
    ' set caption
    cap = Value & "%"
    
    With picProgress
        .Cls
        ' center the caption
        .CurrentX = (.ScaleWidth - .TextWidth(cap)) \ 2
        .CurrentY = (.ScaleHeight - .TextHeight(cap)) \ 2
        picProgress.Print cap
        ' draw a filled rect
        picProgress.Line (0, 0)-(Interval * Value, .ScaleHeight), RGB(0, 0, 200), BF
        .Refresh
    End With
    
    ' stop at 100
    If Value = 100 Then
        Timer1.Enabled = False
        Command1.Enabled = True
    End If
End Sub
