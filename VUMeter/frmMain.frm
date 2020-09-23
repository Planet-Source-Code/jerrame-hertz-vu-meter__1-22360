VERSION 5.00
Begin VB.Form frmMain 
   ClientHeight    =   3855
   ClientLeft      =   9765
   ClientTop       =   2250
   ClientWidth     =   1815
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   1815
   Begin VB.Frame Frame1 
      Caption         =   "VU Meter"
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
      Begin VUMeters.VUMeter VUMeter1 
         Height          =   3000
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   5292
         Border          =   0
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   1215
         Left            =   720
         Max             =   51
         Min             =   1
         TabIndex        =   1
         Top             =   1200
         Value           =   51
         Width           =   135
      End
      Begin VUMeters.VUMeter VUMeter1 
         Height          =   3000
         Index           =   1
         Left            =   960
         TabIndex        =   3
         Top             =   360
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   5292
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then Call FormDrag(Me)
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    End
End Sub

Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then Call FormDrag(Me)
End Sub

Private Sub Frame1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    End
End Sub

Private Sub VScroll1_Change()
Dim i As Integer
For i = 0 To VUMeter1.Count - 1
    VUMeter1(i).Value = VScroll1.Value
Next

End Sub

Private Sub VScroll1_Scroll()
Dim i As Integer
For i = 0 To VUMeter1.Count - 1
    VUMeter1(i).Value = VScroll1.Value
Next
End Sub

