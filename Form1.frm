VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Example By Cause - aka RAD"
   ClientHeight    =   3015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4950
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4950
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4200
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "txt"
      Filter          =   "*.txt"
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save File"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   2400
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Open File"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   2400
      Width           =   2055
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   1980
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   4695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
CommonDialog1.ShowOpen
If CommonDialog1.FileTitle = "" Then ' so theres no text overflow if the user canceles open file
Exit Sub
Else
Dim Doit As String
    On Error Resume Next
    List1.Clear
    directory$ = CommonDialog1.FileTitle
    Open directory$ For Input As #1
        While Not EOF(1)
                Input #1, Doit$
        DoEvents
        List1.AddItem Doit$
    Wend
    Close #1
    End If
End Sub

Private Sub Command2_Click()
CommonDialog1.ShowSave
Dim SaveList1 As Long
    On Error Resume Next
      directory$ = CommonDialog1.FileTitle
        Open directory$ For Output As #1
        For SaveList1& = 0 To List1.ListCount - 1
        Print #1, List1.List(SaveList1&)
       Next SaveList1&
      Close #1
End Sub

