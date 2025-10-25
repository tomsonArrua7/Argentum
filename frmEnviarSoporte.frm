VERSION 5.00
Begin VB.Form frmEnviarSoporte 
   BorderStyle     =   0  'None
   ClientHeight    =   4350
   ClientLeft      =   3450
   ClientTop       =   2040
   ClientWidth     =   6420
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmEnviarSoporte.frx":0000
   ScaleHeight     =   4350
   ScaleWidth      =   6420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSoporte 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   600
      MaxLength       =   300
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmEnviarSoporte.frx":107CB
      Top             =   1680
      Width           =   5295
   End
   Begin VB.Image Image1 
      Height          =   255
      Index           =   1
      Left            =   720
      MouseIcon       =   "frmEnviarSoporte.frx":107DF
      MousePointer    =   99  'Custom
      Top             =   3720
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   255
      Index           =   0
      Left            =   3240
      MouseIcon       =   "frmEnviarSoporte.frx":12161
      MousePointer    =   99  'Custom
      Top             =   3720
      Width           =   2295
   End
End
Attribute VB_Name = "frmEnviarSoporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Picture = LoadPicture(DirGraficos & "SGM.gif")
End Sub

Private Sub Image1_Click(Index As Integer)
Select Case Index
Case 0

If Len(txtSoporte) Then
    Call SendData("/ZOPORTE " & txtSoporte.Text)
End If
txtSoporte.Text = ""
Me.Hide
Case 1
txtSoporte.Text = ""
Me.Hide
End Select
End Sub

