VERSION 5.00
Begin VB.Form frmHonor 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Sistema de Honor"
   ClientHeight    =   7485
   ClientLeft      =   420
   ClientTop       =   315
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   Picture         =   "frm.Honor.frx":0000
   ScaleHeight     =   7485
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   780
      Left            =   240
      ScaleHeight     =   780
      ScaleWidth      =   1005
      TabIndex        =   1
      Top             =   720
      Width           =   1005
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Morpheus"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4530
      Left            =   240
      TabIndex        =   0
      Top             =   1850
      Width           =   2985
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   4680
      Top             =   1200
      Width           =   255
   End
   Begin VB.Image Command1 
      Height          =   735
      Left            =   3480
      Top             =   5760
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   5400
      Top             =   0
      Width           =   615
   End
   Begin VB.Label lblPermisos 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3840
      TabIndex        =   8
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label lblStat 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3720
      TabIndex        =   7
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Label lblPrecio 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3720
      TabIndex        =   6
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Label lblNombre 
      BackStyle       =   0  'Transparent
      Caption         =   "Canjes Escritura"
      BeginProperty Font 
         Name            =   "Morpheus"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   1320
      TabIndex        =   5
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Clases Permitidas"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6840
      TabIndex        =   4
      Top             =   2520
      Width           =   1290
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Stats:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7080
      TabIndex        =   3
      Top             =   3960
      Width           =   465
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Precio:"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   7080
      TabIndex        =   2
      Top             =   3480
      Width           =   555
   End
End
Attribute VB_Name = "frmHonor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

If List1.Text = "Gorro de Navidad" Then Call SendData("/CANJEA T1")
If List1.Text = "Túnica Zagreb Caótica" Then Call SendData("/CANJEA T2")
If List1.Text = "Túnica Zagreb Alianza" Then Call SendData("/CANJEA T3")
If List1.Text = "Armadura Tenebrosa" Then Call SendData("/CANJEA T4")
If List1.Text = "Tunica Milgror" Then Call SendData("/CANJEA T5")
If List1.Text = "Tunica Milgror Amarilla" Then Call SendData("/CANJEA T6")
End Sub
Private Sub Form_Load()
List1.AddItem "Gorro de Navidad"
List1.AddItem "Túnica Zagreb Caótica"
List1.AddItem "Túnica Zagreb Alianza"
List1.AddItem "Armadura Tenebrosa"
List1.AddItem "Tunica Milgror"
List1.AddItem "Tunica Milgror Amarilla"
End Sub

Private Sub Image1_Click()
Unload Me
End Sub

Private Sub Image2_Click()
MsgBox "No Disponible este Tipo de Canjes.."
End Sub

Private Sub List1_Click()

If List1.Text = "Gorro de Navidad" Then
    Picture1.Picture = LoadPicture(DirGraficos & "16023.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "100 Puntos de Retos"
    lblStat.Caption = "Min: 10 / Max: 15"
    lblPermisos.Caption = "Todas las Clases"
    End If
    
    If List1.Text = "Túnica Zagreb Caótica" Then
    Picture1.Picture = LoadPicture(DirGraficos & "16040.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "700 Puntos de Retos"
    lblStat.Caption = "Min: 30 / Max: 40"
    lblPermisos.Caption = "Todas las Clases"
    End If
    
        If List1.Text = "Túnica Zagreb Alianza" Then
    Picture1.Picture = LoadPicture(DirGraficos & "16044.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "700 Puntos de Retos"
    lblStat.Caption = "Min: 30 / Max: 40"
    lblPermisos.Caption = "Todas las Clases"
    End If
    
    
        If List1.Text = "Armadura Tenebrosa" Then
    Picture1.Picture = LoadPicture(DirGraficos & "16048.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "1200 Puntos de Retos"
    lblStat.Caption = "Min: 55 / Max: 60"
    lblPermisos.Caption = "Todas las Clases"
    End If
    
        If List1.Text = "Tunica Milgror" Then
    Picture1.Picture = LoadPicture(DirGraficos & "16314.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "1500 Puntos de Retos"
    lblStat.Caption = "Min: 55 / Max: 60"
    lblPermisos.Caption = "Todas las Clases"
    End If
    
        If List1.Text = "Tunica Milgror Amarilla" Then
    Picture1.Picture = LoadPicture(DirGraficos & "16316.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "1500 Puntos de Retos"
    lblStat.Caption = "Min: 55 / Max: 60"
    lblPermisos.Caption = "Todas las Clases"
    End If
    
End Sub
