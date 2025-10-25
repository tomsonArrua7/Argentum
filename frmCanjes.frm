VERSION 5.00
Begin VB.Form frmCanjes 
   BackColor       =   &H00000000&
   Caption         =   "Sistema de Canje"
   ClientHeight    =   4680
   ClientLeft      =   4395
   ClientTop       =   5460
   ClientWidth     =   7035
   LinkTopic       =   "Form1"
   Picture         =   "frmCanjes.frx":0000
   ScaleHeight     =   4680
   ScaleWidth      =   7035
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   3480
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   1
      Top             =   580
      Width           =   540
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
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
      Height          =   3630
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   3135
   End
   Begin VB.Image Command1 
      Height          =   375
      Left            =   3480
      Top             =   3770
      Width           =   3375
   End
   Begin VB.Label lblPermisos 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Morpheus"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   3360
      TabIndex        =   8
      Top             =   2520
      Width           =   3615
   End
   Begin VB.Label lblStat 
      Height          =   255
      Left            =   3960
      TabIndex        =   7
      Top             =   1800
      Width           =   2895
   End
   Begin VB.Label lblPrecio 
      Height          =   255
      Left            =   4080
      TabIndex        =   6
      Top             =   1320
      Width           =   2775
   End
   Begin VB.Label lblNombre 
      Height          =   255
      Left            =   4080
      TabIndex        =   5
      Top             =   600
      Width           =   2775
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Clases Permitidas"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4560
      TabIndex        =   4
      Top             =   2160
      Width           =   1290
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Stats:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3360
      TabIndex        =   3
      Top             =   1800
      Width           =   465
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Precio:"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   3480
      TabIndex        =   2
      Top             =   1320
      Width           =   555
   End
End
Attribute VB_Name = "frmCanjes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()


If List1.Text = "Ropa de Rey (Altos)" Then Call SendData("/CANJEO T1")
If List1.Text = "Ropa de Rey (E/G)" Then Call SendData("/CANJEO T2")
If List1.Text = "Ropa de la Alianza" Then Call SendData("/CANJEO T3")
If List1.Text = "Ropa del Mal" Then Call SendData("/CANJEO T4")
If List1.Text = "Corona Oscura(RM)" Then Call SendData("/CANJEO T5")
If List1.Text = "Corona de Rey(RM)" Then Call SendData("/CANJEO T6")
If List1.Text = "Pantalon Celeste" Then Call SendData("/CANJEO T7")
If List1.Text = "Pantalon Rojo" Then Call SendData("/CANJEO T8")
If List1.Text = "Pantalon Gris" Then Call SendData("/CANJEO T9")
If List1.Text = "Galera" Then Call SendData("/CANJEO T10")
End Sub
Private Sub Form_Load()


Me.Picture = LoadPicture(DirGraficos & "Canjes.jpg")

List1.AddItem "Ropa de Rey (Altos)"
List1.AddItem "Ropa de Rey (E/G)"
List1.AddItem "Ropa de la Alianza"
List1.AddItem "Ropa del Mal"
List1.AddItem "Corona Oscura(RM)"
List1.AddItem "Corona de Rey(RM)"
List1.AddItem "Pantalon Celeste"
List1.AddItem "Pantalon Rojo"
List1.AddItem "Pantalon Gris"
List1.AddItem "Galera"

End Sub



Private Sub List1_Click()

If List1.Text = "Ropa de Rey (Altos)" Then
    Picture1.Picture = LoadPicture(DirGraficos & "685.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "500 Puntos de Canje"
    lblStat.Caption = "Min: 40 / Max: 40"
    lblPermisos.Caption = "Todas las Clases"
    End If
    
    If List1.Text = "Ropa de Rey (E/G)" Then
    Picture1.Picture = LoadPicture(DirGraficos & "16092.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "500 Puntos de Canje"
    lblStat.Caption = "Min: 40 / Max: 40"
    lblPermisos.Caption = "Todas las Clases"
    End If
    
     If List1.Text = "Ropa de la Alianza" Then
    Picture1.Picture = LoadPicture(DirGraficos & "16038.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "500 Puntos de Canje"
    lblStat.Caption = "Min: 40 / Max: 40"
    lblPermisos.Caption = "Todas las Clases"
    End If
    
     If List1.Text = "Ropa del Mal" Then
    Picture1.Picture = LoadPicture(DirGraficos & "16036.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "500 Puntos de Canje"
    lblStat.Caption = "Min: 40 / Max: 40"
    lblPermisos.Caption = "Todas las Clases"
    End If
    
     If List1.Text = "Corona Oscura(RM)" Then
    Picture1.Picture = LoadPicture(DirGraficos & "2023.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "1800 Puntos de Canje"
    lblStat.Caption = "Min: 10 / Max: 11"
    lblPermisos.Caption = "Todas las Clases(RM)"
    End If
    
     If List1.Text = "Corona de Rey(RM)" Then
    Picture1.Picture = LoadPicture(DirGraficos & "2631.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "2600 Puntos de Canje"
    lblStat.Caption = "Min: 11 / Max: 12"
    lblPermisos.Caption = "Todas las Clases(RM)"
    End If
    
     If List1.Text = "Pantalon Celeste" Then
    Picture1.Picture = LoadPicture(DirGraficos & "16106.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "750 Puntos de Canje"
    lblStat.Caption = "Min: 35 / Max: 40"
    lblPermisos.Caption = "Todas las Clases"
    End If
    
    If List1.Text = "Pantalon Rojo" Then
    Picture1.Picture = LoadPicture(DirGraficos & "16104.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "750 Puntos de Canje"
    lblStat.Caption = "Min: 35 / Max: 40"
    lblPermisos.Caption = "Todas las Clases"
    End If

    If List1.Text = "Pantalon Gris" Then
    Picture1.Picture = LoadPicture(DirGraficos & "16276.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "750 Puntos de Canje"
    lblStat.Caption = "Min: 35 / Max: 40"
    lblPermisos.Caption = "Todas las Clases"
    End If
    
    If List1.Text = "Galera" Then
    Picture1.Picture = LoadPicture(DirGraficos & "464.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "3500 Puntos de Canje"
    lblStat.Caption = "Min: 13 / Max: 13"
    lblPermisos.Caption = "Todas las Clases(RM)"
    End If

    
End Sub

