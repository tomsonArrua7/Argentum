VERSION 5.00
Begin VB.Form frmOpciones 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   4410
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4485
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOpciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmOpciones.frx":0152
   ScaleHeight     =   4410
   ScaleWidth      =   4485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.HScrollBar HScroll2 
      Height          =   255
      Left            =   720
      Max             =   60
      Min             =   30
      TabIndex        =   21
      Top             =   1440
      Value           =   30
      Width           =   1335
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   2160
      Max             =   100
      Min             =   1
      TabIndex        =   20
      Top             =   1440
      Value           =   1
      Width           =   1695
   End
   Begin VB.CommandButton cmdKeys 
      Caption         =   "Config. Teclas"
      Height          =   375
      Left            =   720
      TabIndex        =   19
      Top             =   3280
      Width           =   1335
   End
   Begin VB.PictureBox PictureSanado 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   335
      Left            =   2350
      MouseIcon       =   "frmOpciones.frx":A815
      MousePointer    =   99  'Custom
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   7
      Top             =   2880
      Width           =   335
   End
   Begin VB.PictureBox PictureRecuMana 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   335
      Left            =   840
      MouseIcon       =   "frmOpciones.frx":AB1F
      MousePointer    =   99  'Custom
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   6
      Top             =   2400
      Width           =   335
   End
   Begin VB.PictureBox PictureVestirse 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   335
      Left            =   840
      MouseIcon       =   "frmOpciones.frx":AE29
      MousePointer    =   99  'Custom
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   5
      Top             =   2880
      Width           =   335
   End
   Begin VB.PictureBox PictureMenosCansado 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   335
      Left            =   2350
      MouseIcon       =   "frmOpciones.frx":B133
      MousePointer    =   99  'Custom
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   4
      Top             =   2400
      Width           =   335
   End
   Begin VB.PictureBox PictureNoHayNada 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   335
      Left            =   2350
      MouseIcon       =   "frmOpciones.frx":B43D
      MousePointer    =   99  'Custom
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   3
      Top             =   1920
      Width           =   335
   End
   Begin VB.PictureBox PictureOcultarse 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   335
      Left            =   840
      MouseIcon       =   "frmOpciones.frx":B747
      MousePointer    =   99  'Custom
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   2
      Top             =   1920
      Width           =   335
   End
   Begin VB.PictureBox PictureFxs 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   335
      Left            =   2520
      MouseIcon       =   "frmOpciones.frx":BA51
      MousePointer    =   99  'Custom
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   1
      Top             =   1080
      Width           =   335
   End
   Begin VB.PictureBox PictureMusica 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   335
      Left            =   840
      MouseIcon       =   "frmOpciones.frx":BD5B
      MousePointer    =   99  'Custom
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   0
      Top             =   1080
      Width           =   335
   End
   Begin VB.Label Label11 
      BackColor       =   &H00000000&
      Caption         =   "Manual"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3000
      MouseIcon       =   "frmOpciones.frx":C065
      MousePointer    =   99  'Custom
      TabIndex        =   18
      Top             =   3380
      Width           =   735
   End
   Begin VB.Label Label10 
      BackColor       =   &H00000000&
      Caption         =   "Has sanado"
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
      Height          =   255
      Left            =   2715
      TabIndex        =   17
      Top             =   2940
      Width           =   1095
   End
   Begin VB.Label Label9 
      BackColor       =   &H00000000&
      Caption         =   "Meditaci�n"
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
      Height          =   255
      Left            =   1200
      TabIndex        =   16
      Top             =   2460
      Width           =   975
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000000&
      Caption         =   "No hay nada aqu�"
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
      Height          =   435
      Left            =   2715
      TabIndex        =   15
      Top             =   1920
      Width           =   1140
   End
   Begin VB.Label Label7 
      BackColor       =   &H00000000&
      Caption         =   "Abrigarse"
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
      Height          =   255
      Left            =   1200
      TabIndex        =   14
      Top             =   2940
      Width           =   975
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      Caption         =   "Menos cansado"
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
      Height          =   450
      Left            =   2715
      TabIndex        =   13
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   "Ocultarse"
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
      Height          =   255
      Left            =   1200
      TabIndex        =   12
      Top             =   1980
      Width           =   975
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   "Mostrar carteles"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   720
      TabIndex        =   11
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "Opciones de sonido"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1560
      TabIndex        =   10
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "FXs"
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
      Height          =   255
      Left            =   2880
      TabIndex        =   9
      Top             =   1125
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "M�sica"
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
      Height          =   255
      Left            =   1200
      TabIndex        =   8
      Top             =   1125
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   3600
      MouseIcon       =   "frmOpciones.frx":C36F
      MousePointer    =   99  'Custom
      Top             =   4200
      Width           =   855
   End
End
Attribute VB_Name = "frmOpciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'F�nixAO 1.0
'
'Based on Argentum Online 0.99z
'Copyright (C) 2002 M�rquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'You can contact the original creator of Argentum Online at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 n�mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C�digo Postal 1900
'Pablo Ignacio M�rquez
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'You can contact me at:
'elpresi@fenixao.com.ar
'www.fenixao.com.ar
Private Sub Command2_Click()
Me.Visible = False
End Sub

Private Sub Form_Load()


Me.Picture = LoadPicture(DirGraficos & "OpcionesDelJuego.gif")

If Musica = 0 Then
    PictureMusica.Picture = LoadPicture(DirGraficos & "tick1.gif")
Else
    PictureMusica.Picture = LoadPicture(DirGraficos & "tick2.gif")
End If

If FX = 0 Then
    PictureFxs.Picture = LoadPicture(DirGraficos & "tick1.gif")
Else
    PictureFxs.Picture = LoadPicture(DirGraficos & "tick2.gif")
End If

If CartelOcultarse = 1 Then
    PictureOcultarse.Picture = LoadPicture(DirGraficos & "tick1.gif")
Else
    PictureOcultarse.Picture = LoadPicture(DirGraficos & "tick2.gif")
End If

If CartelMenosCansado = 1 Then
    PictureMenosCansado.Picture = LoadPicture(DirGraficos & "tick1.gif")
Else
    PictureMenosCansado.Picture = LoadPicture(DirGraficos & "tick2.gif")
End If

If CartelVestirse = 1 Then
    PictureVestirse.Picture = LoadPicture(DirGraficos & "tick1.gif")
Else
    PictureVestirse.Picture = LoadPicture(DirGraficos & "tick2.gif")
End If

If CartelNoHayNada = 1 Then
    PictureNoHayNada.Picture = LoadPicture(DirGraficos & "tick1.gif")
Else
    PictureNoHayNada.Picture = LoadPicture(DirGraficos & "tick2.gif")
End If

If CartelRecuMana = 1 Then
    PictureRecuMana.Picture = LoadPicture(DirGraficos & "tick1.gif")
Else
    PictureRecuMana.Picture = LoadPicture(DirGraficos & "tick2.gif")
End If

If CartelSanado = 1 Then
    PictureSanado.Picture = LoadPicture(DirGraficos & "tick1.gif")
Else
    PictureSanado.Picture = LoadPicture(DirGraficos & "tick2.gif")
End If

If (VOLUMEN_FX < HScroll1.min) Or (VOLUMEN_FX > HScroll1.max) Then
VOLUMEN_FX = HScroll1.min
End If
HScroll1.value = VOLUMEN_FX
 
 
If (VOLUMEN_MUSICA < HScroll2.min) Or (VOLUMEN_MUSICA > HScroll2.max) Then
VOLUMEN_MUSICA = HScroll2.min
End If
HScroll2.value = VOLUMEN_MUSICA

End Sub
Private Sub Image1_Click()

Me.Visible = False

End Sub
Private Sub HScroll1_Change()
'FX
Dim s As Integer
If (HScroll1.value < 0) Or (HScroll1.value > 100) Then Exit Sub
 
s = 10 ^ ((HScroll1.value + 900) / 1000 + 1)
 
Audio.SoundVolume = (s)
Call WriteVar(App.Path & "/Init/Opciones.opc", "CONFIG", "Vol_fx", Str(HScroll1.value))
 
 
End Sub
 
Private Sub HScroll2_Change()
'musica, distinto control que fx en valores
Dim s As Integer
 
If (HScroll2.value < 0) Or (HScroll2.value > 100) Then Exit Sub
 
Audio.MusicVolume = (HScroll2.value)
Call WriteVar(App.Path & "/Init/Opciones.opc", "CONFIG", "Vol_music", Str(HScroll2.value))
 
 
End Sub

Private Sub Picture1_Click()

If NoRes = 0 Then
    NoRes = 1
    Picture1.Picture = LoadPicture(DirGraficos & "tick1.gif")
    Call WriteVar(App.Path & "/Init/Opciones.opc", "CONFIG", "ModoVentana", 1)
Else
    NoRes = 0
    Picture1.Picture = LoadPicture(DirGraficos & "tick2.gif")
    Call WriteVar(App.Path & "/Init/Opciones.opc", "CONFIG", "ModoVentana", 0)
End If

MsgBox "Este cambio har� efecto reci�n la pr�xima vez que ejecutes el juego."

End Sub

Private Sub Label11_Click()
ShellExecute Me.hWnd, "open", "https://discord.gg/ewbXhNFyJa", "", "", 1
End Sub

Private Sub PictureFxs_Click()

Select Case FX
    Case 0
        FX = 1
        PictureFxs.Picture = LoadPicture(DirGraficos & "tick2.gif")
    Case 1
        FX = 0
        PictureFxs.Picture = LoadPicture(DirGraficos & "tick1.gif")
End Select

Call WriteVar(App.Path & "/Init/Opciones.opc", "CONFIG", "FX", Trim(Str(FX)))

End Sub
Private Sub PictureMenosCansado_Click()

If CartelMenosCansado = 0 Then
    CartelMenosCansado = 1
    PictureMenosCansado.Picture = LoadPicture(DirGraficos & "tick1.gif")
Else
    CartelMenosCansado = 0
    PictureMenosCansado.Picture = LoadPicture(DirGraficos & "tick2.gif")
End If

Call WriteVar(App.Path & "/Init/Opciones.opc", "CARTELES", "MenosCansado", Trim(Str(CartelMenosCansado)))

End Sub

Private Sub PictureMusica_Click()

Select Case Musica
    Case 0
        Musica = 1
        Audio.StopMidi
        PictureMusica.Picture = LoadPicture(DirGraficos & "tick2.gif")
    Case 1
        Musica = 0
        Audio.PlayMIDI CurMidi
        PictureMusica.Picture = LoadPicture(DirGraficos & "tick1.gif")
End Select
Call WriteVar(App.Path & "/Init/Opciones.opc", "CONFIG", "Musica", Trim(Str(Musica)))

End Sub

Private Sub PictureNoHayNada_Click()
If CartelNoHayNada = 0 Then
    CartelNoHayNada = 1
    PictureNoHayNada.Picture = LoadPicture(DirGraficos & "tick1.gif")
Else
    CartelNoHayNada = 0
    PictureNoHayNada.Picture = LoadPicture(DirGraficos & "tick2.gif")
End If
Call WriteVar(App.Path & "/Init/Opciones.opc", "CARTELES", "NoHayNada", Trim(Str(CartelNoHayNada)))

End Sub

Private Sub PictureOcultarse_Click()

If CartelOcultarse = 0 Then
    CartelOcultarse = 1
    PictureOcultarse.Picture = LoadPicture(DirGraficos & "tick1.gif")
Else
    CartelOcultarse = 0
    PictureOcultarse.Picture = LoadPicture(DirGraficos & "tick2.gif")
End If
Call WriteVar(App.Path & "/Init/Opciones.opc", "CARTELES", "Ocultarse", Trim(Str(CartelOcultarse)))
End Sub

Private Sub PictureRecuMana_Click()
If CartelRecuMana = 0 Then
    CartelRecuMana = 1
    PictureRecuMana.Picture = LoadPicture(DirGraficos & "tick1.gif")
Else
    CartelRecuMana = 0
    PictureRecuMana.Picture = LoadPicture(DirGraficos & "tick2.gif")
End If
Call WriteVar(App.Path & "/Init/Opciones.opc", "CARTELES", "RecuMana", Trim(Str(CartelRecuMana)))

End Sub

Private Sub PictureSanado_Click()
If CartelSanado = 0 Then
    CartelSanado = 1
    PictureSanado.Picture = LoadPicture(DirGraficos & "tick1.gif")
Else
    CartelSanado = 0
    PictureSanado.Picture = LoadPicture(DirGraficos & "tick2.gif")
End If
Call WriteVar(App.Path & "/Init/Opciones.opc", "CARTELES", "Sanado", Trim(Str(CartelSanado)))

End Sub

Private Sub PictureVestirse_Click()
If CartelVestirse = 0 Then
    CartelVestirse = 1
    PictureVestirse.Picture = LoadPicture(DirGraficos & "tick1.gif")
Else
    CartelVestirse = 0
    PictureVestirse.Picture = LoadPicture(DirGraficos & "tick2.gif")
End If
Call WriteVar(App.Path & "/Init/Opciones.opc", "CARTELES", "Vestirse", Trim(Str(CartelVestirse)))

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

   If bmoving = False And Button = vbLeftButton Then

      Dx3 = X

      dy = Y

      bmoving = True

   End If

   

End Sub

 

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

   If bmoving And ((X <> Dx3) Or (Y <> dy)) Then

      Move left + (X - Dx3), top + (Y - dy)

   End If

   

End Sub

 

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

   If Button = vbLeftButton Then

      bmoving = False

   End If

   

End Sub

Private Sub cmdKeys_Click()
Unload Me
    Call frmCustomKeys.Show(vbModeless, frmMain)
End Sub



