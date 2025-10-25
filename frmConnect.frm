VERSION 5.00
Begin VB.Form frmConnect 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Argentum Online"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ClipControls    =   0   'False
   FillColor       =   &H00000040&
   Icon            =   "frmConnect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frmConnect.frx":000C
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPass 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   4560
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2895
      Width           =   2445
   End
   Begin VB.TextBox txtUser 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4560
      MaxLength       =   20
      TabIndex        =   0
      Top             =   1575
      Width           =   2445
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   1560
      MouseIcon       =   "frmConnect.frx":48487
      MousePointer    =   99  'Custom
      Top             =   8400
      Width           =   615
   End
   Begin VB.Image imgWeb 
      Height          =   495
      Left            =   4080
      MouseIcon       =   "frmConnect.frx":48791
      MousePointer    =   99  'Custom
      Top             =   240
      Width           =   3735
   End
   Begin VB.Image imgGetPass 
      Height          =   195
      Left            =   4800
      MouseIcon       =   "frmConnect.frx":48A9B
      MousePointer    =   99  'Custom
      Top             =   5400
      Width           =   1965
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   0
      Left            =   4800
      MouseIcon       =   "frmConnect.frx":48DA5
      MousePointer    =   99  'Custom
      Top             =   4650
      Width           =   1980
   End
   Begin VB.Image Image1 
      Height          =   435
      Index           =   1
      Left            =   4800
      MouseIcon       =   "frmConnect.frx":490AF
      MousePointer    =   99  'Custom
      Top             =   3600
      Width           =   2010
   End
   Begin VB.Image Image1 
      Height          =   195
      Index           =   2
      Left            =   4800
      MouseIcon       =   "frmConnect.frx":493B9
      MousePointer    =   99  'Custom
      Top             =   5040
      Width           =   1965
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FénixAO 1.0
'
'Based on Argentum Online 0.99z
'Copyright (C) 2002 Márquez Pablo Ignacio
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
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'You can contact me at:
'elpresi@fenixao.com.ar
'www.fenixao.com.ar
Option Explicit

Private Sub Command1_Click()
Password.left = RandomNumber(1, 9150)
Password.top = RandomNumber(1, 7500)
Password.Show
Password.SetFocus

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyReturn Then
    Call Audio.PlayWave(SND_CLICK)
            
    If frmMain.Socket1.Connected Then frmMain.Socket1.Disconnect
    
    If frmConnect.MousePointer = 11 Then
    frmConnect.MousePointer = 1
        Exit Sub
    End If
    
    
    UserName = txtUser.Text
    Dim aux As String
    aux = txtPass.Text
    UserPassword = MD5String(aux)
    If CheckUserData(False) = True Then
        frmMain.Socket1.HostName = IPdelServidor
        frmMain.Socket1.RemotePort = PuertoDelServidor
        
        EstadoLogin = Normal
        Me.MousePointer = 11
        frmMain.Socket1.Connect
        frmMain.SendTxt.Visible = False
    End If

End If

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 27 Then
    frmCargando.Show
    frmCargando.Refresh
    AddtoRichTextBox frmCargando.Status, "Cerrando TrhynumAO.", 255, 150, 50, 1, 0, 1
    
    frmConnect.MousePointer = 1
    frmMain.MousePointer = 1
    prgRun = False
    
    AddtoRichTextBox frmCargando.Status, "Liberando recursos..."
    frmCargando.Refresh
    AddtoRichTextBox frmCargando.Status, "Hecho", 255, 150, 50, 1, 0, 1
    AddtoRichTextBox frmCargando.Status, "¡¡Gracias por jugar FenixAO!!", 255, 150, 50, 1, 0, 1
    frmCargando.Refresh
    Call UnloadAllForms
    Call Resolution.ResetResolution
End If

End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)


If KeyCode = vbKeyI And Shift = vbCtrlMask Then
    
    

    
    
    


    
    
    KeyCode = 0
    Exit Sub
End If

End Sub

Private Sub Form_Load()
    
    EngineRun = False
    
    
 Dim j
 For Each j In Image1()
    j.Tag = "0"
 Next

 IntervaloPaso = 0.19
 IntervaloUsar = 1.2
 Picture = LoadPicture(DirGraficos & "conectar.jpg")


 
 
 
 
 
 

End Sub

Private Sub Image1_Click(Index As Integer)

CurServer = 0
Unload Password

Call Audio.PlayWave(SND_CLICK)

Select Case Index
    Case 0

            CurMidi = "7.mid"
            Call Audio.PlayMIDI(CurMidi)

       
        EstadoLogin = dados
        frmMain.Socket1.HostName = IPdelServidor
        frmMain.Socket1.RemotePort = PuertoDelServidor
        Me.MousePointer = 11
        frmMain.Socket1.Connect
        
    Case 1
        
        If frmMain.Socket1.Connected Then frmMain.Socket1.Disconnect
        
        If frmConnect.MousePointer = 11 Then
        frmConnect.MousePointer = 1
            Exit Sub
        End If
        
        
        
        UserName = txtUser.Text
        Dim aux As String
        aux = txtPass.Text
        UserPassword = MD5String(aux)
        If CheckUserData(False) = True Then
            frmMain.Socket1.HostName = IPdelServidor
            frmMain.Socket1.RemotePort = PuertoDelServidor
            
            EstadoLogin = Normal
            Me.MousePointer = 11
            frmMain.Socket1.Connect
        End If
        
    Case 2
       If frmMain.Socket1.Connected Then frmMain.Socket1.Disconnect
 
If frmConnect.MousePointer = 11 Then
frmConnect.MousePointer = 1
Exit Sub
End If
 
frmMain.Socket1.HostName = IPdelServidor
frmMain.Socket1.RemotePort = PuertoDelServidor
EstadoLogin = BorrarPJ
Me.MousePointer = 11
frmMain.Socket1.Connect

End Select

End Sub
Private Sub Image2_Click()

MsgBox "Created By Trhynum AO Team." & vbCrLf & "Codigo Libre" & vbCrLf & vbCrLf & "Web:" & vbCrLf & vbCrLf & "¡Gracias por Jugar nuestro Argentum Online!" & vbCrLf & "Staff Trhynun AO.", vbInformation, "Proyecto Trhynum"

End Sub
Private Sub imgGetPass_Click()

Call ShellExecute(Me.hWnd, "open", "TrhynumAO", "", "", 1)

End Sub
Private Sub imgWeb_Click()

Call ShellExecute(Me.hWnd, "open", "TrhynumAO", "", "", 1)

End Sub

