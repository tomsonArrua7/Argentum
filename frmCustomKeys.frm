VERSION 5.00
Begin VB.Form frmCustomKeys 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5910
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   394
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   542
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   30
      Left            =   3960
      TabIndex        =   59
      Text            =   "Text1"
      Top             =   3600
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar y Salir"
      Height          =   375
      Left            =   3960
      TabIndex        =   38
      Top             =   5400
      Width           =   4095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cargar Teclas por defecto"
      Height          =   375
      Left            =   3960
      TabIndex        =   37
      Top             =   4920
      Width           =   4095
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00000000&
      Caption         =   "Otros"
      ForeColor       =   &H00FFFFFF&
      Height          =   3015
      Left            =   3960
      TabIndex        =   4
      Top             =   1680
      Width           =   4095
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   29
         Left            =   0
         TabIndex        =   58
         Text            =   "Text1"
         Top             =   480
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   28
         Left            =   0
         TabIndex        =   57
         Text            =   "Text1"
         Top             =   720
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   27
         Left            =   0
         TabIndex        =   56
         Text            =   "Text1"
         Top             =   960
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   26
         Left            =   0
         TabIndex        =   55
         Text            =   "Text1"
         Top             =   1200
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   25
         Left            =   0
         TabIndex        =   54
         Text            =   "Text1"
         Top             =   1440
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   24
         Left            =   0
         TabIndex        =   53
         Text            =   "Text1"
         Top             =   1680
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   19
         Left            =   2280
         TabIndex        =   52
         Text            =   "Text1"
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   17
         Left            =   2280
         TabIndex        =   50
         Text            =   "Text1"
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   18
         Left            =   2280
         TabIndex        =   49
         Text            =   "Text1"
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   20
         Left            =   2280
         TabIndex        =   43
         Text            =   "Text1"
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   21
         Left            =   2280
         TabIndex        =   40
         Text            =   "Text1"
         Top             =   1800
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   23
         Left            =   2280
         TabIndex        =   36
         Text            =   "Text1"
         Top             =   2520
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   22
         Left            =   2280
         TabIndex        =   35
         Text            =   "Text1"
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Invisibilidad"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   45
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Party"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   960
         TabIndex        =   44
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Salir"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   42
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Capturar Pantalla"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   39
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Destrabar"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   34
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Meditar"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   33
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Mostrar Opciones"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   32
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00000000&
      Caption         =   "Hablar"
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   5160
      Width           =   3735
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   16
         Left            =   1920
         TabIndex        =   51
         Text            =   "Text1"
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Hablar a Todos"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   600
         TabIndex        =   31
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      Caption         =   "Acciones"
      ForeColor       =   &H00FFFFFF&
      Height          =   3135
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   3735
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   14
         Left            =   1920
         TabIndex        =   48
         Text            =   "Text1"
         Top             =   2400
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   13
         Left            =   1920
         TabIndex        =   47
         Text            =   "Text1"
         Top             =   2040
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   9
         Left            =   1920
         TabIndex        =   46
         Text            =   "Text1"
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   15
         Left            =   1920
         TabIndex        =   30
         Text            =   "Text1"
         Top             =   2760
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   12
         Left            =   1920
         TabIndex        =   29
         Text            =   "Text1"
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   11
         Left            =   1920
         TabIndex        =   28
         Text            =   "Text1"
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   10
         Left            =   1920
         TabIndex        =   27
         Text            =   "Text1"
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   8
         Left            =   1920
         TabIndex        =   26
         Text            =   "Text1"
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Atacar"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   600
         TabIndex        =   25
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Usar"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   600
         TabIndex        =   24
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Tirar"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   600
         TabIndex        =   23
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Ocultar"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   600
         TabIndex        =   22
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Robar"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   600
         TabIndex        =   21
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Domar"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   600
         TabIndex        =   20
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Equipar"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   600
         TabIndex        =   19
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Agarrar"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   600
         TabIndex        =   18
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Opciones Personales"
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   3960
      TabIndex        =   1
      Top             =   120
      Width           =   4095
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   5
         Left            =   2280
         TabIndex        =   41
         Text            =   "Text1"
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   7
         Left            =   2280
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   6
         Left            =   2280
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Mostrar/Ocultar Nombres"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Corregir Posicion"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Activar/Desactivar Musica"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Movimiento"
      ForeColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   4
         Left            =   1920
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   3
         Left            =   1920
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   2
         Left            =   1920
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   1920
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Derecha"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   600
         TabIndex        =   8
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Izquierda"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   600
         TabIndex        =   7
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Abajo"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   600
         TabIndex        =   6
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Arriba"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   600
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmCustomKeys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.11.6
'
'Copyright (C) 2002 M�rquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Mat�as Fernando Peque�o
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 n�mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C�digo Postal 1900
'Pablo Ignacio M�rquez

''
'frmCustomKeys - Allows the user to customize keys.
'Implements class clsCustomKeys
'
'@author Rapsodius
'@date 20070805
'@version 1.0.0
'@see clsCustomKeys

Option Explicit

Private Sub command1_Click()
Call CustomKeys.LoadDefaults
Dim i As Long

For i = 1 To CustomKeys.Count
    Text1(i).Text = CustomKeys.ReadableName(CustomKeys.BindedKey(i))
Next i
End Sub

Private Sub Command2_Click()

Dim i As Long

For i = 1 To CustomKeys.Count
    If LenB(Text1(i).Text) = 0 Then
        Call MsgBox("Hay una o mas teclas no validas, por favor verifique.", vbCritical Or vbOKOnly Or vbApplicationModal Or vbDefaultButton1, "Argentum Online")
        Exit Sub
    Else
        CustomKeys.SaveCustomKeys
    End If
Next i

Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Long
    
    For i = 1 To CustomKeys.Count
        Text1(i).Text = CustomKeys.ReadableName(CustomKeys.BindedKey(i))
    Next i
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim i As Long
    
    If LenB(CustomKeys.ReadableName(KeyCode)) = 0 Then Exit Sub
    'If key is not valid, we exit
    
    Text1(Index).Text = CustomKeys.ReadableName(KeyCode)
    Text1(Index).SelStart = Len(Text1(Index).Text)
    
    For i = 1 To CustomKeys.Count
        If i <> Index Then
            If CustomKeys.BindedKey(i) = KeyCode Then
                Text1(Index).Text = "" 'If the key is already assigned, simply reject it
                Call Beep 'Alert the user
                KeyCode = 0
                Exit Sub
            End If
        End If
    Next i
    
    CustomKeys.BindedKey(Index) = KeyCode
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Call Text1_KeyDown(Index, KeyCode, Shift)
End Sub
