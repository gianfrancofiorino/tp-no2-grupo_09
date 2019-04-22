VERSION 5.00
Begin VB.Form Frm_Mensajes 
   BackColor       =   &H8000000E&
   ClientHeight    =   8985
   ClientLeft      =   720
   ClientTop       =   1530
   ClientWidth     =   11910
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   599
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   794
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label LblCartel 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   99.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9000
      Left            =   0
      TabIndex        =   41
      Top             =   0
      Width           =   12000
   End
   Begin VB.Label LblOrden 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   32
      Left            =   11160
      TabIndex        =   40
      Top             =   4950
      Width           =   750
   End
   Begin VB.Label LblOrden 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   29
      Left            =   11160
      TabIndex        =   39
      Top             =   3600
      Width           =   750
   End
   Begin VB.Label LblOrden 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   23
      Left            =   11160
      TabIndex        =   38
      Top             =   900
      Width           =   750
   End
   Begin VB.Label LblOrden 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   7
      Left            =   10440
      TabIndex        =   37
      Top             =   2700
      Width           =   720
   End
   Begin VB.Label LblNumber 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   300
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9000
      Left            =   0
      TabIndex        =   36
      Top             =   0
      Visible         =   0   'False
      Width           =   10425
   End
   Begin VB.Label LblOrden 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   40
      Left            =   11160
      TabIndex        =   35
      Top             =   8550
      Width           =   750
   End
   Begin VB.Label LblOrden 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   39
      Left            =   11160
      TabIndex        =   34
      Top             =   8100
      Width           =   750
   End
   Begin VB.Label LblOrden 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   38
      Left            =   11160
      TabIndex        =   33
      Top             =   7650
      Width           =   750
   End
   Begin VB.Label LblOrden 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   37
      Left            =   11160
      TabIndex        =   32
      Top             =   7200
      Width           =   750
   End
   Begin VB.Label LblOrden 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   36
      Left            =   11160
      TabIndex        =   31
      Top             =   6750
      Width           =   750
   End
   Begin VB.Label LblOrden 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   35
      Left            =   11160
      TabIndex        =   30
      Top             =   6300
      Width           =   750
   End
   Begin VB.Label LblOrden 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   34
      Left            =   11160
      TabIndex        =   29
      Top             =   5850
      Width           =   750
   End
   Begin VB.Label LblOrden 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   33
      Left            =   11160
      TabIndex        =   28
      Top             =   5400
      Width           =   750
   End
   Begin VB.Label LblOrden 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   31
      Left            =   11160
      TabIndex        =   27
      Top             =   4500
      Width           =   750
   End
   Begin VB.Label LblOrden 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   30
      Left            =   11160
      TabIndex        =   26
      Top             =   4050
      Width           =   750
   End
   Begin VB.Label LblOrden 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   28
      Left            =   11160
      TabIndex        =   25
      Top             =   3150
      Width           =   750
   End
   Begin VB.Label LblOrden 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   27
      Left            =   11160
      TabIndex        =   24
      Top             =   2700
      Width           =   750
   End
   Begin VB.Label LblOrden 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   26
      Left            =   11160
      TabIndex        =   23
      Top             =   2250
      Width           =   750
   End
   Begin VB.Label LblOrden 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   25
      Left            =   11160
      TabIndex        =   22
      Top             =   1800
      Width           =   750
   End
   Begin VB.Label LblOrden 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   24
      Left            =   11160
      TabIndex        =   21
      Top             =   1350
      Width           =   750
   End
   Begin VB.Label LblOrden 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   22
      Left            =   11160
      TabIndex        =   20
      Top             =   450
      Width           =   750
   End
   Begin VB.Label LblOrden 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   21
      Left            =   11160
      TabIndex        =   19
      Top             =   0
      Width           =   750
   End
   Begin VB.Label LblOrden 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   20
      Left            =   10440
      TabIndex        =   18
      Top             =   8550
      Width           =   750
   End
   Begin VB.Label LblOrden 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   19
      Left            =   10440
      TabIndex        =   17
      Top             =   8100
      Width           =   750
   End
   Begin VB.Label LblOrden 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   18
      Left            =   10440
      TabIndex        =   16
      Top             =   7650
      Width           =   750
   End
   Begin VB.Label LblOrden 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   17
      Left            =   10440
      TabIndex        =   15
      Top             =   7200
      Width           =   750
   End
   Begin VB.Label LblOrden 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   16
      Left            =   10440
      TabIndex        =   14
      Top             =   6750
      Width           =   750
   End
   Begin VB.Label LblOrden 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   15
      Left            =   10440
      TabIndex        =   13
      Top             =   6300
      Width           =   750
   End
   Begin VB.Label LblOrden 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   14
      Left            =   10440
      TabIndex        =   12
      Top             =   5850
      Width           =   750
   End
   Begin VB.Label LblOrden 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   13
      Left            =   10440
      TabIndex        =   11
      Top             =   5400
      Width           =   750
   End
   Begin VB.Label LblOrden 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   12
      Left            =   10440
      TabIndex        =   10
      Top             =   4950
      Width           =   750
   End
   Begin VB.Label LblOrden 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   11
      Left            =   10440
      TabIndex        =   9
      Top             =   4500
      Width           =   750
   End
   Begin VB.Label LblOrden 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   10
      Left            =   10440
      TabIndex        =   8
      Top             =   4050
      Width           =   750
   End
   Begin VB.Label LblOrden 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   9
      Left            =   10440
      TabIndex        =   7
      Top             =   3600
      Width           =   750
   End
   Begin VB.Label LblOrden 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   8
      Left            =   10440
      TabIndex        =   6
      Top             =   3150
      Width           =   750
   End
   Begin VB.Label LblOrden 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   6
      Left            =   10440
      TabIndex        =   5
      Top             =   2250
      Width           =   750
   End
   Begin VB.Label LblOrden 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   5
      Left            =   10440
      TabIndex        =   4
      Top             =   1800
      Width           =   750
   End
   Begin VB.Label LblOrden 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   4
      Left            =   10440
      TabIndex        =   3
      Top             =   1350
      Width           =   750
   End
   Begin VB.Label LblOrden 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   3
      Left            =   10440
      TabIndex        =   2
      Top             =   900
      Width           =   750
   End
   Begin VB.Label LblOrden 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   2
      Left            =   10440
      TabIndex        =   1
      Top             =   450
      Width           =   750
   End
   Begin VB.Label LblOrden 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   1
      Left            =   10440
      TabIndex        =   0
      Top             =   0
      Width           =   750
   End
End
Attribute VB_Name = "Frm_Mensajes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_KeyPress(KeyAscii As Integer)
        If cartel Then
        
        If KeyAscii = 27 And salir Then  'esc dos veces salgo programa
        End
               
        Else
            salir = False
        End If
        
        If KeyAscii = 119 And borrartodo Then  'W dos veces borro todos los numeros
        For i = 1 To 90
            Frm_Number.LblNumber(i).BackColor = &HFFFFFF 'pongo todo blanco
            bOrden(i) = False
        Next
        j = 1
        Else
            borrartodo = False
        End If
        
        For x = 1 To 40 'cualquier teclas salgo de cartel
            Frm_Mensajes.LblOrden(x).Visible = True
        Next
        End If
        Unload Frm_Mensajes
End Sub



Private Sub LblCartel_Click()
        For x = 1 To 40
            Frm_Mensajes.LblOrden(x).Visible = True
        Next
        Unload Frm_Mensajes
End Sub


Private Sub LblNumber_Click()
        For x = 1 To 40
            Frm_Mensajes.LblOrden(x).Visible = True
        Next
        Unload Frm_Mensajes
End Sub

Private Sub LblOrden_Click(Index As Integer)
        For x = 1 To 40
            Frm_Mensajes.LblOrden(x).Visible = True
        Next
        Unload Frm_Mensajes
        End Sub
