VERSION 5.00
Begin VB.Form registro 
   Caption         =   "registro"
   ClientHeight    =   3195
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   6405
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   16.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   6405
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton regresar 
      Caption         =   "VOLVER A INICIO"
      Height          =   735
      Left            =   1200
      TabIndex        =   0
      Top             =   360
      Width           =   3135
   End
   Begin VB.Menu salir 
      Caption         =   "SALIR"
   End
   Begin VB.Menu tip 
      Caption         =   "TIPOS"
   End
   Begin VB.Menu peli 
      Caption         =   "PELICULAS"
   End
   Begin VB.Menu dissc 
      Caption         =   "DISCO"
   End
   Begin VB.Menu clien 
      Caption         =   "CLIENTE"
   End
   Begin VB.Menu aut 
      Caption         =   "AUTOR"
   End
   Begin VB.Menu alqu 
      Caption         =   "ALQUILER"
   End
End
Attribute VB_Name = "registro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub alqu_Click()
alquiler.Show
Me.Hide
End Sub

Private Sub aut_Click()
autor.Show
Me.Hide
End Sub

Private Sub clien_Click()
cliente.Show
Me.Hide
End Sub

Private Sub dissc_Click()
disco.Show
Me.Hide
End Sub

Private Sub peli_Click()
pelicula.Show
Me.Hide
End Sub

Private Sub regresar_Click()
inicio.Show
Me.Hide
If Click Then
User = ""
pass = ""
End If
End Sub

Private Sub salir_Click()
End
End Sub


Private Sub tip_Click()
tipodepelicula.Show
Me.Hide

End Sub
