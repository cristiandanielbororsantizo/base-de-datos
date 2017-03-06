VERSION 5.00
Begin VB.Form inicio 
   AutoRedraw      =   -1  'True
   Caption         =   "inicio"
   ClientHeight    =   6765
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7785
   LinkTopic       =   "Form1"
   ScaleHeight     =   6765
   ScaleWidth      =   7785
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   4815
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Width           =   6735
      Begin VB.CommandButton salir 
         Caption         =   "SALIR"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3480
         TabIndex        =   7
         Top             =   3840
         Width           =   2655
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   "E:\EMPRESA_DE_DISCOS\EMPRESAS DE DISCOS\discos.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   4200
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "SLOGAN"
         Top             =   4080
         Width           =   1335
      End
      Begin VB.CheckBox mostrar 
         Caption         =   "MOSTRA LA CONTRASEÑA"
         Height          =   495
         Left            =   3480
         TabIndex        =   6
         Top             =   2640
         Width           =   2655
      End
      Begin VB.CommandButton iniciar 
         Caption         =   "ENTRAR"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   360
         TabIndex        =   5
         Top             =   3840
         Width           =   2655
      End
      Begin VB.TextBox txtpass 
         Alignment       =   2  'Center
         DataField       =   "Password"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   3240
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   1920
         Width           =   2535
      End
      Begin VB.TextBox txtuser 
         Alignment       =   2  'Center
         DataField       =   "Usuarios"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3240
         TabIndex        =   3
         Top             =   720
         Width           =   2535
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "PASSWORD"
         BeginProperty Font 
            Name            =   "Wide Latin"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         TabIndex        =   2
         Top             =   1920
         Width           =   2775
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "USUARIO"
         BeginProperty Font 
            Name            =   "Wide Latin"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         TabIndex        =   1
         Top             =   840
         Width           =   2655
      End
   End
End
Attribute VB_Name = "inicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public con As Integer
Private Sub mostrar_Click()
con = con + 1
If (con / 2) = Int((con / 2)) Then
txtpass.PasswordChar = "*"
Else
txtpass.PasswordChar = ""
End If
End Sub

Private Sub iniciar_Click()
If txtuser.Text = "Axel" And txtpass.Text = "123" Then
registro.Show
Me.Hide
txtuser.Text = ""
txtpass.Text = ""
Else
MsgBox "Usuario o Contraseña incorrecto", , "Error"
End If
End Sub

Private Sub salir_Click()
End
End Sub

