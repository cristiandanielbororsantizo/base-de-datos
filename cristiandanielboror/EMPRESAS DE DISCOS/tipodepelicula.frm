VERSION 5.00
Begin VB.Form tipodepelicula 
   Caption         =   "tipo de pelicula"
   ClientHeight    =   5985
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   9960
   Icon            =   "tipodepelicula.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5985
   ScaleWidth      =   9960
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   1680
      TabIndex        =   8
      Top             =   2640
      Width           =   3135
      Begin VB.CommandButton rigth 
         DragIcon        =   "tipodepelicula.frx":08CA
         Height          =   615
         Left            =   1560
         Picture         =   "tipodepelicula.frx":1194
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton left 
         DragIcon        =   "tipodepelicula.frx":1A5E
         Height          =   615
         Left            =   0
         Picture         =   "tipodepelicula.frx":2328
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame comandos 
      Caption         =   "COMANDOS"
      Height          =   3375
      Left            =   6600
      TabIndex        =   5
      Top             =   360
      Width           =   3015
      Begin VB.CommandButton new 
         Caption         =   "AGREGAR"
         BeginProperty Font 
            Name            =   "Wide Latin"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Width           =   2415
      End
      Begin VB.CommandButton delete 
         Caption         =   "ELIMINAR"
         BeginProperty Font 
            Name            =   "Wide Latin"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         TabIndex        =   6
         Top             =   1920
         Width           =   2415
      End
   End
   Begin VB.Frame principal 
      Caption         =   "PRINCIPAL"
      Height          =   3375
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   6015
      Begin VB.TextBox tipo 
         DataField       =   "Tipo"
         DataSource      =   "data1"
         BeginProperty Font 
            Name            =   "Wide Latin"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2880
         TabIndex        =   2
         Top             =   480
         Width           =   2775
      End
      Begin VB.TextBox cate 
         DataField       =   "Categoria"
         DataSource      =   "data1"
         BeginProperty Font 
            Name            =   "Wide Latin"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2880
         TabIndex        =   1
         Top             =   1200
         Width           =   2775
      End
      Begin VB.Data data1 
         Caption         =   "BASE DE DATOS"
         Connect         =   "Access"
         DatabaseName    =   "F:\EMPRESA_DE_DISCOS\EMPRESAS DE DISCOS\discos.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   1440
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "TIPO DE PELICULA"
         Top             =   2760
         Width           =   2775
      End
      Begin VB.Label Label1 
         Caption         =   "TIPO"
         BeginProperty Font 
            Name            =   "Ravie"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   600
         TabIndex        =   4
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "CATEGORIA"
         BeginProperty Font 
            Name            =   "Ravie"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   2415
      End
   End
   Begin VB.Menu mprincipal 
      Caption         =   "MENU PRINCIPAL"
   End
   Begin VB.Menu volver 
      Caption         =   "VOLVER"
   End
   Begin VB.Menu salir 
      Caption         =   "SALIR"
   End
End
Attribute VB_Name = "tipodepelicula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub delete_Click()
If Data1.Recordset.EOF = False And Data1.Recordset.BOF = False Then
    Data1.Recordset.Delete
    Data1.Recordset.Requery
    End If
End Sub

Private Sub left_Click()
 Data1.Recordset.MovePrevious
If Data1.Recordset.BOF = True Then
    Data1.Recordset.MoveLast
 End If
End Sub

Private Sub mprincipal_Click()
inicio.Show
Me.Hide
End Sub

Private Sub new_Click()
If Data1.Recordset.EOF = False And Data1.Recordset.BOF = False Then
    Data1.Recordset.AddNew
    Data1.Recordset("Tipo") = tipo.Text
    Data1.Recordset("Categoria") = cate.Text
    Data1.Recordset.Update
    End If
End Sub

Private Sub rigth_Click()
Data1.Recordset.MoveNext
If Data1.Recordset.EOF = True Then
Data1.Recordset.MoveFirst
End If
End Sub

Private Sub salir_Click()
End
End Sub
Private Sub volver_Click()
registro.Show
Me.Hide
End Sub
