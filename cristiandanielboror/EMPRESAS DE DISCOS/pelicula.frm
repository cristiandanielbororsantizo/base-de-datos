VERSION 5.00
Begin VB.Form pelicula 
   Caption         =   "pelicula"
   ClientHeight    =   4800
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   10245
   LinkTopic       =   "Form1"
   ScaleHeight     =   4800
   ScaleWidth      =   10245
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame comandos 
      Caption         =   "COMANDOS"
      Height          =   3255
      Left            =   6240
      TabIndex        =   5
      Top             =   720
      Width           =   3015
      Begin VB.CommandButton eliminar 
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
         TabIndex        =   7
         Top             =   1920
         Width           =   2415
      End
      Begin VB.CommandButton agregar 
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
         TabIndex        =   6
         Top             =   480
         Width           =   2415
      End
   End
   Begin VB.Frame principal 
      Caption         =   "PRINCIPAL"
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   6015
      Begin VB.Frame Frame1 
         Height          =   1095
         Left            =   960
         TabIndex        =   8
         Top             =   1560
         Width           =   3135
         Begin VB.CommandButton left 
            DragIcon        =   "pelicula.frx":0000
            Height          =   615
            Left            =   0
            Picture         =   "pelicula.frx":08CA
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   240
            Width           =   1575
         End
         Begin VB.CommandButton rigth 
            DragIcon        =   "pelicula.frx":1194
            Height          =   615
            Left            =   1560
            Picture         =   "pelicula.frx":1A5E
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Data data1 
         Caption         =   "BASE DE DATOS"
         Connect         =   "Access"
         DatabaseName    =   "F:\EMPRESA_DE_DISCOS\EMPRESAS DE DISCOS\discos.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   1200
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "PECICULA"
         Top             =   2040
         Width           =   2775
      End
      Begin VB.TextBox cactor 
         DataField       =   "Cod_actor"
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
         Top             =   1200
         Width           =   2775
      End
      Begin VB.TextBox ctipo 
         DataField       =   "Cod_tipo"
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
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label Label2 
         Caption         =   "COD_ACTOR"
         BeginProperty Font 
            Name            =   "Ravie"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "COD_TIPO"
         BeginProperty Font 
            Name            =   "Ravie"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   2655
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
Attribute VB_Name = "pelicula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub agregar_Click()
If Data1.Recordset.EOF = False And Data1.Recordset.BOF = False Then
    Data1.Recordset.AddNew
    Data1.Recordset("Cod_tipo") = ctipo.Text
    Data1.Recordset("Cod_actor") = cactor.Text
    Data1.Recordset.Update
    End If
End Sub

Private Sub eliminar_Click()
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
