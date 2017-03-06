VERSION 5.00
Begin VB.Form autor 
   Caption         =   "autor"
   ClientHeight    =   6330
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   10875
   LinkTopic       =   "Form1"
   ScaleHeight     =   6330
   ScaleWidth      =   10875
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   1200
      TabIndex        =   10
      Top             =   2640
      Width           =   3135
      Begin VB.CommandButton left 
         DragIcon        =   "autor.frx":0000
         Height          =   615
         Left            =   0
         Picture         =   "autor.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton rigth 
         DragIcon        =   "autor.frx":1194
         Height          =   615
         Left            =   1560
         Picture         =   "autor.frx":1A5E
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame comandos 
      Caption         =   "COMANDOS"
      Height          =   3255
      Left            =   6120
      TabIndex        =   7
      Top             =   0
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
         TabIndex        =   9
         Top             =   1920
         Width           =   2415
      End
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
         TabIndex        =   8
         Top             =   480
         Width           =   2415
      End
   End
   Begin VB.Frame principal 
      Caption         =   "PRINCIPAL"
      Height          =   3975
      Left            =   0
      TabIndex        =   0
      Top             =   -240
      Width           =   6015
      Begin VB.Data data1 
         Caption         =   "BASE DE DATOS"
         Connect         =   "Access"
         DatabaseName    =   "F:\EMPRESA_DE_DISCOS\EMPRESAS DE DISCOS\discos.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   1320
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "AUTOR"
         Top             =   3240
         Width           =   2775
      End
      Begin VB.TextBox fechnac 
         DataField       =   "Fecha de nacimiento"
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
         TabIndex        =   6
         Top             =   2040
         Width           =   2775
      End
      Begin VB.TextBox nombre 
         DataField       =   "Nombre"
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
         TabIndex        =   5
         Top             =   1200
         Width           =   2775
      End
      Begin VB.TextBox cod 
         DataField       =   "Codigo"
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
         TabIndex        =   4
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label Label3 
         Caption         =   "FECHA DE NACIMIENTO"
         BeginProperty Font 
            Name            =   "Ravie"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   3
         Top             =   2040
         Width           =   2295
      End
      Begin VB.Label Label2 
         Caption         =   "NOMBRE"
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
         Left            =   120
         TabIndex        =   2
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "CODIGO"
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
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Menu menu 
      Caption         =   "MENU PRICIPAL"
   End
   Begin VB.Menu volver 
      Caption         =   "VOLVER"
   End
   Begin VB.Menu salir 
      Caption         =   "SALIR"
   End
End
Attribute VB_Name = "autor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub hola_Click()

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

Private Sub menu_Click()
inicio.Show
Me.Hide
End Sub

Private Sub new_Click()
If Data1.Recordset.EOF = False And Data1.Recordset.BOF = False Then
    Data1.Recordset.AddNew
    Data1.Recordset("Codigo") = cod.Text
    Data1.Recordset("Nombre") = nombre.Text
    Data1.Recordset("Fecha de nacimiento") = fechnac.Text
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
