VERSION 5.00
Begin VB.Form alquiler 
   Caption         =   "alquiler"
   ClientHeight    =   7575
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   9555
   LinkTopic       =   "Form1"
   ScaleHeight     =   7575
   ScaleWidth      =   9555
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame comandos 
      Caption         =   "COMANDOS"
      Height          =   6735
      Left            =   6120
      TabIndex        =   9
      Top             =   240
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
         Left            =   120
         TabIndex        =   11
         Top             =   5400
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
         TabIndex        =   10
         Top             =   480
         Width           =   2415
      End
   End
   Begin VB.Frame principal 
      Caption         =   "PRINCIPAL"
      Height          =   6975
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   6015
      Begin VB.Frame Frame1 
         Height          =   1095
         Left            =   1200
         TabIndex        =   18
         Top             =   5760
         Width           =   3135
         Begin VB.CommandButton left 
            DragIcon        =   "alquiler.frx":0000
            Height          =   615
            Left            =   0
            Picture         =   "alquiler.frx":08CA
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   240
            Width           =   1575
         End
         Begin VB.CommandButton rigth 
            DragIcon        =   "alquiler.frx":1194
            Height          =   615
            Left            =   1560
            Picture         =   "alquiler.frx":1A5E
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.TextBox cant 
         DataField       =   "Cantidad"
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
         Left            =   2760
         TabIndex        =   17
         Top             =   5040
         Width           =   2775
      End
      Begin VB.TextBox valalq 
         DataField       =   "valor alquiler"
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
         TabIndex        =   15
         Top             =   4200
         Width           =   2775
      End
      Begin VB.TextBox fecdev 
         DataField       =   "Fecha de devolucion"
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
         TabIndex        =   13
         Top             =   3480
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
         Left            =   1320
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "ALQUILER"
         Top             =   6120
         Width           =   2775
      End
      Begin VB.TextBox ccliente 
         DataField       =   "Cod_cliente"
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
         Top             =   2040
         Width           =   2775
      End
      Begin VB.TextBox cdisco 
         DataField       =   "Cod_disco"
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
         TabIndex        =   3
         Top             =   1200
         Width           =   2775
      End
      Begin VB.TextBox codigo 
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
         TabIndex        =   2
         Top             =   480
         Width           =   2775
      End
      Begin VB.TextBox fealqu 
         DataField       =   "Fecha de alquiler"
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
         Top             =   2760
         Width           =   2775
      End
      Begin VB.Label Label7 
         Caption         =   "CANTIDAD"
         BeginProperty Font 
            Name            =   "Ravie"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   16
         Top             =   5160
         Width           =   2535
      End
      Begin VB.Label Label6 
         Caption         =   "VALOR ALQUILER"
         BeginProperty Font 
            Name            =   "Ravie"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   14
         Top             =   4440
         Width           =   2535
      End
      Begin VB.Label Label5 
         Caption         =   "FECH_DEVOLUCION"
         BeginProperty Font 
            Name            =   "Ravie"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   3600
         Width           =   2535
      End
      Begin VB.Label Label3 
         Caption         =   "COD_CLIENTE"
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
         TabIndex        =   8
         Top             =   2040
         Width           =   2295
      End
      Begin VB.Label Label2 
         Caption         =   "COD_DISCO"
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
         TabIndex        =   7
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "CODIGO"
         BeginProperty Font 
            Name            =   "Ravie"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label4 
         Caption         =   "FECH_ALQUILER"
         BeginProperty Font 
            Name            =   "Ravie"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   2760
         Width           =   2295
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
Attribute VB_Name = "alquiler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub agregar_Click()
If Data1.Recordset.EOF = False And Data1.Recordset.BOF = False Then
    Data1.Recordset.AddNew
    Data1.Recordset("Codigo") = codigo.Text
    Data1.Recordset("Cod_disco") = cdisco.Text
    Data1.Recordset("Cod_cliente") = ccliente.Text
    Data1.Recordset("Fecha de alquiler") = fealqu.Text
    Data1.Recordset("Fecha de devolucion") = fecdev.Text
    Data1.Recordset("valor alquiler") = valalq.Text
    Data1.Recordset("Cantidad") = cant.Text
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
