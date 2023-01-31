VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmBancos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ComparaMDB - Selecionar os Banco de Dados"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9285
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   9285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnAnalisar 
      Caption         =   "&Analisar"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   7080
      TabIndex        =   6
      Top             =   2190
      Width           =   2010
   End
   Begin MSComDlg.CommonDialog cdgMdb 
      Left            =   4710
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Microsoft Access (*.mdb)|*.mdb"
      Orientation     =   2
   End
   Begin VB.Frame Frame1 
      Caption         =   "Banco de Dados (Versão Anterior)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   180
      TabIndex        =   3
      Top             =   1200
      Width           =   8940
      Begin VB.CommandButton btnMdbAntigo 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8400
         TabIndex        =   4
         Top             =   270
         Width           =   330
      End
      Begin VB.Label lblCaminhoAntigo 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   120
         TabIndex        =   5
         Top             =   270
         Width           =   8250
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Banco de Dados (Versão Recente)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   180
      TabIndex        =   0
      Top             =   150
      Width           =   8940
      Begin VB.CommandButton btnMdbNovo 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8400
         TabIndex        =   2
         Top             =   270
         Width           =   330
      End
      Begin VB.Label lblCaminhoNovo 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   120
         TabIndex        =   1
         Top             =   270
         Width           =   8250
      End
   End
End
Attribute VB_Name = "frmBancos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Conectados()
    btnAnalisar.Enabled = False
    If cnnNovo.State <> ObjectStateEnum.adStateOpen Then Exit Sub
    If cnnAntigo.State <> ObjectStateEnum.adStateOpen Then Exit Sub
    btnAnalisar.Enabled = True
End Sub

Private Sub btnAnalisar_Click()
    Me.Visible = False
End Sub

Private Sub btnMdbAntigo_Click()
    On Error GoTo TrataErro
    Dim str As String
    Dim pass As String
    
    cdgMdb.ShowOpen
    lblCaminhoAntigo.Caption = cdgMdb.FileName
    
    pass = InputBox("Digite a senha do Banco de Dados", "ComparaMDB", Empty)
    
    str = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
          "Data Source=" & lblCaminhoAntigo.Caption & ";" & _
          "Jet OLEDB:Database Password=" & pass
    
    Call AbreMDB(cnnAntigo, catAntigo, str)
    Call Conectados
    Exit Sub
    
TrataErro:
    MsgBox Err.Description, vbCritical, "ComparaMDB"
    lblCaminhoAntigo.Caption = ""
End Sub

Private Sub btnMdbNovo_Click()
    On Error GoTo TrataErro
    Dim str As String
    Dim pass As String
    
    cdgMdb.ShowOpen
    lblCaminhoNovo.Caption = cdgMdb.FileName
    
    pass = InputBox("Digite a senha do Banco de Dados", "ComparaMDB", Empty)
    
    str = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
          "Data Source=" & lblCaminhoNovo.Caption & ";" & _
          "Jet OLEDB:Database Password=" & pass
    
    Call AbreMDB(cnnNovo, catNovo, str)
    Call Conectados
    Exit Sub
    
TrataErro:
    MsgBox Err.Description, vbCritical, "ComparaMDB"
    lblCaminhoNovo.Caption = ""
End Sub
