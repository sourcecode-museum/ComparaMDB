VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmAnalisar 
   Caption         =   "ComparaMDB - Comparando Banco de Dados"
   ClientHeight    =   10695
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
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
   ScaleHeight     =   10695
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton btnInverterMDB 
      Caption         =   "&Inverter"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   12780
      TabIndex        =   23
      Top             =   190
      Width           =   2310
   End
   Begin VB.CommandButton btnAnalisar 
      Caption         =   "&Analisar"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   12780
      TabIndex        =   22
      Top             =   870
      Width           =   2310
   End
   Begin MSComDlg.CommonDialog cdgMdb 
      Left            =   7440
      Top             =   30
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "ComparaMDB - Abrir arquivo MDB"
      Filter          =   "Microsoft Access (*.mdb)|*.mdb"
   End
   Begin VB.Frame Frame1 
      Caption         =   "Banco de Dados (Versão Anterior)"
      Height          =   5115
      Left            =   150
      TabIndex        =   7
      Top             =   5190
      Width           =   12480
      Begin MSFlexGridLib.MSFlexGrid grdTabelaAntiga 
         Height          =   3360
         Left            =   165
         TabIndex        =   8
         Top             =   1140
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   5927
         _Version        =   393216
         Cols            =   1
         FixedCols       =   0
         BackColor       =   12648384
         ForeColor       =   16711680
         BackColorBkg    =   12648384
         AllowBigSelection=   0   'False
         ScrollBars      =   2
         SelectionMode   =   1
         Appearance      =   0
         FormatString    =   "Tabelas"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid grdCampoAntigo 
         Height          =   3360
         Left            =   2490
         TabIndex        =   9
         Top             =   1140
         Width           =   4830
         _ExtentX        =   8520
         _ExtentY        =   5927
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         BackColor       =   12648384
         ForeColor       =   16711680
         BackColorBkg    =   12648384
         AllowBigSelection=   0   'False
         SelectionMode   =   1
         Appearance      =   0
         FormatString    =   "Campo | Tipo | Tamanho"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Frame Frame4 
         Caption         =   "Arquivo e Localização"
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
         Left            =   165
         TabIndex        =   16
         Top             =   390
         Width           =   12150
         Begin VB.TextBox txtCaminhoAntigo 
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   360
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   19
            Top             =   270
            Width           =   11520
         End
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
            Left            =   11640
            TabIndex        =   17
            Top             =   270
            Width           =   330
         End
      End
      Begin MSFlexGridLib.MSFlexGrid grdPropriedadeAntiga 
         Height          =   3810
         Left            =   7440
         TabIndex        =   21
         Top             =   1140
         Width           =   4860
         _ExtentX        =   8573
         _ExtentY        =   6720
         _Version        =   393216
         FixedCols       =   0
         BackColor       =   12648384
         ForeColor       =   16711680
         BackColorBkg    =   12648384
         AllowBigSelection=   0   'False
         ScrollBars      =   2
         SelectionMode   =   1
         Appearance      =   0
         FormatString    =   "Propriedade | Valor"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Campos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5400
         TabIndex        =   13
         Top             =   4635
         Width           =   870
      End
      Begin VB.Label lblNroCamposAntigo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   6390
         TabIndex        =   12
         Top             =   4605
         Width           =   930
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tabelas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   180
         TabIndex        =   11
         Top             =   4635
         Width           =   885
      End
      Begin VB.Label lblNroTabelasAntigo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   1440
         TabIndex        =   10
         Top             =   4605
         Width           =   930
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Banco de Dados (Versão Recente)"
      Height          =   5115
      Left            =   150
      TabIndex        =   0
      Top             =   60
      Width           =   12480
      Begin VB.Frame Frame5 
         Caption         =   "Arquivo e Localização"
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
         Left            =   165
         TabIndex        =   14
         Top             =   390
         Width           =   12150
         Begin VB.TextBox txtCaminhoNovo 
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   330
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   18
            Top             =   270
            Width           =   11520
         End
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
            Left            =   11640
            TabIndex        =   15
            Top             =   270
            Width           =   330
         End
      End
      Begin MSFlexGridLib.MSFlexGrid grdTabelaNova 
         Height          =   3360
         Left            =   165
         TabIndex        =   1
         Top             =   1140
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   5927
         _Version        =   393216
         Cols            =   1
         FixedCols       =   0
         BackColor       =   12648384
         ForeColor       =   16711680
         BackColorBkg    =   12648384
         AllowBigSelection=   0   'False
         SelectionMode   =   1
         AllowUserResizing=   3
         Appearance      =   0
         FormatString    =   "Tabelas"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid grdCampoNovo 
         Height          =   3360
         Left            =   2490
         TabIndex        =   4
         Top             =   1140
         Width           =   4830
         _ExtentX        =   8520
         _ExtentY        =   5927
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         BackColor       =   12648384
         ForeColor       =   16711680
         BackColorBkg    =   12648384
         AllowBigSelection=   0   'False
         ScrollBars      =   2
         SelectionMode   =   1
         Appearance      =   0
         FormatString    =   "Campo | Tipo | Tamanho"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid grdPropriedadeNova 
         Height          =   3810
         Left            =   7440
         TabIndex        =   20
         Top             =   1140
         Width           =   4860
         _ExtentX        =   8573
         _ExtentY        =   6720
         _Version        =   393216
         FixedCols       =   0
         BackColor       =   12648384
         ForeColor       =   16711680
         BackColorBkg    =   12648384
         AllowBigSelection=   0   'False
         ScrollBars      =   2
         SelectionMode   =   1
         Appearance      =   0
         FormatString    =   "Propriedade | Valor"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Campos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5400
         TabIndex        =   6
         Top             =   4635
         Width           =   870
      End
      Begin VB.Label lblNroCamposNovo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   6390
         TabIndex        =   5
         Top             =   4605
         Width           =   930
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Tabelas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   180
         TabIndex        =   3
         Top             =   4635
         Width           =   885
      End
      Begin VB.Label lblNroTabelasNovo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   1440
         TabIndex        =   2
         Top             =   4605
         Width           =   930
      End
   End
End
Attribute VB_Name = "frmAnalisar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub grdTabelaNova_Estruturar()
    With grdTabelaNova
        .Clear
        .Cols = 1
        .Rows = 1
        .ColAlignment(0) = 1: .ColWidth(0) = 2000: .TextArray(0) = "Tabela"
    End With
End Sub

Private Sub grdTabelaAntiga_Estruturar()
    With grdTabelaAntiga
        .Clear
        .Cols = 1
        .Rows = 1
        .ColAlignment(0) = 1: .ColWidth(0) = 2000: .TextArray(0) = "Tabela"
    End With
End Sub

Private Sub grdCampoNovo_Estruturar()
    With grdCampoNovo
        .Clear
        .Cols = 3
        .Rows = 1
        .ColAlignment(0) = 1: .ColWidth(0) = 2000: .TextArray(0) = "Campo"
        .ColAlignment(1) = 1: .ColWidth(1) = 1500: .TextArray(1) = "Tipo"
        .ColAlignment(2) = 1: .ColWidth(2) = 1000: .TextArray(2) = "Tamanho"
    End With
End Sub

Private Sub grdCampoAntigo_Estruturar()
    With grdCampoAntigo
        .Clear
        .Cols = 3
        .Rows = 1
        .ColAlignment(0) = 1: .ColWidth(0) = 2000: .TextArray(0) = "Campo"
        .ColAlignment(1) = 1: .ColWidth(1) = 1500: .TextArray(1) = "Tipo"
        .ColAlignment(2) = 1: .ColWidth(2) = 1000: .TextArray(2) = "Tamanho"
    End With
End Sub

Private Sub grdPropriedadeAntiga_Estruturar()
    With grdPropriedadeAntiga
        .Clear
        .Cols = 2
        .Rows = 1
        .ColAlignment(0) = 1: .ColWidth(0) = 3000: .TextArray(0) = "Nome"
        .ColAlignment(1) = 1: .ColWidth(1) = 2000: .TextArray(1) = "Valor"
    End With
End Sub

Private Sub grdPropriedadeNova_Estruturar()
    With grdPropriedadeNova
        .Clear
        .Cols = 2
        .Rows = 1
        .ColAlignment(0) = 1: .ColWidth(0) = 3000: .TextArray(0) = "Nome"
        .ColAlignment(1) = 1: .ColWidth(1) = 2000: .TextArray(1) = "Valor"
    End With
End Sub

Private Sub grdTabelaNova_Carregar()
    On Error Resume Next
    Dim i As Integer
    
    lblNroTabelasNovo.Caption = 0
    For i = 0 To catNovo.Tables.Count - 1
        If catNovo.Tables(i).Type = "TABLE" Or catNovo.Tables(i).Type = "SYNONYM" Then
            grdTabelaNova.AddItem catNovo.Tables.Item(i).Name
        End If
    Next
    lblNroTabelasNovo.Caption = grdTabelaNova.Rows - 1
    grdTabelaNova.Row = 0
End Sub

Private Sub grdTabelaAntiga_Carregar()
    On Error Resume Next
    Dim i As Integer
    
    lblNroTabelasAntigo.Caption = 0
    For i = 0 To catAntigo.Tables.Count - 1
        If catAntigo.Tables(i).Type = "TABLE" Or catAntigo.Tables(i).Type = "SYNONYM" Then
            grdTabelaAntiga.AddItem catAntigo.Tables.Item(i).Name
        End If
    Next
    lblNroTabelasAntigo.Caption = grdTabelaAntiga.Rows - 1
    grdTabelaAntiga.Row = 0
End Sub

Private Sub grdCampoNovo_Carregar(Tabela As String)
    On Error Resume Next
    Dim i As Integer
    
    lblNroCamposNovo.Caption = 0
    For i = 0 To catNovo.Tables(Tabela).Columns.Count - 1
        grdCampoNovo.AddItem catNovo.Tables(Tabela).Columns(i).Name & vbTab & _
            TipoField(catNovo.Tables(Tabela).Columns(i).Type) & vbTab & _
            IIf(catNovo.Tables(Tabela).Columns(i).DefinedSize > 0, catNovo.Tables(Tabela).Columns(i).DefinedSize, "")
    Next
    lblNroCamposNovo.Caption = catNovo.Tables(Tabela).Columns.Count
End Sub

Private Sub grdCampoAntigo_Carregar(Tabela As String)
    On Error Resume Next
    Dim i As Integer
    
    lblNroCamposAntigo.Caption = 0
    For i = 0 To catAntigo.Tables(Tabela).Columns.Count - 1
        grdCampoAntigo.AddItem catAntigo.Tables(Tabela).Columns(i).Name & vbTab & _
            TipoField(catAntigo.Tables(Tabela).Columns(i).Type) & vbTab & _
            IIf(catAntigo.Tables(Tabela).Columns(i).DefinedSize > 0, catAntigo.Tables(Tabela).Columns(i).DefinedSize, "")
    Next
    lblNroCamposAntigo.Caption = catAntigo.Tables(Tabela).Columns.Count
End Sub

Private Sub btnAnalisar_Click()
    Dim i As Integer
    Dim x As Integer

    Load frmResumo
    
    ''// LE as tabelas NOVA
    For i = 1 To grdTabelaNova.Rows - 1
        ''// LE as tabelas ANTIGA
        For x = 1 To grdTabelaAntiga.Rows - 1
            Dim tbNova As ADOX.Table
            Dim tbAntiga As ADOX.Table
            Set tbNova = catNovo.Tables(grdTabelaNova.TextMatrix(i, 0))
            Set tbAntiga = catAntigo.Tables(grdTabelaAntiga.TextMatrix(x, 0))
            
            ''// SE a TABELA existe em ambos BD, ENTAO Verifico os CAMPOS
            If tbNova.Name = tbAntiga.Name Then
                Dim a As Integer
                Dim b As Integer
                ''// LE os campos da tabela NOVA
                For a = 0 To tbNova.Columns.Count - 1
                    ''// LE os campos da tabela ANTIGA
                    For b = 0 To tbAntiga.Columns.Count - 1
                        Dim colNova As ADOX.Column
                        Dim colAntiga As ADOX.Column
                        Set colNova = tbNova.Columns(a)
                        Set colAntiga = tbAntiga.Columns(b)
                        
                        ''// SE o CAMPO existe em ambos BD, ENTAO Verifico as PROPRIEDADES
                        If colNova.Name = colAntiga.Name Then
                            If colNova.Type <> colAntiga.Type Then frmResumo.grdDiferencas.AddItem tbNova.Name & "." & colNova.Name & " {" & TipoField(colNova.Type) & "}" & vbTab & "TIPO DO CAMPO diferente"
                            If colNova.DefinedSize > 0 Then
                                If colNova.DefinedSize <> colAntiga.DefinedSize Then
                                    frmResumo.grdDiferencas.AddItem tbNova.Name & "." & colNova.Name & " {" & TipoField(colNova.Type) & "}" & vbTab & "TAMANHO DO CAMPO diferente"
                                End If
                            End If
                            
                            b = -1
                            Exit For
                        End If
                    Next
                    If b > -1 Then
                        frmResumo.grdDiferencas.AddItem tbNova.Name & "." & colNova.Name & vbTab & "CAMPO não existente na TABELA do BD_ANTIGO"
                        DoEvents
                    End If
                Next

                x = -1
                Exit For
            End If
        Next
        
        If x > -1 And Not tbNova Is Nothing Then
            frmResumo.grdDiferencas.AddItem tbNova.Name & vbTab & "TABELA não existente no BD_ANTIGO"
            DoEvents
        End If
    Next
    frmResumo.Show
End Sub

Private Sub btnInverterMDB_Click()
    If txtCaminhoNovo.Text = Empty Or txtCaminhoAntigo.Text = Empty Then Exit Sub
    Dim tempCaminho As String
    tempCaminho = txtCaminhoAntigo.Text
    txtCaminhoAntigo.Text = txtCaminhoNovo.Text
    txtCaminhoNovo.Text = tempCaminho
    
    Dim tempConn As ADODB.Connection
    Set tempConn = cnnAntigo
    Set cnnAntigo = cnnNovo
    Set cnnNovo = tempConn
    
    Set catNovo.ActiveConnection = cnnNovo
    Set catAntigo.ActiveConnection = cnnAntigo
    
    Call grdTabelas_Preencher
End Sub

Private Sub btnMdbAntigo_Click()
    On Error GoTo TrataErro
    Dim str As String
    
    cdgMdb.DialogTitle = "Selecione o 2º Arquivo.MDB (Versão Anterior)"
    cdgMdb.ShowOpen
    If cdgMdb.FileName = Empty Then Exit Sub
    txtCaminhoAntigo.Text = cdgMdb.FileName

    
    str = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
          "Data Source=" & txtCaminhoAntigo.Text & ";"
    
    Call AbreMDB(cnnAntigo, catAntigo, str)
    If txtCaminhoNovo.Text <> Empty Then Call grdTabelas_Preencher
    Exit Sub
    
TrataErro:
    MsgBox Err.Description, vbCritical, "ComparaMDB"
    txtCaminhoAntigo.Text = ""
End Sub

Private Sub btnMdbNovo_Click()
    On Error GoTo TrataErro
    Dim str As String
    
    cdgMdb.DialogTitle = "Selecione o 1º Arquivo.MDB (ATUALIZADO)"
    cdgMdb.ShowOpen
    If cdgMdb.FileName = Empty Then Exit Sub
    txtCaminhoNovo.Text = cdgMdb.FileName
    
    str = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
          "Data Source=" & txtCaminhoNovo.Text & ";"
    
    Call AbreMDB(cnnNovo, catNovo, str)
    If txtCaminhoAntigo.Text = Empty Then Call btnMdbAntigo_Click
    Exit Sub
    
TrataErro:
    MsgBox Err.Description, vbCritical, "ComparaMDB"
    txtCaminhoNovo.Text = ""
End Sub

Private Sub grdTabelas_Preencher()
    If txtCaminhoNovo.Text = Empty Or txtCaminhoAntigo.Text = Empty Then Exit Sub

    Call grdTabelaNova_Estruturar
    Call grdTabelaAntiga_Estruturar
    Call grdCampoNovo_Estruturar
    Call grdCampoAntigo_Estruturar
    Call grdPropriedadeNova_Estruturar
    Call grdPropriedadeAntiga_Estruturar

    Call grdTabelaNova_Carregar
    Call grdTabelaAntiga_Carregar
End Sub

Private Sub Form_Load()
    Call grdTabelaNova_Estruturar
    Call grdTabelaAntiga_Estruturar
    Call grdCampoNovo_Estruturar
    Call grdCampoAntigo_Estruturar
    Call grdPropriedadeNova_Estruturar
    Call grdPropriedadeAntiga_Estruturar
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call FechaMDB
End Sub

Private Sub grdCampoAntigo_EnterCell()
    If grdCampoAntigo.RowSel < 1 Then Exit Sub
    If grdCampoAntigo.Rows > 1 Then
        Dim i As Integer
        Call grdPropriedadeAntiga_Estruturar
        
        With catAntigo.Tables(grdTabelaAntiga.Text).Columns(grdCampoAntigo.Text)
            For i = 0 To .Properties.Count - 1
                grdPropriedadeAntiga.AddItem .Properties(i).Name & vbTab & _
                    .Properties(i).Value
            Next
        End With
    End If
End Sub

Private Sub grdCampoNovo_EnterCell()
    If grdCampoNovo.RowSel < 1 Then Exit Sub
    If grdCampoNovo.Rows > 1 Then
        Dim i As Integer
        Call grdPropriedadeNova_Estruturar
        
        With catNovo.Tables(grdTabelaNova.Text).Columns(grdCampoNovo.Text)
            For i = 0 To .Properties.Count - 1
                grdPropriedadeNova.AddItem .Properties(i).Name & vbTab & _
                    .Properties(i).Value
            Next
        End With
        
        Call grdPropriedadeAntiga_Estruturar
        For i = 1 To grdCampoAntigo.Rows - 1
            If grdCampoNovo.TextMatrix(grdCampoNovo.RowSel, 0) = grdCampoAntigo.TextMatrix(i, 0) Then
                grdCampoAntigo.Row = i
                grdCampoAntigo.RowSel = i
                Exit For
            End If
        Next
    End If
End Sub

Private Sub grdTabelaNova_EnterCell()
    If grdTabelaNova.RowSel < 1 Then Exit Sub
    If grdTabelaNova.Rows > 1 Then
        Call grdCampoNovo_Estruturar
        Call grdCampoNovo_Carregar(grdTabelaNova.TextMatrix(grdTabelaNova.RowSel, 0))
        
        Dim i As Integer
        Call grdCampoAntigo_Estruturar
        For i = 1 To grdTabelaAntiga.Rows - 1
            If grdTabelaNova.TextMatrix(grdTabelaNova.RowSel, 0) = grdTabelaAntiga.TextMatrix(i, 0) Then
                grdTabelaAntiga.Row = i
                grdTabelaAntiga.RowSel = i
                Exit For
            End If
        Next
    End If
End Sub

Private Sub grdTabelaAntiga_EnterCell()
    If grdTabelaAntiga.RowSel < 1 Then Exit Sub
    If grdTabelaAntiga.Rows > 1 Then
        Call grdCampoAntigo_Estruturar
        Call grdCampoAntigo_Carregar(grdTabelaAntiga.TextMatrix(grdTabelaAntiga.RowSel, 0))
    End If
End Sub

Private Function TipoField(ByVal TypeVal As Long) As String
    Select Case TypeVal
    Case adBigInt                    ' 20
        TipoField = "Big Integer"
    Case adBinary                    ' 128
        TipoField = "Binary"
    Case adBoolean                   ' 11
        TipoField = "Boolean"
    Case adBSTR                      ' 8 i.e. null terminated string
        TipoField = "Text"
    Case adChar                      ' 129
        TipoField = "Text"
    Case adCurrency                  ' 6
        TipoField = "Currency"
    Case adDate                      ' 7
        TipoField = "Date/Time"
    Case adDBDate                    ' 133
        TipoField = "Date/Time"
    Case adDBTime                    ' 134
        TipoField = "Date/Time"
    Case adDBTimeStamp               ' 135
        TipoField = "Date/Time"
    Case adDecimal                   ' 14
        TipoField = "Float"
    Case adDouble                    ' 5
        TipoField = "Float"
    Case adEmpty                     ' 0
        TipoField = "Empty"
    Case adError                     ' 10
        TipoField = "Error"
    Case adGUID                      ' 72
        TipoField = "GUID"
    Case adIDispatch                 ' 9
        TipoField = "IDispatch"
    Case adInteger                   ' 3
        TipoField = "Integer"
    Case adIUnknown                  ' 13
        TipoField = "Unknown"
    Case adLongVarBinary             ' 205
        TipoField = "Binary"
    Case adLongVarChar               ' 201
        TipoField = "Text"
    Case adLongVarWChar              ' 203
        TipoField = "Text"
    Case adNumeric                  ' 131
        TipoField = "Long"
    Case adSingle                    ' 4
        TipoField = "Single"
    Case adSmallInt                  ' 2
        TipoField = "Small Integer"
    Case adTinyInt                   ' 16
        TipoField = "Tiny Integer"
    Case adUnsignedBigInt            ' 21
        TipoField = "Big Integer"
    Case adUnsignedInt               ' 19
        TipoField = "Integer"
    Case adUnsignedSmallInt          ' 18
        TipoField = "Small Integer"
    Case adUnsignedTinyInt           ' 17
        TipoField = "Timy Integer"
    Case adUserDefined               ' 132
        TipoField = "UserDefined"
    Case adVarNumeric                 ' 139
        TipoField = "Long"
    Case adVarBinary                 ' 204
        TipoField = "Binary"
    Case adVarChar                   ' 200
        TipoField = "Text"
    Case adVariant                   ' 12
        TipoField = "Variant"
    Case adVarWChar                  ' 202
        TipoField = "Text"
    Case adWChar                     ' 130
        TipoField = "Text"
    Case Else
        TipoField = "Unknown"
    End Select
End Function
