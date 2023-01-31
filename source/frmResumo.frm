VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form frmResumo 
   Caption         =   "ComparaMDB - Resultado da Análise"
   ClientHeight    =   7905
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11895
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
   ScaleHeight     =   7905
   ScaleWidth      =   11895
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid grdDiferencas 
      Height          =   9735
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   14910
      _ExtentX        =   26300
      _ExtentY        =   17171
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      BackColor       =   12648384
      ForeColor       =   16711680
      BackColorBkg    =   12648384
      AllowBigSelection=   0   'False
      SelectionMode   =   1
      Appearance      =   0
      FormatString    =   $"frmResumo.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmResumo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub grdDiferencas_Estruturar()
    With grdDiferencas
        .Clear
        .Cols = 2
        .Rows = 1
        .ColAlignment(0) = 1: .ColWidth(0) = 4000: .TextArray(0) = "Objeto"
        .ColAlignment(1) = 1: .ColWidth(1) = 10000: .TextArray(1) = "Descrição"
    End With
End Sub


Private Sub Form_Load()
    Call grdDiferencas_Estruturar
End Sub
