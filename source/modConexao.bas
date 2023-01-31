Attribute VB_Name = "modConexao"
Option Explicit

Public cnnNovo As New ADODB.Connection
Public catNovo As New ADOX.Catalog

Public cnnAntigo As New ADODB.Connection
Public catAntigo As New ADOX.Catalog

Public Function AbreMDB(ByRef conn As ADODB.Connection, ByRef cat As ADOX.Catalog, ByVal strConn As String) As Boolean
    On Error GoTo DeuErro
    Dim i As Integer
    
    If conn.State <> ObjectStateEnum.adStateClosed Then conn.Close

    conn.ConnectionString = strConn
    conn.CursorLocation = adUseClient
    conn.Open
    Set cat.ActiveConnection = conn
    AbreMDB = True
    Exit Function
    
DeuErro:
    If Err.Number = -2147217843 Then
        i = i + 1
        If i < 3 Then
            conn.Properties("Jet OLEDB:Database Password").Value = frmLogin.Login
            Resume
        End If
    End If

    Err.Raise 8001, "AbreMDB", Err.Description
End Function

Public Function FechaMDB()
    On Error GoTo ErroTrata
    Set catNovo = Nothing
    If cnnNovo.State <> ObjectStateEnum.adStateClosed Then cnnNovo.Close
    Set cnnNovo = Nothing
    If cnnAntigo.State <> ObjectStateEnum.adStateClosed Then cnnAntigo.Close
    Set cnnAntigo = Nothing
    Exit Function

ErroTrata:
    MsgBox Err.Description
    Resume Next
End Function
