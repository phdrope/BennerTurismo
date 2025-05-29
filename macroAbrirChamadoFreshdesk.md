Public Sub AgenteAmadeus()
    Dim StrTexto As String
    Dim CurValidar As Currency
    CurValidar = 0

    Dim agentevendas As BPesquisa
    Set agentevendas = NewQuery
    agentevendas.Add "SELECT (SELECT NOMEFANTASIA FROM EMPRESAS WHERE HANDLE = EMPRESA) AS EMPRESA, " & _
                     "CASE TIPOERRO " & _
                     "WHEN '1' THEN 'Cliente' WHEN '2' THEN 'Fornecedor' WHEN '3' THEN 'Agente' " & _
                     "WHEN '4' THEN 'Inf Adicionais' WHEN '5' THEN 'Outros' WHEN '6' THEN 'Forma Pag/Rec' " & _
                     "WHEN '7' THEN 'Cancelamento' ELSE 'Não informado' END AS TIPO_ERRO, " & _
                     "CONVERT(VARCHAR, DATAENVIO, 103) AS DT_ENVIO, RLOCCIA AS OS, MENSAGEM AS MENSAGEM, '' AS AST " & _
                     "FROM BB_LOGINTEGRACOES WHERE SITUACAO = 2 AND TIPORESERVA = 5 AND TIPOERRO = 3 " & _
                     "AND DATAENVIO >= CONVERT(DATETIME, DATEADD(DAY, -3, GETDATE()))"
    agentevendas.Active = True

    Do While Not agentevendas.EOF
        CurValidar = 1
        StrTexto = StrTexto & " - DT_ENVIO: " & agentevendas.FieldByName("DT_ENVIO").AsString
        StrTexto = StrTexto & " - EMPRESA: " & agentevendas.FieldByName("EMPRESA").AsString
        StrTexto = StrTexto & " - PNR: " & agentevendas.FieldByName("OS").AsString & " <br>"
        StrTexto = StrTexto & "MENSAGEM: " & agentevendas.FieldByName("MENSAGEM").AsString & " <br>"
        StrTexto = StrTexto & "*: " & agentevendas.FieldByName("AST").AsString & " <br>"
        agentevendas.Next
    Loop

    agentevendas.Active = False
    Set agentevendas = Nothing

    If CurValidar > 0 Then
        Dim m As Mail
        Set m = NewMail
        m.Clear
        m.From = "sistemas@voetur.com.br"
        m.SendTo = "sistemas@voeturturismo.com.br"
        m.Subject = "[INTEGRATUR] AMADEUS - LOG DE AGENTE DE VENDAS"
        m.Priority = 0
        m.IsHtml = True
        m.Text.Add StrTexto
        m.Send
        Set m = Nothing
    End If
End Sub

Public Sub FornecedorAmadeus()
    Dim StrTexto As String
    Dim CurValidar As Currency
    CurValidar = 0

    Dim errofornecedor As BPesquisa
    Set errofornecedor = NewQuery
    errofornecedor.Add "SELECT (SELECT NOMEFANTASIA FROM EMPRESAS WHERE HANDLE = EMPRESA) AS EMPRESA, " & _
                       "CASE TIPOERRO " & _
                       "WHEN '1' THEN 'Cliente' WHEN '2' THEN 'Fornecedor' WHEN '3' THEN 'Agente' " & _
                       "WHEN '4' THEN 'Inf Adicionais' WHEN '5' THEN 'Outros' WHEN '6' THEN 'Forma Pag/Rec' " & _
                       "WHEN '7' THEN 'Cancelamento' ELSE 'Não informado' END AS TIPO_ERRO, " & _
                       "CONVERT(VARCHAR, DATAENVIO, 103) AS DT_ENVIO, RLOCCIA AS OS, MENSAGEM AS MENSAGEM, '' AS AST " & _
                       "FROM BB_LOGINTEGRACOES WHERE SITUACAO = 2 AND TIPORESERVA = 5 AND TIPOERRO = 2 " & _
                       "AND DATAENVIO >= CONVERT(DATETIME, DATEADD(DAY, -3, GETDATE()))"
    errofornecedor.Active = True

    Do While Not errofornecedor.EOF
        CurValidar = 1
        StrTexto = StrTexto & " - DT_ENVIO: " & errofornecedor.FieldByName("DT_ENVIO").AsString
        StrTexto = StrTexto & " - EMPRESA: " & errofornecedor.FieldByName("EMPRESA").AsString
        StrTexto = StrTexto & " - PNR: " & errofornecedor.FieldByName("OS").AsString & " <br>"
        StrTexto = StrTexto & "MENSAGEM: " & errofornecedor.FieldByName("MENSAGEM").AsString & " <br>"
        StrTexto = StrTexto & "*: " & errofornecedor.FieldByName("AST").AsString & " <br>"
        errofornecedor.Next
    Loop

    errofornecedor.Active = False
    Set errofornecedor = Nothing

    If CurValidar > 0 Then
        Dim m As Mail
        Set m = NewMail
        m.Clear
        m.From = "sistemas@voetur.com.br"
        m.SendTo = "sistemas@voeturturismo.com.br"
        m.Subject = "[INTEGRATUR] AMADEUS - LOG DE FORNECEDOR"
        m.Priority = 0
        m.IsHtml = True
        m.Text.Add StrTexto
        m.Send
        Set m = Nothing
    End If
End Sub

Public Sub ErrosAmadeus()
    Dim StrTexto As String
    Dim CurValidar As Currency
    CurValidar = 0

    Dim errosintegracao As BPesquisa
    Set errosintegracao = NewQuery
    errosintegracao.Add "SELECT (SELECT NOMEFANTASIA FROM EMPRESAS WHERE HANDLE = EMPRESA) AS EMPRESA, " & _
                        "CASE TIPOERRO " & _
                        "WHEN '1' THEN 'Cliente' WHEN '2' THEN 'Fornecedor' WHEN '3' THEN 'Agente' " & _
                        "WHEN '4' THEN 'Inf Adicionais' WHEN '5' THEN 'Outros' WHEN '6' THEN 'Forma Pag/Rec' " & _
                        "WHEN '7' THEN 'Cancelamento' ELSE 'Não informado' END AS TIPO_ERRO, " & _
                        "CONVERT(VARCHAR, DATAENVIO, 103) AS DT_ENVIO, RLOCCIA AS OS, MENSAGEM AS MENSAGEM, '' AS AST " & _
                        "FROM BB_LOGINTEGRACOES WHERE SITUACAO = 2 AND TIPORESERVA = 5 AND TIPOERRO NOT IN (1, 2, 3) " & _
                        "AND DATAENVIO >= CONVERT(DATETIME, DATEADD(DAY, -3, GETDATE()))"
    errosintegracao.Active = True

    Do While Not errosintegracao.EOF
        CurValidar = 1
        StrTexto = StrTexto & " - DT_ENVIO: " & errosintegracao.FieldByName("DT_ENVIO").AsString
        StrTexto = StrTexto & " - EMPRESA: " & errosintegracao.FieldByName("EMPRESA").AsString
        StrTexto = StrTexto & " - PNR: " & errosintegracao.FieldByName("OS").AsString & " <br>"
        StrTexto = StrTexto & "MENSAGEM: " & errosintegracao.FieldByName("MENSAGEM").AsString & " <br>"
        StrTexto = StrTexto & "*: " & errosintegracao.FieldByName("AST").AsString & " <br>"
        errosintegracao.Next
    Loop

    errosintegracao.Active = False
    Set errosintegracao = Nothing

    If CurValidar > 0 Then
        Dim m As Mail
        Set m = NewMail
        m.Clear
        m.From = "sistemas@voetur.com.br"
        m.SendTo = "sistemas@voeturturismo.com.br"
        m.Subject = "[INTEGRATUR] AMADEUS - LOG DE ERROS DIVERSOS"
        m.Priority = 0
        m.IsHtml = True
        m.Text.Add StrTexto
        m.Send
        Set m = Nothing
    End If
End Sub

Public Sub AgenteArgo()
    Dim StrTexto As String
    Dim CurValidar As Currency
    CurValidar = 0

    Dim agentevendas As BPesquisa
    Set agentevendas = NewQuery
    agentevendas.Add "SELECT (SELECT NOMEFANTASIA FROM EMPRESAS WHERE HANDLE = EMPRESA) AS EMPRESA, " & _
                     "CASE TIPOERRO " & _
                     "WHEN '1' THEN 'Cliente' WHEN '2' THEN 'Fornecedor' WHEN '3' THEN 'Agente' " & _
                     "WHEN '4' THEN 'Inf Adicionais' WHEN '5' THEN 'Outros' WHEN '6' THEN 'Forma Pag/Rec' " & _
                     "WHEN '7' THEN 'Cancelamento' ELSE 'Não informado' END AS TIPO_ERRO, " & _
                     "CONVERT(VARCHAR, DATAENVIO, 103) AS DT_ENVIO, RLOCCIA AS OS, MENSAGEM AS MENSAGEM, '' AS AST " & _
                     "FROM BB_LOGINTEGRACOES WHERE SITUACAO = 2 AND TIPORESERVA = 11 AND TIPOERRO = 3 " & _
                     "AND DATAENVIO >= CONVERT(DATETIME, DATEADD(DAY, -3, GETDATE()))"
    agentevendas.Active = True

    Do While Not agentevendas.EOF
        CurValidar = 1
        StrTexto = StrTexto & " - DT_ENVIO: " & agentevendas.FieldByName("DT_ENVIO").AsString
        StrTexto = StrTexto & " - EMPRESA: " & agentevendas.FieldByName("EMPRESA").AsString
        StrTexto = StrTexto & " - PNR: " & agentevendas.FieldByName("OS").AsString & " <br>"
        StrTexto = StrTexto & "MENSAGEM: " & agentevendas.FieldByName("MENSAGEM").AsString & " <br>"
        StrTexto = StrTexto & "*: " & agentevendas.FieldByName("AST").AsString & " <br>"
        agentevendas.Next
    Loop

    agentevendas.Active = False
    Set agentevendas = Nothing

    If CurValidar > 0 Then
        Dim m As Mail
        Set m = NewMail
        m.Clear
        m.From = "sistemas@voetur.com.br"
        m.SendTo = "sistemas@voeturturismo.com.br"
        m.Subject = "[INTEGRATUR] ARGO - LOG DE AGENTE DE VENDAS"
        m.Priority = 0
        m.IsHtml = True
        m.Text.Add StrTexto
        m.Send
        Set m = Nothing
    End If
End Sub

Public Sub FornecedorArgo()
    Dim StrTexto As String
    Dim CurValidar As Currency
    CurValidar = 0

    Dim errofornecedor As BPesquisa
    Set errofornecedor = NewQuery
    errofornecedor.Add "SELECT (SELECT NOMEFANTASIA FROM EMPRESAS WHERE HANDLE = EMPRESA) AS EMPRESA, " & _
                       "CASE TIPOERRO " & _
                       "WHEN '1' THEN 'Cliente' WHEN '2' THEN 'Fornecedor' WHEN '3' THEN 'Agente' " & _
                       "WHEN '4' THEN 'Inf Adicionais' WHEN '5' THEN 'Outros' WHEN '6' THEN 'Forma Pag/Rec' " & _
                       "WHEN '7' THEN 'Cancelamento' ELSE 'Não informado' END AS TIPO_ERRO, " & _
                       "CONVERT(VARCHAR, DATAENVIO, 103) AS DT_ENVIO, RLOCCIA AS OS, MENSAGEM AS MENSAGEM, '' AS AST " & _
                       "FROM BB_LOGINTEGRACOES WHERE SITUACAO = 2 AND TIPORESERVA = 11 AND TIPOERRO = 2 " & _
                       "AND DATAENVIO >= CONVERT(DATETIME, DATEADD(DAY, -3, GETDATE()))"
    errofornecedor.Active = True

    Do While Not errofornecedor.EOF
        CurValidar = 1
        StrTexto = StrTexto & " - DT_ENVIO: " & errofornecedor.FieldByName("DT_ENVIO").AsString
        StrTexto = StrTexto & " - EMPRESA: " & errofornecedor.FieldByName("EMPRESA").AsString
        StrTexto = StrTexto & " - PNR: " & errofornecedor.FieldByName("OS").AsString & " <br>"
        StrTexto = StrTexto & "MENSAGEM: " & errofornecedor.FieldByName("MENSAGEM").AsString & " <br>"
        StrTexto = StrTexto & "*: " & errofornecedor.FieldByName("AST").AsString & " <br>"
        errofornecedor.Next
    Loop

    errofornecedor.Active = False
    Set errofornecedor = Nothing

    If CurValidar > 0 Then
        Dim m As Mail
        Set m = NewMail
        m.Clear
        m.From = "sistemas@voetur.com.br"
        m.SendTo = "sistemas@voeturturismo.com.br"
        m.Subject = "[INTEGRATUR] ARGO - LOG DE FORNECEDOR"
        m.Priority = 0
        m.IsHtml = True
        m.Text.Add StrTexto
        m.Send
        Set m = Nothing
    End If
End Sub

Public Sub ErrosArgo()
    Dim StrTexto As String
    Dim CurValidar As Currency
    CurValidar = 0

    Dim errosintegracao As BPesquisa
    Set errosintegracao = NewQuery
    errosintegracao.Add "SELECT (SELECT NOMEFANTASIA FROM EMPRESAS WHERE HANDLE = EMPRESA) AS EMPRESA, " & _
                        "CASE TIPOERRO " & _
                        "WHEN '1' THEN 'Cliente' WHEN '2' THEN 'Fornecedor' WHEN '3' THEN 'Agente' " & _
                        "WHEN '4' THEN 'Inf Adicionais' WHEN '5' THEN 'Outros' WHEN '6' THEN 'Forma Pag/Rec' " & _
                        "WHEN '7' THEN 'Cancelamento' ELSE 'Não informado' END AS TIPO_ERRO, " & _
                        "CONVERT(VARCHAR, DATAENVIO, 103) AS DT_ENVIO, RLOCCIA AS OS, MENSAGEM AS MENSAGEM, '' AS AST " & _
                        "FROM BB_LOGINTEGRACOES WHERE SITUACAO = 2 AND TIPORESERVA = 11 AND TIPOERRO NOT IN (1, 2, 3) " & _
                        "AND DATAENVIO >= CONVERT(DATETIME, DATEADD(DAY, -3, GETDATE()))"
    errosintegracao.Active = True

    Do While Not errosintegracao.EOF
        CurValidar = 1
        StrTexto = StrTexto & " - DT_ENVIO: " & errosintegracao.FieldByName("DT_ENVIO").AsString
        StrTexto = StrTexto & " - EMPRESA: " & errosintegracao.FieldByName("EMPRESA").AsString
        StrTexto = StrTexto & " - PNR: " & errosintegracao.FieldByName("OS").AsString & " <br>"
        StrTexto = StrTexto & "MENSAGEM: " & errosintegracao.FieldByName("MENSAGEM").AsString & " <br>"
        StrTexto = StrTexto & "*: " & errosintegracao.FieldByName("AST").AsString & " <br>"
        errosintegracao.Next
    Loop

    errosintegracao.Active = False
    Set errosintegracao = Nothing

    If CurValidar > 0 Then
        Dim m As Mail
        Set m = NewMail
        m.Clear
        m.From = "sistemas@voetur.com.br"
        m.SendTo = "sistemas@voeturturismo.com.br"
        m.Subject = "[INTEGRATUR] ARGO - LOG DE ERROS DIVERSOS"
        m.Priority = 0
        m.IsHtml = True
        m.Text.Add StrTexto
        m.Send
        Set m = Nothing
    End If
End Sub

Public Sub AgenteB2b()
    Dim StrTexto As String
    Dim CurValidar As Currency
    CurValidar = 0

    Dim agentevendas As BPesquisa
    Set agentevendas = NewQuery
    agentevendas.Add "SELECT (SELECT NOMEFANTASIA FROM EMPRESAS WHERE HANDLE = EMPRESA) AS EMPRESA, " & _
                     "CASE TIPOERRO " & _
                     "WHEN '1' THEN 'Cliente' WHEN '2' THEN 'Fornecedor' WHEN '3' THEN 'Agente' " & _
                     "WHEN '4' THEN 'Inf Adicionais' WHEN '5' THEN 'Outros' WHEN '6' THEN 'Forma Pag/Rec' " & _
                     "WHEN '7' THEN 'Cancelamento' ELSE 'Não informado' END AS TIPO_ERRO, " & _
                     "CONVERT(VARCHAR, DATAENVIO, 103) AS DT_ENVIO, RLOCCIA AS OS, MENSAGEM AS MENSAGEM, '' AS AST " & _
                     "FROM BB_LOGINTEGRACOES WHERE SITUACAO = 2 AND TIPORESERVA = 16 AND TIPOERRO = 3 " & _
                     "AND DATAENVIO >= CONVERT(DATETIME, DATEADD(DAY, -3, GETDATE()))"
    agentevendas.Active = True

    Do While Not agentevendas.EOF
        CurValidar = 1
        StrTexto = StrTexto & " - DT_ENVIO: " & agentevendas.FieldByName("DT_ENVIO").AsString
        StrTexto = StrTexto & " - EMPRESA: " & agentevendas.FieldByName("EMPRESA").AsString
        StrTexto = StrTexto & " - PNR: " & agentevendas.FieldByName("OS").AsString & " <br>"
        StrTexto = StrTexto & "MENSAGEM: " & agentevendas.FieldByName("MENSAGEM").AsString & " <br>"
        StrTexto = StrTexto & "*: " & agentevendas.FieldByName("AST").AsString & " <br>"
        agentevendas.Next
    Loop

    agentevendas.Active = False
    Set agentevendas = Nothing

    If CurValidar > 0 Then
        Dim m As Mail
        Set m = NewMail
        m.Clear
        m.From = "sistemas@voetur.com.br"
        m.SendTo = "sistemas@voeturturismo.com.br"
        m.Subject = "[INTEGRATUR] B2B - LOG DE AGENTE DE VENDAS"
        m.Priority = 0
        m.IsHtml = True
        m.Text.Add StrTexto
        m.Send
        Set m = Nothing
    End If
End Sub

Public Sub FornecedorB2b()
    Dim StrTexto As String
    Dim CurValidar As Currency
    CurValidar = 0

    Dim errofornecedor As BPesquisa
    Set errofornecedor = NewQuery
    errofornecedor.Add "SELECT (SELECT NOMEFANTASIA FROM EMPRESAS WHERE HANDLE = EMPRESA) AS EMPRESA, " & _
                       "CASE TIPOERRO " & _
                       "WHEN '1' THEN 'Cliente' WHEN '2' THEN 'Fornecedor' WHEN '3' THEN 'Agente' " & _
                       "WHEN '4' THEN 'Inf Adicionais' WHEN '5' THEN 'Outros' WHEN '6' THEN 'Forma Pag/Rec' " & _
                       "WHEN '7' THEN 'Cancelamento' ELSE 'Não informado' END AS TIPO_ERRO, " & _
                       "CONVERT(VARCHAR, DATAENVIO, 103) AS DT_ENVIO, RLOCCIA AS OS, MENSAGEM AS MENSAGEM, '' AS AST " & _
                       "FROM BB_LOGINTEGRACOES WHERE SITUACAO = 2 AND TIPORESERVA = 16 AND TIPOERRO = 2 " & _
                       "AND DATAENVIO >= CONVERT(DATETIME, DATEADD(DAY, -3, GETDATE()))"
    errofornecedor.Active = True

    Do While Not errofornecedor.EOF
        CurValidar = 1
        StrTexto = StrTexto & " - DT_ENVIO: " & errofornecedor.FieldByName("DT_ENVIO").AsString
        StrTexto = StrTexto & " - EMPRESA: " & errofornecedor.FieldByName("EMPRESA").AsString
        StrTexto = StrTexto & " - PNR: " & errofornecedor.FieldByName("OS").AsString & " <br>"
        StrTexto = StrTexto & "MENSAGEM: " & errofornecedor.FieldByName("MENSAGEM").AsString & " <br>"
        StrTexto = StrTexto & "*: " & errofornecedor.FieldByName("AST").AsString & " <br>"
        errofornecedor.Next
    Loop

    errofornecedor.Active = False
    Set errofornecedor = Nothing

    If CurValidar > 0 Then
        Dim m As Mail
        Set m = NewMail
        m.Clear
        m.From = "sistemas@voetur.com.br"
        m.SendTo = "sistemas@voeturturismo.com.br"
        m.Subject = "[INTEGRATUR] B2B - LOG DE FORNECEDOR"
        m.Priority = 0
        m.IsHtml = True
        m.Text.Add StrTexto
        m.Send
        Set m = Nothing
    End If
End Sub

Public Sub ErrosB2b()
    Dim StrTexto As String
    Dim CurValidar As Currency
    CurValidar = 0

    Dim errosintegracao As BPesquisa
    Set errosintegracao = NewQuery
    errosintegracao.Add "SELECT (SELECT NOMEFANTASIA FROM EMPRESAS WHERE HANDLE = EMPRESA) AS EMPRESA, " & _
                        "CASE TIPOERRO " & _
                        "WHEN '1' THEN 'Cliente' WHEN '2' THEN 'Fornecedor' WHEN '3' THEN 'Agente' " & _
                        "WHEN '4' THEN 'Inf Adicionais' WHEN '5' THEN 'Outros' WHEN '6' THEN 'Forma Pag/Rec' " & _
                        "WHEN '7' THEN 'Cancelamento' ELSE 'Não informado' END AS TIPO_ERRO, " & _
                        "CONVERT(VARCHAR, DATAENVIO, 103) AS DT_ENVIO, RLOCCIA AS OS, MENSAGEM AS MENSAGEM, '' AS AST " & _
                        "FROM BB_LOGINTEGRACOES WHERE SITUACAO = 2 AND TIPORESERVA = 16 AND TIPOERRO NOT IN (1, 2, 3) " & _
                        "AND DATAENVIO >= CONVERT(DATETIME, DATEADD(DAY, -3, GETDATE()))"
    errosintegracao.Active = True

    Do While Not errosintegracao.EOF
        CurValidar = 1
        StrTexto = StrTexto & " - DT_ENVIO: " & errosintegracao.FieldByName("DT_ENVIO").AsString
        StrTexto = StrTexto & " - EMPRESA: " & errosintegracao.FieldByName("EMPRESA").AsString
        StrTexto = StrTexto & " - PNR: " & errosintegracao.FieldByName("OS").AsString & " <br>"
        StrTexto = StrTexto & "MENSAGEM: " & errosintegracao.FieldByName("MENSAGEM").AsString & " <br>"
        StrTexto = StrTexto & "*: " & errosintegracao.FieldByName("AST").AsString & " <br>"
        errosintegracao.Next
    Loop

    errosintegracao.Active = False
    Set errosintegracao = Nothing

    If CurValidar > 0 Then
        Dim m As Mail
        Set m = NewMail
        m.Clear
        m.From = "sistemas@voetur.com.br"
        m.SendTo = "sistemas@voeturturismo.com.br"
        m.Subject = "[INTEGRATUR] B2B - LOG DE ERROS DIVERSOS"
        m.Priority = 0
        m.IsHtml = True
        m.Text.Add StrTexto
        m.Send
        Set m = Nothing
    End If
End Sub

Public Sub AgenteBuscaIdeal()
    Dim StrTexto As String
    Dim CurValidar As Currency
    CurValidar = 0

    Dim agentevendas As BPesquisa
    Set agentevendas = NewQuery
    agentevendas.Add "SELECT (SELECT NOMEFANTASIA FROM EMPRESAS WHERE HANDLE = EMPRESA) AS EMPRESA, " & _
                     "CASE TIPOERRO " & _
                     "WHEN '1' THEN 'Cliente' WHEN '2' THEN 'Fornecedor' WHEN '3' THEN 'Agente' " & _
                     "WHEN '4' THEN 'Inf Adicionais' WHEN '5' THEN 'Outros' WHEN '6' THEN 'Forma Pag/Rec' " & _
                     "WHEN '7' THEN 'Cancelamento' ELSE 'Não informado' END AS TIPO_ERRO, " & _
                     "CONVERT(VARCHAR, DATAENVIO, 103) AS DT_ENVIO, RLOCCIA AS OS, MENSAGEM AS MENSAGEM, '' AS AST " & _
                     "FROM BB_LOGINTEGRACOES WHERE SITUACAO = 2 AND TIPORESERVA = 28 AND TIPOERRO = 3 " & _
                     "AND DATAENVIO >= CONVERT(DATETIME, DATEADD(DAY, -3, GETDATE()))"
    agentevendas.Active = True

    Do While Not agentevendas.EOF
        CurValidar = 1
        StrTexto = StrTexto & " - DT_ENVIO: " & agentevendas.FieldByName("DT_ENVIO").AsString
        StrTexto = StrTexto & " - EMPRESA: " & agentevendas.FieldByName("EMPRESA").AsString
        StrTexto = StrTexto & " - PNR: " & agentevendas.FieldByName("OS").AsString & " <br>"
        StrTexto = StrTexto & "MENSAGEM: " & agentevendas.FieldByName("MENSAGEM").AsString & " <br>"
        StrTexto = StrTexto & "*: " & agentevendas.FieldByName("AST").AsString & " <br>"
        agentevendas.Next
    Loop

    agentevendas.Active = False
    Set agentevendas = Nothing

    If CurValidar > 0 Then
        Dim m As Mail
        Set m = NewMail
        m.Clear
        m.From = "sistemas@voetur.com.br"
        m.SendTo = "sistemas@voeturturismo.com.br"
        m.Subject = "[INTEGRATUR] BUSCA IDEAL - LOG DE AGENTE DE VENDAS"
        m.Priority = 0
        m.IsHtml = True
        m.Text.Add StrTexto
        m.Send
        Set m = Nothing
    End If
End Sub

Public Sub FornecedorBuscaIdeal()
    Dim StrTexto As String
    Dim CurValidar As Currency
    CurValidar = 0

    Dim errofornecedor As BPesquisa
    Set errofornecedor = NewQuery
    errofornecedor.Add "SELECT (SELECT NOMEFANTASIA FROM EMPRESAS WHERE HANDLE = EMPRESA) AS EMPRESA, " & _
                       "CASE TIPOERRO " & _
                       "WHEN '1' THEN 'Cliente' WHEN '2' THEN 'Fornecedor' WHEN '3' THEN 'Agente' " & _
                       "WHEN '4' THEN 'Inf Adicionais' WHEN '5' THEN 'Outros' WHEN '6' THEN 'Forma Pag/Rec' " & _
                       "WHEN '7' THEN 'Cancelamento' ELSE 'Não informado' END AS TIPO_ERRO, " & _
                       "CONVERT(VARCHAR, DATAENVIO, 103) AS DT_ENVIO, RLOCCIA AS OS, MENSAGEM AS MENSAGEM, '' AS AST " & _
                       "FROM BB_LOGINTEGRACOES WHERE SITUACAO = 2 AND TIPORESERVA = 28 AND TIPOERRO = 2 " & _
                       "AND DATAENVIO >= CONVERT(DATETIME, DATEADD(DAY, -3, GETDATE()))"
    errofornecedor.Active = True

    Do While Not errofornecedor.EOF
        CurValidar = 1
        StrTexto = StrTexto & " - DT_ENVIO: " & errofornecedor.FieldByName("DT_ENVIO").AsString
        StrTexto = StrTexto & " - EMPRESA: " & errofornecedor.FieldByName("EMPRESA").AsString
        StrTexto = StrTexto & " - PNR: " & errofornecedor.FieldByName("OS").AsString & " <br>"
        StrTexto = StrTexto & "MENSAGEM: " & errofornecedor.FieldByName("MENSAGEM").AsString & " <br>"
        StrTexto = StrTexto & "*: " & errofornecedor.FieldByName("AST").AsString & " <br>"
        errofornecedor.Next
    Loop

    errofornecedor.Active = False
    Set errofornecedor = Nothing

    If CurValidar > 0 Then
        Dim m As Mail
        Set m = NewMail
        m.Clear
        m.From = "sistemas@voetur.com.br"
        m.SendTo = "sistemas@voeturturismo.com.br"
        m.Subject = "[INTEGRATUR] BUSCA IDEAL - LOG DE FORNECEDOR"
        m.Priority = 0
        m.IsHtml = True
        m.Text.Add StrTexto
        m.Send
        Set m = Nothing
    End If
End Sub

Public Sub ErrosBuscaIdeal()
    Dim StrTexto As String
    Dim CurValidar As Currency
    CurValidar = 0

    Dim errosintegracao As BPesquisa
    Set errosintegracao = NewQuery
    errosintegracao.Add "SELECT (SELECT NOMEFANTASIA FROM EMPRESAS WHERE HANDLE = EMPRESA) AS EMPRESA, " & _
                        "CASE TIPOERRO " & _
                        "WHEN '1' THEN 'Cliente' WHEN '2' THEN 'Fornecedor' WHEN '3' THEN 'Agente' " & _
                        "WHEN '4' THEN 'Inf Adicionais' WHEN '5' THEN 'Outros' WHEN '6' THEN 'Forma Pag/Rec' " & _
                        "WHEN '7' THEN 'Cancelamento' ELSE 'Não informado' END AS TIPO_ERRO, " & _
                        "CONVERT(VARCHAR, DATAENVIO, 103) AS DT_ENVIO, RLOCCIA AS OS, MENSAGEM AS MENSAGEM, '' AS AST " & _
                        "FROM BB_LOGINTEGRACOES WHERE SITUACAO = 2 AND TIPORESERVA = 28 AND TIPOERRO NOT IN (1, 2, 3) " & _
                        "AND DATAENVIO >= CONVERT(DATETIME, DATEADD(DAY, -3, GETDATE()))"
    errosintegracao.Active = True

    Do While Not errosintegracao.EOF
        CurValidar = 1
        StrTexto = StrTexto & " - DT_ENVIO: " & errosintegracao.FieldByName("DT_ENVIO").AsString
        StrTexto = StrTexto & " - EMPRESA: " & errosintegracao.FieldByName("EMPRESA").AsString
        StrTexto = StrTexto & " - PNR: " & errosintegracao.FieldByName("OS").AsString & " <br>"
        StrTexto = StrTexto & "MENSAGEM: " & errosintegracao.FieldByName("MENSAGEM").AsString & " <br>"
        StrTexto = StrTexto & "*: " & errosintegracao.FieldByName("AST").AsString & " <br>"
        errosintegracao.Next
    Loop

    errosintegracao.Active = False
    Set errosintegracao = Nothing

    If CurValidar > 0 Then
        Dim m As Mail
        Set m = NewMail
        m.Clear
        m.From = "sistemas@voetur.com.br"
        m.SendTo = "sistemas@voeturturismo.com.br"
        m.Subject = "[INTEGRATUR] BUSCA IDEAL - LOG DE ERROS DIVERSOS"
        m.Priority = 0
        m.IsHtml = True
        m.Text.Add StrTexto
        m.Send
        Set m = Nothing
    End If
End Sub

Public Sub AgenteEnvision()
    Dim StrTexto As String
    Dim CurValidar As Currency
    CurValidar = 0

    Dim agentevendas As BPesquisa
    Set agentevendas = NewQuery
    agentevendas.Add "SELECT (SELECT NOMEFANTASIA FROM EMPRESAS WHERE HANDLE = EMPRESA) AS EMPRESA, " & _
                     "CASE TIPOERRO " & _
                     "WHEN '1' THEN 'Cliente' WHEN '2' THEN 'Fornecedor' WHEN '3' THEN 'Agente' " & _
                     "WHEN '4' THEN 'Inf Adicionais' WHEN '5' THEN 'Outros' WHEN '6' THEN 'Forma Pag/Rec' " & _
                     "WHEN '7' THEN 'Cancelamento' ELSE 'Não informado' END AS TIPO_ERRO, " & _
                     "CONVERT(VARCHAR, DATAENVIO, 103) AS DT_ENVIO, RLOCCIA AS OS, MENSAGEM AS MENSAGEM, '' AS AST " & _
                     "FROM BB_LOGINTEGRACOES WHERE SITUACAO = 2 AND TIPORESERVA = 20 AND TIPOERRO = 3 " & _
                     "AND DATAENVIO >= CONVERT(DATETIME, DATEADD(DAY, -3, GETDATE()))"
    agentevendas.Active = True

    Do While Not agentevendas.EOF
        CurValidar = 1
        StrTexto = StrTexto & " - DT_ENVIO: " & agentevendas.FieldByName("DT_ENVIO").AsString
        StrTexto = StrTexto & " - EMPRESA: " & agentevendas.FieldByName("EMPRESA").AsString
        StrTexto = StrTexto & " - PNR: " & agentevendas.FieldByName("OS").AsString & " <br>"
        StrTexto = StrTexto & "MENSAGEM: " & agentevendas.FieldByName("MENSAGEM").AsString & " <br>"
        StrTexto = StrTexto & "*: " & agentevendas.FieldByName("AST").AsString & " <br>"
        agentevendas.Next
    Loop

    agentevendas.Active = False
    Set agentevendas = Nothing

    If CurValidar > 0 Then
        Dim m As Mail
        Set m = NewMail
        m.Clear
        m.From = "sistemas@voetur.com.br"
        m.SendTo = "sistemas@voeturturismo.com.br"
        m.Subject = "[INTEGRATUR] ENVISION - LOG DE AGENTE DE VENDAS"
        m.Priority = 0
        m.IsHtml = True
        m.Text.Add StrTexto
        m.Send
        Set m = Nothing
    End If
End Sub

Public Sub FornecedorEnvision()
    Dim StrTexto As String
    Dim CurValidar As Currency
    CurValidar = 0

    Dim errofornecedor As BPesquisa
    Set errofornecedor = NewQuery
    errofornecedor.Add "SELECT (SELECT NOMEFANTASIA FROM EMPRESAS WHERE HANDLE = EMPRESA) AS EMPRESA, " & _
                       "CASE TIPOERRO " & _
                       "WHEN '1' THEN 'Cliente' WHEN '2' THEN 'Fornecedor' WHEN '3' THEN 'Agente' " & _
                       "WHEN '4' THEN 'Inf Adicionais' WHEN '5' THEN 'Outros' WHEN '6' THEN 'Forma Pag/Rec' " & _
                       "WHEN '7' THEN 'Cancelamento' ELSE 'Não informado' END AS TIPO_ERRO, " & _
                       "CONVERT(VARCHAR, DATAENVIO, 103) AS DT_ENVIO, RLOCCIA AS OS, MENSAGEM AS MENSAGEM, '' AS AST " & _
                       "FROM BB_LOGINTEGRACOES WHERE SITUACAO = 2 AND TIPORESERVA = 20 AND TIPOERRO = 2 " & _
                       "AND DATAENVIO >= CONVERT(DATETIME, DATEADD(DAY, -3, GETDATE()))"
    errofornecedor.Active = True

    Do While Not errofornecedor.EOF
        CurValidar = 1
        StrTexto = StrTexto & " - DT_ENVIO: " & errofornecedor.FieldByName("DT_ENVIO").AsString
        StrTexto = StrTexto & " - EMPRESA: " & errofornecedor.FieldByName("EMPRESA").AsString
        StrTexto = StrTexto & " - PNR: " & errofornecedor.FieldByName("OS").AsString & " <br>"
        StrTexto = StrTexto & "MENSAGEM: " & errofornecedor.FieldByName("MENSAGEM").AsString & " <br>"
        StrTexto = StrTexto & "*: " & errofornecedor.FieldByName("AST").AsString & " <br>"
        errofornecedor.Next
    Loop

    errofornecedor.Active = False
    Set errofornecedor = Nothing

    If CurValidar > 0 Then
        Dim m As Mail
        Set m = NewMail
        m.Clear
        m.From = "sistemas@voetur.com.br"
        m.SendTo = "sistemas@voeturturismo.com.br"
        m.Subject = "[INTEGRATUR] ENVISION - LOG DE FORNECEDOR"
        m.Priority = 0
        m.IsHtml = True
        m.Text.Add StrTexto
        m.Send
        Set m = Nothing
    End If
End Sub

Public Sub ErrosEnvision()
    Dim StrTexto As String
    Dim CurValidar As Currency
    CurValidar = 0

    Dim errosintegracao As BPesquisa
    Set errosintegracao = NewQuery
    errosintegracao.Add "SELECT (SELECT NOMEFANTASIA FROM EMPRESAS WHERE HANDLE = EMPRESA) AS EMPRESA, " & _
                        "CASE TIPOERRO " & _
                        "WHEN '1' THEN 'Cliente' WHEN '2' THEN 'Fornecedor' WHEN '3' THEN 'Agente' " & _
                        "WHEN '4' THEN 'Inf Adicionais' WHEN '5' THEN 'Outros' WHEN '6' THEN 'Forma Pag/Rec' " & _
                        "WHEN '7' THEN 'Cancelamento' ELSE 'Não informado' END AS TIPO_ERRO, " & _
                        "CONVERT(VARCHAR, DATAENVIO, 103) AS DT_ENVIO, RLOCCIA AS OS, MENSAGEM AS MENSAGEM, '' AS AST " & _
                        "FROM BB_LOGINTEGRACOES WHERE SITUACAO = 2 AND TIPORESERVA = 20 AND TIPOERRO NOT IN (1, 2, 3) " & _
                        "AND DATAENVIO >= CONVERT(DATETIME, DATEADD(DAY, -3, GETDATE()))"
    errosintegracao.Active = True

    Do While Not errosintegracao.EOF
        CurValidar = 1
        StrTexto = StrTexto & " - DT_ENVIO: " & errosintegracao.FieldByName("DT_ENVIO").AsString
        StrTexto = StrTexto & " - EMPRESA: " & errosintegracao.FieldByName("EMPRESA").AsString
        StrTexto = StrTexto & " - PNR: " & errosintegracao.FieldByName("OS").AsString & " <br>"
        StrTexto = StrTexto & "MENSAGEM: " & errosintegracao.FieldByName("MENSAGEM").AsString & " <br>"
        StrTexto = StrTexto & "*: " & errosintegracao.FieldByName("AST").AsString & " <br>"
        errosintegracao.Next
    Loop

    errosintegracao.Active = False
    Set errosintegracao = Nothing

    If CurValidar > 0 Then
        Dim m As Mail
        Set m = NewMail
        m.Clear
        m.From = "sistemas@voetur.com.br"
        m.SendTo = "sistemas@voeturturismo.com.br"
        m.Subject = "[INTEGRATUR] ENVISION - LOG DE ERROS DIVERSOS"
        m.Priority = 0
        m.IsHtml = True
        m.Text.Add StrTexto
        m.Send
        Set m = Nothing
    End If
End Sub

Public Sub AgenteLemontech()
    Dim StrTexto As String
    Dim CurValidar As Currency
    CurValidar = 0

    Dim agentevendas As BPesquisa
    Set agentevendas = NewQuery
    agentevendas.Add "SELECT (SELECT NOMEFANTASIA FROM EMPRESAS WHERE HANDLE = EMPRESA) AS EMPRESA, " & _
                     "CASE TIPOERRO " & _
                     "WHEN '1' THEN 'Cliente' WHEN '2' THEN 'Fornecedor' WHEN '3' THEN 'Agente' " & _
                     "WHEN '4' THEN 'Inf Adicionais' WHEN '5' THEN 'Outros' WHEN '6' THEN 'Forma Pag/Rec' " & _
                     "WHEN '7' THEN 'Cancelamento' ELSE 'Não informado' END AS TIPO_ERRO, " & _
                     "CONVERT(VARCHAR, DATAENVIO, 103) AS DT_ENVIO, RLOCCIA AS OS, MENSAGEM AS MENSAGEM, '' AS AST " & _
                     "FROM BB_LOGINTEGRACOES WHERE SITUACAO = 2 AND TIPORESERVA = 21 AND TIPOERRO = 3 " & _
                     "AND DATAENVIO >= CONVERT(DATETIME, DATEADD(DAY, -3, GETDATE()))"
    agentevendas.Active = True

    Do While Not agentevendas.EOF
        CurValidar = 1
        StrTexto = StrTexto & " - DT_ENVIO: " & agentevendas.FieldByName("DT_ENVIO").AsString
        StrTexto = StrTexto & " - EMPRESA: " & agentevendas.FieldByName("EMPRESA").AsString
        StrTexto = StrTexto & " - PNR: " & agentevendas.FieldByName("OS").AsString & " <br>"
        StrTexto = StrTexto & "MENSAGEM: " & agentevendas.FieldByName("MENSAGEM").AsString & " <br>"
        StrTexto = StrTexto & "*: " & agentevendas.FieldByName("AST").AsString & " <br>"
        agentevendas.Next
    Loop

    agentevendas.Active = False
    Set agentevendas = Nothing

    If CurValidar > 0 Then
        Dim m As Mail
        Set m = NewMail
        m.Clear
        m.From = "sistemas@voetur.com.br"
        m.SendTo = "sistemas@voeturturismo.com.br"
        m.Subject = "[INTEGRATUR] LEMONTECH - LOG DE AGENTE DE VENDAS"
        m.Priority = 0
        m.IsHtml = True
        m.Text.Add StrTexto
        m.Send
        Set m = Nothing
    End If
End Sub

Public Sub FornecedorLemontech()
    Dim StrTexto As String
    Dim CurValidar As Currency
    CurValidar = 0

    Dim errofornecedor As BPesquisa
    Set errofornecedor = NewQuery
    errofornecedor.Add "SELECT (SELECT NOMEFANTASIA FROM EMPRESAS WHERE HANDLE = EMPRESA) AS EMPRESA, " & _
                       "CASE TIPOERRO " & _
                       "WHEN '1' THEN 'Cliente' WHEN '2' THEN 'Fornecedor' WHEN '3' THEN 'Agente' " & _
                       "WHEN '4' THEN 'Inf Adicionais' WHEN '5' THEN 'Outros' WHEN '6' THEN 'Forma Pag/Rec' " & _
                       "WHEN '7' THEN 'Cancelamento' ELSE 'Não informado' END AS TIPO_ERRO, " & _
                       "CONVERT(VARCHAR, DATAENVIO, 103) AS DT_ENVIO, RLOCCIA AS OS, MENSAGEM AS MENSAGEM, '' AS AST " & _
                       "FROM BB_LOGINTEGRACOES WHERE SITUACAO = 2 AND TIPORESERVA = 21 AND TIPOERRO = 2 " & _
                       "AND DATAENVIO >= CONVERT(DATETIME, DATEADD(DAY, -3, GETDATE()))"
    errofornecedor.Active = True

    Do While Not errofornecedor.EOF
        CurValidar = 1
        StrTexto = StrTexto & " - DT_ENVIO: " & errofornecedor.FieldByName("DT_ENVIO").AsString
        StrTexto = StrTexto & " - EMPRESA: " & errofornecedor.FieldByName("EMPRESA").AsString
        StrTexto = StrTexto & " - PNR: " & errofornecedor.FieldByName("OS").AsString & " <br>"
        StrTexto = StrTexto & "MENSAGEM: " & errofornecedor.FieldByName("MENSAGEM").AsString & " <br>"
        StrTexto = StrTexto & "*: " & errofornecedor.FieldByName("AST").AsString & " <br>"
        errofornecedor.Next
    Loop

    errofornecedor.Active = False
    Set errofornecedor = Nothing

    If CurValidar > 0 Then
        Dim m As Mail
        Set m = NewMail
        m.Clear
        m.From = "sistemas@voetur.com.br"
        m.SendTo = "sistemas@voeturturismo.com.br"
        m.Subject = "[INTEGRATUR] LEMONTECH - LOG DE FORNECEDOR"
        m.Priority = 0
        m.IsHtml = True
        m.Text.Add StrTexto
        m.Send
        Set m = Nothing
    End If
End Sub

Public Sub ErrosLemontech()
    Dim StrTexto As String
    Dim CurValidar As Currency
    CurValidar = 0

    Dim errosintegracao As BPesquisa
    Set errosintegracao = NewQuery
    errosintegracao.Add "SELECT (SELECT NOMEFANTASIA FROM EMPRESAS WHERE HANDLE = EMPRESA) AS EMPRESA, " & _
                        "CASE TIPOERRO " & _
                        "WHEN '1' THEN 'Cliente' WHEN '2' THEN 'Fornecedor' WHEN '3' THEN 'Agente' " & _
                        "WHEN '4' THEN 'Inf Adicionais' WHEN '5' THEN 'Outros' WHEN '6' THEN 'Forma Pag/Rec' " & _
                        "WHEN '7' THEN 'Cancelamento' ELSE 'Não informado' END AS TIPO_ERRO, " & _
                        "CONVERT(VARCHAR, DATAENVIO, 103) AS DT_ENVIO, RLOCCIA AS OS, MENSAGEM AS MENSAGEM, '' AS AST " & _
                        "FROM BB_LOGINTEGRACOES WHERE SITUACAO = 2 AND TIPORESERVA = 21 AND TIPOERRO NOT IN (1, 2, 3) " & _
                        "AND DATAENVIO >= CONVERT(DATETIME, DATEADD(DAY, -3, GETDATE()))"
    errosintegracao.Active = True

    Do While Not errosintegracao.EOF
        CurValidar = 1
        StrTexto = StrTexto & " - DT_ENVIO: " & errosintegracao.FieldByName("DT_ENVIO").AsString
        StrTexto = StrTexto & " - EMPRESA: " & errosintegracao.FieldByName("EMPRESA").AsString
        StrTexto = StrTexto & " - PNR: " & errosintegracao.FieldByName("OS").AsString & " <br>"
        StrTexto = StrTexto & "MENSAGEM: " & errosintegracao.FieldByName("MENSAGEM").AsString & " <br>"
        StrTexto = StrTexto & "*: " & errosintegracao.FieldByName("AST").AsString & " <br>"
        errosintegracao.Next
    Loop

    errosintegracao.Active = False
    Set errosintegracao = Nothing

    If CurValidar > 0 Then
        Dim m As Mail
        Set m = NewMail
        m.Clear
        m.From = "sistemas@voetur.com.br"
        m.SendTo = "sistemas@voeturturismo.com.br"
        m.Subject = "[INTEGRATUR] LEMONTECH - LOG DE ERROS DIVERSOS"
        m.Priority = 0
        m.IsHtml = True
        m.Text.Add StrTexto
        m.Send
        Set m = Nothing
    End If
End Sub

Public Sub AgenteOmnibees()
    Dim StrTexto As String
    Dim CurValidar As Currency
    CurValidar = 0

    Dim agentevendas As BPesquisa
    Set agentevendas = NewQuery
    agentevendas.Add "SELECT (SELECT NOMEFANTASIA FROM EMPRESAS WHERE HANDLE = EMPRESA) AS EMPRESA, " & _
                     "CASE TIPOERRO " & _
                     "WHEN '1' THEN 'Cliente' WHEN '2' THEN 'Fornecedor' WHEN '3' THEN 'Agente' " & _
                     "WHEN '4' THEN 'Inf Adicionais' WHEN '5' THEN 'Outros' WHEN '6' THEN 'Forma Pag/Rec' " & _
                     "WHEN '7' THEN 'Cancelamento' ELSE 'Não informado' END AS TIPO_ERRO, " & _
                     "CONVERT(VARCHAR, DATAENVIO, 103) AS DT_ENVIO, RLOCCIA AS OS, MENSAGEM AS MENSAGEM, '' AS AST " & _
                     "FROM BB_LOGINTEGRACOES WHERE SITUACAO = 2 AND TIPORESERVA = 27 AND TIPOERRO = 3 " & _
                     "AND DATAENVIO >= CONVERT(DATETIME, DATEADD(DAY, -3, GETDATE()))"
    agentevendas.Active = True

    Do While Not agentevendas.EOF
        CurValidar = 1
        StrTexto = StrTexto & " - DT_ENVIO: " & agentevendas.FieldByName("DT_ENVIO").AsString
        StrTexto = StrTexto & " - EMPRESA: " & agentevendas.FieldByName("EMPRESA").AsString
        StrTexto = StrTexto & " - PNR: " & agentevendas.FieldByName("OS").AsString & " <br>"
        StrTexto = StrTexto & "MENSAGEM: " & agentevendas.FieldByName("MENSAGEM").AsString & " <br>"
        StrTexto = StrTexto & "*: " & agentevendas.FieldByName("AST").AsString & " <br>"
        agentevendas.Next
    Loop

    agentevendas.Active = False
    Set agentevendas = Nothing

    If CurValidar > 0 Then
        Dim m As Mail
        Set m = NewMail
        m.Clear
        m.From = "sistemas@voetur.com.br"
        m.SendTo = "sistemas@voeturturismo.com.br"
        m.Subject = "[INTEGRATUR] OMNIBEES - LOG DE AGENTE DE VENDAS"
        m.Priority = 0
        m.IsHtml = True
        m.Text.Add StrTexto
        m.Send
        Set m = Nothing
    End If
End Sub

Public Sub FornecedorOmnibees()
    Dim StrTexto As String
    Dim CurValidar As Currency
    CurValidar = 0

    Dim errofornecedor As BPesquisa
    Set errofornecedor = NewQuery
    errofornecedor.Add "SELECT (SELECT NOMEFANTASIA FROM EMPRESAS WHERE HANDLE = EMPRESA) AS EMPRESA, " & _
                       "CASE TIPOERRO " & _
                       "WHEN '1' THEN 'Cliente' WHEN '2' THEN 'Fornecedor' WHEN '3' THEN 'Agente' " & _
                       "WHEN '4' THEN 'Inf Adicionais' WHEN '5' THEN 'Outros' WHEN '6' THEN 'Forma Pag/Rec' " & _
                       "WHEN '7' THEN 'Cancelamento' ELSE 'Não informado' END AS TIPO_ERRO, " & _
                       "CONVERT(VARCHAR, DATAENVIO, 103) AS DT_ENVIO, RLOCCIA AS OS, MENSAGEM AS MENSAGEM, '' AS AST " & _
                       "FROM BB_LOGINTEGRACOES WHERE SITUACAO = 2 AND TIPORESERVA = 27 AND TIPOERRO = 2 " & _
                       "AND DATAENVIO >= CONVERT(DATETIME, DATEADD(DAY, -3, GETDATE()))"
    errofornecedor.Active = True

    Do While Not errofornecedor.EOF
        CurValidar = 1
        StrTexto = StrTexto & " - DT_ENVIO: " & errofornecedor.FieldByName("DT_ENVIO").AsString
        StrTexto = StrTexto & " - EMPRESA: " & errofornecedor.FieldByName("EMPRESA").AsString
        StrTexto = StrTexto & " - PNR: " & errofornecedor.FieldByName("OS").AsString & " <br>"
        StrTexto = StrTexto & "MENSAGEM: " & errofornecedor.FieldByName("MENSAGEM").AsString & " <br>"
        StrTexto = StrTexto & "*: " & errofornecedor.FieldByName("AST").AsString & " <br>"
        errofornecedor.Next
    Loop

    errofornecedor.Active = False
    Set errofornecedor = Nothing

    If CurValidar > 0 Then
        Dim m As Mail
        Set m = NewMail
        m.Clear
        m.From = "sistemas@voetur.com.br"
        m.SendTo = "sistemas@voeturturismo.com.br"
        m.Subject = "[INTEGRATUR] OMNIBEES - LOG DE FORNECEDOR"
        m.Priority = 0
        m.IsHtml = True
        m.Text.Add StrTexto
        m.Send
        Set m = Nothing
    End If
End Sub

Public Sub ErrosOmnibees()
    Dim StrTexto As String
    Dim CurValidar As Currency
    CurValidar = 0

    Dim errosintegracao As BPesquisa
    Set errosintegracao = NewQuery
    errosintegracao.Add "SELECT (SELECT NOMEFANTASIA FROM EMPRESAS WHERE HANDLE = EMPRESA) AS EMPRESA, " & _
                        "CASE TIPOERRO " & _
                        "WHEN '1' THEN 'Cliente' WHEN '2' THEN 'Fornecedor' WHEN '3' THEN 'Agente' " & _
                        "WHEN '4' THEN 'Inf Adicionais' WHEN '5' THEN 'Outros' WHEN '6' THEN 'Forma Pag/Rec' " & _
                        "WHEN '7' THEN 'Cancelamento' ELSE 'Não informado' END AS TIPO_ERRO, " & _
                        "CONVERT(VARCHAR, DATAENVIO, 103) AS DT_ENVIO, RLOCCIA AS OS, MENSAGEM AS MENSAGEM, '' AS AST " & _
                        "FROM BB_LOGINTEGRACOES WHERE SITUACAO = 2 AND TIPORESERVA = 27 AND TIPOERRO NOT IN (1, 2, 3) " & _
                        "AND DATAENVIO >= CONVERT(DATETIME, DATEADD(DAY, -3, GETDATE()))"
    errosintegracao.Active = True

    Do While Not errosintegracao.EOF
        CurValidar = 1
        StrTexto = StrTexto & " - DT_ENVIO: " & errosintegracao.FieldByName("DT_ENVIO").AsString
        StrTexto = StrTexto & " - EMPRESA: " & errosintegracao.FieldByName("EMPRESA").AsString
        StrTexto = StrTexto & " - PNR: " & errosintegracao.FieldByName("OS").AsString & " <br>"
        StrTexto = StrTexto & "MENSAGEM: " & errosintegracao.FieldByName("MENSAGEM").AsString & " <br>"
        StrTexto = StrTexto & "*: " & errosintegracao.FieldByName("AST").AsString & " <br>"
        errosintegracao.Next
    Loop

    errosintegracao.Active = False
    Set errosintegracao = Nothing

    If CurValidar > 0 Then
        Dim m As Mail
        Set m = NewMail
        m.Clear
        m.From = "sistemas@voetur.com.br"
        m.SendTo = "sistemas@voeturturismo.com.br"
        m.Subject = "[INTEGRATUR] OMNIBEES - LOG DE ERROS DIVERSOS"
        m.Priority = 0
        m.IsHtml = True
        m.Text.Add StrTexto
        m.Send
        Set m = Nothing
    End If
End Sub

Public Sub AgenteWooba()
    Dim StrTexto As String
    Dim CurValidar As Currency
    CurValidar = 0

    Dim agentevendas As BPesquisa
    Set agentevendas = NewQuery
    agentevendas.Add "SELECT (SELECT NOMEFANTASIA FROM EMPRESAS WHERE HANDLE = EMPRESA) AS EMPRESA, " & _
                     "CASE TIPOERRO " & _
                     "WHEN '1' THEN 'Cliente' WHEN '2' THEN 'Fornecedor' WHEN '3' THEN 'Agente' " & _
                     "WHEN '4' THEN 'Inf Adicionais' WHEN '5' THEN 'Outros' WHEN '6' THEN 'Forma Pag/Rec' " & _
                     "WHEN '7' THEN 'Cancelamento' ELSE 'Não informado' END AS TIPO_ERRO, " & _
                     "CONVERT(VARCHAR, DATAENVIO, 103) AS DT_ENVIO, RLOCCIA AS OS, MENSAGEM AS MENSAGEM, '' AS AST " & _
                     "FROM BB_LOGINTEGRACOES WHERE SITUACAO = 2 AND TIPORESERVA = 24 AND TIPOERRO = 3 " & _
                     "AND DATAENVIO >= CONVERT(DATETIME, DATEADD(DAY, -3, GETDATE()))"
    agentevendas.Active = True

    Do While Not agentevendas.EOF
        CurValidar = 1
        StrTexto = StrTexto & " - DT_ENVIO: " & agentevendas.FieldByName("DT_ENVIO").AsString
        StrTexto = StrTexto & " - EMPRESA: " & agentevendas.FieldByName("EMPRESA").AsString
        StrTexto = StrTexto & " - PNR: " & agentevendas.FieldByName("OS").AsString & " <br>"
        StrTexto = StrTexto & "MENSAGEM: " & agentevendas.FieldByName("MENSAGEM").AsString & " <br>"
        StrTexto = StrTexto & "*: " & agentevendas.FieldByName("AST").AsString & " <br>"
        agentevendas.Next
    Loop

    agentevendas.Active = False
    Set agentevendas = Nothing

    If CurValidar > 0 Then
        Dim m As Mail
        Set m = NewMail
        m.Clear
        m.From = "sistemas@voetur.com.br"
        m.SendTo = "sistemas@voeturturismo.com.br"
        m.Subject = "[INTEGRATUR] WOOBA - LOG DE AGENTE DE VENDAS"
        m.Priority = 0
        m.IsHtml = True
        m.Text.Add StrTexto
        m.Send
        Set m = Nothing
    End If
End Sub

Public Sub FornecedorWooba()
    Dim StrTexto As String
    Dim CurValidar As Currency
    CurValidar = 0

    Dim errofornecedor As BPesquisa
    Set errofornecedor = NewQuery
    errofornecedor.Add "SELECT (SELECT NOMEFANTASIA FROM EMPRESAS WHERE HANDLE = EMPRESA) AS EMPRESA, " & _
                       "CASE TIPOERRO " & _
                       "WHEN '1' THEN 'Cliente' WHEN '2' THEN 'Fornecedor' WHEN '3' THEN 'Agente' " & _
                       "WHEN '4' THEN 'Inf Adicionais' WHEN '5' THEN 'Outros' WHEN '6' THEN 'Forma Pag/Rec' " & _
                       "WHEN '7' THEN 'Cancelamento' ELSE 'Não informado' END AS TIPO_ERRO, " & _
                       "CONVERT(VARCHAR, DATAENVIO, 103) AS DT_ENVIO, RLOCCIA AS OS, MENSAGEM AS MENSAGEM, '' AS AST " & _
                       "FROM BB_LOGINTEGRACOES WHERE SITUACAO = 2 AND TIPORESERVA = 24 AND TIPOERRO = 2 " & _
                       "AND DATAENVIO >= CONVERT(DATETIME, DATEADD(DAY, -3, GETDATE()))"
    errofornecedor.Active = True

    Do While Not errofornecedor.EOF
        CurValidar = 1
        StrTexto = StrTexto & " - DT_ENVIO: " & errofornecedor.FieldByName("DT_ENVIO").AsString
        StrTexto = StrTexto & " - EMPRESA: " & errofornecedor.FieldByName("EMPRESA").AsString
        StrTexto = StrTexto & " - PNR: " & errofornecedor.FieldByName("OS").AsString & " <br>"
        StrTexto = StrTexto & "MENSAGEM: " & errofornecedor.FieldByName("MENSAGEM").AsString & " <br>"
        StrTexto = StrTexto & "*: " & errofornecedor.FieldByName("AST").AsString & " <br>"
        errofornecedor.Next
    Loop

    errofornecedor.Active = False
    Set errofornecedor = Nothing

    If CurValidar > 0 Then
        Dim m As Mail
        Set m = NewMail
        m.Clear
        m.From = "sistemas@voetur.com.br"
        m.SendTo = "sistemas@voeturturismo.com.br"
        m.Subject = "[INTEGRATUR] WOOBA - LOG DE FORNECEDOR"
        m.Priority = 0
        m.IsHtml = True
        m.Text.Add StrTexto
        m.Send
        Set m = Nothing
    End If
End Sub

Public Sub ErrosWooba()
    Dim StrTexto As String
    Dim CurValidar As Currency
    CurValidar = 0

    Dim errosintegracao As BPesquisa
    Set errosintegracao = NewQuery
    errosintegracao.Add "SELECT (SELECT NOMEFANTASIA FROM EMPRESAS WHERE HANDLE = EMPRESA) AS EMPRESA, " & _
                        "CASE TIPOERRO " & _
                        "WHEN '1' THEN 'Cliente' WHEN '2' THEN 'Fornecedor' WHEN '3' THEN 'Agente' " & _
                        "WHEN '4' THEN 'Inf Adicionais' WHEN '5' THEN 'Outros' WHEN '6' THEN 'Forma Pag/Rec' " & _
                        "WHEN '7' THEN 'Cancelamento' ELSE 'Não informado' END AS TIPO_ERRO, " & _
                        "CONVERT(VARCHAR, DATAENVIO, 103) AS DT_ENVIO, RLOCCIA AS OS, MENSAGEM AS MENSAGEM, '' AS AST " & _
                        "FROM BB_LOGINTEGRACOES WHERE SITUACAO = 2 AND TIPORESERVA = 24 AND TIPOERRO NOT IN (1, 2, 3) " & _
                        "AND DATAENVIO >= CONVERT(DATETIME, DATEADD(DAY, -3, GETDATE()))"
    errosintegracao.Active = True

    Do While Not errosintegracao.EOF
        CurValidar = 1
        StrTexto = StrTexto & " - DT_ENVIO: " & errosintegracao.FieldByName("DT_ENVIO").AsString
        StrTexto = StrTexto & " - EMPRESA: " & errosintegracao.FieldByName("EMPRESA").AsString
        StrTexto = StrTexto & " - PNR: " & errosintegracao.FieldByName("OS").AsString & " <br>"
        StrTexto = StrTexto & "MENSAGEM: " & errosintegracao.FieldByName("MENSAGEM").AsString & " <br>"
        StrTexto = StrTexto & "*: " & errosintegracao.FieldByName("AST").AsString & " <br>"
        errosintegracao.Next
    Loop

    errosintegracao.Active = False
    Set errosintegracao = Nothing

    If CurValidar > 0 Then
        Dim m As Mail
        Set m = NewMail
        m.Clear
        m.From = "sistemas@voetur.com.br"
        m.SendTo = "sistemas@voeturturismo.com.br"
        m.Subject = "[INTEGRATUR] WOOBA - LOG DE ERROS DIVERSOS"
        m.Priority = 0
        m.IsHtml = True
        m.Text.Add StrTexto
        m.Send
        Set m = Nothing
    End If
End Sub

Public Sub Main()
    Call AgenteAmadeus
    Call FornecedorAmadeus
    Call ErrosAmadeus
	Call AgenteArgo
    Call FornecedorArgo
    Call ErrosArgo
	Call AgenteB2b
    Call FornecedorB2b
    Call ErrosB2b
	Call AgenteBuscaIdeal
    Call FornecedorBuscaIdeal
    Call ErrosBuscaIdeal
	Call AgenteEnvision
    Call FornecedorEnvision
    Call ErrosEnvision
	Call AgenteLemontech
    Call FornecedorLemontech
    Call ErrosLemontech
	Call AgenteOmnibees
    Call FornecedorOmnibees
    Call ErrosOmnibees
	Call AgenteWooba
    Call FornecedorWooba
    Call ErrosWooba


End Sub
