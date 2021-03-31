Public planilha As Workbook

Public Sub msg_de_nao_preenchido(nome As String, Optional oa As String = "O")
  Call MsgBox("POR FAVOR, INFORME " & oa & " " & nome & "!", vbExclamation, nome & " NÃO INFORMADO")
End Sub

Public Function linha_eh_igual(ByVal itens As Integer, da_lista As MSForms.ListBox, da_tabela As Range, plan As Worksheet) As Boolean
  Dim i As Integer
  linha_eh_igual = True
  For i = 0 To itens - 1
    If da_lista.List(da_lista.ListIndex, i) <> plan.Cells(da_tabela.Row, i + 1) Then
      linha_eh_igual = False
    End If
  Next
End Function

Public Function eh_valido(campo As Object) As Boolean
  eh_valido = True:
  If campo.Value = "" Or campo.Value = False Or campo.Value = 0 Then campo.SetFocus: eh_valido = False
End Function

Public Function ultima_linha(nome_aba As String, Optional coluna As String = "A", Optional plan As Workbook = Nothing) As Integer
  If plan Is Nothing Then Set plan = ThisWorkbook
  ultima_linha = plan.Sheets(nome_aba).Range(coluna & "1048576").End(xlUp).Row + 1
End Function

Public Sub excluir_linha(lista As MSForms.ListBox, nome_aba As String, indices As Variant)
  Dim excluir As String: excluir = lista.List(lista.ListIndex, indices(0))
  Dim valor As Double: valor = lista.List(lista.ListIndex, indices(1))

  If vbYes = MsgBox("DESEJA APAGAR?" & vbNewLine & excluir, vbYesNo + vbExclamation, "REMOVER SAÍDA") Then
    With planilha.Sheets(nome_aba)

      Dim cell_encontrada As Range: Set cell_encontrada = .Range(indices(2) & "1")
      Do While True

        Set cell_encontrada = .Range(indices(2) & ":" & indices(2)) _
            .Find(excluir, .Cells(cell_encontrada.Row, cell_encontrada.Column))

        If linha_eh_igual(indices(3), lista, cell_encontrada, planilha.Sheets(nome_aba)) Then

          Call MsgBox("ITEM APAGADO!" & vbNewLine & excluir & vbNewLine & _
          "VALOR: " & .Cells(cell_encontrada.Row, indices(3)), vbInformation, "SUCESSO")

          cell_encontrada.EntireRow.Delete
          Exit Do
        End If
      Loop
    End With

    Call exclui_extrato_encontrado(excluir, indices(4), valor, indices(5))
    Call listar(lista, nome_aba, indices(3), indices(6))
  End If
End Sub

Public Sub exclui_extrato_encontrado(excluir As String, ByVal col_excluir As String, valor_help As Double, ByVal col_valor As String)
  With planilha.Sheets("EXTRATO")

    Dim cell_extrato As Range: Set cell_extrato = Range(col_excluir & "1")
    Dim cell_valor As Range: Set cell_valor = Range(col_valor & "1")

    Do While True
      Set cell_extrato = .Range(col_excluir & ":" & col_excluir).Find(excluir, .Cells(cell_extrato.Row, cell_extrato.Column))
      Set cell_valor = .Range(col_valor & ":" & col_valor).Find(valor_help, .Cells(cell_valor.Row, cell_valor.Column))

      If cell_extrato.Row = cell_valor.Row Then cell_extrato.EntireRow.Delete: Exit Do
    Loop
  End With
End Sub

Public Sub listar(tela As MSForms.ListBox, nome_aba As String, ByVal numero_de_colunas As Integer, Optional ByVal col_fim As String = "J")
  planilha.Activate
  With tela
    .ColumnCount = numero_de_colunas
    .ColumnHeads = True
    .RowSource = nome_aba & "!A2:" & col_fim & ultima_linha(nome_aba, col_fim, planilha)
  End With
End Sub

Public Sub liga_desliga(on_off As Boolean)
  With Application
    If on_off Then .Calculation = xlCalculationAutomatic
    If Not on_off Then .Calculation = xlCalculationManual
    .ScreenUpdating = on_off
    .DisplayAlerts = on_off
    .Visible = on_off
  End With
End Sub

Public Sub limpar_filtro(nome_aba As String)
  planilha.Sheets("AUXILIAR").Range("A1").CurrentRegion.Clear
  planilha.Sheets(nome_aba).UsedRange.AutoFilter
End Sub

Public Sub filtrar(nome_aba As String, filtro_1 As String, col_1 As Integer, _
                            Optional filtro_2 As String = "", Optional col_2 As Integer = 0, _
                          Optional filtro_3 As String = "", Optional col_3 As Integer = 0, _
                        Optional filtro_4 As String = "", Optional col_4 As Integer = 0)

  Call limpar_filtro(nome_aba)
  With planilha.Sheets(nome_aba)
    If filtro_1 <> "" And filtro_1 <> "**" Then Call .UsedRange.AutoFilter(col_1, filtro_1)
    If filtro_2 <> "" And filtro_2 <> "**" Then Call .UsedRange.AutoFilter(col_2, filtro_2)
    If filtro_3 <> "" Then Call .UsedRange.AutoFilter(col_3, filtro_3)
    If filtro_4 <> "" Then Call .UsedRange.AutoFilter(col_4, filtro_4)

    .Range("A1").CurrentRegion.Copy planilha.Sheets("AUXILIAR").Range("A1")
  End With
End Sub

Public Sub cria_tabela(nome_aba As String, col_posicionamento As Integer, nome_campo_total As String, col_formatar_valor As String, _
                        Optional nome_campo As String = "", Optional posicao_campo As Integer = 0, _
                        Optional outro_total As String = "", Optional outra_col_format As String = "")
    
  Call planilha.PivotCaches.Create(xlDatabase, nome_aba & "!" & planilha.Sheets(nome_aba).UsedRange.Address) _
                                .CreatePivotTable("RESULTADO!R1C" & col_posicionamento, "TOTAL " & nome_aba)

  With planilha.Sheets("RESULTADO").PivotTables("TOTAL " & nome_aba)
    .CompactLayoutRowHeader = nome_aba & "S"
    .DataPivotField.Caption = " "
    
    With .PivotFields("TIPO")
      .Orientation = xlRowField
      .Position = 1
    End With
    If nome_campo <> "" Then With .PivotFields(nome_campo): .Orientation = xlRowField: .Position = posicao_campo: End With
    
    Call .AddDataField(.PivotFields(nome_campo_total), "TOTAIS", xlSum)
    If outro_total <> "" Then Call .AddDataField(.PivotFields(outro_total), "IMPOSTOS", xlSum)

    With .PivotFields(nome_campo_total)
      .Orientation = xlPageField
      .Position = 1
      .CurrentPage = "(All)"
      .PivotItems("$0.00").Visible = False
      .EnableMultiplePageItems = True
    End With

    If outra_col_format <> "" Then Columns(outra_col_format & ":" & outra_col_format).Style = "Currency"
    Columns(col_formatar_valor & ":" & col_formatar_valor).Style = "Currency"
  End With
End Sub