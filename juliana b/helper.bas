Public planilha As Workbook
Public Const ULTIMA_CELULA As String = "1048576"

Public Sub msg_de_nao_preenchido(nome As String, Optional oa As String = "O")
  Call MsgBox("POR FAVOR, INFORME " & oa & " " & nome & "!", vbExclamation, nome & " NÃO INFORMADO")
End Sub

Public Function eh_valido(campo As Object) As Boolean
  eh_valido = True:
  If campo.Value = "" Or campo.Value = False Or campo.Value = 0 Then campo.SetFocus: eh_valido = False
End Function

Public Function ultima_linha(nome_aba As String, Optional coluna As String = "A") As Integer
  ultima_linha = Sheets(nome_aba).Range(coluna & ULTIMA_CELULA).End(xlUp).Row + 1
End Function

Public Sub listar(tela As MSForms.ListBox, nome_aba As String, numero_de_colunas As Integer, Optional col_fim As String = "J")
  With tela
    .ColumnCount = numero_de_colunas
    .ColumnHeads = True
    .RowSource = nome_aba & "!A2:" & col_fim & ultima_linha(nome_aba, col_fim)
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

Public Sub cria_tabela(nome As String, coluna As Integer, col_valor As String, _
                    Optional campo As String = "", Optional posicao As Integer = 0)

  Call planilha.PivotCaches.Create(xlDatabase, nome & "!" & planilha.Sheets(nome).UsedRange.Address) _
                                        .CreatePivotTable("RESULTADO!R1C" & coluna, "TOTAL " & nome)

  With planilha.Sheets("RESULTADO").PivotTables("TOTAL " & nome)
    .CompactLayoutRowHeader = nome & "S"

    With .PivotFields("TIPO")
      .Orientation = xlRowField
      .Position = 1
    End With
    If campo <> "" Then With .PivotFields(campo): .Orientation = xlRowField: .Position = posicao: End With

    Call .AddDataField(.PivotFields("VALOR"), "TOTAIS", xlSum)
    Columns(col_valor & ":" & col_valor).Style = "Currency"
  End With
End Sub