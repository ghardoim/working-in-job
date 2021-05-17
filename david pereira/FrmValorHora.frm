Private Sub spn_diasSemana_SpinDown()
  Call spinChange(Me.txt_diasSemana, 0, False)
End Sub

Private Sub spn_diasSemana_SpinUp()
  Call spinChange(Me.txt_diasSemana, 7)
End Sub

Private Sub spn_horasDia_SpinDown()
  Call spinChange(Me.txt_horasDia, 0, False)
End Sub

Private Sub spn_horasDia_SpinUp()
  Call spinChange(Me.txt_horasDia, 24)
End Sub

Private Sub spn_diasSemana_Change()
  Call calculaValorHora(Me.txt_diasSemana, "F", (Me.txt_diasSemana.value * Me.txt_horasDia.value))
End Sub

Private Sub spn_horasDia_Change()
  Call calculaValorHora(Me.txt_horasDia, "H", (Me.txt_diasSemana.value * Me.txt_horasDia.value))
End Sub