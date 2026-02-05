function resetarCheckboxesPagina14() {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var aba = planilha.getSheetByName("Página14");
  
  if (!aba) return;

  var ultimaLinha = aba.getLastRow();
  if (ultimaLinha < 2) return;

  var range = aba.getRange("K2:K" + ultimaLinha);
  var valores = range.getValues();

  for (var i = 0; i < valores.length; i++) {
    var valor = valores[i][0];

    // Só altera se:
    // - não for FALSE
    // - e não for vazio
    if (valor !== false && valor !== "" && valor !== null) {
      valores[i][0] = false;
    }
  }

  range.setValues(valores);
}
