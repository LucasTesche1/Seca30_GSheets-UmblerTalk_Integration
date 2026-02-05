/**
 * Cria um menu personalizado no Google Sheets ao abrir o arquivo.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🚀 Enviar Mensagens')
      .addItem('Enviar Linha Selecionada (Single)', 'enviarApenasLinhaSelecionada')
      .addSeparator()
      .addItem('Enviar Todas as Linhas (Bulk)', 'iteraExcelEnviaWebhook')
      .addToUi();
}

/**
 * Função BULK: Itera por toda a aba e envia para o webhook.
 */
function iteraExcelEnviaWebhook() {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var aba = planilha.getActiveSheet();
  var dados = aba.getDataRange().getValues();
  var contagemProcessados = 0;
  var urlWebhook = "https://webhook.pluglead.com/webhook/bd1808db-5d6f-4cec-90e2-4fa76c76a166";

  planilha.toast("Iniciando o envio em lote...", "Automação", 3);

  for (var i = 1; i < dados.length; i++) {
    var linha = dados[i];
    
    var nome     = linha[1]; // Coluna B
    var dias     = linha[4]; // Coluna E
    var mensagem = linha[8]; // Coluna I
    var contato  = linha[9]; // Coluna J

    // Validação de dados
    if (!mensagem || mensagem.toString().trim() === "" || !nome || nome.toString().trim() === "") {
      continue; 
    }

    var payload = {
      "nome": nome,
      "dias": dias,
      "mensagem": mensagem,
      "contato": contato,
      "linha_original": i + 1,
      "tipo_envio": "bulk"
    };

    var opcoes = {
      "method": "post",
      "contentType": "application/json",
      "payload": JSON.stringify(payload),
      "muteHttpExceptions": true 
    };

    try {
      UrlFetchApp.fetch(urlWebhook, opcoes);
      contagemProcessados++;

      //Marca checkbox
      aba.getRange(i + 1,11).setValue(true);

      // Toast de progresso a cada 10 envios
      if (contagemProcessados % 10 === 0) {
        planilha.toast("Enviados: " + contagemProcessados + " de " + (dados.length - 1), "Status Bulk", 2);
      }
      
      Utilities.sleep(500); // Pausa para evitar "Too Many Requests"
    } catch (e) {
      console.log("Erro ao enviar linha " + (i + 1) + ": " + e.toString());
    }
  }

  planilha.toast("Finalizado! " + contagemProcessados + " mensagens enviadas. ✅", "Sucesso", 10);
}

/**
 * Função SINGLE: Envia apenas os dados da linha onde o cursor está posicionado.
 */
function enviarApenasLinhaSelecionada() {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var aba = planilha.getActiveSheet();
  var linhaAtiva = aba.getActiveCell().getRow();
  
  // Valida se não está tentando enviar o cabeçalho
  if (linhaAtiva === 1) {
    SpreadsheetApp.getUi().alert("Atenção: Você selecionou o cabeçalho. Escolha uma linha com dados.");
    return;
  }

  // Pega os dados apenas da linha selecionada
  var dadosLinha = aba.getRange(linhaAtiva, 1, 1, aba.getLastColumn()).getValues()[0];

  var nome     = dadosLinha[1]; // Coluna B
  var dias     = dadosLinha[4]; // Coluna E
  var mensagem = dadosLinha[8]; // Coluna I
  var contato  = dadosLinha[9]; // Coluna J
  var urlWebhook = "https://webhook.pluglead.com/webhook/bd1808db-5d6f-4cec-90e2-4fa76c76a166";

  if (!nome || !mensagem) {
    planilha.toast("Erro: Nome ou Mensagem ausentes na linha " + linhaAtiva, "Falha", 5);
    return;
  }

  var payload = {
    "nome": nome,
    "dias": dias,
    "mensagem": mensagem,
    "contato": contato,
    "linha_original": linhaAtiva,
    "tipo_envio": "single"
  };

  var opcoes = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  try {
    planilha.toast("Enviando dados de " + nome + "...", "Status Single", 2);
    UrlFetchApp.fetch(urlWebhook, opcoes);

    aba.getRange(linhaAtiva, 11).setValue(true);

    planilha.toast("Sucesso! Enviado para " + nome + " ✅", "Webhook", 5);
  } catch (e) {
    SpreadsheetApp.getUi().alert("Erro na requisição: " + e.toString());
  }
}
