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
 * Função Auxiliar: Registra o log detalhado em uma aba separada
 */
function registrarLog(nome, contato, dias, mensagem, tipo, status) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var abaLog = ss.getSheetByName("Log_Envios");
  
  if (!abaLog) {
    abaLog = ss.insertSheet("Log_Envios");
    abaLog.appendRow(["Data/Hora", "Nome", "Contato", "Dias", "Mensagem", "Tipo", "Status"]);
    abaLog.getRange("A1:G1").setFontWeight("bold").setBackground("#f3f3f3");
  }
  
  abaLog.appendRow([new Date(), nome, contato, dias, mensagem, tipo, status]);
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

    if (!mensagem || !nome) continue; 

    var payload = {
      "nome": nome, "dias": dias, "mensagem": mensagem,
      "contato": contato, "linha_original": i + 1, "tipo_envio": "bulk"
    };

    var opcoes = {
      "method": "post",
      "contentType": "application/json",
      "payload": JSON.stringify(payload),
      "muteHttpExceptions": true 
    };

    try {
      var resposta = UrlFetchApp.fetch(urlWebhook, opcoes);
      var code = resposta.getResponseCode();
      
      if (code >= 200 && code < 300) {
        aba.getRange(i + 1, 11).setValue(true); // Marca Checkbox
        registrarLog(nome, contato, dias, mensagem, "Bulk", "Enviado ✅");
        contagemProcessados++;
      } else {
        registrarLog(nome, contato, dias, mensagem, "Bulk", "Erro HTTP: " + code);
      }

      if (contagemProcessados % 10 === 0) {
        planilha.toast("Processados: " + contagemProcessados, "Status Bulk", 2);
      }
      Utilities.sleep(500); 
    } catch (e) {
      registrarLog(nome, contato, dias, mensagem, "Bulk", "Falha Crítica: " + e.toString());
    }
  }
  planilha.toast("Processo finalizado!", "Sucesso", 5);
}

/**
 * Função SINGLE: Envia apenas os dados da linha onde o cursor está posicionado.
 */
function enviarApenasLinhaSelecionada() {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var aba = planilha.getActiveSheet();
  var linhaAtiva = aba.getActiveCell().getRow();
  
  if (linhaAtiva === 1) {
    SpreadsheetApp.getUi().alert("Atenção: Selecione uma linha com dados.");
    return;
  }

  var dadosLinha = aba.getRange(linhaAtiva, 1, 1, aba.getLastColumn()).getValues()[0];
  var nome = dadosLinha[1], dias = dadosLinha[4], mensagem = dadosLinha[8], contato = dadosLinha[9];
  var urlWebhook = "https://webhook.pluglead.com/webhook/bd1808db-5d6f-4cec-90e2-4fa76c76a166";

  if (!nome || !mensagem) {
    planilha.toast("Dados incompletos.", "Erro", 5);
    return;
  }

  var opcoes = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify({"nome": nome, "dias": dias, "mensagem": mensagem, "contato": contato, "tipo_envio": "single"}),
    "muteHttpExceptions": true
  };

  try {
    var resposta = UrlFetchApp.fetch(urlWebhook, opcoes);
    var code = resposta.getResponseCode();

    if (code >= 200 && code < 300) {
      aba.getRange(linhaAtiva, 11).setValue(true);
      registrarLog(nome, contato, dias, mensagem, "Single", "Enviado ✅");
      planilha.toast("Enviado com sucesso!", "Webhook", 5);
    } else {
      registrarLog(nome, contato, dias, mensagem, "Single", "Erro HTTP: " + code);
      planilha.toast("Erro no servidor: " + code, "Falha", 5);
    }
  } catch (e) {
    registrarLog(nome, contato, dias, mensagem, "Single", "Erro: " + e.toString());
    SpreadsheetApp.getUi().alert("Erro: " + e.toString());
  }
}
