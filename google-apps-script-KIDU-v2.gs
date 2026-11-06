// Kidu, Apps Script v2, esquema de colunas enviado pela Carol em 2026-11-06.
// Colado por ela em 2026-11-06, URL /exec nova. Esta versão corrige o esquema
// para exatamente as colunas que ela mandou, na ordem dela.
//
// Colunas, na ordem:
// Dados da Criança e Pedido, preenchidos no quiz via Apps Script:
//   ID_Pedido, Data_Criacao, E-mail, Nome_Crianca, Idade, Genero, Pele, Cabelo,
//   Tema, Moral, Dedicatoria
// Dados de Origem e Rastreamento:
//   UTM_Source, UTM_Campaign, UTM_Medium
// Dados Financeiros e Status, preenchidos e atualizados pelo n8n:
//   Bump_Colorir, Valor_Pago, Status, Data_Pagamento
// O Apps Script escreve Status como Pendente na criação, e o n8n atualiza para
// Pago quando o pagamento é aprovado, preenchendo Valor_Pago e Data_Pagamento.
// Bump_Colorir, Valor_Pago e Data_Pagamento começam vazios, o n8n preenche.
// Pele e Cabelo vêm do personagem escolhido no quiz, não há coluna Personagem.
//
// Limpeza automática: Pendente, ou sem status, com mais de 48h vai para a aba
// Temporarios. Temporario com mais de 7 dias desde a criação é apagado, só se
// ainda estiver Pendente ou sem status. Se um pedido em Temporarios virar Pago,
// a limpeza devolve ele para Pedidos-Livros. Pago nunca é apagado.
//
// Regra do n8n: o n8n só gera e envia o PDF quando o Status é Pago, pagamento
// aprovado. Pendente nunca dispara envio. O gatilho é o webhook da Stripe ou
// a leitura da planilha procurando Status Pago, nunca a criação do pedido.
//
// Para implantar, ela cola este arquivo inteiro no Apps Script, salva, e cria
// uma nova versão da mesma implantação, a URL /exec não muda. Se criar uma
// implantação nova, ela manda a URL nova para preencher no index.html.
// Depois, uma única vez, roda a função criarTriggerLimpeza() no editor, para
// criar o gatilho diário.

var ABA_PEDIDOS = "Pedidos-Livros";
var ABA_METRICAS = "Metricas";
var ABA_TEMPORARIOS = "Temporarios";
var HORAS_PARA_TEMPORARIO = 48;
var DIAS_PARA_APAGAR = 7;
var COL_STATUS = 16;
var COL_DATA = 1;

function cabecalhoPedidos() {
  return ["ID_Pedido", "Data_Criacao", "E-mail", "Nome_Crianca", "Idade", "Genero", "Pele", "Cabelo", "Tema", "Moral", "Dedicatoria", "UTM_Source", "UTM_Campaign", "UTM_Medium", "Bump_Colorir", "Valor_Pago", "Status", "Data_Pagamento"];
}

function obterAba(ss, nome, cabecalho) {
  var aba = ss.getSheetByName(nome);
  if (!aba) {
    aba = ss.insertSheet(nome);
    if (cabecalho) aba.appendRow(cabecalho);
  }
  return aba;
}

function doPost(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var data = JSON.parse(e.postData.contents);
    var dataAtual = new Date().toLocaleString("pt-BR", { timeZone: "America/Sao_Paulo" });

    // 1. Evento de métrica, só deve chegar aqui quem já chegou na prévia/checkout.
    if (data.tipo_requisicao === "evento_metrica") {
      var abaMetricas = obterAba(ss, ABA_METRICAS, ["Data/Hora", "Origem", "Meio", "Campanha", "Evento"]);
      abaMetricas.appendRow([
        dataAtual,
        data.utm_source || "direto",
        data.utm_medium || "nenhum",
        data.utm_campaign || "nenhuma",
        data.evento || "visita"
      ]);
      return ContentService.createTextOutput(JSON.stringify({ status: "success" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 2. Pedido, só é criado quando a pessoa chega na etapa final e clica em Garantir.
    var abaPedidos = obterAba(ss, ABA_PEDIDOS, cabecalhoPedidos());
    var pedidoId = data.order_id || "";
    if (!/^KIDU-\d{4}$/.test(pedidoId)) {
      pedidoId = "KIDU-" + Math.floor(1000 + Math.random() * 9000);
    }
    var moral = data.moral || data.valor || data.ensinamento || data.valor_ensinamento || "";
    var idade = data.idade || data.faixa_idade || "";

    abaPedidos.appendRow([
      pedidoId,
      dataAtual,
      data.email || "",
      data.nome_crianca || "",
      idade,
      data.genero || "",
      data.pele || "",
      data.cabelo || "",
      data.tema || "",
      moral,
      data.dedicatoria || "",
      data.utm_source || "direto",
      data.utm_campaign || "nenhuma",
      data.utm_medium || "nenhum",
      "",
      "",
      "Pendente",
      ""
    ]);

    return ContentService.createTextOutput(JSON.stringify({ status: "success", pedido_id: pedidoId }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Leitura para o dashboardpainel.html, que faz GET nesta mesma URL /exec.
function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var abaMetricas = ss.getSheetByName(ABA_METRICAS);
  var abaPedidos = ss.getSheetByName(ABA_PEDIDOS) || ss.getSheetByName("Pedidos - Livros") || ss.getSheets()[0];
  var metricas = abaMetricas ? abaMetricas.getDataRange().getValues().slice(1) : [];
  var pedidos = abaPedidos ? abaPedidos.getDataRange().getValues().slice(1) : [];
  var payload = {
    metricas: metricas.map(function (r) { return { data: r[0], origem: r[1], meio: r[2], campanha: r[3], evento: r[4] }; }),
    pedidos: pedidos.map(function (r) { return { id: r[0], data: r[COL_DATA], email: r[2], nome: r[3], status: r[COL_STATUS], bump: r[14], total: r[15] }; })
  };
  return ContentService.createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

// Limpeza diária, roda por gatilho. Não apaga nada com Status Pago.
function limparPendentesAntigos() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var agora = new Date();
  var limiteTemporario = new Date(agora.getTime() - HORAS_PARA_TEMPORARIO * 60 * 60 * 1000);
  var limiteApagar = new Date(agora.getTime() - DIAS_PARA_APAGAR * 24 * 60 * 60 * 1000);

  var abaTempPrimeiro = ss.getSheetByName(ABA_TEMPORARIOS);
  var abaPedidosPrimeiro = ss.getSheetByName(ABA_PEDIDOS) || ss.getSheetByName("Pedidos - Livros");
  if (abaTempPrimeiro && abaPedidosPrimeiro) {
    var dadosResgate = abaTempPrimeiro.getDataRange().getValues();
    for (var r = dadosResgate.length - 1; r >= 1; r--) {
      if (String(dadosResgate[r][COL_STATUS] || "").trim() === "Pago") {
        abaPedidosPrimeiro.appendRow(dadosResgate[r]);
        abaTempPrimeiro.deleteRow(r + 1);
      }
    }
  }

  var abaPedidos = ss.getSheetByName(ABA_PEDIDOS);
  if (abaPedidos) {
    var abaTemp = obterAba(ss, ABA_TEMPORARIOS, cabecalhoPedidos());
    var dados = abaPedidos.getDataRange().getValues();
    for (var i = dados.length - 1; i >= 1; i--) {
      var linha = dados[i];
      var status = String(linha[COL_STATUS] || "").trim();
      var dataCriacao = new Date(linha[COL_DATA]);
      if ((status === "Pendente" || status === "") && dataCriacao < limiteTemporario) {
        abaTemp.appendRow(linha);
        abaPedidos.deleteRow(i + 1);
      }
    }
  }

  var abaTemp2 = ss.getSheetByName(ABA_TEMPORARIOS);
  if (abaTemp2) {
    var dadosTemp = abaTemp2.getDataRange().getValues();
    for (var j = dadosTemp.length - 1; j >= 1; j--) {
      var statusTemp = String(dadosTemp[j][COL_STATUS] || "").trim();
      var dataTemp = new Date(dadosTemp[j][COL_DATA]);
      if (dataTemp < limiteApagar && statusTemp !== "Pago") abaTemp2.deleteRow(j + 1);
    }
  }

  var abaMetricas = ss.getSheetByName(ABA_METRICAS);
  if (abaMetricas) {
    var dadosMet = abaMetricas.getDataRange().getValues();
    for (var k = dadosMet.length - 1; k >= 1; k--) {
      var dataMet = new Date(dadosMet[k][0]);
      if (dataMet < limiteApagar) abaMetricas.deleteRow(k + 1);
    }
  }
}

// Rodar uma única vez no editor do Apps Script para criar o gatilho diário.
function criarTriggerLimpeza() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "limparPendentesAntigos") return;
  }
  ScriptApp.newTrigger("limparPendentesAntigos").timeBased().everyDays(1).atHour(3).create();
}
