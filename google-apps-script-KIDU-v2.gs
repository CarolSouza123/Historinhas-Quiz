// Kidu, Apps Script v2, preparado em 2026-11-06, AINDA NÃO COLADO NO APPS SCRIPT DELA
// O que muda em relação ao google-apps-script.gs que está no GitHub:
// 1. doGet incluído, o código implantado hoje não tem doGet, por isso o dashboard não carrega.
// 2. Pedido aceita personagem, faixa de idade e valor/ensinamento, além de pele e cabelo.
// 3. Pedido usa o código KIDU que o site manda, nunca cria PED.
// 4. Limpeza automática: Pendente com mais de 48h vai para a aba Temporários,
//    Temporário com mais de 7 dias desde a criação é apagado. Pago nunca é mexido.
// 5. Métricas: o site só deve mandar evento para cá quando chegar na prévia/checkout,
//    ver registrarEvento no index.html. As linhas de métrica também expiram em 7 dias.
//
// Para implantar, ela cola este arquivo inteiro no Apps Script, salva, implanta
// como aplicativo da Web, Executar como Eu, acesso Qualquer pessoa, copia a URL /exec.
// Depois, uma única vez, roda a função criarTriggerLimpeza() no editor, para criar
// o gatilho diário. A URL /exec não muda se for uma nova versão da mesma implantação,
// se criar uma implantação nova, ela manda a URL nova para preencher no index.html.

var ABA_PEDIDOS = "Pedidos-Livros";
var ABA_METRICAS = "Metricas";
var ABA_TEMPORARIOS = "Temporarios";
var HORAS_PARA_TEMPORARIO = 48;
var DIAS_PARA_APAGAR = 7;

function cabecalhoPedidos() {
  return ["ID Pedido", "Data Criação", "Status", "Nome Criança", "Faixa etária", "Gênero", "Pele", "Cabelo", "Personagem", "Tema", "Valor / Ensinamento", "Dedicatória", "Origem UTM", "Comprador", "E-mail", "Order Bump", "Valor Total"];
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

    // 2. Pedido, só é criado quando a pessoa clica em Garantir, na prévia.
    var abaPedidos = obterAba(ss, ABA_PEDIDOS, cabecalhoPedidos());
    var pedidoId = data.order_id || "";
    if (!/^KIDU-\d{4}$/.test(pedidoId)) {
      pedidoId = "KIDU-" + Math.floor(1000 + Math.random() * 9000);
    }
    var valorEnsinamento = data.valor || data.ensinamento || data.valor_ensinamento || "";
    var faixaIdade = data.faixa_idade || data.idade || "";
    var bumpTexto = data.order_bump ? "Sim" : "Não";
    var totalNumero = data.valor_total ? Number(data.valor_total) : (data.order_bump ? 51.9 : 37);

    abaPedidos.appendRow([
      pedidoId,
      dataAtual,
      "Pendente",
      data.nome_crianca || "",
      faixaIdade,
      data.genero || "",
      data.pele || "",
      data.cabelo || "",
      data.personagem || "",
      data.tema || "",
      valorEnsinamento,
      data.dedicatoria || "",
      data.utm_source || "direto",
      data.comprador || "",
      data.email || "",
      bumpTexto,
      totalNumero
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
  var abaPedidos = ss.getSheetByName(ABA_PEDIDOS) || ss.getSheets()[0];
  var metricas = abaMetricas ? abaMetricas.getDataRange().getValues().slice(1) : [];
  var pedidos = abaPedidos ? abaPedidos.getDataRange().getValues().slice(1) : [];
  var payload = {
    metricas: metricas.map(function (r) { return { data: r[0], origem: r[1], meio: r[2], campanha: r[3], evento: r[4] }; }),
    pedidos: pedidos.map(function (r) { return { id: r[0], data: r[1], status: r[2], bump: r[15], total: r[16] }; })
  };
  return ContentService.createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

// Limpeza diária, roda por gatilho. Não apaga nada com Status Pago.
// Pendente com mais de 48h: move para Temporarios.
// Temporario com mais de 7 dias desde a criação: apaga.
// Métrica com mais de 7 dias: apaga, para não ficar dado de quem não pagou para sempre.
function limparPendentesAntigos() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var agora = new Date();
  var limiteTemporario = new Date(agora.getTime() - HORAS_PARA_TEMPORARIO * 60 * 60 * 1000);
  var limiteApagar = new Date(agora.getTime() - DIAS_PARA_APAGAR * 24 * 60 * 60 * 1000);

  var abaPedidos = ss.getSheetByName(ABA_PEDIDOS);
  if (abaPedidos) {
    var abaTemp = obterAba(ss, ABA_TEMPORARIOS, cabecalhoPedidos());
    var dados = abaPedidos.getDataRange().getValues();
    for (var i = dados.length - 1; i >= 1; i--) {
      var linha = dados[i];
      var status = String(linha[2] || "").trim();
      var dataCriacao = new Date(linha[1]);
      if (status === "Pendente" && dataCriacao < limiteTemporario) {
        abaTemp.appendRow(linha);
        abaPedidos.deleteRow(i + 1);
      }
    }
  }

  var abaTemp2 = ss.getSheetByName(ABA_TEMPORARIOS);
  if (abaTemp2) {
    var dadosTemp = abaTemp2.getDataRange().getValues();
    for (var j = dadosTemp.length - 1; j >= 1; j--) {
      var dataTemp = new Date(dadosTemp[j][1]);
      if (dataTemp < limiteApagar) abaTemp2.deleteRow(j + 1);
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
