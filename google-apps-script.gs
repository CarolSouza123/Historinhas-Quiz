function doPost(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const data = JSON.parse(e.postData.contents);
    const dataAtual = new Date().toLocaleString("pt-BR", { timeZone: "America/Sao_Paulo" });

    // 1. EVENTOS DE MÉTRICA / UTM
    if (data.tipo_requisicao === "evento_metrica") {
      let sheetMetricas = ss.getSheetByName("Metricas");
      if (!sheetMetricas) {
        sheetMetricas = ss.insertSheet("Metricas");
        sheetMetricas.appendRow(["Data/Hora", "Origem", "Meio", "Campanha", "Evento"]);
      }
      sheetMetricas.appendRow([
        dataAtual,
        data.utm_source || "direto",
        data.utm_medium || "nenhum",
        data.utm_campaign || "nenhuma",
        data.evento || "visita"
      ]);

      return ContentService.createTextOutput(JSON.stringify({ status: "success" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 2. CRIAÇÃO DE PEDIDO (CHECKOUT PENDENTE)
    // Cabeçalho esperado em "Pedidos - Livros", linha 1:
    // ID Pedido | Data Criação | Status | Nome Criança | Gênero | Pele | Cabelo | Tema | Dedicatória | Origem UTM | Comprador | E-mail | Order Bump | Valor Total
    let sheetPedidos = ss.getSheetByName("Pedidos - Livros") || ss.getSheets()[0];
    const pedidoId = "PED-" + Math.floor(10000 + Math.random() * 90000);
    const bumpTexto = data.order_bump ? "Sim" : "Não";
    const totalNumero = data.valor_total ? Number(data.valor_total) : (data.order_bump ? 51.9 : 37);

    sheetPedidos.appendRow([
      pedidoId,
      dataAtual,
      "Pendente",
      data.nome_crianca || "",
      data.genero || "",
      data.pele || "",
      data.cabelo || "",
      data.tema || "",
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

// Endpoint GET para o Painel ler as métricas da planilha
function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetMetricas = ss.getSheetByName("Metricas");
  const sheetPedidos = ss.getSheetByName("Pedidos - Livros") || ss.getSheets()[0];

  const metricas = sheetMetricas ? sheetMetricas.getDataRange().getValues().slice(1) : [];
  const pedidos = sheetPedidos ? sheetPedidos.getDataRange().getValues().slice(1) : [];

  const payload = {
    metricas: metricas.map(r => ({ data: r[0], origem: r[1], meio: r[2], campanha: r[3], evento: r[4] })),
    pedidos: pedidos.map(r => ({ id: r[0], data: r[1], status: r[2], bump: r[12], total: r[13] }))
  };

  return ContentService.createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
