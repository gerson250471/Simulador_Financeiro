// Remova a linha global antiga e use esta função
function getSpreadsheet() {
  try {
    return SpreadsheetApp.openById('1Ovo-boAo7Zp38z9EJb5g3_HX1lAJIjI37rz1VvddiwY');
  } catch (e) {
    throw new Error("Não foi possível acessar a planilha. Verifique o ID e as permissões.");
  }
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Simulador LQL Consig')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getDadosSimulacao() {
  const ss = getSpreadsheet(); // Chama a planilha de forma segura
  const sheet = ss.getSheetByName("Simulador");
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  return data.slice(1).map(row => ({
    id: row[0],
    banco: row[1],
    parcelas: row[2],
    fator: row[3],
    tipoTabela: row[4]
  }));
}

function getOpcoesConvenio() {
  const dados = getDadosSimulacao();
  return dados.map(item => {
    const fatorFormatado = typeof item.fator === 'number' ? 
                           item.fator.toLocaleString('pt-BR', {minimumFractionDigits: 5}) : 
                           item.fator;

    return {
      texto: `${item.banco} | ${item.tipoTabela} | ${item.parcelas}x | ${fatorFormatado}`,
      fator: item.fator,
      parcelas: item.parcelas
    };
  });
}