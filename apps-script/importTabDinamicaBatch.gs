/**
 * IMPORTAÇÃO DA TABELA DINÂMICA — versão mais rápida
 *
 * Como aplicar no Apps Script:
 * 1. Abra o editor, procure a função "importTabDinamicaBatch" e apague ela
 *    INTEIRA (do "function importTabDinamicaBatch(data) {" até o "}" final,
 *    logo antes do comentário "Recebe um lote (batch) de linhas e SUBSTITUI").
 * 2. No mesmo lugar, cole as duas funções abaixo.
 * 3. Salve e publique uma nova versão (Implantar > Gerenciar implantações >
 *    editar > Nova versão). Pra desfazer, é só voltar pra versão anterior.
 *
 * O que mudou (o resultado na planilha é o mesmo de antes):
 * - Antes, cada pedido que já existia fazia 3 gravações separadas na planilha
 *   (Status Entrega, Receb Cliente, Previsão Transp) — um lote de 2.000
 *   pedidos eram ~6.000 gravações, o que deixava a importação lenta e às
 *   vezes estourava o tempo limite. Agora as 3 colunas são lidas uma vez,
 *   alteradas na memória e gravadas de volta de uma vez só (3 gravações por
 *   lote, não importa quantos pedidos).
 * - Usa a mesma trava (lock) das outras importações, pra duas importações ao
 *   mesmo tempo não se atropelarem.
 * - Regras iguais às de antes: pedido já existente atualiza só essas 3
 *   colunas; pedido novo é adicionado no fim; a aba nunca é limpa.
 */
function importTabDinamicaBatch(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    return importTabDinamicaBatchInterno(data);
  } finally {
    lock.releaseLock();
  }
}

function importTabDinamicaBatchInterno(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(TAB_DINAMICA_SHEET);
  if (!sheet) sheet = ss.insertSheet(TAB_DINAMICA_SHEET);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, TAB_DINAMICA_HEADERS.length).setValues([TAB_DINAMICA_HEADERS]);
  }

  const rows = data.rows || [];
  let novos = 0;
  let atualizados = 0;

  if (rows.length > 0) {
    // Índice Pedido -> posição na lista (0 = linha 2 da planilha), construído
    // a partir do que já existe. Se o pedido aparecer repetido na aba, vale a
    // última ocorrência (igual à versão anterior).
    const numExistentes = sheet.getLastRow() - 1;
    const pedidoParaIndice = Object.create(null);
    let colStatus = [];
    let colPrevisao = [];
    let colReceb = [];
    if (numExistentes > 0) {
      const pedidosExistentes = sheet.getRange(2, TAB_COL_PEDIDO + 1, numExistentes, 1).getValues();
      pedidosExistentes.forEach((p, i) => { if (p[0]) pedidoParaIndice[p[0]] = i; });
      colStatus = sheet.getRange(2, TAB_COL_STATUS_ENTREGA + 1, numExistentes, 1).getValues();
      colPrevisao = sheet.getRange(2, TAB_COL_PREVISAO_TRANSP + 1, numExistentes, 1).getValues();
      colReceb = sheet.getRange(2, TAB_COL_RECEB_CLIENTE + 1, numExistentes, 1).getValues();
    }

    const novasLinhas = [];
    rows.forEach(row => {
      const pedido = row[TAB_COL_PEDIDO];
      const i = pedido ? pedidoParaIndice[pedido] : undefined;
      if (i !== undefined) {
        // "Previsão de Entrega (Transp)" também é atualizada: ela costuma vir
        // vazia na primeira vez que o pedido aparece e só é preenchida numa
        // importação seguinte.
        colStatus[i][0] = row[TAB_COL_STATUS_ENTREGA];
        colReceb[i][0] = row[TAB_COL_RECEB_CLIENTE];
        colPrevisao[i][0] = row[TAB_COL_PREVISAO_TRANSP];
        atualizados++;
      } else {
        novasLinhas.push(row);
      }
    });

    if (atualizados > 0) {
      sheet.getRange(2, TAB_COL_STATUS_ENTREGA + 1, numExistentes, 1).setValues(colStatus);
      sheet.getRange(2, TAB_COL_RECEB_CLIENTE + 1, numExistentes, 1).setValues(colReceb);
      sheet.getRange(2, TAB_COL_PREVISAO_TRANSP + 1, numExistentes, 1).setValues(colPrevisao);
    }

    if (novasLinhas.length > 0) {
      const startRow = sheet.getLastRow() + 1;
      sheet.getRange(startRow, 1, novasLinhas.length, TAB_DINAMICA_HEADERS.length).setValues(novasLinhas);
    }
    novos = novasLinhas.length;
  }

  if (data.isLast) {
    setImportMeta(TAB_DINAMICA_SHEET, sheet.getLastRow() - 1);
    // A reconstrução da mini-tabela (TabDinamicaLookup) continua separada: o
    // site chama a ação "reconstruirTabDinamicaLookup" logo depois.
  }

  return { ok: true, sheet: TAB_DINAMICA_SHEET, batchIndex: data.batchIndex, novos, atualizados, isLast: !!data.isLast };
}
