function onFormSubmit(e) {
  const nomeRelatorio = "RELATÓRIO";
  const nomeFormulario = "Respostas ao formulário 3";
  
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const relatorio = spreadsheet.getSheetByName(nomeRelatorio);
  const formulario = spreadsheet.getSheetByName(nomeFormulario);
  
  if (!relatorio || !formulario) {
    throw new Error("Uma das abas (RELATÓRIO ou Respostas ao formulário 1) não foi encontrada.");
  }
  
  var lastRow = formulario.getLastRow();
  var resposta = formulario.getRange(lastRow, 2, 1, formulario.getLastColumn() - 1).getValues()[0];
  
  const mapaStatus = {
    "Normal": "A3",
    "Agitado": "F3",
    "Crítico!": "K3"
  };
  
  relatorio.getRange("A3").setValue(false);
  relatorio.getRange("F3").setValue(false);
  relatorio.getRange("K3").setValue(false);
  
  var statusTurno = resposta[0];
  if (mapaStatus[statusTurno]) {
    relatorio.getRange(mapaStatus[statusTurno]).setValue(true);
  }
  
  // =======================================================
  // Lógica para o Checklist de Setores (Índice 1)
  // Corrigido para dividir a string por vírgula
  // =======================================================
  var mapaSetores = {
    "Recepção": "F5", "Classificação / NIR": "F6", "SASV": "F7", "UTIP": "F8",
    "Medicação / Consultórios 1, 2, 3": "F9", "Farmácia UTIP": "F10", 
    "PPP / Retaguarda": "F11", "USG": "F12",
    "Recursos Humanos": "F13",
    "Medicina do Trabalho / Segurança do Trabalho / DP": "F14",
    "SAME": "F15",
    "Farmácia Central": "K5", "Lactário": "K6", "Centro Cirurgico / RPA / Farmácia CC": "K7",
    "Enfermarias A, B, C, D": "K8", "CCIH": "K9", "Faturamento": "K10",
    "Cordenacao Administrativa": "K11", "Diretorias": "K12", "Financeiro": "K13",
    "Compras": "K14", "Qualidade": "K15", "Custos": "K16"
  };
  
  for (var setor in mapaSetores) {
    relatorio.getRange(mapaSetores[setor]).setValue(true);
  }
  
  var setoresNaoVisitados = resposta[1]; 
  if (setoresNaoVisitados) {
    // Divide a string por vírgula e remove espaços extras
    var arrayDeSetores = setoresNaoVisitados.split(',').map(function(item) {
      return item.trim();
    });
    
    arrayDeSetores.forEach(function(setor) {
      if (mapaSetores[setor]) {
        relatorio.getRange(mapaSetores[setor]).setValue(false);
      }
    });
  }
  
  // Perguntas de SIM OU NÃO
  var respostaReclamacoes = resposta[3];
  relatorio.getRange("H22").setValue(respostaReclamacoes === "SIM");
  relatorio.getRange("J22").setValue(respostaReclamacoes === "NÃO");

  var respostaTroca = resposta[4];
  relatorio.getRange("H27").setValue(respostaTroca === "SIM");
  relatorio.getRange("J27").setValue(respostaTroca === "NÃO");

  var respostaAtendimentos = resposta[5];
  relatorio.getRange("H32").setValue(respostaAtendimentos === "SIM");
  relatorio.getRange("J32").setValue(respostaAtendimentos === "NÃO");

  var mapaTextos = {
    "A18:O20": resposta[2],
    "A24:O26": resposta[6],
    "A29:O31": resposta[7],
    "A33:O43": resposta[8]
  };

  for (var celula in mapaTextos) {
    relatorio.getRange(celula).setValue(mapaTextos[celula]);
  }

  var mapaTonners = {
    "P4": resposta[9], "P5": resposta[10], "P6": resposta[11],
    "P7": resposta[12], "P8": resposta[13], "P9": resposta[14],
    "P10": resposta[15], "P11": resposta[16], "P12": resposta[17],
    "P13": resposta[18], "P14": resposta[19], "P15": resposta[20],
    "P16": resposta[21]
  };
  
  for (var celula in mapaTonners) {
    var valor = mapaTonners[celula];
    if (valor !== null && valor !== "") {
      relatorio.getRange(celula).setValue(valor);
    }
  }
}
