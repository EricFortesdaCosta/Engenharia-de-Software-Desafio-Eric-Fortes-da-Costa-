function calcularSituacao() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName("engenharia_de_software");

  var rangeNotas = sheet.getRange("D4:F27");
  var rangeFaltas = sheet.getRange("B4:C27");
  var rangeSituacao = sheet.getRange("G4:G27");
  var rangeResultados = sheet.getRange("H4:H27");

  var notasValues = rangeNotas.getValues();
  var faltasValues = rangeFaltas.getValues();

  for (var i = 0; i < notasValues.length; i++) {
    var aluno = "A" + (4 + i);
    var notas = notasValues[i];
    var faltas = faltasValues[i][1];
    var totalAulas = 60;
    var percentualFaltas = (faltas / totalAulas) * 100;
    var somaNotas = notas[0] + notas[1] + notas[2];
    var media = somaNotas / 3;
    var situacao = "";

    // Arredondar a média para o próximo número inteiro (arredondamento para cima)
    media = Math.ceil(media);

    // Verificar se o aluno está reprovado por falta
    if (percentualFaltas > 25) {
      situacao = "Reprovado por Falta";
    } else {
      // Se não estiver reprovado por falta, verificar a situação pelas notas
      if (media < 50) {
        situacao = "Reprovado por Nota";
      } else if (media >= 50 && media < 70) {
        situacao = "Exame Final";

        // Calcular a Nota para Aprovação Final (naf)
        var naf = 2 * (5 - media); // Calculando a nota necessária para aprovação final
        var notaAprovacaoFinal = (media + naf) / 2;

        // Se a nota necessária para aprovação final for maior ou igual a 5, o aluno será aprovado
        if (notaAprovacaoFinal >= 5) {
          situacao = "Aprovado";
        }
      } else {
        situacao = "Aprovado";
      }
    }

    var situacaoCell = rangeSituacao.getCell(i + 1, 1);
    var resultadoCell = rangeResultados.getCell(i + 1, 1);

    situacaoCell.setValue(situacao);
    resultadoCell.setValue(media); // Colocar a média arredondada na coluna H

  }
}
