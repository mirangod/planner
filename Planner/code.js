// Matheus C. @mirangod

// #region GLOBAL VAR
const ss = SpreadsheetApp.getActiveSpreadsheet();
const sheet1 = ss.getSheetByName("Planner Diário");
const sheet2 = ss.getSheetByName("DataBase");

// database range
const range1 = sheet2.getDataRange().getValues();

// planner range
const range2 = sheet1.getRange("B2:C33").getValues();
// prioridades range
const range3 = sheet1.getRange("I2:J11").getValues();
// notas range
const range4 = sheet1.getRange("I13:J22").getValues();
// lembretes range
const range5 = sheet1.getRange("I24:J33").getValues();
// #endregion

/**
 * Função que será chamada sempre que houver uma edição na planilha e que será responsável por verificar haverá uma exportação ou importação.
 * Function that will be called whenever there is an edit to the spreadsheet and that will be responsible for checking whether there will be an export or import.
 */
function Planner() {
  try {
    // a1 é o id que o usuário vai alterar.
    var a1 = Utilities.formatDate(sheet1.getRange("B1").getValue(),"GMT-03:00","dd/MM/yyyy");
    // a2 é o valor comparativo sempre que houver uma edição.
    var a2 = Utilities.formatDate(sheet1.getRange("A2").getValue(),"GMT-03:00","dd/MM/yyyy");
    console.log("Valor do ID: " + a1 + ", valor anterior: " + a2);

    if (a1!==a2) {
      sheet1.getRange("A2").setValue(a1);
      Export(a1);
      return;
    } else {
      Import(a1, sheet1.getRange("B2:H33").getValues(), sheet1.getRange("I2:N11").getValues(), sheet1.getRange("I13:N22").getValues(), sheet1.getRange("I24:N33").getValues());
      return;
    }
  } catch (erro) {
    ErrorAlert(erro);
  }
}

/**
 * Função de importar os dados no Banco de Dados.
 * Function to import data into the Database.
 * @param {date} id - Data que será utilizada como indicador no banco de dados.
 * @param {range} planner - Intervalo do Planner.
 * @param {range} prioridades - Intervalo de Prioridades.
 * @param {range} notas - Intervalo de Notas.
 * @param {range} lembretes - Intervalo do Lembretes.
 */
function Import(id, planner, prioridades, notas, lembretes) {
  try {
    console.log("Importação iniciada!");
    var text = "";
    // Objeto que será 'alimentado'
    var hungry = {
      PLANNER:     Feed(planner, text),
      PRIORIDADES: Feed(prioridades, text),
      NOTAS:       Feed(notas, text),
      LEMBRETES:   Feed(lembretes,text)
    }
    for (var i = 1; i < range1.length; i++) {
      if (Utilities.formatDate(range1[i][0],"GMT-03:00","dd/MM/yyyy") === id) {
        // Utilizar do objeto alimentado para preencher os valores no banco de dados.
        sheet2.getRange(i+1,2).setValue(hungry.PLANNER);
        sheet2.getRange(i+1,3).setValue(hungry.PRIORIDADES);
        sheet2.getRange(i+1,4).setValue(hungry.NOTAS);
        sheet2.getRange(i+1,5).setValue(hungry.LEMBRETES);
        console.log("Valor encontrado e preenchido!");
        break;
      }
    }
    console.log("Importação finalizada!");
  } catch (erro) {
    ErrorAlert(erro);
  }
}

/**
 * Função para exportar dados no Banco de Dados.
 * Function to export data in the Database.
 * @param {date} id - Data que será utilizada como indicador no banco de dados. 
 */
function Export(id) {
  try {
    console.log("Exportação iniciada!");
    for (var i = 1; i < range1.length; i++) {
      if (Utilities.formatDate(range1[i][0],"GMT-03:00","dd/MM/yyyy") === id) {
        // Objeto que será 'alimentado'
        var baby = {
          PLANNER:     Fed(sheet2.getRange(i+1,2).getValue()),
          PRIORIDADES: Fed(sheet2.getRange(i+1,3).getValue()),
          NOTAS:       Fed(sheet2.getRange(i+1,4).getValue()),
          LEMBRETES:   Fed(sheet2.getRange(i+1,5).getValue())
        }
        // #region Seção de alimentação na planilha original
        for (var j = 0; j < range2.length*2; j++) {
          if (j%2===0) {
            // PAR
            sheet1.getRange(2+Math.floor(j/2),2).setValue(baby.PLANNER[j]);
          } else {
            // ÍMPAR
            sheet1.getRange(2+Math.floor(j/2),3).setValue(baby.PLANNER[j]);
          }
        }
        for (var k = 0; k < range3.length*2; k++) {
          if (k%2===0) {
            // PAR
            sheet1.getRange(2+Math.floor(k/2),9).setValue(baby.PRIORIDADES[k]);
          } else {
            // ÍMPAR
            sheet1.getRange(2+Math.floor(k/2),10).setValue(baby.PRIORIDADES[k]);
          }
        }
        for (var l = 0; l < range4.length*2; l++) {
          if (l%2===0) {
            // PAR
            sheet1.getRange(13+Math.floor(l/2),9).setValue(baby.NOTAS[l]);
          } else {
            // ÍMPAR
            sheet1.getRange(13+Math.floor(l/2),10).setValue(baby.NOTAS[l]);
          }
        }
        for (var m = 0; m < range5.length*2; m++) {
          if (m%2===0) {
            // PAR
            sheet1.getRange(24+Math.floor(m/2),9).setValue(baby.LEMBRETES[m]);
          } else {
            // ÍMPAR
            sheet1.getRange(24+Math.floor(m/2),10).setValue(baby.LEMBRETES[m]);
          }
        }
        // #endregion
        console.log("Valor encontrado para exportação!");
        break;
      }
    }
    console.log("Exportação concluída!");
  } catch(erro) {
    ErrorAlert(erro);
  }
}

/**
 * Função de alimentação no objeto hungry.
 * Feeding function on the hungry object.
 * @param {range} meal - Intervalo dos valores a serem preenchidos no objeto hungry
 * @param {string} hungry - String do objeto que será alimentado
 */
function Feed(meal, hungry) {
  try {
    console.log("Feed iniciado!");
    hungry = "";
    for (var i = 0; i < meal.length; i++) {
      if (meal[i][1] !== "") {
        // Cria um delimitador ';' para cada célula preenchida
        hungry += meal[i][0] + ";" + meal[i][1] + ";";
      }
    }
    console.log("Retorno da função: " + hungry);
    console.log("Feed concluído!");
    return hungry;
  } catch(erro) {
    ErrorAlert(erro);
  }
}

/**
 * Função que extrai o valor preenchido no banco de dados e retira os delimitadores para inserir em um array.
 * Function that extracts the value filled in the database and removes the delimiters to insert into an array.
 * @param {range} meal - Intervalo do banco de dados onde comporta o valor com base no ID.
 */
function Fed(meal) {
  try {
    console.log("Fed iniciado!");
    // Array onde vai ficar os valores encontrados no banco de dados
    var array = new Array();
    var cont = 0;
    for (var i = 0; i < meal.length; i++) {
      if (meal.charAt(i) === ";") {
        // Extrai o valor
        var substring = meal.substring(cont,i);
        // Insere o valor encontrado no array
        array.push(substring);
        cont = i + 1;
      }
    }
    console.log("Retorno da função: " + array);
    console.log("Fed concluído!");
    return array;
  } catch(erro) {
    ErrorAlert(erro);
  }
}

/**
 * Função que retorna um erro quando acontece e retorna um alerta para o usuário.
 * Function that returns an error when it happens and returns an alert to the user.
 * @param {object} erro - Erro enviado pelo console.
 */
function ErrorAlert(erro) {
  Logger.log("Ocorreu o seguinte erro: " + erro.toString());
  // Envia um e-mail com o erro informado
  MailApp.sendEmail({
    to: "erro@ativanautica.com.br",
    subject: "Erro: " + SpreadsheetApp.getActiveSpreadsheet().getName(),
    body: "Ocorreu o seguinte erro: " + erro.toString()
  });
  Logger.log("Emails restantes: " + MailApp.getRemainingDailyQuota());
}
