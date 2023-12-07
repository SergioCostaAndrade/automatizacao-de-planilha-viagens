//
// classifica as colunas da planilha banco de dados
//
function onEdit(){
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var folha = planilha.getSheets()[0];
  var intervalo1 = folha.getRange('A4:A200');
  var intervalo2 = folha.getRange('B4:B200');
  var intervalo3 = folha.getRange('C4:C200');
  var intervalo4 = folha.getRange('F4:F200');
  intervalo1.sort({column:1, ascending:true});
  intervalo2.sort({column:2, ascending:true});
  intervalo3.sort({column:3, ascending:true});
  intervalo4.sort({column:6, ascending:true});
}
//
// Trata registros com marcação X na planilha Programação de viagens
//
// função que abre uma planilha pelo nome
function obterplanilhapelonome(nome) {
let sheet = SpreadsheetApp.getActiveSpreadsheet()
return sheet.getSheetByName(nome)
}
function Classificacolunas (){

let programacaodeviagens = obterplanilhapelonome('Programação de Viagens')
let ultimalinha = programacaodeviagens.getLastRow()
let geroupedido = 'n'
let valortotal = 0
let controlalinha = 0 
let controlagrupodeviagem = 3
let pedidodediaria = programacaodeviagens.getRange(1,7).getValue()
let pedidodediariamaisum = pedidodediaria + 1
// 
// Varrer as linhas do início ao fim procurando por um "X" ou "x"
//
for (var i = 5; i < ultimalinha;i++) {
// 
let marca = programacaodeviagens.getRange(i,1).getValue()
//
let dataatual = programacaodeviagens.getRange(1,5).getValue()
 //
 // Se coluna 1 marcado com 'X' ou 'x' adiciona qtde de diárias ao motorista e prepara formulário para envio de e-mail 
 //
 if (marca == 'X' || marca == 'x') {
  //
  // atribui valores de variaveis
  //
  let motorista = programacaodeviagens.getRange(i,11).getValue()
  let qtddiaria = programacaodeviagens.getRange(i,10).getValue()
  let valor = qtddiaria * programacaodeviagens.getRange(1,9).getValue()
  let roteiro = programacaodeviagens.getRange(i,5).getValue()
  let datainicio = programacaodeviagens.getRange(i,7).getValue()
  let datafim = programacaodeviagens.getRange(i,8).getValue()
  let justificativa = programacaodeviagens.getRange(i,12).getValue()
  //
  // totaliza o numero de diarias do motorista
  //
  somadiaria(motorista,qtddiaria)
  geroupedido = 's'
  //
  //
  preparaformulario(pedidodediariamaisum,dataatual,controlagrupodeviagem,motorista,qtddiaria,valor,roteiro,datainicio,datafim,justificativa,controlalinha)
  valortotal += valor
  //
  controlagrupodeviagem += 7 
  //
  programacaodeviagens.getRange(i,1).setValue('Feito')
  programacaodeviagens.getRange(i,9).setValue(pedidodediariamaisum)
  //
  }  else {
      if (marca == 'D' || marca == 'd' || marca == 'Feito') { 
           ultimalinha = ultimalinha
      }
      else {
         // 
      // encerra o loop
      //
        ultimalinha = i
      }
      
    }
    
}
if (geroupedido == 's') {
      programacaodeviagens.getRange(1,7).setValue(pedidodediariamaisum)
 //
 //   totaliza o valor final do pedido de diaria 
 //
 let solicitapagamentodiaria = obterplanilhapelonome('Solicita pagamento diaria')
 let ultimalinhapagamentodiaria = solicitapagamentodiaria.getLastRow()
 solicitapagamentodiaria.getRange(ultimalinhapagamentodiaria,73).setValue(valortotal+','+'00')
}
function somadiaria(motorista,qtddiaria) {
  // SpreadsheetApp.getUi().alert('dentro de somardiaria' + motorista + qtddiaria)
  
  let motoristasxdiarias = obterplanilhapelonome('Motoristas x Diárias')

  for (var s = 2; s < 50; s++) {
    let qtdeantigadiaria = motoristasxdiarias.getRange(s,2).getValue()

    if (motorista == motoristasxdiarias.getRange(s,1).getValue()){
    //  SpreadsheetApp.getUi().alert('motorista' + '  '  + motorista + 'somar' + qtddiaria + 'a' + ' ' + 
    //  qtdeantigadiaria)
      // atualiza o numero de diarias
      qtdeantigadiaria += qtddiaria
      motoristasxdiarias.getRange(s,2,1).setValue(qtdeantigadiaria)
    }
  } 
}
function preparaformulario(pedidodediariamaisum,dataatual,controlagrupodeviagem,motorista,qtddiaria,valor,roteiro,datainicio,datafim,justificativa) {
  let solicitapagamentodiaria = obterplanilhapelonome('Solicita pagamento diaria')
  let ultimalinhapagamentodiaria = solicitapagamentodiaria.getLastRow()
  if (controlalinha == 0) {
    controlalinha += 1
  } else {
    ultimalinhapagamentodiaria -= 1
  }
  solicitapagamentodiaria.getRange(ultimalinhapagamentodiaria+1,1).setValue(pedidodediariamaisum)
  solicitapagamentodiaria.getRange(ultimalinhapagamentodiaria+1,2).setValue(dataatual)
  solicitapagamentodiaria.getRange(ultimalinhapagamentodiaria+1,controlagrupodeviagem).setValue(motorista)
  solicitapagamentodiaria.getRange(ultimalinhapagamentodiaria+1,controlagrupodeviagem+1).setValue(qtddiaria)
  solicitapagamentodiaria.getRange(ultimalinhapagamentodiaria+1,controlagrupodeviagem+2).setValue(valor+','+'00')
  solicitapagamentodiaria.getRange(ultimalinhapagamentodiaria+1,controlagrupodeviagem+3).setValue(roteiro)
  solicitapagamentodiaria.getRange(ultimalinhapagamentodiaria+1,controlagrupodeviagem+4).setValue(datainicio)
  solicitapagamentodiaria.getRange(ultimalinhapagamentodiaria+1,controlagrupodeviagem+5).setValue(datafim)
  solicitapagamentodiaria.getRange(ultimalinhapagamentodiaria+1,controlagrupodeviagem+6).setValue(justificativa)
}
}
