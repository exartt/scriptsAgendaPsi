
function formConsulta() {
    var starta = autoInc();
    var clientes = starta;
    var form = HtmlService.createTemplateFromFile("Consultas");
    form.clientes = clientes.map(
      function(r2){
      return r2[0];
      });
    var mostraForm = form.evaluate();
    mostraForm.setTitle("Marcar Consulta").setHeight(250).setWidth(350);
    SpreadsheetApp.getUi().showModalDialog(mostraForm, "Cadastro Consulta");
  }
  
  function verificaControle(ID){
    
    if(isMatch(ID,2, 'controleCliente') !== false){
     var match = isMatch(ID,2, 'controleCliente');
      return match;
    }else if(isMatch(ID,3, 'controleCliente') !== false){
     var match = isMatch(ID,3,'controleCliente');
      return match;
    }else{
  return false;
    }
  }
  
  function verificadorDuplicidade(dados){
    
    if(isMatch([dados.Data],4,'consultas') !== false &&
       isMatch([dados.Nome],3,'consultas') !== false){
      return true;
    }else{
      return false;
    }
    
  }
  
  function insereConsulta(dados,quantidade){
  var sh = SpreadsheetApp.getActiveSpreadsheet();
    var shC = sh.getSheetByName('consultas');
    var shD = sh.getSheetByName('controleCliente');
    var lastRow = nxRow('consultas');
    var match = rowMatch('controleCliente',[dados.Nome],2,3);
    console.log('nome: ' + [dados.Nome]);
    console.log('rowMatch: ' + match);
    if(match !== false){
        var id = shD.getRange(match, 2).getValue();
        var nome = shD.getRange(match, 3).getValue();
        shC.getRange(lastRow, 2).setValue(id) ; 
        shC.getRange(lastRow, 3).setValue(nome);
        shC.getRange(lastRow, 4).setValue([dados.Data]);
        shC.getRange(lastRow, 5).setValue([dados.Hora]);
      gerarID('consultas');
      retornaDataBR();
      Datavalidation();
      selecao(quantidade);
    }else{
    Browser.msgBox("Erro no agendamento", "Não foi possivel achar o controle desse cliente, verifique se ele já foi criado.", Browser.Buttons.OK);
    
    }
    
  }
  function testes(){
  var teste = 4;
    var x = selecao(teste);
    console.log(x);
  }
  function selecao(quantidade){
    console.log(quantidade);
    var qnt = filterInt(quantidade);
    console.log(qnt);
    if(qnt === 0){
      agendarVarios(quantidade);
    }else if(qnt !== 0){
      agendarVarios(quantidade);
      repeteFuncao(qnt);
    }
    
  }
  
  function _hyperlink(mySheet){
  var sh = SpreadsheetApp.getActiveSpreadsheet();
  var shss = sh.getSheetByName(mySheet);
  SpreadsheetApp.setActiveSheet(shss);
  }
  
  function btnconsulta(){
    _hyperlink('consultas');
    formConsulta();
  }
  
  
  function repeteFuncao(quantidade){
    var sh = SpreadsheetApp.getActiveSpreadsheet();
    var ssh = sh.getSheetByName('consultas');
    var contador = quantidade-1;
    var diasAtuais=0;
    var row = ssh.getLastRow();
      for(;contador !== 0 ;contador--){
         diasAtuais = diasAtuais +7; 
        var dataAtual = ssh.getRange(row, 4).getValue();
        dataAtual = manipulaData(dataAtual,diasAtuais);
        console.log(dataAtual);
        var x = insereConsultaMultiplas(dataAtual);
        console.log(x);
      }
  }
  
  function insereConsultaMultiplas(dataAtual){
  var sh = SpreadsheetApp.getActiveSpreadsheet();
    var shC = sh.getSheetByName('consultas');
    var lastRow = nxRow('consultas');
    var row = shC.getLastRow();
    var spreadsheet = SpreadsheetApp.getActive();
    spreadsheet.getRange('B'+lastRow+':C'+lastRow).activate();
    spreadsheet.getRange('B'+row+':C'+row).copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    shC.getRange(lastRow, 4).setValue(dataAtual);
    spreadsheet.getRange('E'+lastRow).activate();
    spreadsheet.getRange('E'+row).copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    shC.getRange(lastRow, 6).setValue('Agendado');
    var x = shC.getRange(row, 8).getValue();
    shC.getRange(lastRow,8).setValue(x);
    Datavalidation();
    gerarID('consultas');
  }
  
  //att
  function baixarConsulta(){
    var sh = SpreadsheetApp.getActiveSpreadsheet();
    var ssh = sh.getSheetByName('consultas');
    var csh = sh.getSheetByName('controleCliente');
    var row = ssh.getLastRow();
    console.log(row);
    //ssh.getRange(row, 7).getValue() !== 'Atendido&Contabilizado'
    for(;row>1;row--){
      var x = ssh.getRange(row,7).getValue();
      var y = ssh.getRange(row,2).getValue();
      console.log('valor do y = '+y);
      console.log('valor do x = '+x);
      switch (x){
        case 'Atendido':
       var linha = isMatch(y,2,'controleCliente');
          console.log('linha '+linha);
          if(csh.getRange(linha, 9).getValue() === ''){
            var novoValor = 1;
            
            }else{
              var novoValor = filterInt(csh.getRange(linha, 9).getValue())+1;
            }
          csh.getRange(linha, 9).setValue(novoValor);
          console.log('novo valor: '+novoValor);
          ssh.getRange(row, 7).setValue('Atendido&Contabilizado');
          var nx = nxRow('HCAT');
          copiaLinha('consultas',row,nx);
            break;
          default:
          console.log('erro');
          break;
      }
     } 
    }
  //att
  function copiaLinha(plan,row,nxrow){
    var sh  = SpreadsheetApp.getActiveSpreadsheet();
    var shp = sh.getSheetByName(plan);
    //var x = shp.getRange(row, col).getValue();
    sh.getRange('consultas!A'+row+':I'+row).copyTo(sh.getRange('HCAT!A'+nxrow+':I'+nxrow), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    removeLinhas('consultas',row);
  }
  //att
  function removeLinhas(plan,row){
    var sh  = SpreadsheetApp.getActiveSpreadsheet();
    var shp = sh.getSheetByName(plan);
    shp.deleteRow(row);
  }
  
  
  function marcaBaixasPagas(id){
    var sh = SpreadsheetApp.getActiveSpreadsheet();
    var ssh = sh.getSheetByName('controleCliente');
    var row = ssh.getLastRow();
    var acima = shh.getRange('i'+(row-1)).getValue();
    var posicaoAtual = shh.getRange('i'+row).getValue();
    
  }
  
  
  