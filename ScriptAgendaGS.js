//testado
function testemanpulaData(){
    var data = "2020-10-16";
    var dias = 4;
    //var resultado = manipulaData(data,dias);
    var resultado = calculaDias(data,dias);
    console.log(resultado);
  }
  
  
  function manipulaData(data,dias){
    var splitD = data.split("-");
    var aaaa = splitD[0];
    var mmi = splitD[1] -1;
    var ddi = splitD[2];
    var inicial = new Date(aaaa, mmi, ddi);
    var final = new Date(inicial);
    final.setDate(final.getDate() +dias);
    
    
    var dd = ("0" + final.getDate()).slice(-2);
    var mm = ("0" + (final.getMonth() + 1)).slice(-2);
    var y = final.getFullYear();
    var newData = mm + "-" + dd + "-" + y;  
    return newData;
  }
  function testemanipulaHora(){
    var sh = SpreadsheetApp.getActiveSpreadsheet();
    var ssh = sh.getActiveSheet();
    var horario = ssh.getRange(9, 5).getValue();
    console.log(horario);
  
  }
  
  //testado
  function manipulaHora(horarios){
    var splitC = horarios.split(":");
    var hh = filterInt(splitC[0]);
    var mm = splitC[1];
    hh = hh +1;
    var newHora = hh + ":" + mm;
    return newHora;
  }
  
  function calculaDias(data,quantidade){
    var semanas = quantidade * 7;
    var datafinal = manipulaData(data,semanas);
    console.log(datafinal);
    return datafinal;
  }
  
  function agendarVarios(quantidade){
    
    var sh = SpreadsheetApp.getActiveSpreadsheet();
    var ssh = sh.getSheetByName('consultas');
    var row = ssh.getLastRow();
  
    var nomeP = ssh.getRange(row,3).getValue();
    var horario = ssh.getRange(row,5).getValue();
    var horarioFim = manipulaHora(horario);
    var diaTabela = ssh.getRange(row,4).getValue();
    var dia = manipulaData(diaTabela,0);
    var diaHoraI = dia + " " + horario;
    var diaHoraF = dia + " " + horarioFim;
    var diaSem = diaSemana(diaTabela);
    var datafinal = calculaDias(diaTabela,quantidade);
   
    
    
    switch (diaSem){
    case 'SUNDAY':
        var event = CalendarApp.getDefaultCalendar().createEventSeries(nomeP,
                                                  new Date(diaHoraI),
                                                  new Date(diaHoraF),
                                                  CalendarApp.newRecurrence().addWeeklyRule()
                                                  .onlyOnWeekdays([CalendarApp.Weekday.SUNDAY])
                                                  .until(new Date(datafinal)));
          break;
        case 'MONDAY':
        var event = CalendarApp.getDefaultCalendar().createEventSeries(nomeP,
                                                  new Date(diaHoraI),
                                                  new Date(diaHoraF),
                                                  CalendarApp.newRecurrence().addWeeklyRule()
                                                  .onlyOnWeekdays([CalendarApp.Weekday.MONDAY])
                                                  .until(new Date(datafinal)));
          break;
        case 'TUESDAY':
        var event = CalendarApp.getDefaultCalendar().createEventSeries(nomeP,
                                                  new Date(diaHoraI),
                                                  new Date(diaHoraF),
                                                  CalendarApp.newRecurrence().addWeeklyRule()
                                                  .onlyOnWeekdays([CalendarApp.Weekday.TUESDAY])
                                                  .until(new Date(datafinal)));
          break;
        case 'WEDNESDAY':
        var event = CalendarApp.getDefaultCalendar().createEventSeries(nomeP,
                                                  new Date(diaHoraI),
                                                  new Date(diaHoraF),
                                                  CalendarApp.newRecurrence().addWeeklyRule()
                                                  .onlyOnWeekdays([CalendarApp.Weekday.WEDNESDAY])
                                                  .until(new Date(datafinal)));
          break;
        case 'THURSDAY':
        var event = CalendarApp.getDefaultCalendar().createEventSeries(nomeP,
                                                  new Date(diaHoraI),
                                                  new Date(diaHoraF),
                                                  CalendarApp.newRecurrence().addWeeklyRule()
                                                  .onlyOnWeekdays([CalendarApp.Weekday.THURSDAY])
                                                  .until(new Date(datafinal)));
          break;
        case 'FRIDAY':
        var event = CalendarApp.getDefaultCalendar().createEventSeries(nomeP,
                                              new Date(diaHoraI),
                                                  new Date(diaHoraF),
                                                  CalendarApp.newRecurrence().addWeeklyRule()
                                                  .onlyOnWeekdays([CalendarApp.Weekday.FRIDAY])
                                                  .until(new Date(datafinal)));
          break;
        case 'SATURDAY':
         var event = CalendarApp.getDefaultCalendar().createEventSeries(nomeP,
                                                  new Date(diaHoraI),
                                                  new Date(diaHoraF),
                                                  CalendarApp.newRecurrence().addWeeklyRule()
                                                  .onlyOnWeekdays([CalendarApp.Weekday.SATURDAY])
                                                  .until(new Date(datafinal)));
          break;
        default:
        var x = false;
        console.log('error');
          break;
    }
    if(x !== false)
      var marcado = ssh.getRange(row,6).setValue("Agendado");
  
    Logger.log('Event ID: '+ ' dia da semana: ' + diaSem);
  }
  /*
  function testediaSemana() {
    var diaTabela = "2020-12-30";
    var diaSem = diaSemana(diaTabela);
    console.log(diaSem);
  }
  */
  function diaSemana (data){
  var semana = ["MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY", "SATURDAY", "SUNDAY"];
  var d = new Date(data);
  var ddSemana = semana[d.getDay()];
  return ddSemana;
  }
  
  function calcularVariosInsercoes(dados,quantidade){
    var sh = SpreadsheetApp.getActiveSpreadsheet();
    var shC = sh.getSheetByName('consultas');
    var shD = sh.getSheetByName('controleCliente');
    var match = rowMatch('controleCliente',[dados.Nome],2,3);
    var qnt = quantidade
    console.log(qnt);
    for(;qnt>0;qnt--){
    var lastRow = nxRow('consultas');
    if(match !== false){
        var id = shD.getRange(match, 2).getValue();
        var nome = shD.getRange(match, 3).getValue();
        shC.getRange(lastRow, 2).setValue(id) ; 
        shC.getRange(lastRow, 3).setValue(nome);
        shC.getRange(lastRow, 4).setValue([dados.Data]);
        shC.getRange(lastRow, 5).setValue([dados.Hora]);
    }
    gerarID('consultas');
    }
  }
  
  
  
  
  
  