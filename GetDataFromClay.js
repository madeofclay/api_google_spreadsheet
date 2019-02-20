function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Clay')
      .addItem('Actualizar Facturas','getDTE')
      .addToUi()
}

function callAPIDTE(offset){
  offset = !offset?0:offset;
  var conf = SpreadsheetApp.getActive().getSheetByName('Configuración');
  var token = conf.getRange(3,2).getValue();
  var rut_analisis = conf.getRange(8,2).getValue();
  var dv_analisis = conf.getRange(8,3).getValue();
  var start_handler = conf.getRange(4,2).getValue().getTime();  
  var start = Utilities.formatDate(new Date(start_handler), "GMT+1", "YYYY-MM-dd");
  var end_handler = conf.getRange(6,2).getValue().getTime();  
  var end = Utilities.formatDate(new Date(end_handler), "GMT+1", "YYYY-MM-dd");
  var uri = 'https://api.clay.cl/v1/obligaciones/documentos_tributarios/';
  var response = UrlFetchApp.fetch(uri+'?rut_empresa='+rut_analisis+'&dv_empresa='+dv_analisis+'&fecha_desde='+start+'&fecha_hasta='+end+'&recibida=false&offset='+offset, {
    headers: {
      Token: token
    }
  });
  var data = JSON.parse(response);
  return data;
}

function getDTE(offset) {
  var dte = SpreadsheetApp.getActive().getSheetByName('DTE');
  dte.getRange("A2:K").clearContent(); //limpiamos todo antes de volver a cargar
  var conf = SpreadsheetApp.getActive().getSheetByName('Configuración');  
  
  var initial = callAPIDTE(0); //hacemos la primera llamada para obtener el total_records
  if(initial.status == false){
    dte.getRange(2,1).setValue(initial.message); //si algo falla ponemos el mensaje en pantalla
  }
  else{
    var total_records = parseInt(initial.data.records.total_records);
    conf.getRange(7,2).setValue(total_records); //lo mostramos en la página de configuración

    var cursor = 50;
    var i = 0, index = 0;
    while(cursor*i <= total_records){ //de 50 en 50 hasta llegar al total_records
      var data = callAPIDTE(cursor*i);
      for (var key in data.data.items) {
        if(!data.data.items[parseInt(key)].pagado){ //solo mostramos los que no están pagados
          var pos = index + 2;
//          Logger.log(pos + ' va el valor ' + data.data.items[parseInt(key)].numero + ' con index ' + index);
        
          dte.getRange(pos,1).setValue(data.data.items[parseInt(key)].tipo); //col 1
          dte.getRange(pos,2).setValue(data.data.items[parseInt(key)].codigo); //col 2
          dte.getRange(pos,3).setValue(data.data.items[parseInt(key)].numero);
          dte.getRange(pos,4).setValue(data.data.items[parseInt(key)].fecha_humana_emision);
      
          var d = new Date(data.data.items[parseInt(key)].fecha_emision*1000); //pasamos del timestamp al año / mes
          dte.getRange(pos,5).setValue(d.getFullYear()+'-'+("0" + (d.getMonth() + 1)).slice(-2)); //formato 2018-01
          dte.getRange(pos,6).setValue(d.getFullYear()+''+("0" + (d.getMonth() + 1)).slice(-2)+''+("0" + (d.getDay() + 1)).slice(-2)); //formato 20180101
      
          dte.getRange(pos,7).setValue(data.data.items[parseInt(key)].total.total);
          dte.getRange(pos,8).setValue(data.data.items[parseInt(key)].saldo_isoluto);
          dte.getRange(pos,9).setValue(data.data.items[parseInt(key)].total.exento);
          dte.getRange(pos,10).setValue(data.data.items[parseInt(key)].total.impuesto);
          dte.getRange(pos,11).setValue(data.data.items[parseInt(key)].total.otros_impuestos);
          index = index + 1;
        }
      }
      i = i + 1; //siguiente página (offset)
    }
  }
}
