/**
 * Função para Exportar ao Google Calendar
 */
function calExport() {
    var sheet = SpreadsheetApp.getActiveSheet();
    /**
     * Carregar a célula o número de linhas a serem processadas
     */
    var cellNum = sheet.getRange("G1");
    /**
     * Linha onde se iniciará a exportação dos Dados
     */
    var startRow = 2;
    /**
     * Número de linhas a serem processadas
     */
    var numRows = cellNum.getValue(); // Number of rows to process
    /**
     * Range dos Dados
     */
    var dataRange = sheet.getRange(startRow, 1, numRows, 7);
    var data = dataRange.getValues();
    /**
     * String de compartilhamento do Calendário
     */
    //var cal = CalendarApp.getDefaultCalendar();  
    //var cal = CalendarApp.getCalendarById("pbhbo7sb0n74q4u9rmf6q69gmc@group.calendar.google.com");
    var cal = CalendarApp.getCalendarById("37asl18t7bdukms55e3pupgt0o@group.calendar.google.com");
    var numero = 1;
    /**
     * Iterando para buscar as informações e lançar no Google Calendar
     */
    for (i in data) {
        var row = data[i];
        numero = numero + 1;
        if (row[6] == "") {
            var title = row[0] + '-' + row[3];
            var desc = 'Descrição:' + row[4] + '\nTotal:' + row[5];
            var tstart = row[1];
            var tstop = row[2];

            /**
             * Criando o evento
             */
            cal.createEvent(title, tstart, tstop,{description:desc});
            //cal.createAllDayEvent(title, tstart,  {description:desc});

            var rangeStr = 'G' + numero;
            var cell = sheet.getRange(rangeStr);
            cell.setValue("Lançado");
        }
    }
}

/**
 * Função para adicionar o Menu a Planilha Google
 */
function onOpen() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var menuEntries = [{
        name: "Enviar para Google Calendar",
        functionName: "calExport"
    }];
    ss.addMenu("Integrar", menuEntries);
}