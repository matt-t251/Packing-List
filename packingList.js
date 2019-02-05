/** @OnlyCurrentDoc */

function confirmOrder() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1').activate();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.NEXT).activate();
  spreadsheet.getCurrentCell().offset(0, 1).setValue('0');
  
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Variants'), true);
  spreadsheet.getRange('A1').activate();
  spreadsheet.getRange('\'Enter Order\'!C1:Z1').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  
  // make Arrays
  
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Variants'), true);
  var productArrayT = spreadsheet.getRange('A1:M1').getValues();
  var productArray = productArrayT[0];
  
  // make an array of cell address'.
  var cellArray = ['A3', 'B3', 'C3', 'D3', 'E3', 'F3', 'G3', 'H3', 'I3','J3', 'K3', 'L3', 'M3', 'N3', 'O3', 'P3', 'Q3', 'R3', 'S3', 'T3', 'U3', 'V3','W3', 'X3', 'Y3', 'Z3'];
  var cellArrayPos = 0;
  var currentListCell = cellArray[cellArrayPos];
  
  var productArrayLength = productArray.length;
  
  var product = productArray[0];
  var i = 0;
  for (i = 0; i <= productArrayLength; i++) {
    
    var PAT = i + 1;
    
    if(product == '1') { // Product 1
      spreadsheet.getRange(currentListCell).activate();                                                                             // go to start location.
      spreadsheet.getRange('Extra!AB1:AC2').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false); //copy extra and paste to variant sheet.
      // create dropdown 1.
      spreadsheet.getRange(currentListCell).offset(2, 0).setDataValidation(SpreadsheetApp.newDataValidation().setAllowInvalid(true).requireValueInRange(spreadsheet.getRange('Extra!$AB$3:$AB$4'), true).build());
      // create dropdown 2.
      spreadsheet.getRange(currentListCell).offset(2, 1).setDataValidation(SpreadsheetApp.newDataValidation().setAllowInvalid(true).requireValueInRange(spreadsheet.getRange('Extra!$AC$3:$AC$4'), true).build());
      var cellArrayPos = cellArrayPos + 2;
      var currentListCell = cellArray[cellArrayPos];
    };
    
    if(product == '2') { // Product 2
      spreadsheet.getRange(currentListCell).activate();
      spreadsheet.getRange('Extra!AI1:AJ2').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
      spreadsheet.getRange(currentListCell).offset(2, 0).setDataValidation(SpreadsheetApp.newDataValidation().setAllowInvalid(true).requireValueInRange(spreadsheet.getRange('Extra!$AI$3:$AI$4'), true).build());
      spreadsheet.getRange(currentListCell).offset(2, 1).setDataValidation(SpreadsheetApp.newDataValidation().setAllowInvalid(true).requireValueInRange(spreadsheet.getRange('Extra!$AJ$3:$AJ$4'), true).build());
      var cellArrayPos = cellArrayPos + 2;
      var currentListCell = cellArray[cellArrayPos];
    };
    
    if(product == '3') { // Product 3
      spreadsheet.getRange(currentListCell).activate();
      spreadsheet.getRange('Extra!AP1:AR2').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
      spreadsheet.getRange(currentListCell).offset(2, 0).setDataValidation(SpreadsheetApp.newDataValidation().setAllowInvalid(true).requireValueInRange(spreadsheet.getRange('Extra!$AP$3:$AP$4'), true).build());
      spreadsheet.getRange(currentListCell).offset(2, 1).setDataValidation(SpreadsheetApp.newDataValidation().setAllowInvalid(true).requireValueInRange(spreadsheet.getRange('Extra!$AQ$3:$AQ$4'), true).build());
      spreadsheet.getRange(currentListCell).offset(2, 2).setDataValidation(SpreadsheetApp.newDataValidation().setAllowInvalid(true).requireValueInRange(spreadsheet.getRange('Extra!$AR$3:$AR$6'), true).build());
      var cellArrayPos = cellArrayPos + 3;
      var currentListCell = cellArray[cellArrayPos];
    };
    
    if(product == '4') { // Product 4
      spreadsheet.getRange(currentListCell).activate();
      spreadsheet.getRange('Extra!AX1:AX2').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
      spreadsheet.getRange(currentListCell).offset(2, 0).setDataValidation(SpreadsheetApp.newDataValidation().setAllowInvalid(true).requireValueInRange(spreadsheet.getRange('Extra!$AX$3:$AX$5'), true).build());
      var cellArrayPos = cellArrayPos + 1;
      var currentListCell = cellArray[cellArrayPos];
    };
    
    if(product == '5') { // Product 5
      spreadsheet.getRange(currentListCell).activate();
      spreadsheet.getRange('Extra!BD1:BE2').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
      spreadsheet.getRange(currentListCell).offset(2, 0).setDataValidation(SpreadsheetApp.newDataValidation().setAllowInvalid(true).requireValueInRange(spreadsheet.getRange('Extra!$BD$3:$BD$6'), true).build());
      spreadsheet.getRange(currentListCell).offset(2, 1).setDataValidation(SpreadsheetApp.newDataValidation().setAllowInvalid(true).requireValueInRange(spreadsheet.getRange('Extra!$BE$3:$BE$6'), true).build());
      var cellArrayPos = cellArrayPos + 1;
      var currentListCell = cellArray[cellArrayPos];
    };
    
    if(product == '7') { // Product 7
      spreadsheet.getRange(currentListCell).activate();
      spreadsheet.getRange('Extra!M1:O2').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
      spreadsheet.getRange(currentListCell).offset(2, 0).setDataValidation(SpreadsheetApp.newDataValidation().setAllowInvalid(true).requireValueInRange(spreadsheet.getRange('Extra!$M$3:$M$4'), true).build());
      spreadsheet.getRange(currentListCell).offset(2, 1).setDataValidation(SpreadsheetApp.newDataValidation().setAllowInvalid(true).requireValueInRange(spreadsheet.getRange('Extra!$N$3:$N$4'), true).build()); 
      spreadsheet.getRange(currentListCell).offset(2, 2).setDataValidation(SpreadsheetApp.newDataValidation().setAllowInvalid(true).requireValueInRange(spreadsheet.getRange('Extra!$O$3:$O$4'), true).build());
      var cellArrayPos = cellArrayPos + 3;
      var currentListCell = cellArray[cellArrayPos];
    };
    
    if(product == '11') { // Product 11
      spreadsheet.getRange(currentListCell).activate();
      spreadsheet.getRange('Extra!U1:V2').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
      spreadsheet.getRange(currentListCell).offset(2, 0).setDataValidation(SpreadsheetApp.newDataValidation().setAllowInvalid(true).requireValueInRange(spreadsheet.getRange('Extra!$U$3:$U$6'), true).build());
      spreadsheet.getRange(currentListCell).offset(2, 1).setDataValidation(SpreadsheetApp.newDataValidation().setAllowInvalid(true).requireValueInRange(spreadsheet.getRange('Extra!$V$3:$V$5'), true).build());
      var cellArrayPos = cellArrayPos + 2;
      var currentListCell = cellArray[cellArrayPos];
    };
    
    if(product == '0') {
      spreadsheet.getRange(currentListCell).setValue('END');
      var cellArrayPos = cellArrayPos + 1;
      var currentListCell = cellArray[cellArrayPos];
    };
    
    if(product == '') {
      spreadsheet.getRange(currentListCell).setValue('');
      var cellArrayPos = cellArrayPos + 1;
      var currentListCell = cellArray[cellArrayPos];
    };
    
    var product = productArray[PAT];
  };
};

function confirmVariants() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Variants'), true);
  var titleArrayT = spreadsheet.getRange('A3:Z3').getValues();
  var titleArray = titleArrayT[0];
  Logger.log(titleArray);
  var variantArrayT = spreadsheet.getRange('A5:Z5').getValues();
  var variantArray = variantArrayT[0];
  Logger.log(variantArray);
  var titleArrayLength = titleArray.length;
  Logger.log(titleArrayLength);
  
  var currentListCellArray = ['A1',  'A2',  'A3',  'A4',  'A5',  'A6',  'A7',  'A8',  'A9',  'A10',  'A11',  'A12',  'A13',  'A14',  'A15',  'A16',  'A17',  'A18',  'A19',  'A20',  'A21',  'A22',  'A23',  'A24',  'A25',  'A26',  'A27',  'A28',  'A29',  'A30',  'A31',  'A32',  'A33',  'A34',  'A35',  'A36',  'A37',  'A38',  'A39',  'A40',  'A41',  'A42',  'A43',  'A44',  'A45',  'A46',  'A47',  'A48',  'A49',  'A50',  'A51',  'A52',  'A53',  'A54',  'A55',  'A56',  'A57',  'A58',  'A59',  'A60',  'A61',  'A62',  'A63',  'A64',  'A65',  'A66',  'A67',  'A68',  'A69',  'A70',  'A71',  'A72',  'A73',  'A74',  'A75',  'A76',  'A77',  'A78',  'A79',  'A80',  'A81',  'A82',  'A83',  'A84',  'A85',  'A86',  'A87',  'A88',  'A89',  'A90',  'A91',  'A92',  'A93',  'A94',  'A95',  'A96',  'A97',  'A98',  'A99',  'A100',  'A101',  'A102',  'A103',  'A104',  'A105',  'A106',  'A107',  'A108',  'A109',  'A110',  'A111',  'A112',  'A113',  'A114',  'A115',  'A116',  'A117',  'A118',  'A119',  'A120',  'A121',  'A122',  'A123',  'A124',  'A125',  'A126',  'A127',  'A128',  'A129',  'A130',  'A131',  'A132',  'A133',  'A134',  'A135',  'A136',  'A137',  'A138',  'A139',  'A140',  'A141',  'A142',  'A143',  'A144',  'A145',  'A146',  'A147',  'A148',  'A149',  'A150',  'A151',  'A152',  'A153',  'A154',  'A155',  'A156',  'A157',  'A158',  'A159',  'A160',  'A161',  'A162',  'A163',  'A164',  'A165',  'A166',  'A167',  'A168',  'A169',  'A170',  'A171',  'A172',  'A173',  'A174',  'A175',  'A176',  'A177',  'A178',  'A179',  'A180',  'A181',  'A182',  'A183',  'A184',  'A185',  'A186',  'A187',  'A188',  'A189',  'A190',  'A191',  'A192',  'A193',  'A194',  'A195',  'A196',  'A197',  'A198',  'A199',  'A200',  'A201',  'A202',  'A203',  'A204',  'A205',  'A206',  'A207',  'A208',  'A209',  'A210',  'A211',  'A212',  'A213',  'A214',  'A215',  'A216',  'A217',  'A218',  'A219',  'A220',  'A221',  'A222',  'A223',  'A224',  'A225',  'A226',  'A227',  'A228',  'A229',  'A230',  'A231',  'A232',  'A233',  'A234',  'A235',  'A236',  'A237',  'A238',  'A239',  'A240',  'A241',  'A242',  'A243',  'A244',  'A245',  'A246',  'A247',  'A248',  'A249',  'A250',  'A251',  'A252',  'A253',  'A254',  'A255',  'A256',  'A257',  'A258',  'A259',  'A260',  'A261',  'A262',  'A263',  'A264',  'A265',  'A266',  'A267',  'A268',  'A269',  'A270',  'A271',  'A272',  'A273',  'A274',  'A275',  'A276',  'A277',  'A278',  'A279',  'A280',  'A281',  'A282',  'A283',  'A284',  'A285',  'A286',  'A287',  'A288',  'A289',  'A290',  'A291',  'A292',  'A293',  'A294',  'A295',  'A296',  'A297',  'A298',  'A299',  'A300',  'A301',  'A302',  'A303',  'A304',  'A305',  'A306',  'A307',  'A308',  'A309',  'A310',  'A311',  'A312',  'A313',  'A314',  'A315',  'A316',  'A317',  'A318',  'A319',  'A320',  'A321',  'A322',  'A323',  'A324',  'A325',  'A326',  'A327',  'A328',  'A329',  'A330',  'A331',  'A332',  'A333',  'A334',  'A335',  'A336',  'A337',  'A338',  'A339',  'A340',  'A341',  'A342',  'A343',  'A344',  'A345',  'A346',  'A347',  'A348',  'A349',  'A350',  'A351',  'A352',  'A353',  'A354',  'A355',  'A356',  'A357',  'A358',  'A359',  'A360',  'A361',  'A362',  'A363',  'A364',  'A365',  'A366',  'A367',  'A368',  'A369',  'A370',  'A371',  'A372',  'A373',  'A374',  'A375',  'A376',  'A377',  'A378',  'A379',  'A380',  'A381',  'A382',  'A383',  'A384',  'A385',  'A386',  'A387',  'A388',  'A389',  'A390',  'A391',  'A392',  'A393',  'A394',  'A395',  'A396',  'A397',  'A398',  'A399',  'A400',  'A401',  'A402',  'A403',  'A404',  'A405',  'A406',  'A407',  'A408',  'A409',  'A410',  'A411',  'A412',  'A413',  'A414',  'A415',  'A416',  'A417',  'A418',  'A419',  'A420',  'A421',  'A422',  'A423',  'A424',  'A425',  'A426',  'A427',  'A428',  'A429',  'A430',  'A431',  'A432',  'A433',  'A434',  'A435',  'A436',  'A437',  'A438',  'A439',  'A440',  'A441',  'A442',  'A443',  'A444',  'A445',  'A446',  'A447',  'A448',  'A449',  'A450',  'A451',  'A452',  'A453',  'A454',  'A455',  'A456',  'A457',  'A458',  'A459',  'A460',  'A461',  'A462',  'A463',  'A464',  'A465',  'A466',  'A467',  'A468',  'A469',  'A470',  'A471',  'A472',  'A473',  'A474',  'A475',  'A476',  'A477',  'A478',  'A479',  'A480',  'A481',  'A482',  'A483',  'A484',  'A485',  'A486',  'A487',  'A488',  'A489',  'A490',  'A491',  'A492',  'A493',  'A494',  'A495',  'A496',  'A497',  'A498',  'A499',  'A500'];
  var TAP = 0;
  var VAP = 0;
  var currentPos = 0;
  var product = titleArray[TAP];
  Logger.log(product);
  
  for ( i = 0; i <= 1000; i++ ) {
    
    if( product == 'Product 1') {
      spreadsheet.setActiveSheet(spreadsheet.getSheetByName('List'), true);
      spreadsheet.getRange(currentListCellArray[currentPos]).activate();
      spreadsheet.getRange('\'Master Checklist\'!A27:C28').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
      var option1 = variantArray[VAP];
      var option2 = variantArray[VAP + 1];
      var currentPos = currentPos + 2;
      if (option1 == 'Standard' ) {};
      if (option1 == 'Autoplay' ) {
        var currentPos = currentPos - 1;
        spreadsheet.getRange(currentListCellArray[currentPos]).activate()
        spreadsheet.getRange('\'Master Checklist\'!A131:C133').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
        var currentPos = currentPos + 3;
      };
      if (option2 == 'Yes' ) {
        spreadsheet.getRange(currentListCellArray[currentPos]).activate()
        spreadsheet.getRange('\'Master Checklist\'!A30:C30').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
        var currentPos = currentPos + 1;
      };
      if (option2 == 'No' ) {};
      
      
      
      var TAP = TAP + 2;
      var VAP = VAP + 2;
    };
    
    if ( product == 'Product 2') {
      spreadsheet.setActiveSheet(spreadsheet.getSheetByName('List'), true);
      spreadsheet.getRange(currentListCellArray[currentPos]).activate();
      spreadsheet.getRange('\'Master Checklist\'!A36:C37').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
      var option1 = variantArray[VAP];
      var option2 = variantArray[VAP + 1];
      var currentPos = currentPos + 2;
      if (option1 == 'Standard' ) {};
      if (option1 == 'Autoplay' ) {
        var currentPos = currentPos - 1;
        spreadsheet.getRange(currentListCellArray[currentPos]).activate()
        spreadsheet.getRange('\'Master Checklist\'!A138:C39').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
        var currentPos = currentPos + 2;
      };
      if (option2 == 'Yes' ) {
        spreadsheet.getRange(currentListCellArray[currentPos]).activate()
        spreadsheet.getRange('\'Master Checklist\'!A30:C30').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
        var currentPos = currentPos + 1;
      };
      if (option2 == 'No' ) {};
      
      var TAP = TAP + 2;
      var VAP = VAP + 1;
    };
    
    if( product == 'Product 3') {
      spreadsheet.setActiveSheet(spreadsheet.getSheetByName('List'), true);
      spreadsheet.getRange(currentListCellArray[currentPos]).activate();
      spreadsheet.getRange('\'Master Checklist\'!A86:C88').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
      var option1 = variantArray[VAP];
      var option2 = variantArray[VAP + 1];
      var option3 = variantArray[VAP + 2];
      var currentPos = currentPos + 3;
      if (option1 == 'Stainless' ) {
        var currentPos = currentPos - 2;
        spreadsheet.getRange(currentListCellArray[currentPos]).setValue('Stainless');
        var currentPos = currentPos + 2;
      };
      if (option1 == 'Black' ) {
        var currentPos = currentPos - 2;
        spreadsheet.getRange(currentListCellArray[currentPos]).setValue('Black');
        var currentPos = currentPos + 2;
      };
      if (option2 == '14' ) {  
        var currentPos = currentPos - 1;
        spreadsheet.getRange(currentListCellArray[currentPos]).setValue('14');
        var currentPos = currentPos + 1;
      };
      if (option2 == '21' ) {    
        var currentPos = currentPos - 1;
        spreadsheet.getRange(currentListCellArray[currentPos]).setValue('21');
        var currentPos = currentPos + 1;
      };
      if (option3 == 'Blue' ) {    
        var currentPos = currentPos - 1;
        spreadsheet.getRange(currentListCellArray[currentPos]).activate();
        spreadsheet.getRange('\'Master Checklist\'!A90:C90').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
        spreadsheet.getRange(currentListCellArray[currentPos]).setValue('Blue');
        var currentPos = currentPos + 1;
      };
      if (option3 == 'Green' ) {    
        var currentPos = currentPos - 1;
        spreadsheet.getRange(currentListCellArray[currentPos]).activate();
        spreadsheet.getRange('\'Master Checklist\'!A90:C90').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
        spreadsheet.getRange(currentListCellArray[currentPos]).setValue('Green');
        var currentPos = currentPos + 1;
      };
      if (option3 == 'Red' ) {    
        var currentPos = currentPos - 1;
        spreadsheet.getRange(currentListCellArray[currentPos]).activate();
        spreadsheet.getRange('\'Master Checklist\'!A90:C90').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
        spreadsheet.getRange(currentListCellArray[currentPos]).setValue('Red');
        var currentPos = currentPos + 1;
      };
      if (option3 == 'White' ) {    
        var currentPos = currentPos - 1;
        spreadsheet.getRange(currentListCellArray[currentPos]).activate();
        spreadsheet.getRange('\'Master Checklist\'!A90:C90').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
        spreadsheet.getRange(currentListCellArray[currentPos]).setValue('White');
        var currentPos = currentPos + 1;
      };
      var TAP = TAP + 3;
      var VAP = VAP + 1;
    };
    if( product == 'Product 3') {
      spreadsheet.setActiveSheet(spreadsheet.getSheetByName('List'), true);
      spreadsheet.getRange(currentListCellArray[currentPos]).activate();
      spreadsheet.getRange('\'Master Checklist\'!A95:C97').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
      var option1 = variantArray[VAP];
      var currentPos = currentPos + 4;
      if (option1 == 'Green' ) {
        var currentPos = currentPos - 3;
        spreadsheet.getRange(currentListCellArray[currentPos]).setValue('Green');
        var currentPos = currentPos + 3;
      };
      if (option1 == 'White' ) {
        var currentPos = currentPos - 3;
        spreadsheet.getRange(currentListCellArray[currentPos]).setValue('White');
        var currentPos = currentPos + 3;
      };
      if (option1 == 'Red' ) {
        var currentPos = currentPos - 3;
        spreadsheet.getRange(currentListCellArray[currentPos]).setValue('Red');
        var currentPos = currentPos + 3;
      };
      
      var TAP = TAP + 1;
      var VAP = VAP + 1;
    };
    if( product == 'Product 4') {
      spreadsheet.setActiveSheet(spreadsheet.getSheetByName('List'), true);
      spreadsheet.getRange(currentListCellArray[currentPos]).activate();
      var option1 = variantArray[VAP];
      var option2 = variantArray[VAP + 1];
      if (option1 == 'Through' ) {
        spreadsheet.getRange('\'Master Checklist\'!A57:C58').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
        var currentPos = currentPos + 2;
      };
      if (option1 == 'Behind' ) {
        spreadsheet.getRange('\'Master Checklist\'!A64:C65').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
        var currentPos = currentPos + 2;
      };
      if (option1 == 'Style 1' ) {
        spreadsheet.getRange('\'Master Checklist\'!A71:C72').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
        var currentPos = currentPos + 2;
      };
      if (option1 == 'Style 2' ) {
      spreadsheet.getRange('\'Master Checklist\'!A79:C79').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
        var currentPos = currentPos + 1;
      };
      if (option2 == '3"' ) {
        spreadsheet.getRange(currentListCellArray[currentPos]).offset(-1, 1).setValue('3"');
      };
      if (option2 == '4.5"' ) {
        spreadsheet.getRange(currentListCellArray[currentPos]).offset(-1, 1).setValue('4.5"');
      };
      if (option2 == '20W' ) {
        spreadsheet.getRange(currentListCellArray[currentPos]).offset(-1, 1).setValue('20W');
      };
      if (option2 == '25W' ) {
        spreadsheet.getRange(currentListCellArray[currentPos]).offset(-1, 1).setValue('25W');
      };
      
      var TAP = TAP + 2;
      var VAP = VAP + 1;
    };
    Logger.log(product);
    var product = titleArray[TAP];
    Logger.log(product);
  };
};

function mac1() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1').activate();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.NEXT).activate();
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  spreadsheet.getCurrentCell().setValue('1');
};

function test1() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getCurrentCell().offset(-18, -1).activate();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.NEXT).activate();
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  spreadsheet.getCurrentCell().setValue('1');
  spreadsheet.getCurrentCell().offset(1, 0).activate();
};

function mac2() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1').activate();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.NEXT).activate();
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  spreadsheet.getCurrentCell().setValue('2');
};

function mac3() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1').activate();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.NEXT).activate();
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  spreadsheet.getCurrentCell().setValue('3');
};

function mac4() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1').activate();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.NEXT).activate();
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  spreadsheet.getCurrentCell().setValue('4');
};

function mac5() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1').activate();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.NEXT).activate();
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  spreadsheet.getCurrentCell().setValue('5');
};

function mac6() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1').activate();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.NEXT).activate();
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  spreadsheet.getCurrentCell().setValue('6');
};

function mac7() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1').activate();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.NEXT).activate();
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  spreadsheet.getCurrentCell().setValue('7');
};

function mac8() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1').activate();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.NEXT).activate();
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  spreadsheet.getCurrentCell().setValue('8');
};

function mac9() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1').activate();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.NEXT).activate();
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  spreadsheet.getCurrentCell().setValue('9');
};

function mac10() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1').activate();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.NEXT).activate();
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  spreadsheet.getCurrentCell().setValue('10');
};

function mac11() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1').activate();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.NEXT).activate();
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  spreadsheet.getCurrentCell().setValue('11');
};

function mac12() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1').activate();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.NEXT).activate();
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  spreadsheet.getCurrentCell().setValue('12');
};

function mac13() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1').activate();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.NEXT).activate();
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  spreadsheet.getCurrentCell().setValue('13');
};

function mac14() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1').activate();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.NEXT).activate();
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  spreadsheet.getCurrentCell().setValue('14');
};

function mac15() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1').activate();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.NEXT).activate();
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  spreadsheet.getCurrentCell().setValue('15');
};

function mac16() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1').activate();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.NEXT).activate();
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  spreadsheet.getCurrentCell().setValue('16');
};

function mac17() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1').activate();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.NEXT).activate();
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  spreadsheet.getCurrentCell().setValue('17');
};

