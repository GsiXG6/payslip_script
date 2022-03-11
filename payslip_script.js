Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  await context.sync()
    .then(function () {
      console.log("======= Sheet repair initiated =======");
    });

  let i, range, conditionalFormula;
  let rowMin = 3;
  let rowMax = 33;

  //----------------------- --------------------------------------//
  //--------------------    INSERT FORMULA    --------------------//
  //------------------------ -------------------------------------//
  console.log("Adding formula");
  let lunchOut = 1.25;
  let counteR = 1;
  
  for (i = rowMin; i <= rowMax; i++) {
    sheet.getRange("A" + i).formulas = `=IF(C${i}<>"",IF(WEEKDAY(C${i})=1,1,""),"")`;
    //sheet.getRange("B" + i).formulas = ``;
    sheet.getRange("C" + i).formulas = `=IF($D$1>(${counteR}-1),DATE($B$1,$C$1,${counteR}),"")`;
    //sheet.getRange("D" + i).formulas = ``;
    //sheet.getRange("E" + i).formulas = ``;
    //sheet.getRange("F" + i).formulas = ``;
    sheet.getRange("G" + i).formulas = `=IF(AND(C${i}<>"",D${i}<>""),ROUNDDOWN(D${i},-2)/2400+MOD(D${i},100)/1440,"")`;
    sheet.getRange("H" + i).formulas = `=IF(AND(C${i}<>"",E${i}<>""),ROUNDDOWN(E${i},-2)/2400+MOD(E${i},100)/1440,"")`;;
    sheet.getRange("I" + i).formulas = `=IF(AND(G${i}<>"",H${i}<>""),MOD(H${i}-G${i},1)*24,"")`;
    sheet.getRange("j" + i).formulas = `=IF(AND(A${i}<>1,B${i}<>1,D${i}<>"",E${i}<>"",I${i}<>""),IFS(I${i}<4.25,0,I${i}<=4.25,0.5,I${i}<9.25,I${i}/9.25,I${i}>=9.25,1,TRUE,"error-default"),"")`;
    sheet.getRange("K" + i).formulas = `=IF(AND(A${i}<>1,B${i}<>1,D${i}<>"",E${i}<>"",I${i}<>"",I${i}>9.25),MROUND(I${i}-9.25,0.5),"")`;
    sheet.getRange("L" + i).formulas = `=IF(AND(A${i}=1,G${i}<>"",H${i}<>""),IF((H${i}-G${i})*24<=4.5,(H${i}-G${i})*24,(H${i}-G${i})*24-1.25),"")`;
    sheet.getRange("M" + i).formulas = `=IF(AND(B${i}=1,G${i}<>"",H${i}<>""),IF((H${i}-G${i})*24<=4.5,(H${i}-G${i})*24,(H${i}-G${i})*24-1.25),"")`;
    sheet.getRange("N" + i).formulas = `=IF(AND(K${i}>=3,K${i}<>"",H${i}<>""),1,"")`;
    sheet.getRange("O" + i).formulas = `=IF(AND(D${i}<>"",E${i}<>"",D${i}>E${i},J${i}>=1),1,"")`;
    counteR++;
  }
  await context.sync();
  console.log("Formula added!!!!");





  //----------------------- --------------------------------------//
  //-------------    CLEAR CONDITIONAL FORMATTING    -------------//
  //------------------------ -------------------------------------//
  console.log("Removing conditional formatting");
  range = sheet.getRange("A" + rowMin + ":R" + rowMax);
  conditionalFormula = range.conditionalFormats.clearAll();
  await context.sync();
  console.log("Conditional formatting removed!!!!");

  //----------------------- --------------------------------------//
  //------------    INSERT CONDITIONAL FORMATTING    -------------//
  //------------------------ -------------------------------------//
  console.log("Inserting new conditional formatting");
  let sndyColor = "yellow";
  let pbdyColor = "red";
  let errrColor = "#7030A0";
  let blanColor = "#A6A6A6";

  for (i = rowMin; i <= rowMax; i++) {
    //=WEEKDAY($C$3)=1
    range = sheet.getRange("A" + i + ":E" + i);
    conditionalFormula = range.conditionalFormats.add(Excel.ConditionalFormatType.custom);
    conditionalFormula.custom.rule.formula = "=$A$" + i + "=1";
    conditionalFormula.custom.format.fill.color = sndyColor;
    conditionalFormula.custom.format.font.color = "blue";
  }
  for (i = rowMin; i <= rowMax; i++) {
    //=WEEKDAY($C$3)=1
    range = sheet.getRange("G" + i + ":O" + i);
    conditionalFormula = range.conditionalFormats.add(Excel.ConditionalFormatType.custom);
    conditionalFormula.custom.rule.formula = "=$A$" + i + "=1";
    conditionalFormula.custom.format.fill.color = sndyColor;
    conditionalFormula.custom.format.font.color = "blue";
  }
  for (i = rowMin; i <= rowMax; i++) {
    range = sheet.getRange("A" + i + ":O" + i);
    conditionalFormula = range.conditionalFormats.add(Excel.ConditionalFormatType.custom);
    conditionalFormula.custom.rule.formula = `=AND($B$${i}=1,$C$${i}<>"")`;
    conditionalFormula.custom.format.font.color = pbdyColor;
  }
  for (i = rowMin; i <= rowMax; i++) {
    range = sheet.getRange("J" + i);
    conditionalFormula = range.conditionalFormats.add(Excel.ConditionalFormatType.custom);
    conditionalFormula.custom.rule.formula = `=AND($J$${i}<>"",$J$${i}<>1,$J$${i}<>0.5)`;
    conditionalFormula.custom.format.fill.color = errrColor;
  }
  for (i = rowMin; i <= rowMax; i++) {
    range = sheet.getRange("A" + i + ":O" + i);
    conditionalFormula = range.conditionalFormats.add(Excel.ConditionalFormatType.custom);
    conditionalFormula.custom.rule.formula = `=$C$${i}=""`;
    conditionalFormula.custom.format.fill.color = blanColor;
  }
  await context.sync();
  console.log("New conditional formatting added!!!!");

  //----------------------- --------------------------------------//
  //-----------------    ECO REPAIR COMPLETED    -----------------//
  //------------------------ -------------------------------------//
  sheet.load("name");
  await context.sync()
    .then(function () {
      console.log(`======= "${sheet.name}" repair completed =======`);
    });
})