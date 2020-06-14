function Daytrading() {  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
 
  var NumRG=ss.getRangeByName("NumRG").getValue();
   var ok=ss.getRangeByName("OKG").getValue();
  
   Logger.log("OK="+ok); 
  if(ok==1) {
  sheetOp=ss.getActiveSheet();
  Logger.log("OK1111d="+ok+" NumRG="+NumRG);
    // ss.getSheetByName("Giornaliero").deleteRow(5);
    if(NumRG==1200) ss.getSheetByName("Giornaliero").deleteRow(5);
 
//  sheet = ss.getSheetByName("Giornaliero"); //Store
  var NumRG=ss.getRangeByName("NumRG").getValue();
  sorg=ss.getRangeByName("DeposGior");
    var dest=sheetOp.getRange(NumRG,1,1,7);
  ss.getSheetByName("Giornaliero").getRange(NumRG,1,1,7).setValues(sorg.getValues());
  
  }
  else {
    if (ss.getRangeByName("OKG").getValue()==8) sheet.getRange('A5:f30').clear();
  }  
 
 
   
}


///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
function dollprec_celle() {
  var OutP = new Array(4);
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet1 = ss.getSheetByName("Dollaro"); //Get values
  var source=sheet1.getRange("A1:A2").getValues();
  NumR=source[0];
  Doll=source[1];
   Logger.log("sour="+source+" NumR="+NumR+"  Doll"+Doll);
  sheet1.getRange("A3").setValue(Doll);
  var ora=new Date();
  sheet1.getRange("A4").setValue(ora);


 sheet1.getRange("b"+NumR).setValue(ora);

 sheet1.getRange("C"+NumR).setValue(Doll);



}
//////////

function Visual() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
    ss.setActiveRange(ss.getRangeByName("Saldi"));
//  sheet=ss.getSheetByName("Menu");
  //SpreadsheetApp.setActiveSheet(sheet);
  
}

function Menu() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  sheet=ss.getSheetByName("Menu");
  SpreadsheetApp.setActiveSheet(sheet);
  
}

function formatta() {
 var ss = SpreadsheetApp.getActiveSpreadsheet();
 ss.getSheetByName("depos").getRange(1, 1, 800, 7).clear();
 
};




function aggiornaQuot(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
shdep=ss.getActiveSheet();
  Source=ss.getRangeByName("deposito").getValues();
 Logger.log("ss="+Source);
  Titolo=Source[0][0];
  NumR=Source[0][1];
  Logger.log("ss="+Titolo+"NumR= "+NumR);
   var shquot=ss.getSheetByName(Titolo);

  NumRDest=shquot.getRange("H1").getValue();
     Logger.log("ss="+shquot.getSheetName()+"NumRDest="+NumRDest);
  valori=shdep.getRange(7,23,NumR,3).getValues();
  shquot.getRange(NumRDest,1, NumR,3).setValues(valori);
  
   

}




function AggDati()
{
  var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet1 = ss.getSheetByName("Foglio40");
    
  source=sheet1.getRange("R3:R5").getValues();
  var Data=new Date(source[1]);
  
 // Data=source[1];
  Logger.log("Data="+Data);
  var GSett=Data.getDay();
  Logger.log("GSett="+GSett);
  
  if((GSett!=1) || (GSett!=7)) {
    
   var resto=source[0];
    
    resto++;
    if(resto==4) resto=0;
 //   Data.setDate(Data.getDate()+1);
    if(resto==0) sheet1.getRange("R4").setValue(source[2]);
    sheet1.getRange("R3").setValue(resto);
  }
  else  sheet1.getRange("R4").setValue(source[2]);
} 





function Copia() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('P:P').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('P10'));
  spreadsheet.getRange('P:P').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
}


//////////////////////////////////////////////////////////////////////////////////////////////////////////////////

function contenuto(Rig, Col) {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var sheet=ss.getActiveSheet();
  var valor=sheet.getRange(Rig,Col).getValue();
  return valor;
}  


function lancia() {
Logger.log("val="+ contenuto(39,3));
}





///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
function dollprec_celle() {
  var OutP = new Array(4);
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet1 = ss.getSheetByName("Dollaro"); //Get values
  var source=sheet1.getRange("A1:A2").getValues();
  NumR=source[0];
  Doll=source[1];
   Logger.log("sour="+source+" NumR="+NumR+"  Doll"+Doll);
  sheet1.getRange("A3").setValue(Doll);
  var ora=new Date();
  sheet1.getRange("A4").setValue(ora);


 sheet1.getRange("b"+NumR).setValue(ora);

 sheet1.getRange("C"+NumR).setValue(Doll);



}
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

function Resetta() {
   var sh = SpreadsheetApp.getActiveSheet(); 
  
  sh.getRange("A3:E3").clear();
}

function Modifica(NumRDest) {
  var sh = SpreadsheetApp.getActiveSheet(); 
  
//   var NumRDest=sh.getRange("W5").getValue();
  Logger.log("1 "+sh.getRange("L6:T6").getValues()+" NumR "+NumRDest);
      
   var   sheetDest=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Operazioni");
  sheetDest.getRange(NumRDest,1,1,1).setValues(sh.getRange("A3").getValues());
   sheetDest.getRange(NumRDest,3,1,1).setValues(sh.getRange("B3").getValues());
  sheetDest.getRange(NumRDest,5,1,2).setValues(sh.getRange("C3:D3").getValues());
   sheetDest.getRange(NumRDest,8,1,1).setValues(sh.getRange("E3").getValues());
    sh.getRange("G3:I3").setValues(sh.getRange("G4:I4").getValues());
  Resetta();
  
      }  

function onEdit(e) {  
   Logger.log("riga?");
  if (!e || e.value === undefined)     return;
  const edited = e.range;
  const ss = edited.getSheet();
  var sh = SpreadsheetApp.getActiveSheet();
   var riga=sh.getActiveCell().getRow();
     //  sheet.getActiveCell().getRow();
     var col=sh.getActiveCell().getColumn();
 Logger.log("riga?"+riga+" col="+col+" nome="+ss.getName());
if (ss.getName() == 'Operazioni per titolo') {   // nascondi Verifiche
   if((col==7) && (riga==3)) Modifica(sh.getRange("J1").getValue()); 
  if((col==8) && (riga==3)) {

     var currentRow = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Operazioni").getLastRow();
      var sourceRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Operazioni").getRange(currentRow, 18);
         var sourceFormulas = sourceRange.getFormulasR1C1();
    Modifica(sh.getRange("J3").getValue());  
    currentRow++;
         var targetRange =SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Operazioni").getRange(currentRow, 10);
         targetRange.setFormulasR1C1(sourceFormulas);
  }  
    if((col==9) && (riga==3)) {
       
         var NumRDest=sh.getRange("J1").getValue();
         Logger.log("NumRDest="+NumRDest);      
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Operazioni").deleteRow(NumRDest);
       Resetta();
      sh.getRange("G3:I3").setValues(sh.getRange("G4:I4").getValues());
     
      } 
  if((col==5) && (riga>4)) {
  //   sh.getRange(sh.getRange("J6").getValue(),col).setValue(sh.getRange("G2").getValue());  
    var ID = sh.getRange(riga,col-1).getValue();
        sh.getRange("J1").setValue(ID);
  // sh.getRange(8,5,riga-7,1).setValues(sh.getRange(8,6,riga-8,1).getValues());
    //  sh.getRange(riga+1,5,100-riga,1).setValues(sh.getRange(riga+1,6,100-riga,1).getValues());
      sh.getRange("E8:E50").setValues(sh.getRange("F8:F50").getValues());
        sh.getRange(riga,col).setValue(sh.getRange("J2").getValue());
    sh.getRange("A3:E3").setValues(sh.getRange("A2:E2").getValues());
    
 //  sh.getRange("J6").setValue(riga);
    

                                                                    
    }

}
}


