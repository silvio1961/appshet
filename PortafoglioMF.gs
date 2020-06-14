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


//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
function DataMax(ColData,ColVal,numrig) {

 var ss = SpreadsheetApp.getActiveSpreadsheet();
      sheet=ss.getActiveSheet(); //Store

  var Data=sheet.getRange(2, ColData, numrig, 1).getValues();
  var valori=sheet.getRange(2, ColVal, numrig, 1).getValues();

  
  var Max=-10000000;
  var Datamax=0;
  
  for(i=0;i<numrig;i++) {  

    
     var segno=valori[i]-Max;
    if(segno>0)  {

        Max=valori[i];
      
      Datamax=Data[i];
    }  
  
  }
  return Datamax;
}  

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
function DataMin(ColData,ColVal,numrig) {

 var ss = SpreadsheetApp.getActiveSpreadsheet();
      sheet=ss.getActiveSheet(); //Store
  Logger.log("NUmr="+numrig);
  var Data=sheet.getRange(2, ColData, numrig, 1).getValues();
  var valori=sheet.getRange(2, ColVal, numrig, 1).getValues();

  
  var Min=10000000;
  var Dataminimo=0;
  
  for(i=0;i<numrig;i++) {  
  var segno=valori[i]-Min;
    if(segno<0)  {  
      Logger.log("i="+ i+"Min="+Min+" Val="+valori[i]+" Data ="+Dataminimo);
        Min=valori[i];
    
      Dataminimo=Data[i];
    }  
  
  }
  return Dataminimo;
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


function onEdit(e) {  

  ss = SpreadsheetApp.getActiveSpreadsheet();  
  var sheet = ss.getActiveSheet(); 

      riga=e.range.getRow();
     col=e.range.getColumn();
  Logger.log("riga"+riga+" colonna="+col);
 
if (e.range.getSheet().getName() == 'Operazioni') {   // nascondi Verifiche


}
}



//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
function CancellaDatiGior() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheetSorg = ss.getSheetByName("guad-perd5m"); 
    var NumrSorg=sheetSorg.getRange("A1").getValue()+1;
 sheetSorg.getRange(2, 2,NumrSorg,10).clear();
}


//////////////////////////////////////////////////////////////////////////////////////////////////////////////////

function Datigiorn3() {
 var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheetSorg = ss.getSheetByName("guad-perd5m");
  
  Logger.log("Numr="+Numr);
  var source=sheetSorg.getRange("J10:X10").getValues();
  Logger.log(source);
    var sheetDest = ss.getSheetByName("DatiGiorn");
  var Numr=sheetDest.getRange("A1").getValue();
  sheetDest.getRange(Numr+2,2,1,15).setValues(source);
  
}
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
  
  
  
  


function guadagn5min() {
  var result = new Array(12); //Sort by date
  result[0] = new Date();
 var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetgua = ss.getSheetByName("guad-perd5m"); //Store
  var ok=ss.getRangeByName("OraOk").getValue();
  
  var sheetrag=ss.getSheetByName("raggrup");   
 
  if(ok) {
    
    var iniz=sheetgua.getRange("A1").getValue();

    var valo = sheetrag.getRange("G23:N23").getValues();
        Logger.log("sour="+valo);
    ss.getSheetByName("guad-perd5m").getRange(2+iniz,2,1,8).setValues(valo);
  }  
  
}
//////////////////////////////////
function maximo2() {
  
var ss = SpreadsheetApp.getActiveSpreadsheet();
 var sheet = ss.getSheetByName("Operaz Aperte"); //Store
 
    range=ss.getRangeByName("RecOpAp");
  range1=ss.getRangeByName("RecOpAp1");
  MinR=ss.getRangeByName("RigaMinOp").getValue();
  MaxR=ss.getRangeByName("RigaMaxOp").getValue();
 
  MinRsw=sheet.getRange("I1").getValue();
  MaxRsw=sheet.getRange("J1").getValue();
  NumRsw=MaxRsw-MinRsw+1;
   NumR=MaxR-MinR+1;
 
  source=range.getValues();
     Logger.log("NumR="+NumR+" NumRsw="+NumRsw);
 dest=ss.getRangeByName("RecOpAp1").getValues();
  for(k=0;k<2;k++) {
    if(k==0) {
         range=ss.getRangeByName("RecOpAp");
         range1=ss.getRangeByName("RecOpAp1");
         MinR=ss.getRangeByName("RigaMinOp").getValue();
         MaxR=ss.getRangeByName("RigaMaxOp").getValue();   
    }  
    else {
       range=ss.getRangeByName("RecOpApSw");
       range1=ss.getRangeByName("RecOpApSw1");
       MinR=sheet.getRange("I1").getValue();
       MaxR=sheet.getRange("J1").getValue();
        
    }  
    dest=range1.getValues();
    source=range.getValues();
    
    NumR=MaxR-MinR+1;
    for(i=0;i<NumR;i++) {
    
      valatt=source[i][0];
      maxim=source[i][1];
      minim=source[i][2];
//    Logger.log("i="+i+" valatt="+valatt+" minim="+minim +" maxim="+maxim);
      if(valatt>maxim) dest[i][0]=valatt;
      if(valatt<minim) dest[i][1]=valatt;
      minim=dest[i][1];
      maxim=dest[i][0];
  //      Logger.log("dopo i="+i+" minim="+minim +" maxim="+maxim);
    
    }
    if(k==0) range1=ss.getRangeByName("RecOpAp1");
    else range1=ss.getRangeByName("RecOpApSw1");
    range1.setValues(dest);

  }   
}


/////////////////////////////////////////////////////////
function aggiorna() {
var d1 = new Date();
var d2 = new Date(2017,3, 18, 17, 30, 0, 0);
var d3 = new Date(2017,3, 18, 9, 31, 0, 0);
ora1=d1.getHours();
min1=d1.getMinutes();
totor1=ora1*60+min1;
ora2=d2.getHours();
min2=d2.getMinutes();
totor2=ora2*60+min2;
ora3=d3.getHours();
min3=d3.getMinutes();
totor3=ora3*60+min3;
Logger.log(" ora att"+totor1);
Logger.log(" ora iniz="+totor2);
Logger.log(" ora fin="+totor3);
   maximo();
if((totor1>totor3) && (totor1<totor2)) {
Logger.log( "ehhjksllo");
   
}    
}






