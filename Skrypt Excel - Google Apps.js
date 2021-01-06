var sebadoc = DocumentApp.openByUrl("private url to google docs");
var janekdoc = DocumentApp.openByUrl("private url to google docs");
var sheeturl = SpreadsheetApp.openByUrl("private url to google spreadsheet");
var mainsheet = sheeturl.getSheetByName("Refile");
var statsheet = sheeturl.getSheetByName("Statystyki");
var variablesheet = sheeturl.getSheetByName("Variables");


function onOpen() {
  // Odwołujemy się do naszego arkusza i jego metody
  // dla pobrania całego UI
  SpreadsheetApp.getUi()
  .createMenu('Skrypty') // Tworzymy nową pozycję w głównym menu
  .addItem('Przekopiuj sebe', 'eba') // Dodajemy opcję która uruchomi wskazaną w drugim parametrze funkcję 
  .addItem('Przekopiuj janka', 'janek')
  .addItem('Przekopiuj obu', 'both')
  .addItem('Usun puste wiersze', 'clear')
  //.addItem('test', 'test1')
  .addToUi();
}

var xpt = "XP Gain: ";
var loott = "Loot: ";
var wastet = "Supplies: ";
var balancest = "Balance: -";
var balancejt = "Balance: ";
var damaget = "Damage: ";
var healingt = "Healing: ";
var refilt = "Refil: ";
var frazt = "x frazzlemaw";
var guzt = "x guzzlemaw";
var silt = "x silencer";

function getmobs(body) {
  var frazel = body.findText(frazt).getElement();
  var guzel = body.findText(guzt).getElement();
  var silel = body.findText(silt, body.findText(damaget)).getElement();
  
  var fraz = frazel.asText().getText().replace(frazt, '');
  var guz = guzel.asText().getText().replace(guzt, '');
  var sil = silel.asText().getText().replace(silt, '');
  
  frazel.removeFromParent();
  guzel.removeFromParent();
  silel.removeFromParent();
  
  var mobs = parseInt(fraz) + parseInt(guz) + parseInt(sil);
  return mobs;
  
}


function getmobsv2(body) {
  var startTag = 'Killed';
  var endTag = 'Looted'
  var para = body.getParagraphs();
  var name;
      var killed = 0;
    var max = 0;
  
  for (var i=0; i<para.length; i++) {
    var from = para[i].findText(startTag);
    var to = para[i].findText(endTag, from);


    //Logger.log(para[i].getText());
    //Logger.log(i);
    if (from != null) {
      
      i++;
      while (para[i].findText(endTag) == null) {
        var tempnum = parseInt(para[i].getText().substring(0, para[i].getText().indexOf("x")));
        if (tempnum > max) {
          var words = para[i].getText().substring(2).split(" ");
          
          name = words[1];
          max = tempnum;
        }
        
        killed += tempnum;
        
        para[i].removeFromParent();
        i++;
        
        
    } 
      body.findText(endTag).getElement().removeFromParent();
      body.findText(startTag).getElement().removeFromParent();
      break;
    }
  }
  var array = new Array(killed, name);
  return array;
}



function datefind(date, hour, kolumna, wierszyk, again){
  var kolumienka = mainsheet.getRange(kolumna + ":" + kolumna);  // like A:A
  var values = kolumienka.getValues();
  
  
  do {
    wierszyk++;
    var tempcell = values[wierszyk-1][0];
    
    if(!(tempcell == "")) {
      var convdate  =  Utilities.formatDate(tempcell, "GMT+2", "yyyy-MM-dd");
    }
    else {
      break;
    }
  }
  while (convdate != date)
    
    
    if (convdate == date) {
      //Browser.msgBox("tu");
      
      if (hour >= 10) {
        //Browser.msgBox("znaleziono date " + convdate + " w wierszyku " + wierszyk);
        return wierszyk;
      }
      else {
        //Browser.msgBox("znaleziono date " + convdate + " w wierszyku " + (wierszyk-1));
        return (wierszyk-1);
      }
    }
  else {
    if (!again) {
      wierszyk = 5;
      again = true;
      return datefind(date, hour, kolumna, wierszyk, again);
    }
    else {
      Browser.msgBox("aborting...");
      return 0;
    }
  }
}

function eba() {
  var body = sebadoc.getBody();
  Browser.msgBox("Wpisujemy sebe");
  var lastrowcell = variablesheet.getRange('B2');
  var lastrow = lastrowcell.getValue();
  
  var searchdoc = true;
  var ilerazy = 0;
  while (searchdoc) {
    var fromre = body.findText("From ");
    
    if (fromre != null) {
      var fromtext = fromre.getElement().asText().getText();
      var fromel = fromre.getElement();
      
      var datetofind = fromtext.substring(fromtext.indexOf("From") + 5, fromtext.indexOf("From") + 15);
      var hour = fromtext.substring(fromtext.indexOf("From") + 17, fromtext.indexOf("From") + 19);
      var wierszyk = datefind(datetofind, hour, "A", lastrow, false);
      
      var actualxpc = mainsheet.getRange('D' + wierszyk);
      var actualwastec = mainsheet.getRange('E' + wierszyk);
      var actualdamagec = mainsheet.getRange('F' + wierszyk);
      var actualhealingc = mainsheet.getRange('G' + wierszyk);
      var actualrefilc = statsheet.getRange('J' + wierszyk);
      
      
      var xpel = body.findText(xpt).getElement();
      var balanceel = body.findText(balancest).getElement();
      var damageel = body.findText(damaget).getElement();
      var healingel = body.findText(healingt).getElement();
      var refilel = body.findText(refilt).getElement();
      
      var xp = xpel.asText().getText().replace(xpt, '');
      var balance = balanceel.asText().getText().replace(balancest, '');
      var damage = damageel.asText().getText().replace(damaget, '');
      var healing = healingel.asText().getText().replace(healingt, '');
      var refil = refilel.asText().getText().replace(refilt, '');
      
      if (xp.match(/,.*,/)) { // Check if there are 2 commas
        xp = xp.replace(',', ''); // Remove the first one
      }
      if (balance.match(/,.*,/)) { // Check if there are 2 commas
        balance = balance.replace(',', ''); // Remove the first one
      }
      if (damage.match(/,.*,/)) { // Check if there are 2 commas
        damage = damage.replace(',', ''); // Remove the first one
      }
      if (healing.match(/,.*,/)) { // Check if there are 2 commas
        healing = healing.replace(',', ''); // Remove the first one
      }
      
      actualxpc.setValue("=" + actualxpc.getDisplayValue() + "+" + xp);
      actualwastec.setValue("=FLOOR(" + actualwastec.getDisplayValue() + "+" + balance + ")");
      actualdamagec.setValue("=" + actualdamagec.getDisplayValue() + "+" + damage);
      actualhealingc.setValue("=" + actualhealingc.getDisplayValue() + "+" + healing);
      actualrefilc.setValue("=" + actualrefilc.getDisplayValue() + "+" + refil);
      
      
      
      
      xpel.removeFromParent();
      balanceel.removeFromParent();
      damageel.removeFromParent();
      healingel.removeFromParent();
      refilel.removeFromParent();
      fromel.removeFromParent();
      
      //Browser.msgBox(wierszyk);
      
      lastrow = wierszyk - 1;
      ilerazy++;
      //searchdoc = false; //to wykasowac potem bo narazie tescik na jednego whilea
    }
    else {
      Browser.msgBox("Zakonczono kopiowanie, przekopiowano " + ilerazy + " wpisow.");
      body.clear();
      searchdoc = false;
    }
    lastrowcell.setValue(lastrow);
  }
}


function janek() {
  var body = janekdoc.getBody();
  Browser.msgBox("Wpisujemy janka");
  var lastrowcell = variablesheet.getRange('A2');
  var lastrow = lastrowcell.getValue();
  
  var searchdoc = true;
  var ilerazy = 0;
  while (searchdoc) {
    var fromre = body.findText("From ");
    
    if (fromre != null) {
      var fromtext = fromre.getElement().asText().getText();
      var fromel = fromre.getElement();
      
      var datetofind = fromtext.substring(fromtext.indexOf("From") + 5, fromtext.indexOf("From") + 15);
      var hour = fromtext.substring(fromtext.indexOf("From") + 17, fromtext.indexOf("From") + 19);
      
      var wierszyk = datefind(datetofind, hour, "A", lastrow, false);
      
      var actualxpc = mainsheet.getRange('M' + wierszyk);
      var actualwastec = mainsheet.getRange('N' + wierszyk);
      var actualdamagec = mainsheet.getRange('O' + wierszyk);
      var actuallootc = statsheet.getRange('K' + wierszyk);
      var mobsc = statsheet.getRange('Y' + wierszyk);
      var mobsavgc = statsheet.getRange('AA' + wierszyk);
      
      var mobsnamec = mainsheet.getRange('B' + wierszyk);
      
      var mobsarr = getmobsv2(body);
      
      var mobs = mobsarr[0];
      var mobname = mobsarr[1];
      
      
      var xpel = body.findText(xpt).getElement();
      var wasteel = body.findText(wastet).getElement();
      var damageel = body.findText(damaget).getElement();
      var lootel = body.findText(loott).getElement();
      
      
      var xp = xpel.asText().getText().replace(xpt, '');
      var waste = wasteel.asText().getText().replace(wastet, '');
      var damage = damageel.asText().getText().replace(damaget, '');
      var loot = lootel.asText().getText().replace(loott, '');
      
      if (xp.match(/,.*,/)) { // Check if there are 2 commas
        xp = xp.replace(',', ''); // Remove the first one
      }
      if (loot.match(/,.*,/)) { // Check if there are 2 commas
        loot = loot.replace(',', ''); // Remove the first one
      }
      if (waste.match(/,.*,/)) { // Check if there are 2 commas
        waste = waste.replace(',', ''); // Remove the first one
      }
      if (damage.match(/,.*,/)) { // Check if there are 2 commas
        damage = damage.replace(',', ''); // Remove the first one
      }
      
      var mobsloot = loot.replace(',', '')
      var mobsavg = parseInt(mobsloot / mobs);
      
      
      actualxpc.setValue("=" + actualxpc.getDisplayValue() + "+" + xp);
      actualwastec.setValue("=FLOOR(" + actualwastec.getDisplayValue() + "+" + waste + ")");
      actualdamagec.setValue("=" + actualdamagec.getDisplayValue() + "+" + damage);
      actuallootc.setValue("=FLOOR(" + actuallootc.getDisplayValue() + "+" + loot + ")");
      mobsc.setValue("=" + mobsc.getDisplayValue() + "+" + mobs);
      mobsavgc.setValue("=\"" + mobsavgc.getValue() + mobsavg + "\"&REPT(CHAR(160);1)");
      
      
      xpel.removeFromParent();
      wasteel.removeFromParent();
      damageel.removeFromParent();
      lootel.removeFromParent();
      fromel.removeFromParent();
     
      
      mobsnamec.setValue(mobname);
      
      
      //Logger.log(lastrow);
      //Logger.log(ilerazy);
      
      lastrow = wierszyk - 1;
      ilerazy++;
    }
    else {
      Browser.msgBox("Zakonczono kopiowanie, przekopiowano " + ilerazy + " wpisow.");
      body.clear();
      searchdoc = false;
    }
    lastrowcell.setValue(lastrow);
  }
}

function both() {
  Browser.msgBox("Przygotowuję kopiowanie obu...");
  eba();
  janek();
}

function test1() {
  var lolek = statsheet.getRange('U' + 5).getValue();
  if (lolek == 0) {
    Browser.msgBox("fak");
    Browser.msgBox(lolek);
  }
  else {
    Browser.msgBox("ju");
    Browser.msgBox(lolek);
    
  }
  
}

function clear() {
  var janeklastrow = variablesheet.getRange('A2').getValue();
  var sebalastrow = variablesheet.getRange('B2').getValue();
  var lastrowcell = variablesheet.getRange('C2');
  //could use MAX but comparing there also gives me variable
  if (sebalastrow < janeklastrow) {
    var zakres = 'E';
    var lastrow = sebalastrow;
  }
  else {
    var zakres = 'N';
    var lastrow = janeklastrow;
  }
  var strink = '';
  var deletedcount = 0;
  
  for (var row = lastrow; row >= Math.max(lastrowcell.getValue(), 5); row--) {
    if (mainsheet.getRange('E' + row).getValue() == 0 && mainsheet.getRange('N' + row).getValue() == 0 && statsheet.getRange('U' + row).getValue() == 0) {
      
      strink = strink + row + ' ';
      deletedcount++;
      mainsheet.deleteRow(row);
      statsheet.deleteRow(row);
    }
  }
  lastrowcell.setValue(lastrow - deletedcount);
  variablesheet.getRange('A2').setValue(janeklastrow - deletedcount);
  variablesheet.getRange('B2').setValue(sebalastrow - deletedcount);
  
  
  Browser.msgBox("Usunieto wiersze: " + strink);
  
}

