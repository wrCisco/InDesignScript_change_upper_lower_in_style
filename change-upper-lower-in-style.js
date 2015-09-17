//Lo script opera sul documento corrente.
var myDocument = app.activeDocument;

//Inizializzo una variabile che contenga un riferimento a tutti gli stili di 
//paragrafo del documento e un'altra che contenga una lista di stringhe 
//dei rispettivi nomi.
var paragrafi = myDocument.allParagraphStyles.sort();
var listaparagrafi = [];
for (indice = 0; indice < paragrafi.length; indice++) {
  listaparagrafi.push(paragrafi[indice].name);
}
//Creo la finestra di dialogo per l'interfaccia utente.
var myDialog = app.dialogs.add(
{name:"Modifica maiuscole/minuscole in uno stile"});

//Creo i contenitori e gli elementi di controllo all'interno della finestra di 
//dialogo.
with(myDialog.dialogColumns.add()){
  with (borderPanels.add()) {
    with (dialogColumns.add()) {
      staticTexts.add({staticLabel:"Seleziona lo stile:"});
    }  
    with (dialogColumns.add()) {
      var lista_drop_down = dropdowns.add(
      {stringList: listaparagrafi, selectedIndex: 0});
      var radiobutt = radiobuttonGroups.add();
      with (radiobutt) {
        var maiusc = radiobuttonControls.add(
        {staticLabel:"Converti in maiuscole", checkedState:true});
        var minusc = radiobuttonControls.add(
        {staticLabel:"Converti in minuscole"});
      }
    }
  }
  with (borderPanels.add({name:"Opzioni"})) {
    with(dialogColumns.add()) {
      //Inizializzo le variabili legate alle espressioni regolari da utilizzare
      //per selezionare il testo all'interno dello stile su cui applicare 
      //le conversioni. Le espressioni regolari sono definite nella funzione
      //applica_conversione
      var opzione1 = checkboxControls.add({staticLabel:"Iniziali di paragrafo"});
      var opzione2 = checkboxControls.add(
      {staticLabel:"Iniziali di frase (precedute da un punto e uno spazio)"});
      var opzione3 = checkboxControls.add({staticLabel:
      "Iniziali di frase (precedute da un segno fra ? ! . … e da uno spazio)"});
      var opzione4 = checkboxControls.add({staticLabel:
      "Iniziali di frase (precedute da un segno fra ? ! . … e da zero o più spazi)"});
      var opzione5 = checkboxControls.add({staticLabel:
      "Iniziali di riga (precedute da un'interruzione forzata di riga)"});
      var opzione6 = checkboxControls.add({staticLabel:"Iniziali di parola"});
      var opzione7 = checkboxControls.add({staticLabel:"Tutte le lettere"});
      var opzione8 = checkboxControls.add(
      {staticLabel:"Tutte le lettere tranne le iniziali di parola"});
      with(dialogRows.add()) {
        var opzione9 = checkboxControls.add();
        var valore_opzione9 = textEditboxes.add(
        {editContents:"Scrivi una tua espressione regolare", minWidth:300});
      }
    }
  }
}

//Attivo la finestra di dialogo per l'utente.
var finestra = myDialog.show();

//Inizializzo due variabili che verranno passate come argomenti alla funzione
//applica_conversione e i cui valori sono decisi dall'utente nella finestra di
//dialogo: lo stile di paragrafo selezionato e il verso della conversione 
//(da maiuscole a minuscole o viceversa).
var stile_selezionato = paragrafi[lista_drop_down.selectedIndex];
if (radiobutt.selectedButton == 0) {
  var conversione = ChangecaseMode.uppercase;
} else {
  var conversione = ChangecaseMode.lowercase;
}
//Se l'utente clicca sul pulsante "ok", viene eseguita la funzione 
//applica_conversione.
if (finestra == true) {
  applica_conversione(conversione, opzione1.checkedState, opzione2.checkedState, 
    opzione3.checkedState, opzione4.checkedState, opzione5.checkedState,
    opzione6.checkedState, opzione7.checkedState, opzione8.checkedState,
    opzione9.checkedState, valore_opzione9.editContents, stile_selezionato);
}

//Cancella l'oggetto finestra di dialogo.
myDialog.destroy();


function applica_conversione(change, init_par, init_period, init_period_2,
init_period_3, init_row, init_word, all_chars, all_not_init, custom_check,
custom_value, stile_sel) {
  app.findGrepPreferences = NothingEnum.nothing;
  app.changeGrepPreferences = NothingEnum.nothing;
  app.findChangeGrepOptions.includeFootnotes = false;
  app.findChangeGrepOptions.includeHiddenLayers = false;
  app.findChangeGrepOptions.includeLockedLayersForFind = false;
  app.findChangeGrepOptions.includeLockedStoriesForFind = false;
  app.findChangeGrepOptions.includeMasterPages = false;
  app.findGrepPreferences.appliedParagraphStyle = stile_sel;
  
  var elencoGrep = {
    inizio_paragrafo: "(?<!\\n)^[^\\w]*?\\b\\w", //[[:punct:]]\\s*
    inizio_frase_dopo_punto: "\\.\\s[«\\-\"]?\\w",
    inizio_frase: "[\\.?!…]\\s[«\\-—\"]?\\w",
    inizio_frase_anche_senza_spazio: "[\\.?!…]\\s*?[«\\-—\"]?\\w",
    inizio_riga: "\\n[^\\w]*?\\w",
    inizio_parola: "\\b\\w",
    tutti_i_caratteri: ".+",
    tutte_tranne_iniziali: "(?<!\\b)\\w",
    regex_personalizzata: custom_value
  }
  if (init_par == true) {
    app.findGrepPreferences.findWhat = elencoGrep["inizio_paragrafo"]; 
    find_and_updown(change);
    alert("Trovate "+myFoundItems.length+" iniziali di paragrafo.");
  }
  if (init_period == true) {
    app.findGrepPreferences.findWhat = elencoGrep["inizio_frase_dopo_punto"];
    find_and_updown(change);
    alert("Trovate "+myFoundItems.length+" iniziali di frase.");
  }
  if (init_period_2 == true) {
    app.findGrepPreferences.findWhat = elencoGrep["inizio_frase"];
    find_and_updown(change);
    alert("Trovate "+myFoundItems.length+" iniziali di frase.");
  }
  if (init_period_3 == true) {
    app.findGrepPreferences.findWhat = elencoGrep["inizio_frase_anche_senza_spazio"];
    find_and_updown(change);
    alert("Trovate "+myFoundItems.length+" iniziali di frase.");
  }
  if (init_row == true) {
    app.findGrepPreferences.findWhat = elencoGrep["inizio_riga"];
    find_and_updown(change);
    alert("Trovate "+myFoundItems.length+" iniziali di riga.");
  }
  if (init_word == true) {
    app.findGrepPreferences.findWhat = elencoGrep["inizio_parola"];
    find_and_updown(change);
    alert("Trovate "+myFoundItems.length+" iniziali di parola.");
  }
  if (all_chars == true) {
    app.findGrepPreferences.findWhat = elencoGrep["tutti_i_caratteri"];
    find_and_updown(change);
    alert("Modifica effettuata su tutti i caratteri.");
  }
  if (all_not_init == true) {
    app.findGrepPreferences.findWhat = elencoGrep["tutte_tranne_iniziali"];
    find_and_updown(change);
    alert("Trovate "+myFoundItems.length+" lettere non iniziali.");
  }
  if (custom_check == true) {
    app.findGrepPreferences.findWhat = elencoGrep["regex_personalizzata"];
    find_and_updown(change);
    alert("Trovate "+myFoundItems.length+" occorrenze.");
  }
  if ((init_par || init_period || init_period_2 || init_period_3 || init_row ||
  init_word || all_chars || all_not_init || custom_check) == false) {
    alert("Nessun cambiamento effettuato.");
  }
  
  app.findGrepPreferences = NothingEnum.nothing;
  app.changeGrepPreferences = NothingEnum.nothing;
}

function find_and_updown(converti) {
  myFoundItems = myDocument.findGrep();
  for (index = 0; index < myFoundItems.length; index++) {
    myFoundItems[index].changecase(converti);
  }
}
