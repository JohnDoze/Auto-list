
function onEdit(e) {
  var sheet = e.source.getActiveSheet();
  var range = e.range;
  var value = range.getValue().toLowerCase(); // Convertir en minuscules pour une correspondance insensible à la casse

  // Vérifier si la cellule modifiée est dans la première colonne et non vide
  if (range.getColumn() == 1 && value != "") {
    var rule = SpreadsheetApp.newDataValidation()
      .requireValueInList([
        "lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi", "dimanche", // Jours de la semaine
        "janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre", "octobre", "novembre", "décembre", // Mois
        "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z", // Alphabet complet
        "do", "do#", "ré", "ré#", "mi", "fa", "fa#", "sol", "sol#", "la", "la#", "si", // Gamme avec dièse (#)
        "do", "ré♭", "ré", "mi♭", "mi", "fa", "solb", "sol", "la", "la♭", "si", "si♭", // Gamme avec bémol (♭)
        "A", "A#", "B", "C", "C#", "D", "D#", "E", "F", "F#", "G", "G#", // Gamme dièse (#) en majuscules
        "A", "B♭", "B", "C", "D♭", "D", "E♭", "E", "F", "G♭", "G", "A♭", // Gamme bémol (♭) en majuscules
        "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", // Chiffres de 1 à 10
        "true", "false" // Valeurs booléennes
      ], true)
      .build();
    range.setDataValidation(rule);
  }
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Autocomplétion')
      .addItem('Configurer', 'configureAutoComplete')
      .addToUi();
}

function configureAutoComplete() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var range = sheet.getDataRange();
  var rule = SpreadsheetApp.newDataValidation()
      .requireValueInList([
        "lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi", "dimanche", // Jours de la semaine
        "janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre", "octobre", "novembre", "décembre", // Mois
        "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z", // Alphabet complet
        "do", "do#", "ré", "ré#", "mi", "fa", "fa#", "sol", "sol#", "la", "la#", "si", // Gamme avec dièse (#)
        "do", "ré♭", "ré", "mi♭", "mi", "fa", "solb", "sol", "la", "la♭", "si", "si♭", // Gamme avec bémol (♭)
        "A", "A#", "B", "C", "C#", "D", "D#", "E", "F", "F#", "G", "G#", // Gamme dièse (#) en majuscules
        "A", "B♭", "B", "C", "D♭", "D", "E♭", "E", "F", "G♭", "G", "A♭", // Gamme bémol (♭) en majuscules
        "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", // Chiffres de 1 à 10
        "true", "false" // Valeurs booléennes
      ], true)
      .build();
  range.setDataValidation(rule);
}
