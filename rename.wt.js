const tableId = "MultiRenameTable";
const oldNameClassName = "x-grid3-col x-grid3-cell x-grid3-td-objName ";
const newNameClassName = "x-grid3-col x-grid3-cell x-grid3-td-NEWNAME_JSID ";

function getOldNames() {
  var oldNamesList =
    document.getElementById(tableId).getElementsByClassName(oldNameClassName);

  var oldNames = [];
  let i = 0;
  [].forEach.call(oldNamesList, function(v, i) {
    oldNames[i] = v.children[0].innerText;
  });

  return oldNames;
}

function getNewNames() {
  var oldNames = getOldNames();
  var newNames = [];
  for (var i = 0; i < oldNames.length; i++) {
    newNames[i] = changeLetters(oldNames[i]);
  }

  return newNames;
}

function changeLetters(oldName) {
  var polishLetters = ["Ą", "Ć", "Ę", "Ł", "Ń", "Ó", "Ś", "Ź", "Ż", "ą", "ć", "ę", "ł", "ń", "ó", "ś", "ź", "ż"];
  var englishLetters = ["A", "C", "E", "L", "N", "O", "S", "Z", "Z", "a", "c", "e", "l", "n", "o", "s", "z", "z"];

  var newName = oldName;
  for (var i = 0; i < polishLetters.length; i++) {
    var re = new RegExp(polishLetters[i], "g");
    newName = newName.replace(re, englishLetters[i]);
  }

  return newName;
}

function setNewNames() {
  var newNamesFields =
    document.getElementById(tableId).getElementsByClassName(newNameClassName);

  var newNames = getNewNames();

  let i = 0;
  [].forEach.call(newNamesFields, function(v, i) {
    if (typeof v.children[0].children[0] != "undefined") {
      v.children[0].children[0].value = newNames[i];
      v.children[0].children[0].onblur;
    }
  });
}

setNewNames();
