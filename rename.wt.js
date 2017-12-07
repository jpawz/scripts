const oldNameClassName = "x-grid3-col x-grid3-cell x-grid3-td-objName ";
const newNameClassName = "x-grid3-col x-grid3-cell x-grid3-td-NEWNAME_JSID ";
const iframeName = "lbContentIframe";
const gridClass = "x-grid3-body";

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
  var rows = window.frames[iframeName].document
    .getElementsByClassName(gridClass)[0].childNodes;

  for (var i = 0; i < rows.length; i++) {
    var oldName = rows[i].getElementsByClassName(oldNameClassName)[0].innerText;
    var newName = changeLetters(oldName);
    var newNameField = rows[i].getElementsByClassName(newNameClassName)[0]
      .children[0].children[0];

    if ((typeof newNameField != "undefined") && oldName != newName) {
      newNameField.value = newName;
      newNameField.onblur();
    }
  }
}

setNewNames();
