javascript: (function () {
    var clipboard = window.clipboardData.getData('text');
    var withoutDashes = clipboard.replace(/-/g, '') + '*';
    var number = withoutDashes.replace(/\s/g, '');
    document.getElementById("gloabalSearchField").value = number;
    var globalSearch = document.getElementById("globalSearch");
    globalSearch.getElementsByTagName("img")[0].click();
})();