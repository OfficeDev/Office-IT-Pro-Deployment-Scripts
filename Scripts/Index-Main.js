$(document).ready(function() {

    var gitHubReadme = null;

    $("#btViewOnGitHub").click(function () {
        window.open("https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts");
        return false;
    });

    $("#btDownloadZip").click(function () {
        window.open("https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/zipball/master");
        return false;
    });

    $("#gitHubImg").click(function () {
        window.open("https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/blob/master/README.md");
        return false;
    });

    $("#xmlEditorLabel").click(function() {
        window.open("./XmlEditor.html");
        return false;
    });

    $(window).resize(function () {
        resizeWindow();
    });

    resizeWindow();
});

function resizeWindow() {
    var bodyHeight = window.innerHeight;
    var bodyWidth = window.innerWidth;

    var bodyDiv = $("#bodyDiv");

    var d = document.getElementById('bodyDiv');
    var newLeft = ((bodyWidth / 2) - (bodyDiv.width() / 2));
    if (newLeft <= 0) {
        newLeft = 0;
    }
    d.style.left = newLeft + "px";

    var textDiv = document.getElementById('textDiv');
    var imgDiv = document.getElementById('imgDiv');

    var imgWidth = imgDiv.clientWidth;

    var textWidth = (bodyWidth - imgWidth);
    if (textWidth > 700) {
        textWidth = 700;
    }

    if (textWidth <= 500) {
        textWidth = 500;
    }

    var imageLeft = (bodyWidth - textWidth);
    if (imageLeft <= textWidth) {
        imageLeft = textWidth;
    }

    if (imageLeft >= textWidth) {
        imageLeft = textWidth;
    }

    imgDiv.style.position = "absolute";
    imgDiv.style.left = (imageLeft) + "px";

    textDiv.style.width = (textWidth - 20) + "px";
}
