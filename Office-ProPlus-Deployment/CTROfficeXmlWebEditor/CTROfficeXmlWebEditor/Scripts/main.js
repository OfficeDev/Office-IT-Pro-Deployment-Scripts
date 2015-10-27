
var selectDate;

$(document).ready(function () {
    
    $('#txtDeadline').datetimepicker({
        sideBySide: true
    }).on("dp.change", function (e) {
        var date = e.date; //e.date is a moment object
        if (date) {
            selectDate = date.format("MM/DD/YYYY HH:mm");
            var target = $(e.target).attr('name');
        }
    });

    var finput = document.getElementById('fileInput');
    finput.addEventListener('change', function (e) {
        fileUploaded(e);

        document.getElementById("fileUploadForm").reset();

    });

    document.getElementById("collapseOne").style.display = "block";
    document.getElementById("collapseProperties").style.display = "block";

    document.getElementById("pidkeySignal").style.display = "none";
    document.getElementById("targetversionSignal").style.display = "none";
    document.getElementById("updatepathSignal").style.display = "none";

    if (isInternetExplorer()) {
        document.getElementById("txtVersion").style.lineHeight = "0px";
        document.getElementById("txtTargetVersion").style.lineHeight = "0px";
    }

    $("#btRemoveProduct").prop("disabled", true);
    $("#btAddLanguage").prop("disabled", true);
    $("#btRemoveLanguage").prop("disabled", true);

    toggleTextBox("txtUpdatePath", false);
    toggleTextBox("txtTargetVersion", false);
    $("#inputDeadline").prop("disabled", true);

    var collapse = $.cookie("optionalcollapse");
    if (collapse == "true") {
        $("#toggleExpand").removeClass('glyphicon-collapse-down');
        $("#toggleExpand").addClass('glyphicon-collapse-up');
        $("#collapseOne").collapse('show');
        $("#collapseOne").css("height", "");
    } else {
        $("#toggleExpand").removeClass('glyphicon-collapse-up');
        $("#toggleExpand").addClass('glyphicon-collapse-down');
        $("#collapseOne").collapse('hide');
        $("#collapseOne").css("height", "0");
    }

    var collapseProperties = $.cookie("propertiescollapse");
    if (collapseProperties == "true") {
        $("#togglePropertiesExpand").removeClass('glyphicon-collapse-down');
        $("#togglePropertiesExpand").addClass('glyphicon-collapse-up');
        $("#collapseProperties").collapse('show');
        $("#collapseProperties").css("height", "");
    } else {
        $("#togglePropertiesExpand").removeClass('glyphicon-collapse-up');
        $("#togglePropertiesExpand").addClass('glyphicon-collapse-down');
        $("#collapseProperties").collapse('hide');
        $("#collapseProperties").css("height", "0");
    }

    $('#templateList li').click(function (e) {
        e.preventDefault();
        var $that = $(this);
        var url = document.getElementById(this.id).getAttribute("href");
        
        var rawFile = new XMLHttpRequest();
        rawFile.open("GET", url, true);
        rawFile.onreadystatechange = function () {
            if (rawFile.readyState === 4) {
                var allText = rawFile.responseText;
                if (allText) {
                    $('textarea#xmlText').val(allText);
                    loadUploadXmlFile();
                }
            }
        }

        rawFile.send();

    });

    $("#collapseOne").prop("height", "auto");
    $("#collapseProperties").prop("height", "auto");

    setActiveTab();

    resizeWindow();

    var xmlOutput = $.cookie("xmlcache");
    $('textarea#xmlText').val(xmlOutput);
    loadUploadXmlFile();

    $(window).resize(function () {
        resizeWindow();
    });

    //$(".alert").addClass("in").fadeOut(4500);

    $('#txtPidKey').keydown(function (e) {
        var currentText = this.value;
        var code = e.keyCode || e.which;

        var start = document.getElementById("txtPidKey").selectionStart;
        var end = document.getElementById("txtPidKey").selectionEnd;
        
        if (code == 189) {
            if (start != 5 && start != 11 && start != 17 && start != 23) {
                e.preventDefault();
            }
        }
        
        if (code == 8 || code == 46) {
            if (end < currentText.length) {
                var selPart = currentText.substring(start - 1, end);
                if (selPart.indexOf("-") > -1) {
                    e.preventDefault();
                }
            }
        }

        if (code == 8 || (code >= 37 && code <= 40)) return;
        if (code == 46 || code == 16) return;
        
        var strSplit = currentText.split('-');
        for (var t = 0; t < strSplit.length; t++) {
            var part = strSplit[t];
            if (part.length > 5) {
                //e.preventDefault();
            }
        }

        if (currentText.length > 28) {
            e.preventDefault();
        }
    });

    $('#txtPidKey').keyup(function () {
        validatePidKey(this);

        var currentText = this.value;
        if (currentText.length >= 27) {

            while (currentText.indexOf("-") > -1) {
                currentText = currentText.replace("-", "");
            }

            var newCode = currentText.substring(0, 5) + "-" +
                          currentText.substring(5, 10) + "-" +
                          currentText.substring(10, 15) + "-" +
                          currentText.substring(15, 20) + "-" +
                          currentText.substring(20, 25);

            var start = document.getElementById("txtPidKey").selectionStart;
            var end = document.getElementById("txtPidKey").selectionEnd;
            this.value = newCode;
            document.getElementById("txtPidKey").selectionStart = start;
            document.getElementById("txtPidKey").selectionEnd = end;
        }
    });


    $('#txtPACKAGEGUID').keydown(function (e) {
        var currentText = this.value;
        var code = e.keyCode || e.which;

        var start = document.getElementById("txtPACKAGEGUID").selectionStart;
        var end = document.getElementById("txtPACKAGEGUID").selectionEnd;

        if (code == 189) {
            if (start != 5 && start != 11 && start != 17 && start != 23) {
                e.preventDefault();
            }
        }

        if (code == 8 || code == 46) {
            if (end < currentText.length) {
                var selPart = currentText.substring(start - 1, end);
                if (selPart.indexOf("-") > -1) {
                    e.preventDefault();
                }
            }
        }

        if (code == 8 || (code >= 37 && code <= 40)) return;
        if (code == 46 || code == 16) return;

        var strSplit = currentText.split('-');
        for (var t = 0; t < strSplit.length; t++) {
            var part = strSplit[t];
            if (part.length > 5) {
                //e.preventDefault();
            }
        }

        if (currentText.length > 36) {
            e.preventDefault();
        }
    });

    $('#txtPACKAGEGUID').keyup(function () {
        validatePackageGuid(this);

        var currentText = this.value;
        if (currentText.length >= 31) {

            while (currentText.indexOf("-") > -1) {
                currentText = currentText.replace("-", "");
            }

            var newCode = currentText.substring(0, 8) + "-" +
                          currentText.substring(8, 12) + "-" +
                          currentText.substring(12, 16) + "-" +
                          currentText.substring(16, 20) + "-" +
                          currentText.substring(20, 32);

            var start = document.getElementById("txtPACKAGEGUID").selectionStart;
            var end = document.getElementById("txtPACKAGEGUID").selectionEnd;
            this.value = newCode;
            document.getElementById("txtPACKAGEGUID").selectionStart = start;
            document.getElementById("txtPACKAGEGUID").selectionEnd = end;
        }
    });


    $('txtPidKey').on('input propertychange paste focus click', function () {
        if (this.value.length == 0) {
            document.getElementById("pidkeySignal").style.display = "none";
        } else {
            document.getElementById("pidkeySignal").style.display = "block";
        }
    });

    $('#txtVersion').on('input propertychange paste focus click', function () {
        if (this.value.length == 0) {
            document.getElementById("versionSignal").style.display = "none";
        } else {
            document.getElementById("versionSignal").style.display = "block";
        }
    });

    $('#txtPACKAGEGUID').on('input propertychange paste focus click', function () {
        if (this.value.length == 0) {
            document.getElementById("PACKAGEGUIDSignal").style.display = "none";
        } else {
            document.getElementById("PACKAGEGUIDSignal").style.display = "block";
        }
    });

    $('#txtSourcePath').on('input propertychange paste focus click', function () {
        if (this.value.length == 0) {
            document.getElementById("sourcepathSignal").style.display = "none";
        } else {
            document.getElementById("sourcepathSignal").style.display = "block";
        }
    });

    $('#txtUpdatePath').on('input propertychange paste focus click', function () {
        if (this.value.length == 0) {
            document.getElementById("updatepathSignal").style.display = "none";
        } else {
            document.getElementById("updatepathSignal").style.display = "block";
        }
    });

    $('#txtTargetVersion').on('input propertychange paste focus click', function () {
        if (this.value.length == 0) {
            document.getElementById("targetversionSignal").style.display = "none";
        } else {
            document.getElementById("targetversionSignal").style.display = "block";
        }
    });

    $('#txtLoggingUpdatePath').on('input propertychange paste focus click', function () {
        if (this.value.length == 0) {
            document.getElementById("logupdatepathSignal").style.display = "none";
        } else {
            document.getElementById("logupdatepathSignal").style.display = "block";
        }
    });

    $("#btAddProduct").button().click(function () {
        var xmlDoc = getXmlDocument();

        odtAddProduct(xmlDoc);

        displayXml(xmlDoc);

        $("#btAddProduct").text('Edit Product');

        return false;
    });

    $("#btRemoveProduct").button().click(function () {
        var xmlDoc = getXmlDocument();

        odtRemoveProduct(xmlDoc);

        displayXml(xmlDoc);

        return false;
    });

    $("#btAddLanguage").button().click(function () {
        var xmlDoc = getXmlDocument();

        odtAddLanguage(xmlDoc);

        displayXml(xmlDoc);
        return false;
    });

    $("#btRemoveLanguage").button().click(function () {
        var xmlDoc = getXmlDocument();

        odtRemoveLanguage(xmlDoc);

        displayXml(xmlDoc);
        return false;
    });

    $("#cbProduct").change(function () {
        var end = this.value;
        changeSelectedProduct();
    });

    $("#cbLanguage").change(function () {
        var end = this.value;
        changeSelectedLanguage();
    });

    $('a[data-toggle="tab"]').on('shown.bs.tab', function (e) {
        //e.target // activated tab
        //e.relatedTarget // previous tab
        scrollXmlEditor();

        $.cookie("activeTab", e.target);

        var mainTabs = document.getElementById("myTab");
        if (mainTabs) {
            var target = $(e.target).attr("href");
            var liItems = mainTabs.getElementsByTagName("li");
            if (liItems) {
                
                for (var i = 0; i < liItems.length; i++) {
                    var liItem = liItems[i];
                    if ($("#" + liItem.id).hasClass("active")) {
                        
                    }
                }
            }
        }
    });

    $("#btAddExcludeApp").button().click(function () {
        var xmlDoc = getXmlDocument();

        odtAddExcludeApp(xmlDoc);

        displayXml(xmlDoc);
        return false;
    });

    $("#btRemoveExcludeApp").button().click(function () {
        var xmlDoc = getXmlDocument();

        odtRemoveExcludeApp(xmlDoc);

        displayXml(xmlDoc);
        return false;
    });

    $("#btAddRemoveProduct").button().click(function () {
        var xmlDoc = getXmlDocument();

        odtAddRemoveApp(xmlDoc);

        displayXml(xmlDoc);
        return false;
    });

    $("#btDeleteRemoveProduct").button().click(function () {
        var xmlDoc = getXmlDocument();

        odtDeleteRemoveApp(xmlDoc);

        displayXml(xmlDoc);
        return false;
    });

    $("#btAddRemoveLanguage").button().click(function () {
        var xmlDoc = getXmlDocument();

        odtAddRemoveLanguage(xmlDoc);

        displayXml(xmlDoc);
        return false;
    });

    $("#btRemoveRemoveLanguage").button().click(function () {
        var xmlDoc = getXmlDocument();

        odtRemoveRemoveLanguage(xmlDoc);

        displayXml(xmlDoc);
        return false;
    });

    $("#btSaveUpdates").button().click(function () {
        var xmlDoc = getXmlDocument();

        odtSaveUpdates(xmlDoc);

        displayXml(xmlDoc);
        return false;
    });

    $("#btRemovesUpdates").button().click(function () {
        var xmlDoc = getXmlDocument();

        odtRemoveUpdates(xmlDoc);

        displayXml(xmlDoc);
        return false;
    });

    $("#btSaveDisplay").button().click(function () {
        var xmlDoc = getXmlDocument();

        odtSaveDisplay(xmlDoc);

        displayXml(xmlDoc);
        return false;
    });

    $("#btRemoveDisplay").button().click(function () {
        var xmlDoc = getXmlDocument();

        odtRemoveDisplay(xmlDoc);

        displayXml(xmlDoc);
        return false;
    });

    $("#btSaveLogging").button().click(function () {
        var xmlDoc = getXmlDocument();

        odtSaveLogging(xmlDoc);

        displayXml(xmlDoc);
        return false;
    });

    $("#btRemoveLogging").button().click(function () {
        var xmlDoc = getXmlDocument();

        odtRemoveLogging(xmlDoc);

        displayXml(xmlDoc);
        return false;
    });

    $("#btSaveProperties").button().click(function () {
        var xmlDoc = getXmlDocument();

        odtSaveProperties(xmlDoc);

        displayXml(xmlDoc);
        return false;
    });

    $("#btRemoveProperties").button().click(function () {
        var xmlDoc = getXmlDocument();

        odtRemoveProperties(xmlDoc);

        displayXml(xmlDoc);
        return false;
    });

    $("#btViewOnGitHub").button().click(function () {
        window.open("https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/tree/master/Office-ProPlus-Deployment/CTROfficeXmlWebEditor");

        return false;
    });

    $(window).scroll(function() {
        scrollXmlEditor();
    });
    
    $('#the-basics .typeahead').typeahead({
        hint: true,
        highlight: true,
        minLength: 1
    },
    {
        name: 'versions',
        source: substringMatcher(versions)
    });

    $('#txtVersion').keydown(function (e) {
        restrictToVersion(e);
    });

    $('#txtTargetVersion').keydown(function (e) {
        restrictToVersion(e);
    });

    setScrollBar();

});

function setScrollBar() {
    var optionDiv = document.getElementById("optionDiv");
    if (optionDiv) {
        var optionDivHeight = optionDiv.clientHeight;
        var bodyHeight = window.innerHeight;

        if (optionDivHeight > bodyHeight - 100) {
            document.body.style.overflow = "auto";
        } else {
            if (isInternetExplorer()) {
                document.body.style.overflow = "hidden";
            }
        }
    }
}

function scrollXmlEditor() {
    var scrollTop = $(window).scrollTop();
    var bodyWidth = window.innerWidth;

    var xmlEditorDiv = document.getElementById("xmlEditorDiv");
    if (xmlEditorDiv) {
        var clientLeft = getPos(xmlEditorDiv).x;

        if (clientLeft > 50) {
            if (scrollTop > 50) {
                xmlEditorDiv.style.top = (scrollTop - 50) + "px";
            } else {
                xmlEditorDiv.style.top = (0) + "px";
            }
        } else {
            xmlEditorDiv.style.top = (0) + "px";
        }
    }
}

function getPos(el) {
    // yay readability
    for (var lx = 0, ly = 0;
         el != null;
         lx += el.offsetLeft, ly += el.offsetTop, el = el.offsetParent);
    return { x: lx, y: ly };
}

function isInternetExplorer() {
    if (window.ActiveXObject || "ActiveXObject" in window) {
        return true;
    }
    return false;
}

var substringMatcher = function (strs) {
    return function findMatches(q, cb) {
        try {
            var matches, substringRegex;

            // an array that will be populated with substring matches
            matches = [];

            // regex used to determine if a string contains the substring `q`
            substrRegex = new RegExp(q, 'i');

            // iterate through the pool of strings and for any string that
            // contains the substring `q`, add it to the `matches` array
            $.each(strs, function(i, str) {
                if (substrRegex.test(str)) {
                    matches.push(str);
                }
            });

            cb(matches);
        } catch(ex) {}
    };
};

function restrictToVersion(e) {
    var currentText = this.value;
    var code = e.keyCode || e.which;

    var start = document.getElementById("txtPidKey").selectionStart;
    var end = document.getElementById("txtPidKey").selectionEnd;

    if ((code >= 48 && code <= 57) || code == 190 || code == 8
       || code == 46 || (e.ctrlKey && code == 67) || code == 17
       || (e.ctrlKey && code == 86) || (code >= 37 && code <= 40)) {

    } else {
        e.preventDefault();
    }
}

function setActiveTab() {
    var activeTab = $.cookie("activeTab");

    if (activeTab) {
        if (activeTab.indexOf('#') > -1) {
            var tabSplit = activeTab.split('#');
            activeTab = tabSplit[tabSplit.length - 1];
        }
        $('[data-toggle="tab"][href="#' + activeTab + '"]').tab('show');
    }
}

function clickUpload() {
    var finput = document.getElementById('fileInput');
    finput.click();
}

function fileUploaded(e) {
    var control = document.getElementById('fileInput');

    var i = 0,
    files = control.files;
    var file = files[i];

    var reader = new FileReader();
    reader.onload = function (event) {
        var contents = event.target.result;
        var xmlOutput = vkbeautify.xml(contents);

        $('textarea#xmlText').val(xmlOutput);

        loadUploadXmlFile();
    };
    reader.onerror = function (event) {
        throw "File could not be read! Code " + event.target.error.code;
    };
    reader.readAsText(file);
}

function toggleExpandOptional(source) {

    if ($("#toggleExpand").hasClass('glyphicon-collapse-up')) {
        $("#toggleExpand").removeClass('glyphicon-collapse-up');
        $("#toggleExpand").addClass('glyphicon-collapse-down');
        $.cookie("optionalcollapse", "false");
    } else {
        $("#toggleExpand").addClass('glyphicon-collapse-up');
        $("#toggleExpand").removeClass('glyphicon-collapse-down');
        $.cookie("optionalcollapse", "true");
    }

    setTimeout(setScrollBar, 500);
}

function toggleExpandProperties(source) {

    if ($("#togglePropertiesExpand").hasClass('glyphicon-collapse-up')) {
        $("#togglePropertiesExpand").removeClass('glyphicon-collapse-up');
        $("#togglePropertiesExpand").addClass('glyphicon-collapse-down');
        $.cookie("propertiescollapse", "false");
    } else {
        $("#togglePropertiesExpand").addClass('glyphicon-collapse-up');
        $("#togglePropertiesExpand").removeClass('glyphicon-collapse-down');
        $.cookie("propertiescollapse", "true");
    }

    setTimeout(setScrollBar, 500);
}

function download() {
    var xmlDoc = getXmlDocument();
    var xmlString = (new XMLSerializer().serializeToString(xmlDoc.documentElement));
    var xmlOutput = vkbeautify.xml(xmlString);

    xmlOutput = xmlOutput.replace(/\n/g, "\r\n");

    var blob = new Blob([xmlOutput], { type: "text/xml" });
    saveAs(blob, "configuration.xml");
}

function validatePidKey(t) {
    //if (!this.value.match(/[0-9]/)) {
    //    this.value = this.value.replace(/[^0-9]/g, '');
    //}

    var firstPart = "";
    var secondPart = "";
    var thirdPart = "";
    var fourthPart = "";
    var fifthPart = "";

    var currentText = t.value;

    if (currentText.indexOf("--") > -1) {
        var start = document.getElementById("txtPidKey").selectionStart;
        var end = document.getElementById("txtPidKey").selectionEnd;
        t.value = t.value.replace("--", "-");
        document.getElementById("txtPidKey").selectionStart = start;
        document.getElementById("txtPidKey").selectionEnd = end;
        return;
    }

    if (currentText.length > 5) {
        firstPart = currentText.substring(0, 5);
        if (firstPart.indexOf("-") > -1) return;

        var dash1 = currentText.substring(5, 6);
        if (dash1 != "-") {
            var firstPart1 = currentText.substring(0, 5);
            var restPart1 = currentText.substring(5, currentText.length);

            firstPart1 = firstPart1.replace("-", "");

            firstPart = firstPart1 + "-" + restPart1;

            var startPos = document.getElementById("txtPidKey").selectionStart;

            t.value = firstPart;
            t.value = t.value.replace("--", "-");

            document.getElementById("txtPidKey").selectionStart = startPos + 1;
            document.getElementById("txtPidKey").selectionEnd = startPos + 1;
        }
    }

    if (currentText.length > 11) {
        secondPart = currentText.substring(6, 11);
        if (secondPart.indexOf("-") > -1) return;

        var dash2 = currentText.substring(11, 12);
        if (dash2 != "-" & dash2 != "") {
            var firstPart2 = currentText.substring(11, 15);
            var restPart2 = currentText.substring(15, currentText.length);

            if (restPart2) {
                thirdPart = firstPart2 + "-" + restPart2;
            } else {
                thirdPart = firstPart2;
            }
            
            var startPos = document.getElementById("txtPidKey").selectionStart;
            
            t.value = firstPart + "-" + secondPart + "-" + thirdPart;
            t.value = t.value.replace("--", "-");

            document.getElementById("txtPidKey").selectionStart = startPos + 1;
            document.getElementById("txtPidKey").selectionEnd = startPos + 1;
        }
    }

    if (currentText.length > 17) {
        thirdPart = currentText.substring(12, 17);;
        if (thirdPart.indexOf("-") > -1) return;

        var dash3 = currentText.substring(17, 18);
        if (dash3 != "-" & dash3 != "") {
            var firstPart3 = currentText.substring(17, 21);
            var restPart3 = currentText.substring(21, currentText.length);

            if (restPart3) {
                fourthPart = firstPart3 + "-" + restPart3;
            } else {
                fourthPart = firstPart3;
            }

            var startPos = document.getElementById("txtPidKey").selectionStart;

            t.value = firstPart + "-" + secondPart + "-" + thirdPart + "-" + fourthPart;
            t.value = t.value.replace("--", "-");

            document.getElementById("txtPidKey").selectionStart = startPos + 1;
            document.getElementById("txtPidKey").selectionEnd = startPos + 1;
        }
    }

    if (currentText.length > 23) {
        fourthPart = currentText.substring(18, 23);;
        if (fourthPart.indexOf("-") > -1) return;

        var dash4 = currentText.substring(23, 24);
        if (dash4 != "-" & dash4 != "") {
            var firstPart4 = currentText.substring(23, 27);
            var restPart4 = currentText.substring(27, currentText.length);

            if (restPart4) {
                fifthPart = firstPart4 + "-" + restPart4;
            } else {
                fifthPart = firstPart4;
            }

            var startPos = document.getElementById("txtPidKey").selectionStart;

            t.value = firstPart + "-" + secondPart + "-" + thirdPart + "-" + fourthPart + "-" + fifthPart;
            t.value = t.value.replace("--", "-");

            document.getElementById("txtPidKey").selectionStart = startPos + 1;
            document.getElementById("txtPidKey").selectionEnd = startPos + 1;
        }
    }


   

}

function validatePackageGuid(t) {
    //if (!this.value.match(/[0-9]/)) {
    //    this.value = this.value.replace(/[^0-9]/g, '');
    //}

    var firstPart = "";
    var secondPart = "";
    var thirdPart = "";
    var fourthPart = "";
    var fifthPart = "";

    var currentText = t.value;

    if (currentText.indexOf("--") > -1) {
        var start = document.getElementById("txtPACKAGEGUID").selectionStart;
        var end = document.getElementById("txtPACKAGEGUID").selectionEnd;
        t.value = t.value.replace("--", "-");
        document.getElementById("txtPACKAGEGUID").selectionStart = start;
        document.getElementById("txtPACKAGEGUID").selectionEnd = end;
        return;
    }

    if (currentText.length > 8) {
        firstPart = currentText.substring(0, 8);
        if (firstPart.indexOf("-") > -1) return;

        var dash1 = currentText.substring(8, 9);
        if (dash1 != "-") {
            var firstPart1 = currentText.substring(0, 8);
            var restPart1 = currentText.substring(8, currentText.length);

            firstPart1 = firstPart1.replace("-", "");

            firstPart = firstPart1 + "-" + restPart1;

            var startPos = document.getElementById("txtPACKAGEGUID").selectionStart;

            t.value = firstPart;
            t.value = t.value.replace("--", "-");

            document.getElementById("txtPACKAGEGUID").selectionStart = startPos + 1;
            document.getElementById("txtPACKAGEGUID").selectionEnd = startPos + 1;
        }
    }

    if (currentText.length > 13) {
        secondPart = currentText.substring(9, 13);
        if (secondPart.indexOf("-") > -1) return;

        var dash2 = currentText.substring(13, 14);
        if (dash2 != "-" & dash2 != "") {
            var firstPart2 = currentText.substring(13, 17);
            var restPart2 = currentText.substring(17, currentText.length);

            if (restPart2) {
                thirdPart = firstPart2 + "-" + restPart2;
            } else {
                thirdPart = firstPart2;
            }

            var startPos = document.getElementById("txtPACKAGEGUID").selectionStart;

            t.value = firstPart + "-" + secondPart + "-" + thirdPart;
            t.value = t.value.replace("--", "-");

            document.getElementById("txtPACKAGEGUID").selectionStart = startPos + 1;
            document.getElementById("txtPACKAGEGUID").selectionEnd = startPos + 1;
        }
    }

    if (currentText.length > 18) {
        thirdPart = currentText.substring(14, 18);;
        if (thirdPart.indexOf("-") > -1) return;

        var dash3 = currentText.substring(18, 19);
        if (dash3 != "-" & dash3 != "") {
            var firstPart3 = currentText.substring(18, 22);
            var restPart3 = currentText.substring(22, currentText.length);

            if (restPart3) {
                fourthPart = firstPart3 + "-" + restPart3;
            } else {
                fourthPart = firstPart3;
            }

            var startPos = document.getElementById("txtPACKAGEGUID").selectionStart;

            t.value = firstPart + "-" + secondPart + "-" + thirdPart + "-" + fourthPart;
            t.value = t.value.replace("--", "-");

            document.getElementById("txtPACKAGEGUID").selectionStart = startPos + 1;
            document.getElementById("txtPACKAGEGUID").selectionEnd = startPos + 1;
        }
    }

    if (currentText.length > 23) {
        fourthPart = currentText.substring(19, 23);
        if (fourthPart.indexOf("-") > -1) return;

        var dash4 = currentText.substring(23, 24);
        if (dash4 != "-" & dash4 != "") {
            var firstPart4 = currentText.substring(23, 27);
            var restPart4 = currentText.substring(27, currentText.length);

            if (restPart4) {
                fifthPart = firstPart4 + "-" + restPart4;
            } else {
                fifthPart = firstPart4;
            }

            var startPos = document.getElementById("txtPACKAGEGUID").selectionStart;

            t.value = firstPart + "-" + secondPart + "-" + thirdPart + "-" + fourthPart + "-" + fifthPart;
            t.value = t.value.replace("--", "-");

            document.getElementById("txtPACKAGEGUID").selectionStart = startPos + 1;
            document.getElementById("txtPACKAGEGUID").selectionEnd = startPos + 1;
        }
    }




}


function changeSelectedLanguage() {
    var selectedProduct = $("#cbProduct").val();
    var selectLanguage = $("#cbLanguage").val();

    var xmlDoc = getXmlDocument();
            
    $("#btAddLanguage").prop("disabled", false);

    var addNode = null;
    var nodes = xmlDoc.documentElement.getElementsByTagName("Add");
    if (nodes.length > 0) {
        addNode = xmlDoc.documentElement.getElementsByTagName("Add")[0];
    }

    if (addNode) {
        var productNode = getProductNode(addNode, selectedProduct);

        var langNode = getLanguageNode(productNode, selectLanguage);
        if (langNode) {
            $("#btAddLanguage").prop("disabled", true);
        }
    }
}

function changeSelectedProduct() {
    var productId = $('#cbProduct').val();

    var xmlDoc = getXmlDocument();

    $("#txtPidKey").val("");

    var productCount = getAddProductCount(xmlDoc);
    if (productCount > 0) {
        var productNode = getProductNode(xmlDoc, productId);
        if (productNode) {
            $("#btAddProduct").text('Edit Product');
            $("#btRemoveProduct").prop("disabled", false);

            var pidKey = productNode.getAttribute("PIDKEY");
            if (pidKey) {
                $("#txtPidKey").val(pidKey);
            }

            var excludeApps = productNode.getElementsByTagName("ExcludeApp");
            if (excludeApps.length == 0) {
                $("#btRemoveExcludeApp").prop("disabled", true);
                $("select#cbExcludeApp").prop('selectedIndex', 0);
            } else {
                $("#btRemoveExcludeApp").prop("disabled", false);

                var excludeApp = excludeApps[0];
                if (excludeApp) {
                    $("#cbExcludeApp").val(excludeApp.getAttribute("ID"));
                }
            }

        } else {
            $("#btAddProduct").text('Add Product');
            //$("#btRemoveProduct").prop("disabled", true);
            $("#btRemoveExcludeApp").prop("disabled", true);
            $("select#cbExcludeApp").prop('selectedIndex', 0);
        }
    } else {
        $("#btRemoveProduct").prop("disabled", true);
    }

    var langCount = getLanguageNodeCount(xmlDoc, productId);
    $("#btRemoveLanguage").prop("disabled", !(langCount > 1));
}


function resizeWindow() {
    var bodyHeight = window.innerHeight;
    var bodyWidth = window.innerWidth;
    var leftPaneHeight = bodyHeight - 180;

    var rightPaneHeight = bodyHeight - 100;

    var scrollTop = $(window).scrollTop();
    if (scrollTop > 50) {
        rightPaneHeight = rightPaneHeight + 50;
    }

    $("#xmlText").height(rightPaneHeight - 90);

    setScrollBar();
}

function cacheNodes(xmlDoc) {
    var propertyNameList = ["Remove", "Display", "Logging", "Property", "Updates"];
    var nodeList = [];

    for (var p = 0; p < propertyNameList.length; p++) {
        var propertyName = propertyNameList[p];

        var nodes = xmlDoc.documentElement.getElementsByTagName(propertyName);
        if (nodes.length > 0) {
            var nodeCount = nodes.length;
            for (var n = 0; n < nodeCount; n++) {
                var propNode = xmlDoc.documentElement.getElementsByTagName(propertyName)[0];
                nodeList.push(propNode);
                xmlDoc.documentElement.removeChild(propNode);
            }
        }

    }

    return nodeList;
}

function readdNodes(xmlDoc, nodeList) {
    for (var t = 0; t < nodeList.length; t++) {
        var addPropNode = nodeList[t];
        xmlDoc.documentElement.appendChild(addPropNode);
    }
}

function odtAddProduct(xmlDoc) {
    var selectedProduct = $("#cbProduct").val();
    var selectBitness = $("#cbEdition").val();
    var selectVersion = $("#txtVersion").val();
    var selectSourcePath = $("#txtSourcePath").val();
    var selectLanguage = $("#cbLanguage").val();
    var selectPidKey = $("#txtPidKey").val();

    var addNode = xmlDoc.createElement("Add");
    var nodes = xmlDoc.documentElement.getElementsByTagName("Add");
    if (nodes.length > 0) {
        addNode = xmlDoc.documentElement.getElementsByTagName("Add")[0];
    } else {
        var nodeList = cacheNodes(xmlDoc);

        xmlDoc.documentElement.appendChild(addNode);

        readdNodes(xmlDoc, nodeList);
    }

    if (selectSourcePath) {
        addNode.setAttribute("SourcePath", selectSourcePath);
    } else {
        addNode.removeAttribute("SourcePath");
    }

    if (selectVersion) {
        addNode.setAttribute("Version", selectVersion);
    } else {
        addNode.removeAttribute("Version");
    }

    addNode.setAttribute("OfficeClientEdition", selectBitness);

    var productNode = getProductNode(addNode, selectedProduct);
    if (!(productNode)) {
        productNode = xmlDoc.createElement("Product");
        productNode.setAttribute("ID", selectedProduct);
        addNode.appendChild(productNode);
    }

    if (selectPidKey) {
        productNode.setAttribute("PIDKEY", selectPidKey);
    } else {
        productNode.removeAttribute("PIDKEY");
    }

    var langNode = getLanguageNode(productNode, selectLanguage);
    if (!(langNode)) {
        langNode = xmlDoc.createElement("Language");
        langNode.setAttribute("ID", selectLanguage);
        productNode.appendChild(langNode);
    }

    var removeNode = null;
    var removeNodes = xmlDoc.documentElement.getElementsByTagName("Remove");
    if (removeNodes.length > 0) {
        removeNode = xmlDoc.documentElement.getElementsByTagName("Remove")[0];
    }

    if (removeNode) {
        var existingRemoveProduct = checkForRemoveProductNode(xmlDoc, selectedProduct);
        if (existingRemoveProduct) {
            removeNode.removeChild(existingRemoveProduct);
        }

        if (removeNode.childElementCount == 0) {
            xmlDoc.documentElement.removeChild(removeNode);
        }
    }

    var productCount = getAddProductCount(xmlDoc);
    if (productCount == 0) {
        $("#btRemoveProduct").prop("disabled", true);
        $("#btAddLanguage").prop("disabled", true);
        $("#btRemoveLanguage").prop("disabled", true);
    } else {
        $("#btRemoveProduct").prop("disabled", false);
        $("#btAddLanguage").prop("disabled", true);
    }
}

function odtRemoveProduct(xmlDoc) {
    var selectedProduct = $("#cbProduct").val();

    var addNode = null;

    var nodes = xmlDoc.documentElement.getElementsByTagName("Add");
    if (nodes.length > 0) {
        addNode = xmlDoc.documentElement.getElementsByTagName("Add")[0];

        var productNode = getProductNode(addNode, selectedProduct);
        if (productNode) {
            addNode.removeChild(productNode);
        }

        var products = addNode.getElementsByTagName("Product");
        if (products.length == 0) {
            addNode.parentNode.removeChild(addNode);
        }
    }

    var productCount = getAddProductCount(xmlDoc);
    if (productCount == 0) {
        $("#btRemoveProduct").prop("disabled", true);
        $("#btAddLanguage").prop("disabled", true);
        $("#btRemoveLanguage").prop("disabled", true);
    } else {
        $("#btRemoveProduct").prop("disabled", false);
        $("#btAddLanguage").prop("disabled", false);
    }

    //$("#removeAllProducts").removeClass('btn-primary');
    //$("#removeSelectProducts").removeClass('btn-primary');
    //$("#removeAllProducts").removeClass('active');
    //$("#removeSelectProducts").removeClass('active');

    $("#btAddProduct").text('Add Product');
}


function odtAddLanguage(xmlDoc) {
    var selectedProduct = $("#cbProduct").val();
    var selectLanguage = $("#cbLanguage").val();

    var addNode = xmlDoc.createElement("Add");

    var nodes = xmlDoc.documentElement.getElementsByTagName("Add");
    if (nodes.length > 0) {
        addNode = xmlDoc.documentElement.getElementsByTagName("Add")[0];

        var productNode = getProductNode(addNode, selectedProduct);
        if (productNode) {
            var langNode = getLanguageNode(productNode, selectLanguage);
            if (!(langNode)) {

                var langs = productNode.getElementsByTagName("Language");
                var lastLang = langs[langs.length - 1];

                var langList = [];
                for (var p = 0; p < langs.length; p++) {
                    var langChkNode1 = langs[p];
                    langList.push(langChkNode1);
                }

                for (var l = 0; l < langs.length; l++) {
                    var langChkNode2 = langs[l];
                    productNode.removeChild(langChkNode2);
                }

                for (var t = 0; t < langList.length ; t++) {
                    var langChkNode = langList[t];
                    productNode.appendChild(langChkNode);
                }

                langNode = xmlDoc.createElement("Language");
                langNode.setAttribute("ID", selectLanguage);
                productNode.appendChild(langNode, lastLang);

                $("#btAddLanguage").prop("disabled", true);
            }
        }
    }

    var langCount = getLanguageNodeCount(xmlDoc, selectedProduct);
    $("#btRemoveLanguage").prop("disabled", !(langCount > 1));
}

function odtRemoveLanguage(xmlDoc) {
    var selectedProduct = $("#cbProduct").val();
    var selectLanguage = $("#cbLanguage").val();

    var addNode = null;

    var nodes = xmlDoc.documentElement.getElementsByTagName("Add");
    if (nodes.length > 0) {
        addNode = xmlDoc.documentElement.getElementsByTagName("Add")[0];

        var productNode = getProductNode(addNode, selectedProduct);
        if (productNode) {
            if (getLanguageNodeCount(xmlDoc, selectedProduct) > 1) {
                var langNode = getLanguageNode(productNode, selectLanguage);
                if (langNode) {
                    productNode.removeChild(langNode);
                    $("#btAddLanguage").prop("disabled", false);
                }
            }
        }
    }

    var langCount = getLanguageNodeCount(xmlDoc, selectedProduct);
    $("#btRemoveLanguage").prop("disabled", !(langCount > 1));
}


function odtAddRemoveLanguage(xmlDoc) {
    var selectedProduct = $("#cbRemoveProduct").val();
    var selectLanguage = $("#cbRemoveLanguage").val();

    var removeNode = xmlDoc.createElement("Remove");

    var nodes = xmlDoc.documentElement.getElementsByTagName("Remove");
    if (nodes.length > 0) {
        removeNode = xmlDoc.documentElement.getElementsByTagName("Remove")[0];

        var productNode = getProductNode(removeNode, selectedProduct);
        if (productNode) {
            var langNode = getLanguageNode(productNode, selectLanguage);
            if (!(langNode)) {

                var langs = productNode.getElementsByTagName("Language");
                var lastLang = langs[langs.length - 1];

                langNode = xmlDoc.createElement("Language");
                langNode.setAttribute("ID", selectLanguage);
                productNode.insertBefore(langNode, lastLang);
            }
        }
    }

    var langCount = getRemoveLanguageNodeCount(xmlDoc, selectedProduct);
    $("#btRemoveRemoveLanguage").prop("disabled", !(langCount > 1));
}

function odtRemoveRemoveLanguage(xmlDoc) {
    var selectedProduct = $("#cbRemoveProduct").val();
    var selectLanguage = $("#cbRemoveLanguage").val();

    var removeNode = null;

    var nodes = xmlDoc.documentElement.getElementsByTagName("Remove");
    if (nodes.length > 0) {
        removeNode = xmlDoc.documentElement.getElementsByTagName("Remove")[0];

        var productNode = getProductNode(removeNode, selectedProduct);
        if (productNode) {
            if (getRemoveLanguageNodeCount(xmlDoc, selectedProduct) > 1) {
                var langNode = getLanguageNode(productNode, selectLanguage);
                if (langNode) {
                    productNode.removeChild(langNode);
                }
            }
        }
    }

    var langCount = getRemoveLanguageNodeCount(xmlDoc, selectedProduct);
    $("#btRemoveLanguage").prop("disabled", !(langCount > 1));
}

function removeAllSections() {
    document.getElementById("btRemoveDisplay").click();
    document.getElementById("btRemoveExcludeApp").click();
    document.getElementById("btRemoveLanguage").click();
    document.getElementById("btRemoveLogging").click();
    document.getElementById("btRemoveProduct").click();
    document.getElementById("btRemoveProperties").click();
    document.getElementById("btRemoveRemoveLanguage").click();
    document.getElementById("btRemovesUpdates").click();
}

function odtAddRemoveApp(xmlDoc) {
    var selectedProduct = $("#cbRemoveProduct").val();
    var selectLanguage = $("#cbRemoveLanguage").val();

    var $removeAll = $("#removeAllProducts");
    if ($removeAll.hasClass('btn-primary')) {
 
    }

    var removeNode = xmlDoc.createElement("Remove");
    var nodes = xmlDoc.documentElement.getElementsByTagName("Remove");
    if (nodes.length > 0) {
        removeNode = xmlDoc.documentElement.getElementsByTagName("Remove")[0];
    } else {
        xmlDoc.documentElement.appendChild(removeNode);
    }

    var removeAll = false;

    var $removeSelect = $("#removeSelectProducts");
    if ($removeSelect.hasClass('btn-primary')) {
        removeNode.removeAttribute("All");

        var productNode = getProductNode(removeNode, selectedProduct);
        if (!(productNode)) {
            productNode = xmlDoc.createElement("Product");
            productNode.setAttribute("ID", selectedProduct);
            removeNode.appendChild(productNode);
        }

        var langNode = getLanguageNode(productNode, selectLanguage);
        if (!(langNode)) {
            langNode = xmlDoc.createElement("Language");
            langNode.setAttribute("ID", selectLanguage);
            productNode.appendChild(langNode);
        }
    } else {
        removeAll = true;

        removeNode.setAttribute("All", "TRUE");
        if (removeNode.childElementCount > 0) {
            var products = removeNode.getElementsByTagName("Product");
            var prodLength = products.length;
            for (var v = 0; v < prodLength; v++) {
                removeNode.removeChild(products[0]);
            }
        }
    }

    var addNode = null;
    var addNodes = xmlDoc.documentElement.getElementsByTagName("Add");
    if (addNodes.length > 0) {
        addNode = xmlDoc.documentElement.getElementsByTagName("Add")[0];
    }

    if (addNode) {
        var existingAddProduct = checkForAddProductNode(xmlDoc, selectedProduct);
        if (existingAddProduct) {
            addNode.removeChild(existingAddProduct);
        }

        if (addNode.childElementCount == 0 || removeAll) {
            xmlDoc.documentElement.removeChild(addNode);
        }
    }
}

function odtDeleteRemoveApp(xmlDoc) {
    var selectedProduct = $("#cbRemoveProduct").val();
    var selectLanguage = $("#cbRemoveLanguage").val();

    var removeNode = xmlDoc.createElement("Remove");
    var nodes = xmlDoc.documentElement.getElementsByTagName("Remove");
    if (nodes.length > 0) {
        removeNode = xmlDoc.documentElement.getElementsByTagName("Remove")[0];
    } else {
        xmlDoc.documentElement.appendChild(removeNode);
    }

    var $removeSelect = $("#removeSelectProducts");
    if ($removeSelect.hasClass('btn-primary')) {
        removeNode.removeAttribute("All");

        var productNode = getProductNode(removeNode, selectedProduct);
        if (productNode) {

            removeNode.removeChild(productNode);
        }
    }

    var products = removeNode.getElementsByTagName("Product");
    if (products.length == 0) {
        removeNode.parentNode.removeChild(removeNode);
    }
}


function odtAddExcludeApp(xmlDoc) {
    var selectedProduct = $("#cbProduct").val();
    var selectExcludeApp = $("#cbExcludeApp").val();

    var addNode = null;

    var nodes = xmlDoc.documentElement.getElementsByTagName("Add");
    if (nodes.length > 0) {
        addNode = xmlDoc.documentElement.getElementsByTagName("Add")[0];

        var productNode = getProductNode(addNode, selectedProduct);
        if (productNode) {
            var exNode = getExcludeAppNode(productNode, selectExcludeApp);
            if (!(exNode)) {
                var excludeApps = productNode.getElementsByTagName("ExcludeApp");
                var excludeNode = excludeApps[excludeApps.length - 1];

                exNode = xmlDoc.createElement("ExcludeApp");
                exNode.setAttribute("ID", selectExcludeApp);

                if (excludeNode) {
                    productNode.insertBefore(exNode, excludeNode);
                } else {
                    productNode.appendChild(exNode);
                }
            }
        }
    }

    var exCount = getExcludeAppNodeCount(xmlDoc, selectedProduct);
    $("#btRemoveExcludeApp").prop("disabled", !(exCount > 0));
}

function odtRemoveExcludeApp(xmlDoc) {
    var selectedProduct = $("#cbProduct").val();
    var selectExcludeApp = $("#cbExcludeApp").val();

    var addNode = null;

    var nodes = xmlDoc.documentElement.getElementsByTagName("Add");
    if (nodes.length > 0) {
        addNode = xmlDoc.documentElement.getElementsByTagName("Add")[0];

        var productNode = getProductNode(addNode, selectedProduct);
        if (productNode) {
            if (getExcludeAppNodeCount(xmlDoc, selectedProduct) > 0) {
                var exNode = getExcludeAppNode(productNode, selectExcludeApp);
                if (exNode) {
                    productNode.removeChild(exNode);
                }
            }
        }
    }

    var langCount = getExcludeAppNodeCount(xmlDoc, selectedProduct);
    $("#btRemoveExcludeApp").prop("disabled", !(langCount > 0));
}


function getExcludeAppNode(excludeAppNode, selectedExcludeApp) {
    var exNode = null;
    var excludeApps = excludeAppNode.getElementsByTagName("ExcludeApp");
    for (var i = 0; i < excludeApps.length; i++) //looping xml childnodes
    {
        var excludeApp = excludeApps[i];
        var excludeAppId = excludeApp.getAttribute("ID");

        if (excludeAppId == selectedExcludeApp) {
            exNode = excludeApp;
        }
    }
    return exNode;
}

function getExcludeAppNodeCount(xmlDoc, productId) {
    var addNode = xmlDoc.createElement("Add");

    var nodes = xmlDoc.documentElement.getElementsByTagName("Add");
    if (nodes.length > 0) {
        addNode = xmlDoc.documentElement.getElementsByTagName("Add")[0];

        var productNode = getProductNode(addNode, productId);
        if (productNode) {
            var excludeApps = productNode.getElementsByTagName("ExcludeApp");
            return excludeApps.length;
        }
    }

    return 0;
}


function odtSaveUpdates(xmlDoc) {
    var selectUpdatePath = $("#txtUpdatePath").val();
    var selectTargetVersion = $("#txtTargetVersion").val();

    var $btUpdatesEnabled = $("#btupdatesEnabled");
    var $btUpdatesDisabled = $("#btupdatesDisabled");

    if (!$btUpdatesEnabled.hasClass('btn-primary') && !$btUpdatesDisabled.hasClass('btn-primary')) {
        $btUpdatesEnabled.addClass('btn-primary');

        $("#txtUpdatePath").prop("disabled", false);
        //$("#txtTargetVersion").prop("disabled", false);
        $("#inputDeadline").prop("disabled", false);
    }

    if ($btUpdatesEnabled.hasClass('btn-primary') || $btUpdatesDisabled.hasClass('btn-primary')) {

        var updateNode = xmlDoc.createElement("Updates");
        var nodes = xmlDoc.documentElement.getElementsByTagName("Updates");
        if (nodes.length > 0) {
            updateNode = xmlDoc.documentElement.getElementsByTagName("Updates")[0];
        } else {
            xmlDoc.documentElement.appendChild(updateNode);
        }

        if (selectUpdatePath) {
            updateNode.setAttribute("UpdatePath", selectUpdatePath);
        } else {
            updateNode.removeAttribute("UpdatePath");
        }

        if (selectTargetVersion) {
            updateNode.setAttribute("TargetVersion", selectTargetVersion);
        } else {
            updateNode.removeAttribute("TargetVersion");
        }

        if (selectDate) {
            updateNode.setAttribute("Deadline", selectDate);
        } else {
            updateNode.removeAttribute("Deadline");
        }

        if ($btUpdatesEnabled.hasClass('btn-primary')) {
            updateNode.setAttribute("Enabled", "TRUE");
        }

        if ($btUpdatesDisabled.hasClass('btn-primary')) {
            updateNode.setAttribute("Enabled", "FALSE");
            updateNode.removeAttribute("UpdatePath");
            updateNode.removeAttribute("TargetVersion");
            updateNode.removeAttribute("Deadline");
        }

    }
}

function odtRemoveUpdates(xmlDoc) {
    var addNode = xmlDoc.createElement("Updates");
    var nodes = xmlDoc.documentElement.getElementsByTagName("Updates");
    if (nodes.length > 0) {
        addNode = xmlDoc.documentElement.getElementsByTagName("Updates")[0];
        if (addNode) {
            xmlDoc.documentElement.removeChild(addNode);
        }
    }

    $("#btupdatesDisabled").removeClass('btn-primary');
    $("#btupdatesEnabled").removeClass('btn-primary');
    $("#btupdatesDisabled").removeClass('active');
    $("#btupdatesEnabled").removeClass('active');

    $("#inputDeadline").prop("disabled", true);
    toggleTextBox("txtUpdatePath", false);
    toggleTextBox("txtTargetVersion", false);
}


function odtSaveDisplay(xmlDoc) {
    var addNode = xmlDoc.createElement("Display");
    var nodes = xmlDoc.documentElement.getElementsByTagName("Display");
    if (nodes.length > 0) {
        addNode = xmlDoc.documentElement.getElementsByTagName("Display")[0];
    } else {
        xmlDoc.documentElement.appendChild(addNode);
    }

    var $displayLevelNone = $("#btLevelNone");
    var $displayLevelFull = $("#btLevelFull");
    var $AcceptEulaEnabled = $("#btAcceptEULAEnabled");
    var $AcceptEulaDisabled = $("#btAcceptEULADisabled");

    if (!$displayLevelNone.hasClass('btn-primary') && !$displayLevelFull.hasClass('btn-primary') &&
        !$AcceptEulaEnabled.hasClass('btn-primary') && !$AcceptEulaDisabled.hasClass('btn-primary')) {
        $displayLevelNone.addClass('btn-primary');
        $AcceptEulaEnabled.addClass('btn-primary');
    }

    if ($displayLevelNone.hasClass('btn-primary')) {
        addNode.setAttribute("Level", "None");
    }
    
    if ($displayLevelFull.hasClass('btn-primary')) {
        addNode.setAttribute("Level", "Full");
    }
    
    if ($AcceptEulaEnabled.hasClass('btn-primary')) {
        addNode.setAttribute("AcceptEULA", "TRUE");
    }
    
    if ($AcceptEulaDisabled.hasClass('btn-primary')) {
        addNode.setAttribute("AcceptEULA", "FALSE");
    }
}

function odtRemoveDisplay(xmlDoc) {
    var addNode = xmlDoc.createElement("Display");
    var nodes = xmlDoc.documentElement.getElementsByTagName("Display");
    if (nodes.length > 0) {
        addNode = xmlDoc.documentElement.getElementsByTagName("Display")[0];
        if (addNode) {
            xmlDoc.documentElement.removeChild(addNode);
        }
    }

    $("#btLevelNone").removeClass('btn-primary');
    $("#btLevelFull").removeClass('btn-primary');
    $("#btLevelNone").removeClass('active');
    $("#btLevelFull").removeClass('active');

    $("#btAcceptEULAEnabled").removeClass('btn-primary');
    $("#btAcceptEULAEnabled").removeClass('active');
    $("#btAcceptEULADisabled").removeClass('btn-primary');
    $("#btAcceptEULADisabled").removeClass('active');
}


function odtSaveProperties(xmlDoc) {
    var autoActivateNode = null;
    var forceShutDownNode = null;
    var sharedComputerLicensingNode = null;
    var packageguidNode = null;

    var nodes = xmlDoc.documentElement.getElementsByTagName("Property");
    if (nodes.length > 0) {
        for (var n = 0; n < nodes.length; n++) {
            propNode = xmlDoc.documentElement.getElementsByTagName("Property")[n];
            if (propNode) {
                var attrValue = propNode.getAttribute("Name");
                if (attrValue) {
                    if (propNode.getAttribute("Name").toUpperCase() == "AUTOACTIVATE") {
                        autoActivateNode = propNode;
                    }
                    if (propNode.getAttribute("Name").toUpperCase() == "FORCEAPPSHUTDOWN") {
                        forceShutDownNode = propNode;
                    }
                    if (propNode.getAttribute("Name").toUpperCase() == "SHAREDCOMPUTERLICENSING") {
                        sharedComputerLicensingNode = propNode;
                    }
                    if (propNode.getAttribute("Name").toUpperCase() == "PACKAGEGUID") {
                        packageguidNode = propNode;
                    }
                }
            }
        }
    }

    var $btAutoActivateYes = $("#btAutoActivateYes");
    var $btAutoActivateNo = $("#btAutoActivateNo");
    var $btForceAppShutdownTrue = $("#btForceAppShutdownTrue");
    var $btForceAppShutdownFalse = $("#btForceAppShutdownFalse");
    var $btSharedComputerLicensingYes = $("#btSharedComputerLicensingYes");
    var $btSharedComputerLicensingNo = $("#btSharedComputerLicensingNo");

    var packageguidVal = $("#txtPACKAGEGUID").val();
    if (packageguidVal) {
        if (packageguidVal.length > 0) {
            if (IsGuid(packageguidVal)) {
                if (!(packageguidNode)) {
                    packageguidNode = xmlDoc.createElement("Property");
                    xmlDoc.documentElement.appendChild(packageguidNode);
                }

                packageguidNode.setAttribute("Name", "PACKAGEGUID");
                packageguidNode.setAttribute("Value", packageguidVal);
            }
        }   
    }

    if ($btAutoActivateYes.hasClass('btn-primary') || $btAutoActivateNo.hasClass('btn-primary')) {
        if (!(autoActivateNode)) {
            autoActivateNode = xmlDoc.createElement("Property");
            xmlDoc.documentElement.appendChild(autoActivateNode);
        }

        if ($btAutoActivateYes.hasClass('btn-primary')) {
            autoActivateNode.setAttribute("Name", "AUTOACTIVATE");
            autoActivateNode.setAttribute("Value", "1");
        }

        if ($btAutoActivateNo.hasClass('btn-primary')) {
            autoActivateNode.setAttribute("Name", "AUTOACTIVATE");
            autoActivateNode.setAttribute("Value", "0");
        }
    }

    if ($btForceAppShutdownTrue.hasClass('btn-primary') || $btForceAppShutdownFalse.hasClass('btn-primary')) {
        if (!(forceShutDownNode)) {
            forceShutDownNode = xmlDoc.createElement("Property");
            xmlDoc.documentElement.appendChild(forceShutDownNode);
        }

        if ($btForceAppShutdownTrue.hasClass('btn-primary')) {
            forceShutDownNode.setAttribute("Name", "FORCEAPPSHUTDOWN");
            forceShutDownNode.setAttribute("Value", "TRUE");
        }

        if ($btForceAppShutdownFalse.hasClass('btn-primary')) {
            forceShutDownNode.setAttribute("Name", "FORCEAPPSHUTDOWN");
            forceShutDownNode.setAttribute("Value", "FALSE");
        }
    }

    if ($btSharedComputerLicensingYes.hasClass('btn-primary') || $btSharedComputerLicensingNo.hasClass('btn-primary')) {
        if (!(sharedComputerLicensingNode)) {
            sharedComputerLicensingNode = xmlDoc.createElement("Property");
            xmlDoc.documentElement.appendChild(sharedComputerLicensingNode);
        }

        if ($btSharedComputerLicensingYes.hasClass('btn-primary')) {
            sharedComputerLicensingNode.setAttribute("Name", "SharedComputerLicensing");
            sharedComputerLicensingNode.setAttribute("Value", "1");
        }

        if ($btSharedComputerLicensingNo.hasClass('btn-primary')) {
            sharedComputerLicensingNode.setAttribute("Name", "SharedComputerLicensing");
            sharedComputerLicensingNode.setAttribute("Value", "0");
        }
    }
}

function odtRemoveProperties(xmlDoc) {
    var propNode = null;
    var nodes = xmlDoc.documentElement.getElementsByTagName("Property");
    if (nodes.length > 0) {
        var nodeCount = nodes.length;
        for (var n = 0; n < nodeCount; n++) {
            propNode = xmlDoc.documentElement.getElementsByTagName("Property")[0];
            if (propNode) {
                xmlDoc.documentElement.removeChild(propNode);
            }
        }
    }

    $("#btAutoActivateYes").removeClass('btn-primary');
    $("#btAutoActivateNo").removeClass('btn-primary');
    $("#btAutoActivateYes").removeClass('active');
    $("#btAutoActivateNo").removeClass('active');

    $("#btForceAppShutdownTrue").removeClass('btn-primary');
    $("#btForceAppShutdownTrue").removeClass('active');
    $("#btForceAppShutdownFalse").removeClass('btn-primary');
    $("#btForceAppShutdownFalse").removeClass('active');

    $("#btSharedComputerLicensingYes").removeClass('btn-primary');
    $("#btSharedComputerLicensingYes").removeClass('active');
    $("#btSharedComputerLicensingNo").removeClass('btn-primary');
    $("#btSharedComputerLicensingNo").removeClass('active');
}


function odtSaveLogging(xmlDoc) {
    var loggingUpdatePath = $("#txtLoggingUpdatePath").val();
    var $displayLevelNone = $("#btLoggingLevelOff");
    var $displayLevelStandard = $("#btLoggingLevelStandard");

    if (!$displayLevelNone.hasClass('btn-primary') && !$displayLevelStandard.hasClass('btn-primary')) {
        $displayLevelNone.addClass('btn-primary');

        $("#txtLoggingUpdatePath").prop("disabled", true);
    }

    if ($displayLevelNone.hasClass('btn-primary') || $displayLevelStandard.hasClass('btn-primary')) {
        var addNode = xmlDoc.createElement("Logging");
        var nodes = xmlDoc.documentElement.getElementsByTagName("Logging");
        if (nodes.length > 0) {
            addNode = xmlDoc.documentElement.getElementsByTagName("Logging")[0];
        } else {
            xmlDoc.documentElement.appendChild(addNode);
        }

        if ($displayLevelNone.hasClass('btn-primary')) {
            addNode.setAttribute("Level", "Off");
            addNode.removeAttribute("Path");
        }

        if ($displayLevelStandard.hasClass('btn-primary')) {
            addNode.setAttribute("Level", "Standard");

            if (loggingUpdatePath) {
                addNode.setAttribute("Path", loggingUpdatePath);
            }
        }
    }
}

function odtRemoveLogging(xmlDoc) {
    var addNode = xmlDoc.createElement("Logging");
    var nodes = xmlDoc.documentElement.getElementsByTagName("Logging");
    if (nodes.length > 0) {
        addNode = xmlDoc.documentElement.getElementsByTagName("Logging")[0];
        if (addNode) {
            xmlDoc.documentElement.removeChild(addNode);
        }
    }

    $("#btLoggingLevelOff").removeClass('btn-primary');
    $("#btLoggingLevelOff").removeClass('active');
    $("#btLoggingLevelStandard").removeClass('btn-primary');
    $("#btLoggingLevelStandard").removeClass('active');
}


function getProductNode(addNode, selectedProduct) {
    var productNode = null;
    var products = addNode.getElementsByTagName("Product");
    for (var i = 0; i < products.length; i++) //looping xml childnodes
    {
        var product = products[i];
        var productId = product.getAttribute("ID");

        if (productId == selectedProduct) {
            productNode = product;
        }
    }
    return productNode;
}

function getAddProductCount(xmlDoc) {
    var nodes = xmlDoc.documentElement.getElementsByTagName("Add");
    if (nodes.length > 0) {
        var addNode = xmlDoc.documentElement.getElementsByTagName("Add")[0];

        var products = addNode.getElementsByTagName("Product");
        return products.length;
    }
    return 0;
}

function checkForAddProductNode(xmlDoc, selectedProduct) {
    var addNode = xmlDoc.createElement("Add");
    var nodes = xmlDoc.documentElement.getElementsByTagName("Add");
    if (nodes.length > 0) {
        addNode = xmlDoc.documentElement.getElementsByTagName("Add")[0];
    } else {
        xmlDoc.documentElement.appendChild(addNode);
    }

    var productNode = getProductNode(addNode, selectedProduct);
    return productNode;
}

function checkForRemoveProductNode(xmlDoc, selectedProduct) {
    var removeNode = xmlDoc.createElement("Remove");
    var nodes = xmlDoc.documentElement.getElementsByTagName("Remove");
    if (nodes.length > 0) {
        removeNode = xmlDoc.documentElement.getElementsByTagName("Remove")[0];
    } else {
        xmlDoc.documentElement.appendChild(removeNode);
    }

    var productNode = getProductNode(removeNode, selectedProduct);
    return productNode;
}



function getLanguageNode(productNode, selectedLanguage) {
    var langNode = null;
    var languages = productNode.getElementsByTagName("Language");
    for (var i = 0; i < languages.length; i++) //looping xml childnodes
    {
        var language = languages[i];
        var languageId = language.getAttribute("ID");

        if (languageId == selectedLanguage) {
            langNode = language;
        }
    }
    return langNode;
}

function getLanguageNodeCount(xmlDoc, productId) {
    var addNode = xmlDoc.createElement("Add");

    var nodes = xmlDoc.documentElement.getElementsByTagName("Add");
    if (nodes.length > 0) {
        addNode = xmlDoc.documentElement.getElementsByTagName("Add")[0];

        var productNode = getProductNode(addNode, productId);
        if (productNode) {
            var languages = productNode.getElementsByTagName("Language");
            return languages.length;
        }
    }

    return 0;
}

function getRemoveLanguageNodeCount(xmlDoc, productId) {
    var addNode = xmlDoc.createElement("Remove");

    var nodes = xmlDoc.documentElement.getElementsByTagName("Remove");
    if (nodes.length > 0) {
        addNode = xmlDoc.documentElement.getElementsByTagName("Remove")[0];

        var productNode = getProductNode(addNode, productId);
        if (productNode) {
            var languages = productNode.getElementsByTagName("Language");
            return languages.length;
        }
    }

    return 0;
}


function loadUploadXmlFile(inXmlDoc) {
    var xmlDoc = inXmlDoc;
    if (!(xmlDoc)) {
        xmlDoc = getXmlDocument();
    }

    var addNode = null;
    var nodes = xmlDoc.documentElement.getElementsByTagName("Add");
    if (nodes.length > 0) {
        addNode = xmlDoc.documentElement.getElementsByTagName("Add")[0];

        var selectBitness = addNode.getAttribute("OfficeClientEdition");
        $("#cbEdition").val(selectBitness);

        var products = addNode.getElementsByTagName("Product");
        if (products.length > 0) {
            var product = products[0];
            var productId = product.getAttribute("ID");

            $("#cbProduct").val(productId);

            var pidKey = product.getAttribute("PIDKEY");
            $("#txtPidKey").val(pidKey);

            var exApps = product.getElementsByTagName("ExcludeApp");
            if (exApps.length > 0) {
                var exApp = exApps[0];
                var excludeAppId = exApp.getAttribute("ID");
                $("#cbExcludeApp").val(excludeAppId);

                $("#btRemoveExcludeApp").prop("disabled", false);
            } else {
                $("#btRemoveExcludeApp").prop("disabled", true);
            }
        }

        var version = addNode.getAttribute("Version");
        $("#txtVersion").val(version);

        var version = addNode.getAttribute("SourcePath");
        $("#txtSourcePath").val(version);
    }

    var removeNode = null;
    var remvoeNodes = xmlDoc.documentElement.getElementsByTagName("Remove");
    if (remvoeNodes.length > 0) {
        removeNode = xmlDoc.documentElement.getElementsByTagName("Remove")[0];
        if (removeNode) {
            var removeProducts = removeNode.getElementsByTagName("Product");
            if (removeProducts.length > 0) {
                var removeproduct = removeProducts[0];
                var removeproductId = removeproduct.getAttribute("ID");

                $("#cbRemoveProduct").val(removeproductId);

                var removeLangs = removeproduct.getElementsByTagName("Language");
                if (removeLangs.length > 0) {
                    var removeLangId = removeLangs[0].getAttribute("ID");
                    $("#cbRemoveLanguage").val(removeLangId);
                }

                toggleRemove("removeSelectProducts");
            } else {
                toggleRemove("removeallproducts");
            }
        }
    }

    var updateNodes = xmlDoc.documentElement.getElementsByTagName("Updates");
    if (updateNodes.length > 0) {
        var updateNode = xmlDoc.documentElement.getElementsByTagName("Updates")[0];

        var updatesEnabled = updateNode.getAttribute("Enabled");
        var selectUpdatePath = updateNode.getAttribute("UpdatePath");
        var selectTargetVersion = updateNode.getAttribute("TargetVersion");
        var selectDeadline = updateNode.getAttribute("Deadline");

        if (updatesEnabled == "TRUE") {
            toggleUpdatesEnabled("btupdatesEnabled");
            $("#txtUpdatePath").val(selectUpdatePath);
            $("#txtTargetVersion").val(selectTargetVersion);
            $("#txtDeadline").val(selectDeadline);
        } else {
            toggleUpdatesEnabled("btupdatesDisabled");
            $("#txtUpdatePath").val("");
            $("#txtTargetVersion").val("");
            $("#txtDeadline").val("");
        }
    }

    var displayNodes = xmlDoc.documentElement.getElementsByTagName("Display");
    if (displayNodes.length > 0) {
        var displayNode = xmlDoc.documentElement.getElementsByTagName("Display")[0];

        var logLevel = displayNode.getAttribute("Level");
        var acceptEula = displayNode.getAttribute("AcceptEULA");

        if (logLevel == "None") {
            toggleDisplayLevelEnabled("btLevelNone");
        } else {
            toggleDisplayLevelEnabled("btLevelFull");
        }

        if (acceptEula == "TRUE") {
            toggleDisplayEULAEnabled("btAcceptEULAEnabled");
        } else {
            toggleDisplayEULAEnabled("btAcceptEULADisabled");
        }
    }

    var propertyNodes = xmlDoc.documentElement.getElementsByTagName("Property");
    if (propertyNodes.length > 0) {
        var autoActivateNode = null;
        var forceShutDownNode = null;
        var sharedComputerLicensingNode = null;
        var packageguidNode = null;

        nodes = xmlDoc.documentElement.getElementsByTagName("Property");
        if (nodes.length > 0) {
            for (var n = 0; n < nodes.length; n++) {
                var propNode = xmlDoc.documentElement.getElementsByTagName("Property")[n];
                if (propNode) {
                    var attrValue = propNode.getAttribute("Name");
                    if (attrValue) {
                        if (propNode.getAttribute("Name").toUpperCase() == "AUTOACTIVATE") {
                            autoActivateNode = propNode;
                        }
                        if (propNode.getAttribute("Name").toUpperCase() == "FORCEAPPSHUTDOWN") {
                            forceShutDownNode = propNode;
                        }
                        if (propNode.getAttribute("Name").toUpperCase() == "SHAREDCOMPUTERLICENSING") {
                            sharedComputerLicensingNode = propNode;
                        }
                        if (propNode.getAttribute("Name").toUpperCase() == "PACKAGEGUID") {
                            packageguidNode = propNode;
                        }
                    }
                }
            }
        }

        var autoActivate = "";
        if (autoActivateNode) {
            autoActivate = autoActivateNode.getAttribute("Value");
        }

        var forceShutDown = "";
        if (forceShutDownNode) {
            forceShutDown = forceShutDownNode.getAttribute("Value");
        }

        var sharedComputerLicensing = "";
        if (sharedComputerLicensingNode) {
            sharedComputerLicensing = sharedComputerLicensingNode.getAttribute("Value");
        }

        var packageguid = "";
        if (packageguidNode) {
            packageguid = packageguidNode.getAttribute("Value");
        }

        if (packageguid) {
            if (packageguid.length > 0) {
                $("#txtPACKAGEGUID").val(packageguid);
            }
        }

        if (autoActivate == "1") {
            toggleAutoActivateEnabled("btAutoActivateYes");
        } else {
            toggleAutoActivateEnabled("btAutoActivateNo");
        }

        if (forceShutDown == "TRUE") {
            toggleForceAppShutdownEnabled("btForceAppShutdownTrue");
        } else {
            toggleForceAppShutdownEnabled("btForceAppShutdownFalse");
        }

        if (sharedComputerLicensing == "1") {
            toggleSharedComputerLicensing("btSharedComputerLicensingYes");
        } else {
            toggleSharedComputerLicensing("btSharedComputerLicensingNo");
        }
    } else {
        document.getElementById("btRemoveProduct").click();
    }

    var loggingNodes = xmlDoc.documentElement.getElementsByTagName("Logging");
    if (loggingNodes.length > 0) {
        var loggingNode = xmlDoc.documentElement.getElementsByTagName("Logging")[0];

        var logLevel = loggingNode.getAttribute("Level");
        var path = loggingNode.getAttribute("Path");

        if (logLevel == "Off") {
            toggleLoggingEnabled("btLoggingLevelOff");
        } else {
            toggleLoggingEnabled("btLoggingLevelStandard");
        }

        $("#txtLoggingUpdatePath").val(path);
    }

    var productCount = getAddProductCount(xmlDoc);
    if (productCount == 0) {
        $("#btRemoveProduct").prop("disabled", true);
    } else {
        $("#btRemoveProduct").prop("disabled", false);
    }

}

function sendMail() {
    var xmlSource = $('textarea#xmlText').val();

    var link = "mailto:"
             + "&subject=" + escape("Office Click-To-Run Configuration XML")
             + "&body=" + escape(xmlSource)
    ;

    window.location.href = link;
}


function clearXml() {
    $('textarea#xmlText').val("");
    $("#txtDeadline").val("");
    $("#txtLoggingUpdatePath").val("");
    $("#txtPidKey").val("");
    $("#txtSourcePath").val("");
    $("#txtTargetVersion").val("");
    $("#txtUpdatePath").val("");
    $("#txtVersion").val("");

    var resetDropDowns = document.getElementsByTagName("select");
    for (var t = 0; t < resetDropDowns.length; t++) {
        var dropDown = resetDropDowns[t];
        $("#" + dropDown.id).prop('selectedIndex', 0);
    }

    toggleRemove("removeAllProducts");

    var clearButtons = ["#btLevelNone", "#btLevelFull", "#btLoggingLevelOff",
        "#btLoggingLevelStandard", "#btAcceptEULAEnabled", "#btAcceptEULADisabled",
        "#btLoggingLevelOff", "#btLoggingLevelStandard", "#btAutoActivateYes",
        "#btAutoActivateNo", "#btForceAppShutdownTrue", "#btForceAppShutdownFalse",
        "#btSharedComputerLicensingYes", "#btSharedComputerLicensingNo"];

    for (var b = 0; b < clearButtons.length; b++) {
        var buttonName = clearButtons[b];
        $(buttonName).removeClass('active');
        $(buttonName).removeClass('btn-primary');
    }

    $.cookie("xmlcache", "");

    $("#btAddProduct").text('Add Product');
}

function getXmlDocument() {
    var xmlSource = $('textarea#xmlText').val();
    if (!(xmlSource)) {
        xmlSource = "<Configuration></Configuration>";
    }
    var xmlDoc = createXmlDocument(xmlSource);
    return xmlDoc;
}

function createXmlDocument(string) {
    var doc;
    if (window.DOMParser) {
        parser = new DOMParser();
        doc = parser.parseFromString(string, "application/xml");
    }
    else // Internet Explorer
    {
        doc = new ActiveXObject("Microsoft.XMLDOM");
        doc.async = "false";
        doc.loadXML(string);
    }
    return doc;
}

function displayXml(xmlDoc) {
    var xmlString = (new XMLSerializer().serializeToString(xmlDoc.documentElement));
    var xmlOutput = vkbeautify.xml(xmlString);
    $('textarea#xmlText').val(xmlOutput);
    $.cookie("xmlcache", xmlOutput);
}


function toggleRemove(sourceId) {

    if (sourceId.toLowerCase() == "removeallproducts") {
        $("#removeSelectProducts").removeClass('active');
        $("#removeSelectProducts").removeClass('btn-primary');

        var $this = $("#removeAllProducts");
        if (!$this.hasClass('active')) {
            $this.addClass('active');
        }

        if (!$this.hasClass('btn-primary')) {
            $this.addClass('btn-primary');
        }

        $("#cbRemoveProduct").prop("disabled", true);
        $("#cbRemoveLanguage").prop("disabled", true);
        $("#btAddRemoveLanguage").prop("disabled", true);
    } else {
        $("#removeAllProducts").removeClass('active');
        $("#removeAllProducts").removeClass('btn-primary');

        var $this = $("#removeSelectProducts");
        if (!$this.hasClass('active')) {
            $this.addClass('active');
        }

        if (!$this.hasClass('btn-primary')) {
            $this.addClass('btn-primary');
        }

        $("#cbRemoveProduct").prop("disabled", false);
        $("#cbRemoveLanguage").prop("disabled", false);
        $("#btAddRemoveLanguage").prop("disabled", false);
    }

}

function toggleUpdatesEnabled(sourceId) {

    if (sourceId.toLowerCase() == "btupdatesenabled") {
        //$("#btupdatesDisabled").removeClass('active');
        $("#btupdatesDisabled").removeClass('btn-primary');

        var $this = $("#btupdatesEnabled");
        //if (!$this.hasClass('active')) {
        //    $this.addClass('active');
        //}

        if (!$this.hasClass('btn-primary')) {
            $this.addClass('btn-primary');
        }

        $("#inputDeadline").prop("disabled", false);

        toggleTextBox("txtUpdatePath", true);
        toggleTextBox("txtTargetVersion", true);
    } else {
        //$("#btupdatesEnabled").removeClass('active');
        $("#btupdatesEnabled").removeClass('btn-primary');

        var $this = $("#btupdatesDisabled");
        //if (!$this.hasClass('active')) {
        //    $this.addClass('active');
        //}

        if (!$this.hasClass('btn-primary')) {
            $this.addClass('btn-primary');
        }

        $("#inputDeadline").prop("disabled", true);

        toggleTextBox("txtUpdatePath", false);
        toggleTextBox("txtTargetVersion", false);

    }
    return false;
}

function toggleTextBox(id, enabled) {
    if (enabled) {
        $("#" + id).prop("disabled", false);
        $("#" + id).css("background-color", "");
        $("#" + id).css("border-color", "");
    } else {
        $("#" + id).prop("disabled", true);
        $("#" + id).css("background-color", "#eeeeee");
        $("#" + id).css("border-color", "#e1e1e1");
    }
}

function toggleAutoActivateEnabled(sourceId) {

    if (sourceId.toLowerCase() == "btautoactivateyes") {
        $("#btAutoActivateNo").removeClass('active');
        $("#btAutoActivateNo").removeClass('btn-primary');

        var $this = $("#btAutoActivateYes");
        if (!$this.hasClass('active')) {
            $this.addClass('active');
        }

        if (!$this.hasClass('btn-primary')) {
            $this.addClass('btn-primary');
        }
    } else {
        $("#btAutoActivateYes").removeClass('active');
        $("#btAutoActivateYes").removeClass('btn-primary');

        var $this = $("#btAutoActivateNo");
        if (!$this.hasClass('active')) {
            $this.addClass('active');
        }

        if (!$this.hasClass('btn-primary')) {
            $this.addClass('btn-primary');
        }
    }

}

function toggleForceAppShutdownEnabled(sourceId) {

    if (sourceId.toLowerCase() == "btforceappshutdowntrue") {
        $("#btForceAppShutdownFalse").removeClass('active');
        $("#btForceAppShutdownFalse").removeClass('btn-primary');

        var $this = $("#btForceAppShutdownTrue");
        if (!$this.hasClass('active')) {
            $this.addClass('active');
        }

        if (!$this.hasClass('btn-primary')) {
            $this.addClass('btn-primary');
        }
    } else {
        $("#btForceAppShutdownTrue").removeClass('active');
        $("#btForceAppShutdownTrue").removeClass('btn-primary');

        var $this = $("#btForceAppShutdownFalse");
        if (!$this.hasClass('active')) {
            $this.addClass('active');
        }

        if (!$this.hasClass('btn-primary')) {
            $this.addClass('btn-primary');
        }
    }

}

function toggleSharedComputerLicensing(sourceId) {

    if (sourceId.toLowerCase() == "btsharedcomputerlicensingyes") {
        $("#btSharedComputerLicensingNo").removeClass('active');
        $("#btSharedComputerLicensingNo").removeClass('btn-primary');

        var $this = $("#btSharedComputerLicensingYes");
        if (!$this.hasClass('active')) {
            $this.addClass('active');
        }

        if (!$this.hasClass('btn-primary')) {
            $this.addClass('btn-primary');
        }
    } else {
        $("#btSharedComputerLicensingYes").removeClass('active');
        $("#btSharedComputerLicensingYes").removeClass('btn-primary');

        var $this = $("#btSharedComputerLicensingNo");
        if (!$this.hasClass('active')) {
            $this.addClass('active');
        }

        if (!$this.hasClass('btn-primary')) {
            $this.addClass('btn-primary');
        }
    }

}

function toggleDisplayLevelEnabled(sourceId) {

    if (sourceId.toLowerCase() == "btlevelnone") {
        $("#btLevelFull").removeClass('active');
        $("#btLevelFull").removeClass('btn-primary');

        var $this = $("#btLevelNone");
        if (!$this.hasClass('active')) {
            $this.addClass('active');
        }

        if (!$this.hasClass('btn-primary')) {
            $this.addClass('btn-primary');
        }
    } else {
        $("#btLevelNone").removeClass('active');
        $("#btLevelNone").removeClass('btn-primary');

        var $this = $("#btLevelFull");
        if (!$this.hasClass('active')) {
            $this.addClass('active');
        }

        if (!$this.hasClass('btn-primary')) {
            $this.addClass('btn-primary');
        }
    }

}

function toggleDisplayEULAEnabled(sourceId) {

    if (sourceId.toLowerCase() == "btaccepteulaenabled") {
        $("#btAcceptEULADisabled").removeClass('active');
        $("#btAcceptEULADisabled").removeClass('btn-primary');

        var $this = $("#btAcceptEULAEnabled");
        if (!$this.hasClass('active')) {
            $this.addClass('active');
        }

        if (!$this.hasClass('btn-primary')) {
            $this.addClass('btn-primary');
        }
    } else {
        $("#btAcceptEULAEnabled").removeClass('active');
        $("#btAcceptEULAEnabled").removeClass('btn-primary');

        var $this = $("#btAcceptEULADisabled");
        if (!$this.hasClass('active')) {
            $this.addClass('active');
        }

        if (!$this.hasClass('btn-primary')) {
            $this.addClass('btn-primary');
        }
    }

}

function toggleLoggingEnabled(sourceId) {

    if (sourceId.toLowerCase() == "btloggingleveloff") {
        $("#btLoggingLevelStandard").removeClass('active');
        $("#btLoggingLevelStandard").removeClass('btn-primary');

        var $this = $("#btLoggingLevelOff");
        if (!$this.hasClass('active')) {
            $this.addClass('active');
        }

        if (!$this.hasClass('btn-primary')) {
            $this.addClass('btn-primary');
        }

        $("#txtLoggingUpdatePath").prop("disabled", true);
    } else {
        $("#btLoggingLevelOff").removeClass('active');
        $("#btLoggingLevelOff").removeClass('btn-primary');

        var $this = $("#btLoggingLevelStandard");
        if (!$this.hasClass('active')) {
            $this.addClass('active');
        }

        if (!$this.hasClass('btn-primary')) {
            $this.addClass('btn-primary');
        }

        $("#txtLoggingUpdatePath").prop("disabled", false);
    }

}


function IsGuid(value) {
    var rGx = new RegExp("\\b(?:[A-F0-9]{8})(?:-[A-F0-9]{4}){3}-(?:[A-F0-9]{12})\\b");
    return rGx.exec(value) != null;
}


var versions = [
'15.0.4745.1001',
'15.0.4727.1003',
'15.0.4719.1002',
'15.0.4711.1003',
'15.0.4701.1002',
'15.0.4693.1002',
'15.0.4693.1001',
'15.0.4675.1002',
'15.0.4667.1002',
'15.0.4659.1001',
'15.0.4649.1003',
'15.0.4649.1001',
'15.0.4641.1003',
'15.0.4631.1004',
'15.0.4631.1002',
'15.0.4623.1003',
'15.0.4615.1002',
'15.0.4605.1003',
'15.0.4569.1508',
'15.0.4569.1507',
'15.0.4551.1512',
'15.0.4551.1011',
'15.0.4551.1005',
'15.0.4535.1511',
'15.0.4535.1004',
'15.0.4517.1509',
'15.0.4517.1005',
'15.0.4505.1510',
'15.0.4505.1006',
'15.0.4481.1510'
];

