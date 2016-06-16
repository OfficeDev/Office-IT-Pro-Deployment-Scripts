var productToInstall = "";
var versionToInstall = "";
var languages = "";
var languageDictionary = {
    "English": "en-us",
    "Arabic": "ar-sa",
    "Bulgarian": "bg-bg",
    "Chinese (Simplified)": "zh-cn",
    "Chinese": "zh-tw",
    "Croatian": "hr-hr",
    "Czech": "cs-cz",
    "Croatian": "hr-hr",
    "Danish": "da-dk",
    "Estonian": "et-ee",
    "Finnish": "fi-fi",
    "French": "fr-fr",
    "German": "de-de",
    "Greek": "el-gr",
    "Hebrew": "he-il",
    "Hindi": "hi-in",
    "Hungarian": "hu-hu",
    "Indonesian": "id-id",
    "Italian": "it-it",
    "Japanese": "ja-jp",
    "Kazakh": "kk-kh",
    "Korean": "ko-kr",
    "Latvian": "lv-lv",
    "Lithuanian": "lt-lt",
    "Malay": "ms-my",
    "Norwegian (Bokm�l)": "nb-no",
    "Polish": "pl-pl",
    "Portuguese (Brazil)": "pt-br",
    "Portuguese (Portugal)": "pt-pt",
    "Romanian": "ro-ro",
    "Russian": "ru-ru",
    "Serbian (Latin)": "sr-latn-rs",
    "Slovak": "sk-sk",
    "Slovenian": "sl-si",
    "Spanish": "es-es",
    "Swedish": "sv-se",
    "Thai": "th-th",
    "Turkish": "tr-tr",
    "Ukrainian": "uk-ua",

    "en-us": "English",
    "ar-sa": "Arabic",
    "bg-bg": "Bulgarian",
    "zh-cn": "Chinese (Simplified)",
    "zh-tw": "Chinese",
    "hr-hr": "Croatian",
    "cs-cz": "Czech",
    "hr-hr": "Croatian",
    "da-dk": "Danish",
    "et-ee": "Estonian",
    "fi-fi": "Finnish",
    "fr-fr": "French",
    "de-de": "German",
    "el-gr": "Greek",
    "he-il": "Hebrew",
    "hi-in": "Hindi",
    "hu-hu": "Hungarian",
    "id-id": "Indonesian",
    "it-it": "Italian",
    "ja-jp": "Japanese",
    "kk-kh": "Kazakh",
    "ko-kr": "Korean",
    "lv-lv": "Latvian",
    "lt-lt": "Lithuanian",
    "ms-my": "Malay",
    "nb-no": "Norwegian (Bokm�l)",
    "pl-pl": "Polish",
    "pt-br": "Portuguese (Brazil)",
    "pt-pt": "Portuguese (Portugal)",
    "ro-ro": "Romanian",
    "ru-ru": "Russian",
    "sr-latn-rs": "Serbian (Latin)",
    "sk-sk": "Slovak",
    "sl-si": "Slovenian",
    "es-es": "Spanish",
    "sv-se": "Swedish",
    "th-th": "Thai",
    "tr-tr": "Turkish",
    "uk-ua": "Ukrainian"
}

var availableFilters = [];
var searchBoxTaggle;
var currentLocation;
var currentFilter;

var appliedFilters = [];
var previousSearch = "";

function setProduct(product, build) {
    buildID = build;
    getLanguages();
    productToInstall = product;
    $('#productSpan').text(product);
    showModal('languageModal');
}

function setVersion(version) {
    versionToInstall = version;
    $('#versionSpan').text(version);
    showModal('productModal');
}

function setLanguage() {
    var checkboxes = null;
    languages = null;
    checkboxes = $(".languageCheckBox:checked");
    $('#languageSpan').text(languageDictionary[checkboxes[0].id]);
    languages = [checkboxes[0].id];
    if (checkboxes.length > 1) {
        for (var i = 1; i < checkboxes.length; i++) {
            languages[i] = checkboxes[i].id
            $('#languageSpan')[0].innerText += ", "+languageDictionary[checkboxes[i].id];
        }
    }
    showModal('confirmationModal');
}

function startInstall() {

}

function showModal(modalId) {
    $(".custom-Dialog").removeClass("hidden").addClass("hidden");
    $("#" + modalId).removeClass("hidden");
    if (modalId == "downloadModal") {
        $('#directDL').text(versionToInstall);
    }

    if (modalId === 'productModal')
    {
        resetFilters();
    }
}

function resetFilters() {
    searchBoxTaggle.removeAll();
    appliedFilters = [];
    $(searchBoxTaggle.getInput()).val('');
    searchBoxFilter();
    $('#ul-Location li:first').click();
}

function verifyLanguageInput() {
    sl = $('.languageCheckBox:checked');
    if (sl.length > 0) {
        $('#languageButton').prop('disabled', false);
    } else {
        $('#languageButton').prop('disabled', true);
    }
}

function getLanguages() {

    $.ajax({
        type: "GET",
        url: "SelfServiceConfig.xml",
        datatype: "xml",
        success:
            function (xml) {
                $('#languagesGrid li').remove();
                $xml = $(xml);
                var languages = $xml.find("[ID='" + buildID + "']").attr('Languages').split(",");
                $.each(languages, function (index, value) {
                    var label = value;
                    var id = value.split(" ").pop().replace(")",'').replace("(",'');
                    $('#languagesGrid > ul').append("<li class='languageli'><label> <input type='checkbox' id='" + id + "' class='languageCheckBox' onclick='verifyLanguageInput()'/> \
                                    <span class='ms-Label checkboxLabel'>" + label + "</span></label></li>");
                });  
            }
    });
}

function getVersions() {

    $.ajax({
        type: "GET",
        url: "SelfServiceConfig.xml",
        datatype: "xml",
        success:
            function (xml) {
                var optional = $(xml).find('Versions').attr('Optional');
                if (optional === "true") {
                    $(xml).find('Version').each(function () {
                        var version = $(this).attr('ID');
                        if (version === "2013") {
                            $('#versions').prepend("<li class='squareButton'>\
                                        <div class='ms-Dialog-action ms-Button ms-Button--primary ms-bgColor-orangeLight version-box' onclick='setVersion(\"2013\")'>\
                                        <img class='version-img' src='Content/imgs/office-icon-white.png'/>\
                                        <p class='ms-font-xl ms-fontColor-white version-box-text'>Office 365 ProPlus (2013)</p>\
                                        </div>\
                                        </li>");
                        }
                        if (version === "2016") {
                            $('#versions').prepend("<li class='squareButton'>\
                                    <div class='ms-Dialog-action ms-Button ms-Button--primary version-box' onclick='setVersion(\"2016\")'>\
                                    <img  class='version-img'  src='Content/imgs/office-icon-white.png'>\
                                    <p class='ms-font-xl ms-fontColor-white'>Office 365 ProPlus (2016)</p>\
                                    </div>\
                                    </li>");
                        }
                    });
                }
                else
                {
                    $('#versions').prepend("<li class='squareButton'>\
                                    <div class='ms-Dialog-action ms-Button ms-Button--primary version-box' onclick='setVersion(\"365 ProPlus\")'>\
                                    <img  class='version-img'  src='Content/imgs/office-icon-white.png'/>\
                                    <p class='ms-font-xl ms-fontColor-white'>Office 365 ProPlus</p>\
                                    </div>\
                                    </li>");
                }
            }
    });
}

function getBuild() {

    $.ajax({
        type: "GET",
        url: "SelfServiceConfig.xml",
        datatype: "xml",
        success:
            function (xml) {
                $(xml).find('Build').each(function () {
                    var buildType = $(this).attr('Type');
                    var filters = $(this).attr('Filters').split(',');
                    var classString = "";
                    var textString = "";
                    if (Array.isArray(filters)) {
                        filters.forEach(function (element, index, array) {
                            classString += element.toLocaleLowerCase() + "-filter ";
                            textString += ", " + element;
                        });
                    } else {
                        if (filters) {
                            classString += filters + "-filter ";
                            textString += ", " + filters;
                        }
                    }
                    $("#buildsGrid").append("<li class='squareButton-build shown " + classString + $(this).attr('Location').toLocaleLowerCase() + "-filter'>\
                                    <button class='ms-Dialog-action ms-Button' onclick='setProduct(versionToInstall,\""+ $(this).attr('ID') + "\")'>\
                                    <i class='ms-Icon ms-Icon--people' style='font-size:125px'></i>\
                                    <p class='ms-font-xl filter-field' style='display:block'>" + $(this).attr('Type') + "\
                                    <br>" + $(this).attr('Location') + textString + "</p>\
                                    </button>\
                                    </li>");
                });
            }
    });
}

function getLocations(callback) {

    var locations = [];
    $.ajax({
        type: "GET",
        url: "SelfServiceConfig.xml",
        datatype: "xml",
        success:
            function (xml) {
                $("#ddl-Location").siblings('span.ms-Dropdown-title').text("Location Filter");
                $("#ddl-Location").siblings('ul').append("<li class='ms-Dropdown-item'>Location Filter</li>");

                $(xml).find('Build').each(function () {
                    var location = $(this).attr('Location');
                    if($.inArray(location,locations) === -1)
                    {
                        availableFilters.push(location);
                        locations.push(location);
                        $("#ddl-Location").siblings('ul').attr('id','ul-Location');
                        $("#ddl-Location").siblings('ul').append("<li class='ms-Dropdown-item'>" + location + "</li>");
                    }
                });
                updateAutocomplete();
                callback();
            }
    });
}

function getFilters() {

    $.ajax({
        type: "GET",
        url: "SelfServiceConfig.xml",
        datatype: "xml",
        success:
            function (xml) {
                $(xml).find('Build').each(function () {

                    var filter = $(this).attr('Filters').split(',');
                    if (Array.isArray(filter)) {
                        filter.forEach(function (element, index, array) {
                            if (availableFilters.indexOf(element.toLocaleLowerCase()) < 0) {
                                availableFilters.push(element.toLocaleLowerCase());
                            }
                        });
                    } else {
                        if (filter) {
                            if (availableFilters.indexOf(filter.toLocaleLowerCase()) < 0) {
                                availableFilters.push(filter.toLocaleLowerCase());
                            }
                        }
                    }
                });
                updateAutocomplete();
            }
    });
}

function getHelp() {

    $.ajax({
        type: "GET",
        url: "SelfServiceConfig.xml",
        datatype: "xml",
        success:
            function (xml) {
                $(xml).find('Help').find('Item').each(function () {

                    var qtext = $(this).find('Question').text();
                    var atext = $(this).find('Answer').text();
                    $('#helpContent').append(
                    "<div class='questionDiv'>\
                        <h4 class='ms-font-xl'>"+ qtext + "</h4>\
                        <p class='ms-font-m questionAnswer'>"
                            + atext +
                        "</p>\
                    </div>");
                });
            }
    });
}

function getCompanyInfo() {

    $.ajax({
        type: "GET",
        url: "SelfServiceConfig.xml",
        datatype: "xml",
        success:
            function (xml) {
                $(xml).find('Company').each(function () {
                    $('.companyName').text($(this).attr('Name'));
                    if ($(this).attr('LogoSrc')) {
                        $('.companyLogo').src($(this).attr('LogoSrc'));
                    } else {
                        $('.companyLogo').addClass('hidden');
                    }
                    
                });
            }
    });
}

function searchBoxFilter() {
    var searchTerm = searchBoxTaggle.getInput().value;
    searchTerm = searchTerm.toLocaleLowerCase();
    $(".squareButton-build").removeClass('search-filter');
    removeFilter("search");
    if (searchTerm) {
        $(".squareButton-build p").each(function () {
            if ($(this).text().toLocaleLowerCase().indexOf(searchTerm) >= 0) {
                $(this).parent().parent().addClass('search-filter');
            }
        })
        addFilter("search");
    }
    applyFilters();
}

function setTaggleFilters() {
    taggles = searchBoxTaggle.getTagValues();
    taggles.forEach(function (element, index, array) {
        addFilter(element);
    })
    applyFilters();
}

function locationFilter(location) {
    if (location) {
        removeFilter(currentLocation);
        if (location.toLocaleLowerCase().indexOf('filter') < 0) {
            addFilter(location);
        }
        applyFilters();
    }
    currentLocation = location;
}

function addFilter(filter) {
    if (appliedFilters.indexOf(filter) == -1) {
        appliedFilters.push(filter);
    }
}

function applyFilters() {
    var filterString = ".squareButton-build";
    appliedFilters.forEach(function (element, index, array) {
        filterString += "." + element + "-filter";
    });
    $(".squareButton-build").addClass("hidden");
    $(filterString).removeClass("hidden").addClass('shown');
}

function removeFilter(filter) {
    if (appliedFilters.indexOf(filter) >= 0) {
        appliedFilters.splice(appliedFilters.indexOf(filter), 1);
    }
}

function addLocationClick() {
    $('#ul-Location li').each(function () {
        $(this).attr('onclick', "locationFilter('"+$(this).text().toLocaleLowerCase()+"')");
    });
            }

function prepTags() {
    searchBoxTaggle = new Taggle('outerSearchBox',
        {
            saveOnBlur: true,
            placeholder: "search...",
            onTagAdd: function (event, tag) {
                $(searchBoxTaggle.getInput()).val('');
                addFilter(tag);
                applyFilters();
            },
            onTagRemove: function (event, tag) {
                removeFilter(tag);
                applyFilters();
            }
        });
}

function updateAutocomplete() {
    var container = searchBoxTaggle.getContainer();
    var input = searchBoxTaggle.getInput();
    searchBoxTaggle.settings.allowedTags = availableFilters;
    $(input).autocomplete({
        source: availableFilters,
        appendTo: container,
        position: { at: "left bottom", my: "left top" },
        select: function (event, data) {
            event.preventDefault();
            //Add the tag if user clicks
            if (event.which === 1) {
                searchBoxTaggle.add(data.item.value);
            }
        }
    });
}

function getVersionDescription(){
    $.ajax({
        type: "GET",
        url: "SelfServiceConfig.xml",
        datatype: "xml",
        success:
            function (xml) {
                var description = $(xml).find('Description').text();
                $('#versionDescription').text(description);
            }
    });
}

function getUserInfo(){
    $.ajax({
        type: "GET",
        url: "/SelfService/UserInfo",
        success:
            function (response) {
                console.log(response);
            }
    });
}


$(document).ready(function () {

    getCompanyInfo();
    getLocations(addLocationClick);
    getFilters();
    getVersions();
    getBuild();
    getVersionDescription();
    getHelp();
    getUserInfo();
    //searchbox filter
    $("#outerSearchBox").keyup(function (e) {
        searchBoxFilter(e);
    });

    //filter reset
    $('#btn-Reset').click(function () {
        resetFilters();
    })

    prepTags();
});