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
};

var availableFilters = [];
var searchBoxTaggle;
var primaryLanguage;
var currentLocation;
var currentFilter;
var listView = 0;
var xmlConfigPath, exePath, setupPath, manifestPath;
var appliedFilters = [];
var previousSearch = "";


function setProduct(product, build) {
    buildID = build;
    getPrimaryLanguages();
    productToInstall = product;
    $('#productSpan').text(product);
    showModal('primaryLanguageModal');
}

function setVersion(version) {
    versionToInstall = version;
    $('#versionSpan').text(version);
}

function setPrimaryLanguage() {
    primaryLanguage = $("input[name=radio1]:checked").attr('id');
    getLanguages();
    showModal('languageModal');
}

function setLanguage() {
    var checkboxes = null;
    languages = null;
    checkboxes = $(".languageCheckBox:checked");
    $('#languageSpan').text('');
    $('#languageSpan').text(languageDictionary[primaryLanguage]);
    if (checkboxes.length >= 1) {
        languages = [checkboxes[0].id];
        for (var i = 0; i < checkboxes.length; i++) {
            languages[i] = checkboxes[i].id;
            $('#languageSpan').append(", " + languageDictionary[checkboxes[i].id]);
        }
    }
    showModal('confirmationModal');
}


function buildQueryString() {
    location.hash = '';
    var params = {
        xml: xmlConfigPath,
        installer: setupPath
    },
    query = $.param(params);
    location.hash = query;

}

function generateXML() {

    $.ajax(
        {
            type: "POST",
            url: ServerSide.GenerateXML,
            data: { buildName: buildID, languageList: languages, uiLanguage: primaryLanguage },
            traditional: true,
            success:
                function (xhr) {
                    xmlConfigPath = xhr.xml;
                    exePath = xhr.exe;
                    setupPath = xhr.setup;
                    manifestPath = xhr.manifest;


                    window.open(exePath + "?xml=" + xmlConfigPath + "&installer=" + setupPath);

                    $('#directDL').attr({ target: "_blank", href: exePath });
                    buildQueryString();
                    showModal('downloadModal');

                },
            error:
                function (xhr) {
                    $('#errorMessage').removeClass('hidden');
                    $('#errorText').text(xhr.responseText);
                }
        })
}

function showModal(modalId) {
    $(".custom-Dialog").removeClass("hidden").addClass("hidden");
    $("#" + modalId).removeClass("hidden");
    if (modalId === "downloadModal") {
        $('#directDL').text(versionToInstall + " click here");
    }

    if (modalId === 'productModal') {
        resetFilters();
        $('#buildsTable').empty();
        $('#buildsGrid').empty();
        $('#languageButton').prop('disabled', 'true');
        getBuild();
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
    sl = $('input[name=radio1]:checked');
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
                $('#languagesGrid div').remove();
                $xml = $(xml);
                var languages = $xml.find("[ID='" + buildID + "']").attr('Languages').split(",");
                $.each(languages, function (index, value) {
                    var label = value;
                    var id = value.split(" ").pop().replace(")", '').replace("(", '');
                    if (id !== primaryLanguage) {
                        $('#languagesGrid ').append("<div class='ms-Grid-col ms-u-sm6 ms-u-md6 ms-u-lg3 ms-u-xl3 languageli'><label><input type='checkbox' id='" + id + "' class='languageCheckBox' />\
                                     <span class='ms-font-m checkboxLabel'>" + label + "</span></label></div>");
                    }
                });
            }
    });
}

function getPrimaryLanguages() {
    $.ajax({
        type: "GET",
        url: "SelfServiceConfig.xml",
        datatype: "xml",
        success:
            function (xml) {
                $('#primaryLanguagesGrid div').remove();
                $xml = $(xml);
                var languages = $xml.find("[ID='" + buildID + "']").attr('Languages').split(",");
                $.each(languages, function (index, value) {
                    var label = value;
                    var id = value.split(" ").pop().replace(")", '').replace("(", '');
                    $('#primaryLanguagesGrid').append("<div class='ms-Grid-col ms-u-sm6 ms-u-md6 ms-u-lg3 ms-u-xl3 ms-ChoiceField '>\
                        <input id='" + id + "' class='ms-ChoiceField-input' type='radio' name='radio1' value='" + id + "'onclick='verifyLanguageInput()'>\
                        <label for='" + id + "' class='ms-ChoiceField-field'><span class='ms-Label' >" + label + "</span></label>\
                        </div>")
                });
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
                if (listView === 1) {
                    $("#buildsTable").empty();
                    $("#buildsTable").append("<div class='ms-Table-row'>\
                            <span class='ms-Table-cell custom-cell' style='padding-left:4%'>Name</span>\
                            <span class='ms-Table-cell custom-cell' style=''>Location</span>\
                            <span class='ms-Table-cell custom-cell'>Tags</span>\
                            <span class='ms-Table-cell custom-cell'></span>\
                            </div>");
                }
                $(xml).find('Build').each(function () {
                    var buildType = $(this).attr('Type');
                    var filters = $(this).attr('Filters').split(',');
                    var classString = "";
                    var textString = "";



                    if (listView === 1) {

                        if (Array.isArray(filters)) {
                            filters.forEach(function (element) {
                                classString += element.toLocaleLowerCase().replace(/\W+/g, "-").replace(/\ /g, "-") + "-filter ";
                                textString += " " + element + ",";
                            });
                        } else {
                            if (filters) {
                                classString += filters.replace(/\W+/g, "-").replace(/\ /g, "-") + "-filter ";
                                textString += " " + element + ",";
                            }
                        }


                        $("#buildsTable").append("<div class='ms-Table-row custom-table-row shown " + $(this).attr('Location').toLocaleLowerCase().replace(/\W+/g, " ").replace(/\ /g, "-") + "-filter " + classString + "location-filter " + buildType.toLowerCase().replace(/\ /g, "-").replace(/\W+/g, "-") + "-filter'>\
                            <span class='ms-Table-cell ms-font-l custom-first-cell custom-cell filter-field'><i class='ms-Icon ms-Icon--people package-people-table'></i>" + buildType + "</span>\
                            <span class='ms-Table-cell custom-cell'>"+ $(this).attr('Location') + "</span>\
                            <span class='ms-Table-cell custom-cell'><i class='ms-Icon ms-Icon--tag custom-table-tag'></i>"+ textString + "</span>\
                            <span class='ms-Table-cell custom-cell custom-last-cell' onclick='setProduct(\"2016\",\""+ $(this).attr('ID') + "\")'><i class='ms-Icon ms-Icon--download custom-table-tag' ></i><a class='ms-link'>Install</a></span>\
                        </div>");
                    }
                    else {
                        if (Array.isArray(filters)) {
                            filters.forEach(function (element) {
                                classString += element.toLocaleLowerCase().replace(/\W+/g, "-").replace(/\ /g, "-") + "-filter ";
                                textString += "<li class='ms-font-m " + classString + "'>" + element + "</li>";
                            });
                        } else {
                            if (filters) {
                                classString += filters.replace(/\W+/g, "-").replace(/\ /g, "-") + "-filter ";
                                textString += "<li class='ms-font-m " + classString + "'>" + filters + "</li>";
                            }
                        }

                        $("#buildsGrid").append("<div class='ms-Grid-col ms-u-sm12 ms-u-md6 ms-u-lg4 ms-u-xl3  package-group shown " + $(this).attr('Location').toLocaleLowerCase().replace(/\W+/g, " ").replace(/\ /g, "-") + "-filter " + classString + " location-filter " + buildType.toLowerCase().replace(/\W+/g, "-").replace(/\ /g, "-") + "-filter'>\
                                                        <div id='custom-callout' class='ms-Callout ms-Callout--OOBE ms-Callout--arrowLeft hidden'>\
                                                            <div class='ms-Callout-main'>\
                                                                <div class='ms-Callout-header custom-callout-header'>\
                                                                    <i class='ms-Icon ms-Icon--x custom-x' onclick='closeCallout(event)'></i>\
                                                                    <div class='ms-Callout-title ms-font-xl ms-fontWeight-regular'>Tags</div>\
                                                                </div>\
                                                                <div class='ms-Callout-inner custom-callout-inner'>\
                                                                    <div class='ms-Callout-content'>\
                                                                        <ul id='tags-list' class='tags-list'>"
                                                                        + textString + "\
                                                                        </ul>\
                                                                    </div>\
                                                                </div>\
                                                            </div>\
                                                        </div>\
                                                         <div class='package package-main'>\
                                                         <div class='ms-Grid-row package-inner'>\
                                                            <div class='ms-Grid-col ms-u-sm2 ms-u-md4 ms-u-lg2 ms-u-xl1'>\
                                                                <i class='ms-Icon ms-Icon--people package-people'></i>\
                                                            </div>\
                                                            <div class='ms-Grid-col ms-u-sm6 ms-u-smPush4 ms-u-md4 ms-u-mdPush2 ms-u-lg4 ms-u-lgPush4 ms-u-xl3 ms-u-xlPush4'>\
                                                                 <div class='ms-Grid'>\
                                                                    <div class='ms-Grid-row'>\
                                                                        <p class='ms-font-l type-label filter-field'>" + buildType + "</b></p><br />\
                                                                    </div>\
                                                                    <div class='ms-Grid-row'>\
                                                                        <p class='location-label filter-field' >"+ $(this).attr('Location') + "</p>\
                                                                    </div>\
                                                                </div>\
                                                            </div>\
                                                            <div style='display:relative;' class='ms-Grid-col ms-u-sm1 ms-u-smPush3 ms-u-md1 ms-u-mdPush3 ms-u-lg1 ms-u-lgPush5 ms-u-xl1 ms-u-xlPush7'>\
                                                                <i id='calloutTag' class='ms-Icon ms-Icon--tag package-tag' onclick='toggleCallout(event)'></i>\
                                                            </div>\
                                                        </div>\
                                                        <div class='ms-Grid-row'>\
                                                            <span class='package-bottom' onclick='setProduct(\"2016\",\""+ $(this).attr('ID') + "\")'>\
                                                                <i class=' ms-Icon ms-Icon--download package-download'></i>\
                                                                <a class=' ms-font-m ms-link'>Install</a>\
                                                            </span>\
                                                        </div>\
                                                    </div>\
                                                </div> \
                                            </div>");


                    }
                });
            }
    });
}

function toggleCallout(event) {
    $(event.target).toggleClass('callout-open');
    $(event.target).parents().eq(3).find('#custom-callout').toggleClass('hidden');
}

function closeCallout(event) {

    $(event.target).parents().eq(3).find('#calloutTag').toggleClass('callout-open');
    $(event.target).parents().eq(3).find('#custom-callout').addClass('hidden');
}

function getLocations(callback) {

    var locations = [];
    $.ajax({
        type: "GET",
        url: "SelfServiceConfig.xml",
        datatype: "xml",
        success:
            function (xml) {
                $("#ddl-Location").siblings('span.ms-Dropdown-title').text("Location");
                $("#ddl-Location").siblings('ul').append("<li class='ms-Dropdown-item'>Location</li>");

                $(xml).find('Build').each(function () {
                    var location = $(this).attr('Location');
                    if ($.inArray(location, locations) === -1) {
                        availableFilters.push(location.replace(/\,/g, " "));
                        locations.push(location);
                        $("#ddl-Location").siblings('ul').attr('id', 'ul-Location');
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
                    var type = $(this).attr("Type");

                    if (type) {
                        if (availableFilters.indexOf(type.toLocaleLowerCase()) < 0) {
                            availableFilters.push(type.toLocaleLowerCase());
                        }
                    }

                    if (Array.isArray(filter)) {
                        filter.forEach(function (element) {
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
    if (listView === 0) {
        $(".package-group").removeClass('search-filter');
        removeFilter("search");
        if (searchTerm) {

            $(".package-main p").each(function () {
                var parent = $(this).parent().parent().parent();
                if ($(this).text().toLocaleLowerCase().indexOf(searchTerm) >= 0 &&
                    $(parent).attr('class').indexOf('hidden') >= 0) {

                    $(parent).addClass('search-filter');
                }
            });

            $(".type-label").each(function () {
                var parent = $(this).parent().parent().parent().parent().parent().parent();
                if ($(this).text().toLocaleLowerCase().indexOf(searchTerm) >= 0 &&
                    $(parent).attr('class').indexOf('hidden') < 0) {

                    $(parent).addClass('search-filter');

                }
            });

            $(".tags-list li").each(function () {
                var parent = $(this).parent().parent().parent().parent().parent().parent();
                if ($(this).text().toLocaleLowerCase().indexOf(searchTerm) >= 0 &&
                    $(parent).attr('class').indexOf('hidden') < 0) {

                    $(parent).addClass('search-filter');

                }
            });

            $(".location-label").each(function () {
                var parent = $(this).parent().parent().parent().parent().parent().parent();
                if ($(this).text().toLocaleLowerCase().indexOf(searchTerm) >= 0 &&
                    $(parent).attr('class').indexOf('hidden') < 0) {

                    $(parent).addClass('search-filter');

                }
            });


            addFilter("search");

        } else if (searchTerm === "" && $('.taggle_list').children('.taggle').length > 0) {

            $('.taggle_list').children('.taggle').each(function () {

                var tagTerm = $(this).children('input')[0].value;


                $(".package-main p").each(function () {
                    var parent = $(this).parent().parent().parent();

                    if ($(this).text().toLocaleLowerCase().indexOf(tagTerm) >= 0) {
                        $(parent).addClass('search-filter');
                        $(parent).removeClass('hidden');
                    }
                });

                $(".type-label").each(function () {
                    var parent = $(this).parent().parent().parent().parent().parent().parent();

                    if ($(this).text().toLocaleLowerCase().indexOf(tagTerm) >= 0) {
                        $(parent).addClass('search-filter');
                        $(parent).removeClass('hidden');

                    }
                });

                $(".tags-list li").each(function () {
                    var parent = $(this).parent().parent().parent().parent().parent().parent();

                    if ($(this).text().toLocaleLowerCase().indexOf(tagTerm) >= 0) {
                        $(parent).addClass('search-filter');
                        $(parent).removeClass('hidden');

                    }
                });

                $(".location-label").each(function () {
                    var parent = $(this).parent().parent().parent().parent().parent().parent();

                    if ($(this).text().toLocaleLowerCase().indexOf(tagTerm) >= 0) {
                        $(parent).addClass('search-filter');
                        $(parent).removeClass('hidden');

                    }
                });

            });

            $('#ul-Location').children('li').each(function () {
                if ($(this).attr('class').indexOf('is-selected') >= 0) {
                    $(this).click();
                }
            });

            addFilter("search");

        } else {
            $(".package-main p").each(function () {
                var parent = $(this).parent().parent().parent();
                $(parent).removeClass('search-filter');
                $(parent).removeClass('hidden');

            });

            $(".type-label").each(function () {
                var parent = $(this).parent().parent().parent().parent().parent().parent();

                $(parent).removeClass('search-filter');
                $(parent).removeClass('hidden')
            });

            $(".tags-list li").each(function () {
                var parent = $(this).parent().parent().parent().parent().parent().parent();

                $(parent).removeClass('search-filter');
                $(parent).removeClass('hidden')
            });

            $(".location-label").each(function () {
                var parent = $(this).parent().parent().parent().parent().parent().parent();

                $(parent).removeClass('search-filter');
                $(parent).removeClass('hidden')
            });


            $('#ul-Location').children('li').each(function () {
                if ($(this).attr('class').indexOf('is-selected') >= 0) {
                    $(this).click();
                }
            });
        }
    } else {

        $(".custom-table-row").removeClass('search-filter');
        removeFilter("search");
        if (searchTerm) {

            $(".custom-table-row span").each(function () {
                if ($(this).text().toLocaleLowerCase().indexOf(searchTerm) >= 0 &&
                    $(this).parent().attr('class').indexOf('hidden') < 0) {
                    $(this).parent().addClass('search-filter');
                }
            });

            addFilter("search");
        }

        $('#ul-Location').children('li').each(function () {
            if ($(this).attr('class').indexOf('is-selected') >= 0) {
                $(this).click();
            }
        });



    }

    applyFilters();
}

function setTaggleFilters() {
    taggles = searchBoxTaggle.getTagValues();
    taggles.forEach(function (element) {
        addFilter(element);
    });
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
    if (appliedFilters.indexOf(filter.replace(/\W+/g, "-").replace(/\ /g, "-")) === -1) {
        appliedFilters.push(filter);
    }
}

function applyFilters() {

    if (listView === 0) {
        var filterString = ".package-group";
        appliedFilters.forEach(function (element) {
            filterString += "." + element.replace(/\W+/g, "-").replace(/\ /g, "-") + "-filter";
        });


        $(".package-group").addClass("hidden");
        $(filterString).removeClass("hidden").addClass('shown');
    }
    else {

        var filterString = ".custom-table-row";
        appliedFilters.forEach(function (element) {
            filterString += "." + element.replace(/\W+/g, "-").replace(/\ /g, "-") + "-filter";
        });


        $(".custom-table-row").addClass("hidden");
        $(filterString).removeClass("hidden").addClass('shown');
    }

}

function removeFilter(filter) {
    appliedFilters.splice(appliedFilters.indexOf(filter), 1);
}

function addLocationClick() {
    $('#ul-Location li').each(function () {
        $(this).attr('onclick', "locationFilter('" + $(this).text().toLocaleLowerCase() + "')");
    });
}

function prepTags() {
    searchBoxTaggle = new Taggle('outerSearchBox',
        {
            saveOnBlur: true,
            placeholder: "Search",
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
    $('.taggle_placeholder').prepend('<i class="ms-SearchBox-icon ms-Icon ms-Icon--search"></i>');
}

function updateAutocomplete() {

    jQuery.fn.extend({
        propAttr: $.fn.prop || $.fn.attr
    });

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

function isListView() {
    listView = 1;
    $('#buildsTable').empty();
    $('#buildsGrid').empty();
    resetFilters();
    $('#tileViewToggle').attr('background-color', '#EFF6FC');
    $('#listViewToggle').attr('background-color', '#C7E0F4');
    getBuild();


}

function isTileView() {
    listView = 0;
    $('#buildsTable').empty();
    $('#buildsGrid').empty();
    $('#tileViewToggle').attr('background-color', '#C7E0F4');
    $('#listViewToggle').attr('background-color', '#EFF6FC');
    resetFilters();
    getBuild();

}

function focusDialog() {
    var winHeight = $(window).height();

    $("#productModal").height(winHeight);
    $("#productList").height(winHeight - 80);

    $('html,body').animate({



        scrollTop: $('.custom-mini-banner').offset().top + 500
    }, 500);

}

function toggleBanner() {
    $('#banner').toggleClass('hidden');
    $('#mini-banner').toggleClass('hidden');

    resizeWindow();
}

function closeDialog() {
    $('#errorMessage').addClass('hidden');
}

function directDL() {
    window.location.href = xmlConfigPath;
}

function resizeWindow() {
    var topBarHeight = $("#topBar");
    var banner = $("#banner");
    var miniBanner = $("#mini-banner");
    var searchBar = $("searchBar");

    var prodModelHeight = 0;
    var buildListHeight = 0;
    var winHeight = $(window).height();
    var winWidth = $(window).width();

    if (miniBanner.is(':visible')) {
        prodModelHeight = winHeight - topBarHeight.height() - miniBanner.height();
        buildListHeight = winHeight - topBarHeight.height() - searchBar.height() - miniBanner.height() - 80;
    } else {
        prodModelHeight = winHeight - topBarHeight.height() - banner.height();
        buildListHeight = winHeight - topBarHeight.height() - searchBar.height() - banner.height() - 80;
    }

    $("#productModal").height(prodModelHeight);
    $("#productList").height(buildListHeight);
    //$("#productContainer").width(winWidth - 50);
}

$(document).resize(function() {
    resizeWindow();
});

$(document).ready(function () {

    $("#btAbout").mouseover(function () {
        $(this).animate({ backgroundColor: "#005a9e" }, 'slow');
    });
    $("#btAbout").mouseout(function () {
        $(this).animate({ backgroundColor: "#0078d7" }, 'slow');
    });

    setVersion('2016');
    getCompanyInfo();
    getLocations(addLocationClick);
    getFilters();
    getBuild();
    getHelp();

    //searchbox filter
    $("#outerSearchBox").keyup(function (e) {
        if ((e.keyCode >= 48 && e.keyCode <= 57) || (e.keyCode >= 65 && e.keyCode <= 90) || e.keyCode == 8) {
            searchBoxFilter(e);
        }
    });

    prepTags();

    resizeWindow();
});