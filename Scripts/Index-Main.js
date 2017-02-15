$(document).ready(function() {

    checkAddress();
   sideBarHeight();
});

function openInNewTab(url) {
      var win = window.open(url, '_blank');
      win.focus();
    }

function downloadZip(){
      window.open("https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/zipball/master");
        return false;
}

function toggleSection(item){

    var section_body = $(item).parent().children('.Section-Body') 
    var chevron = $(item).parent().children('.Section-Header').children('.section-Header-Chevron'); 

    if(section_body.hasClass('ms-u-slideDownIn10')){
        section_body.removeClass('ms-u-slideDownIn10');
        section_body.addClass('ms-u-slideUpOut10');
        section_body.addClass('hidden');

        chevron.removeClass('open-Chevron');
        chevron.addClass('close-Chevron');


    }else if(section_body.hasClass('hidden') || section_body.hasClass('ms-u-slideUpOut10')){
        var parentId =  $(item).parent().children('.Section-Header').attr('id') 
        section_body.removeClass('hidden');
        section_body.addClass('ms-u-slideDownIn10');
        section_body.removeClass('ms-u-slideUpOut10');

        chevron.removeClass('close-Chevron');
        chevron.addClass('open-Chevron');
    }
    sideBarHeight();
}

function focusSection(sectionId,item){

    location.hash = '';
    location.hash = '#'+sectionId;   

     $('html, body').animate({
        scrollTop: $('#'+sectionId).offset().top
    }, 500);
    

    $('#'+sectionId).click();

    $('.Nav-Option').each(function(i,obj){
        $(obj).removeClass('selected');
    });

    $(item).addClass('selected');
    sideBarHeight();
}

function loadSection(sectionId,item){

    location.hash = '';
    location.hash = '#'+sectionId;   

    $('.Nav-Option').each(function(i,obj){
        $(obj).removeClass('selected');
    });

    $('#partial-views').empty();
    $('#partial-views').load('./Partials/'+sectionId+'.html')

    $(item).addClass('selected');
    sideBarHeight();
}

function sideBarHeight(){
     $('#siteNav').parents().each(function(i,obj){ 
        $(obj).height('100%');
    });

    $('#siteNav').height($('body').height());
    $('#siteNav').height($(document).height());
}

function checkAddress(){
    var sectionId = location.hash.replace('#','');

    if(sectionId){
         $('.Nav-Option').each(function(i,obj){
        var onclickvalue = ($(obj).attr('onClick'))
        
        if(onclickvalue.indexOf(sectionId) > -1){
            $(obj).click();
        }
    });
    }  
}


