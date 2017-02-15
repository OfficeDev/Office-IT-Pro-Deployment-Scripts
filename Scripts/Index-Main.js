$(document).ready(function() {
         $('#Hamburger').remove();

     checkAddress();
    sideBarHeight();

});

window.onresize = function(){
    sideBarHeight();
}

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

function loadSection(sectionId,item){

    location.hash = '';
    location.hash = '#'+sectionId;   

    $('.Nav-Option').each(function(i,obj){
        $(obj).removeClass('selected');
    });

    $('#partial-views').empty();
    $('#partial-views').load('./Partials/'+sectionId+'.html')

    $(item).addClass('selected');
    //sideBarHeight();
}

function sideBarHeight(){
    var windowWidth = $(window).outerWidth(); 

    if(windowWidth > 450){
         $('#Hamburger').remove();
         $('#siteNav').parents().each(function(i,obj){ 
            $(obj).height('100%');
        });

        $('#siteNav').height($('body').height());
        $('#siteNav').height($(document).height());

        $('#siteNav').children().children().children().children('.Nav-Option').each(function(i,obj){
        $(obj).removeClass('hidden')
        $(obj).removeClass('ms-u-slideUpOut10');
        $(obj).removeClass('ms-u-slideDownIn10');
       });
    }
    else{

        if( $('#Hamburger').length === 0){
            var navHtml = "<div id='Hamburger' class='ms-Grid-row'>\
                          <div class='ms-Grid-col ms-u-sm12'>\
                              <div class='Nav-Option ms-font-l ms-fontWeight-regular ms-u-textAlignLeft' onclick='toggleHamburger()'>\
                                  <i class='ms-Icon ms-Icon--GlobalNavButton ms-fontWeight-regular'></i>\
                              </div>\
                          </div>\
                       </div>  "

           $('#siteNav').children().children().children().children('.Nav-Option').each(function(i,obj){
            $(obj).addClass('hidden')
           });

           navHtml += $('#siteNav').html();
           $('#siteNav').html(navHtml); 
        }

        $('#siteNav').parents().each(function(i,obj){ 
            $(obj).height('auto');
        });

        $('#siteNav').height('auto');
        $('#siteNav').height('auto');
    }
}

function toggleHamburger(){
     $('#siteNav').children().children().children().children('.Nav-Option').each(function(i,obj){
        if($(obj).hasClass('hidden') || $(obj).hasClass('ms-u-slideUpOut10')){
            $(obj).removeClass('hidden');
            $(obj).removeClass('ms-u-slideUpOut10');
            $(obj).addClass('ms-u-slideDownIn10');
        }
        else{
            $(obj).removeClass('ms-u-slideDownIn10');
            $(obj).addClass('ms-u-slideUpOut10');           
            $(obj).addClass('hidden');
        }     
   });
}

function checkAddress(){
    var sectionId = location.hash.replace('#','');
    if(!sectionId){
        sectionId = 'Home' 
    }  
    
    $('.Nav-Option').each(function(i,obj){
         var onclickvalue = ($(obj).attr('onClick'))
        
        if(onclickvalue.indexOf(sectionId) > -1){
            $(obj).click();
        }
    });
}


