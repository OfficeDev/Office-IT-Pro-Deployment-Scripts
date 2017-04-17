$(document).ready(function() {    
    loadPage();
});

window.onresize = function(){
    resizePage();
}

// $(window).bind( 'hashchange', function(e) 
// {
//     loadPage(); 
// });


function loadPage(){
    $('#Hamburger').remove();
    addHamburger();
    checkAddress();
}

function resizePage(){

    if(window.innerWidth >= 640){
        $('#Hamburger').remove();
        $('.Nav-Option').each(function(i,obj){
            $(obj).css('display','initial');
            $(obj).removeClass('ms-u-slideUpOut10');
        }); 
        
        $('.Site-Content').height('auto');
        $('#partialViews').height('100%'); 
    }
    else{
        $('.Site-Content').height('auto');
        $('#Nav').height('auto'); 
        $('#partialViews').height('auto'); 
        $('#trendingTopics').height('auto');  

        addHamburger();
    }

}

function openInNewTab(url) {
      var win = window.open(url, '_blank');
      win.focus();
}

function downloadZip(){
      window.open("https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/zipball/master");
        return false;
}

function loadSection(sectionId){
    window.location.hash = sectionId; 

    if (location.hash) {
      setTimeout(function() {

        window.scrollTo(0, 0);
      }, 1);
    }

    $('.Nav-Option').each(function(i,obj){
        $(obj).removeClass('Selected'); 
    });

    $('#'+sectionId).addClass('Selected');

    $('#partial-views').empty();
    $.ajax({
        type: 'GET',
        url: './Partials/'+sectionId+'.html',    
        dataType: 'html', 
        success:function(result)
        {
           $('#partialViews').html(result);   
        }
     });  
}         

function toggleHamburger(){
     $('.Nav-Option').each(function(i,obj){

        if($(obj).hasClass('ms-u-slideUpOut10')){
            $(obj).css('display','initial');
            $(obj).removeClass('ms-u-slideUpOut10');
            $(obj).addClass('ms-u-slideDownIn10');
        }
        else if (!$(obj).css('display') != 'initial' && $(obj).attr('onclick') != "toggleHamburger()"){
            $(obj).removeClass('ms-u-slideDownIn10');
            $(obj).addClass('ms-u-slideUpOut10');           
            $(obj).css('display','none');
        }});
}

function addHamburger(){
    if( $('#Hamburger').length === 0 && window.innerWidth < 640){
            var navHtml = "<div id='Hamburger' class='ms-Grid-row'>\
                          <div class='ms-Grid-col ms-u-sm12'>\
                              <div class='Nav-Option ms-font-l ms-fontWeight-regular ms-u-textAlignLeft' onclick='toggleHamburger()'>\
                                  <i class='ms-Icon ms-Icon--GlobalNavButton ms-fontWeight-regular'></i>\
                              </div>\
                          </div>\
                       </div>  "

           $('.Nav-Option').each(function(i,obj){
            $(obj).css('display','none');
           });

           $('#Nav').find('.ms-Grid').prepend(navHtml);
        }

}

function checkAddress(){
    var pageId = location.hash.split('#')[1];
    var sectionId = location.hash.split('#')[2];

    if(pageId === undefined){
        pageId = 'home' 
                loadSection(pageId);
    }  
    else{
        loadSection(pageId);
        if(sectionId != undefined){
            location.href += '#' +sectionId; 
            $("html, body").delay(2000).animate({scrollTop: $(sectionId).offset().top()}, 2000);
        }
    }

    resizePage();
}

function toggleCopyLink(section){
    var textFieldParent = $(section).parent().siblings()[0];
    var copyButton = $(section).parent().siblings()[1];
    var textField = $(textFieldParent).children('input'); 
    var url = location.href; 
    var sectionId = $(section).attr('Id');

    $(textField).val(url+"#"+sectionId);
    $(textFieldParent).toggleClass('hidden');
    $(copyButton).toggleClass('hidden');
}


function copyToClipboard(icon){
   var textFieldParent = $(icon).parent().siblings()[1];
   var textField = $(textFieldParent).children()[0];

  textField.select();
  document.execCommand('copy');
}

