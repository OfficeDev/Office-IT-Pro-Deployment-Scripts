$(document).ready(function() {

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

    section_body.toggleClass('ms-u-slideUpIn10 ms-u-slideUpOut10')
}


