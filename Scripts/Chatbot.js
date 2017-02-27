var directLine = new DirectLine.DirectLine(({
    secret: "XluYsZYN9N8.cwA.oFg.Hecj3zMKJbyrtEAgz78Jz1u4__-Bu-_ba4O0eLTw138"
}));

directLine.activity$
.filter(activity => activity.type === 'message' && (activity.from.id === 'Office_Lifecycle_Bot_Dev' || activity.from.id === 'Office_Lifecycle_Bot'))
.subscribe(
    activity => processResponse(activity)
);

function sendMessage(messageText){
	directLine.postActivity({
    from: { id: 'github' }, 
    type: 'message',
    text: messageText
}).subscribe(
    id => console.log("Posted activity, assigned ID ", id),
    activity => console.log("activity", activity),
    error => console.log("Error posting activity"));
}

function processResponse(activity){
	if(typeof(activity.attachments) !== 'undefined' && activity.attachments.length > 0){
	var responseType = activity.attachments[0].contentType;
	var card = activity.attachments[0];

	switch(responseType){

		case "application/vnd.microsoft.card.thumbnail":
			buildThumbNailCard(activity);
			break; 
		case "application/vnd.microsoft.card.hero":
			buildHeroCard(activity);
			break;
		default:
			var cardBody = card.content.text; 
			var response = "<div class='ms-Grid-row ChatResponse'><div class='ms-Grid-col col-u-sm8'>"+cardBody+"</div></div>"; 
			$('.ChatArea').append(response); 
			break;
		}
	}
	else{
		var response = "<div class='ms-Grid-row ChatResponse'><div class='ms-Grid-col col-u-sm8'>"+activity.text+"</div></div>"; 
		$('.ChatArea').append(response); 
	}
	updateScroll();
}

function messageKeyDown(event){
	if(event.keyCode == 13){
		var text = $('#Message-txtbx').val(); 
		if(text.trim() !== ''){
			processPrompt(text);
			$('#Message-txtbx').val(''); 
		}
	}
}

function processPrompt(messageText){		
		if(messageText.trim() != ''){
			sendMessage(messageText);
			var response = "<div class='ms-Grid-row ChatPrompt'><div class='ms-Grid-col col-u-sm8'>"+messageText+"</div></div>"; 
			var html = $('.ChatArea').html(); 
		    $('#Message-txtbx').val(''); 
			$('.ChatArea').append(response); 
			updateScroll();
		}
}

function buildThumbNailCard(activity){
	var card = activity.attachments[0];
	var imageUrl = card.content.images[0].url;
	var responseText = card.content.text; 
	var response = "<div class='ms-Grid-row ChatResponse'><img class='ms-Grid-col col-u-sm8' src='"+imageUrl+"'>"+responseText+"</img></div>"
	$('.ChatArea').append(response); 
}

function buildHeroCard(activity)
{		
	var card = activity.attachments[0];
	var text = card.content.text;
	var buttons = card.content.buttons; 
	var response = "<div class='ms-Grid-row ChatResponse'><div class='ms-Grid'><div class='ms-Grid-row'><div class='ms-Grid-col col-u-sm12'>"+text+"</div></div><div class='ms-u-Grid-row'>";

	buttons.forEach(function(currentValue,index, arr){
		var button = '<button class="ms-Button CardButton" onclick="processPrompt(\''+currentValue.title+'\')"><span class="ms-Button-label">'+currentValue.title+'</span></button>'
		response += button; 
	});
	response += "</div></div></div>"; 
	$('.ChatArea').append(response); 
}

function updateScroll(){
   $(".ChatArea").animate({ scrollTop: $('.ChatArea').prop("scrollHeight")}, 1000);

}