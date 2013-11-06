$( document ).ready(function() {
	var separatorLineElement = $( ".UICalendarPortlet .calendarWorkingWorkspace .uiActionBar .btnRight .separatorLine" );
	if(!separatorLineElement || !separatorLineElement.length || separatorLineElement.length == 0) {
		return;
	}
	separatorLineElement.before("<a href='#' class='ExchangeSettingsButton pull-right'><img src='/exchange-resources/skin/images/exchange.png' width='24px' height='24px'/></a>");

	$('.ExchangeSettingsButton').click(function(e) {
    	if($('.ExchangeSettingsWindow').is(':visible')) {
    		return;
    	}
		$('.ExchangeSettingsWindow .ExchangeSettingsContent').html("<div class='ExchangeSettingsLoading'>Loading...</div>");
    	$.getJSON("/portal/rest/exchange/calendars", function(data){
    		$('.ExchangeSettingsWindow .ExchangeSettingsContent').html("");
        	if(!data || data.length == 0) {
    			$('.ExchangeSettingsWindow .ExchangeSettingsContent').html("<div class='ExchangeSettingsError'>User seems not connected to Exchange</div>");
    		} else {
	    	    $.each(data, function(i,item){
	    	    	$('.ExchangeSettingsWindow .ExchangeSettingsContent').append(""+item.name+"<input type='checkbox' "+(item.synchronizedFolder?"checked":"")+" name='"+item.name+"' value='"+item.id+"' /><BR/>");
	    	    });
	        	$('.ExchangeSettingsWindow input[type="checkbox"]').click(function(){
	        	    if($(this).is(':checked')){
	        	    	$.get("/portal/rest/exchange/sync?"+$.param({folderId : $(this).val()}));
	        	    } else {
	        	    	$.get("/portal/rest/exchange/unsync?"+$.param({folderId : $(this).val()}));
	        	    }
	        	});
    		}
    	});
    	$('.ExchangeSettingsWindow').css('top', (separatorLineElement.position().top + 25) + 'px');
    	$('.ExchangeSettingsWindow').css('right', ($(window).width() - separatorLineElement.position().left - 37) + 'px');

    	$('.ExchangeSettingsMask').show();
	    $('.ExchangeSettingsWindow').show();
	});

    $("body").append("<div class='ExchangeSettingsWindow' />");
	$('.ExchangeSettingsWindow').hide();
	$('.ExchangeSettingsWindow').html("<div class='ExchangeSettingsTitle'><h6>My Exchange Calendars</h6><div class='ExchangeSettingsInfo'>Sync with eXo</div></div><div class='ExchangeSettingsContent'></div>");

    $("body").append("<div class='ExchangeSettingsMask' />");
    $('.ExchangeSettingsMask').hide();
	$('.ExchangeSettingsMask').click(function(e) {
		if (e.target.id == 'ExchangeSettingsMask') {
			return true;
		} else {
		    $('.ExchangeSettingsMask').hide();
			$('.ExchangeSettingsWindow').hide();
		}
	});
});