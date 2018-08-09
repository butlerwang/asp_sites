jQuery.fn.rater = function(url, options , callback)
{
	var settings = {
		url			: url ,
		start		: 1 ,
		step		: 1 ,
		maxvalue	: 3 ,
		curvalue	: 0 ,
		//title		: null ,
		enabled		: true
	};
	
	if(options) { jQuery.extend(settings, options); };
	jQuery.extend(settings, {cancel: (settings.maxvalue > 1) ? true : false});
	
	var container = jQuery(this);
	jQuery.extend(container, { averageRating: settings.curvalue, url: settings.url });

	var starWidth	= 25;
	var raterWidth	= (settings.maxvalue - settings.start + settings.step)/settings.step * starWidth;
	var curvalueWidth	= (settings.curvalue - settings.start + settings.step)/settings.step * starWidth;
	
	var title	= '';
	if (typeof settings.title == 'object' && typeof settings.title[settings.curvalue] == 'string') {
		title	= settings.title[settings.curvalue];
	} else {
		title	= settings.curvalue+'/'+settings.maxvalue;
	}
	
	var ratingParent	= '<ul class="rating" style="width:'+raterWidth+'px" title="'+title+'">';
	container.html(ratingParent);
	
	var listItems	= '<li class="current" style="width:'+curvalueWidth+'px"></li>';
	
	if (settings.enabled){
		var k = 0;
		for (var i = settings.start;i <= settings.maxvalue;i = i + settings.step) {
			k++;
			if (typeof settings.title == 'object' && typeof settings.title[i] == 'string') {
				title	= settings.title[i];
			} else {
				//title	= i+'/'+settings.maxvalue;
				title	= i;
			}
			
			listItems	+= '<li class="star" style="width:'+(k * starWidth)+'px;z-index:'+((settings.maxvalue - i)  / settings.step + 1)+'" title="'+title+'"></li>';
			
		}
	}
	//alert(listItems);
	container.find('.rating').html(listItems);
	container.find('.rating').find('.star').hover(function() {
		container.find('.rating').find('.current').hide();
		this.className	= 'star_hover';
	} , function() {
		container.find('.rating').find('.current').show();
		this.className	= 'star';
	});
	
	container.find('.rating').find('.star').click(function() {
		//var value	= settings.maxvalue - $(this).css('z-index') + 1;
		var value	= $(this).attr('title');
		container.find('.rating').find('.current').width((value - settings.start + settings.step)/settings.step * starWidth);
		
		if (url) {
			$.post(url , {value:value} , function(response) {
				if (typeof callback == 'function') {
					callback(container , value , response);	
				}
			});
			return;
		}
		
		if (typeof callback == 'function') {
			callback(container , value);
			return ;
		}
		
	});
}