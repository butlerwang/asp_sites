
//信息切换
$(".hjone").hide();
$(".hjone:eq(0)").show();
$(".hjnavleft li").each(function(index){
       $(this).mouseover(
	   	  function(){
			  $(".hjnavleft li").removeClass();
			  $(this).addClass("hover"+index+"");
			  $(".hjone:visible").hide();
			  $(".hjone:eq(" + index + ")").show();
	  })
   })

