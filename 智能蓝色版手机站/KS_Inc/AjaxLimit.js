
function calculagraph(){
	this._id=null;
	//this._sT=null;
	this._cT=null;
	this._eT=null;
	this._lT=null;
	this._tasktype=null;
	this.timerRunning=false;
	this.NowYear=null;
	this.NowDate=null;
	this.NowMonth=null;
	this.NowHour=null;
	this.NowMinute=null;
	this.NowSecond=null;
	this._gT=function(){
		if (this._tasktype==1){  //限时抢购
				//if (this._lT==null){
					var t=this._cT.split(' ')[0];
					var tarr=t.split('-');
					if (this.timerRunning==false){
						this.NowYear = tarr[0];
						this.NowMonth = tarr[1];   
						this.NowDate = tarr[2]; 
						t=this._cT.split(' ')[1];
						if (t==''||t==null) t='00:00:00'
						tarr=t.split(':');
						this.NowHour = tarr[0];   
						this.NowMinute = tarr[1]; 
						this.NowSecond = tarr.length==3?tarr[2]:0;  
					}else{
						if (this.NowMonth>12){
							this.NowMonth=0;
							this.NowYear++;
						}
						
						if(this.NowDate>=parseInt(getDaysInMonth(this.NowYear,this.NowMonth))){
							this.NowDate=0;
							this.NowMonth++;
						}

						if (this.NowHour>=24){
							this.NowHour=0;
							this.NowDate++;
						}
						if (this.NowMinute>=59){
							this.NowMinute=0;
							this.NowHour++;
						}
						if (this.NowSecond>=59){
							this.NowSecond=0;
							this.NowMinute++;
						}
						this.NowSecond++;
					}
					
					
					if (this.NowYear <2000)   
					this.NowYear=1900+this.NowYear;  
					var t=this._eT.split(' ')[0];
					var tarr=t.split('-');
					Yearleft = tarr[0] - this.NowYear   
					Monthleft = tarr[1] - this.NowMonth  
					Dateleft = tarr[2] - this.NowDate
					
					
					t=this._eT.split(' ')[1];
					if (t==''||t==null) t='00:00:00'
					tarr=t.split(':');
					var eHour = tarr[0];   
					var eMinute = tarr[1];   
					var eSecond = tarr.length==3?tarr[2]:0;   
					
					Hourleft = eHour - this.NowHour   
					Minuteleft = eMinute - this.NowMinute   
					Secondleft = eSecond - this.NowSecond   
					if (Secondleft<0)   
					{   
					Secondleft=60+Secondleft;   
					Minuteleft=Minuteleft-1;   
					}   
					if (Minuteleft<0)   
					{    
					Minuteleft=60+Minuteleft;   
					Hourleft=Hourleft-1;   
					}   
					if (Hourleft<0)   
					{   
					Hourleft=24+Hourleft;   
					Dateleft=Dateleft-1;   
					}   
					if (Dateleft<0)   
					{   
					Dateleft=31+Dateleft;   
					Monthleft=Monthleft-1;   
					}   
					if (Monthleft<0)   
					{   
					Monthleft=12+Monthleft;   
					Yearleft=Yearleft-1;   
					} 
					var Temp='';
					if (Yearleft>0){
						Temp='<strong>'+Yearleft+'</strong>年'
					}
					if (Monthleft>0){
						Temp+='<strong>'+Monthleft+'</strong>月';
					}
					Temp+='<strong>'+Dateleft+'</strong>天<strong>'+Hourleft+'</strong>小时<strong>'+Minuteleft+'</strong>分<strong>'+Secondleft+'</strong>秒';	
					if (getMinuteInDates(this._sT.replace(/(-)/g,'/'),this.NowYear+'/'+this.NowMonth+'/'+this.NowDate+' '+this.NowHour+':'+this.NowMinute+':'+this.NowSecond)<0){
					  document.getElementById(this._id).innerHTML='<span style="color:red">开始时间：'+this._sT+'</span>';   
					}
					else if (getMinuteInDates(this.NowYear+'/'+this.NowMonth+'/'+this.NowDate+' '+this.NowHour+':'+this.NowMinute+':'+this.NowSecond,this._eT.replace(/(-)/g,'/'))>=0){
					  document.getElementById(this._id).innerHTML=Temp;   
					}else{
					  document.getElementById(this._id).innerHTML="<strong style='font-size:14px;'>抢购结束</strong>";  
					  clearInterval(timerID);
					}
					this.timerRunning = true;   
					var timerID = setTimeout("function(){var oo=this;oo.gT()}",1000);   

			
		}else{  //限量抢购
					document.getElementById(this._id).innerHTML="<strong style='font-size:14px;'>限量抢购</strong>";
					clearInterval(this._interval);
		}
	}
	this._interval=function(){
		var o=this;
		this._interval=setInterval(function(){o._gT()},1000)
	}
}

function getDaysInMonth(year,month){
      month = parseInt(month,10)+1;
      var temp = new Date(year+"/"+month+"/0");
      return temp.getDate();
}

/**
 * 判断两个时间这间间隔几分钟

 * date1与date2格式：yyyy/MM/dd hh:mm:ss ，它们是字符串类型
 * */
function getMinuteInDates(date1,date2){
    var beginDate= new Date(date1);
    var endDate = new Date(date2);
    
    var date = endDate.getTime() - beginDate.getTime();
    
    var time = Math.floor(date / (1000 * 60   ));
    return time;
}


function GetHtmlStr(id,num){
    $.ajax({
      type: "get",
      url: "/shop/limitBuy.asp",
      data: "id="+id+"&num="+num+"&fresh=" + Math.random(),
      cache:false, 
      success: function(result){
		    result=unescape(result);
            $("#loading"+id).hide();
            $("#hasQiangGou"+id).show();
            eval(result.split('|')[0]);
            $("#qianggou"+id).html(result.split('|')[1]);
      }
    });
}

//调用限时/限量抢购
function getLimitBuy(taskid,num)
{
	document.writeln('<div id="loading'+taskid+'" class="loading"><img src="/images/loading.gif" /></div>');
	document.writeln('<div id="hasQiangGou'+taskid+'" style="display:none;">');
	document.writeln('	 <div class="timeBox" id="time'+taskid+'">正在加载…</div>');
	document.writeln('	 <div class="Product_List_S" id="qianggou'+taskid+'"></div>');
	document.writeln('</div>');
	GetHtmlStr(taskid,num);  //异步调用主方法
}



