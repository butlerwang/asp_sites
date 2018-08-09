function time1()
{   today=new Date();
	function initArray(){  
	this.length=initArray.arguments.length
	for(var i=0;i<this.length;i++)
	this[i+1]=initArray.arguments[i]  }
	document.write(
	" ",
	today.getFullYear(),"年", 
	today.getMonth()+1,"月",
	today.getDate(),"日",
	"");
}
time1();