var SEslider = function(container, slider, thisname)
{
	this.Container = document.getElementById(container)
	this.Container.style.overflow = "hidden";
	this.Container.style.position = "relative";
	this.Slider = document.getElementById(slider);
	this.Slider.style.position = "absolute";
	this.Items = this.Slider.getElementsByTagName("img");
	this.Count = this.Items.length;
	this.Width = parseInt(this.Container.style.width);
	this.Height = parseInt(this.Container.style.height);
	this.Name = thisname;
	this.BindView();
	
	this.Index = 0;
	this.PreviousIndex = 0;
	this.Pause = 2000;
	this.Timer = null;
	this.Auto = true;
}
SEslider.prototype.BindView = function()
{
	this.codeArray = new Array();
	var _codeDiv = document.createElement("div");
	var _codeSpan = null;
	for(var i=0; i<this.Count; i++)
	{
		this.Items[i].style.border = "none";
		this.Items[i].style.width = this.Width + "px";
		this.Items[i].style.height = this.Height + "px";
		_codeSpan = document.createElement("span");
		_codeSpan.style.fontSize = "12px";
		_codeSpan.style.color = "#FF6600";
		_codeSpan.style.cursor = "pointer";
		_codeSpan.style.border = "1px solid #FF6600";
		_codeSpan.style.backgroundColor = "#FFFFFF";
		_codeSpan.style.padding = "2px 5px 2px 5px";
		_codeSpan.style.marginLeft = "5px";
		_codeSpan.onmouseover = new Function(this.Name + '.Mover(' + i + ')');
		_codeSpan.onmouseout = new Function(this.Name + '.Mout()');
		_codeSpan.innerHTML = (i + 1).toString();
		
		_codeDiv.appendChild(_codeSpan);
		this.codeArray.push(_codeSpan);
	}
	_codeDiv.style.position = "absolute";
	_codeDiv.style.right = "10px";
	_codeDiv.style.bottom = "10px";
	this.Container.appendChild(_codeDiv);
}
SEslider.prototype.Move = function()
{
	if (this.Index >= this.Count) {
		this.Index = 0;
	}
	this.codeArray[this.PreviousIndex].style.color = "#FF6600";
	this.codeArray[this.PreviousIndex].style.backgroundColor= "#FFFFFF";
	this.codeArray[this.Index].style.color = "#FFFFFF";
	this.codeArray[this.Index].style.backgroundColor= "#FF6600";
	this.PreviousIndex = this.Index;
	this.Run(-1 * this.Height * this.Index);
}
SEslider.prototype.Run = function(targetValue)
{
	var _instance = this;
	clearTimeout(_instance.Timer);
	var currentValue = parseInt(_instance.Slider.style["top"]) || 0;
	var stepValue = _instance.GetStep(currentValue, targetValue);
	if(stepValue != 0)
	{
		_instance.Slider.style["top"] = (currentValue + stepValue) + "px";
		_instance.Timer = setTimeout(function(){ _instance.Run(targetValue); }, 20);
	}
	else
	{
		_instance.Slider.style["top"] = targetValue + "px";
		if (_instance.Auto)
		{
			_instance.Index++;
			_instance.Timer = setTimeout(function(){ _instance.Move(); }, _instance.Pause);
		}
	}
}
SEslider.prototype.Mover = function(index)
{
	this.Auto = false;
	this.Index = index;
	this.Move();
}
SEslider.prototype.Mout = function()
{
	this.Auto = true;
	this.Move();
}
SEslider.prototype.GetStep = function(currentValue, targetValue)
{
	var _step = (targetValue - currentValue) / 5;
	if (_step == 0) return 0;
	if (Math.abs(_step) < 1) return (_step > 0 ? 1 : -1);
	return _step;
}