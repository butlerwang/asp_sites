function LoadFlash(url,wmode,width,Height,param)
{ 
document.write(
  '<embed src="' + url + '" FlashVars="'+param+'" wmode=' + wmode +
  ' quality="high" pluginspage=http://www.macromedia.com/go/getflashplayer type="application/x-shockwave-flash" width="' + width + 
  '" height="' + Height + '"></embed>');   
}