var a=/\.(avi|asf|fla|flv|mov|rm|rmvb|ra|mp3|mp4|mpg|mpeg|qt|swf|wma|wmv)(?:$|\?)/i,b=/^\d+(?:\.\d+)?$/;function c(f){if(b.test(f))return f+'px';return f;};function d(f){var g=f.attributes;return a.test(g.src||'');};function e(f,g){var h=f.createFakeParserElement(g,'cke_media','flvPlayer',true),i=h.attributes.style||'',j=g.attributes.width,k=g.attributes.height;if(typeof j!='undefined')i=h.attributes.style=i+'width:'+c(j)+';';if(typeof k!='undefined')i=h.attributes.style=i+'height:'+c(k)+';';return h;};

CKEDITOR.plugins.add('flvPlayer',
{
    init: function(editor)    
    {        
        //plugin code goes here
        var pluginName = 'flvPlayer';        
        CKEDITOR.dialog.add(pluginName, this.path + 'dialogs/flvPlayer.js');        
        editor.addCommand(pluginName, new CKEDITOR.dialogCommand(pluginName));        
   
        editor.ui.addButton('flvPlayer',
        {               
            label: '插入Flv等视频',
			icon:this.path+'images/media.gif',
            command: pluginName
        });
		
		editor.addCss('img.cke_media{background-image: url('+CKEDITOR.getUrl(this.path+'images/placeholder.png')+');'+'background-position: center center;'+'background-repeat: no-repeat;'+'border: 1px solid #a9a9a9;'+'width: 80px;'+'height: 80px;'+'}');   //背景
		
    },
	afterInit:function(f){var g=f.dataProcessor,h=g&&g.dataFilter;if(h)h.addRules({elements:{'cke:embed':function(i){if(!d(i))return null;return e(f,i);}}},4);}
	
});
