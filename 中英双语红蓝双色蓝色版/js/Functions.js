
// nav menu
var timeout	= 500;
var closetimer	= 0;
var ddmenuitem	= 0;

// open hidden layer
function mopen(id)
{	
	// cancel close timer
	mcancelclosetime();

	// close old layer
	if(ddmenuitem) ddmenuitem.style.visibility = 'hidden';

	// get new layer and show it
	ddmenuitem = document.getElementById(id);
	ddmenuitem.style.visibility = 'visible';

}
// close showed layer
function mclose()
{
	if(ddmenuitem) ddmenuitem.style.visibility = 'hidden';
}

// go close timer
function mclosetime()
{
	closetimer = window.setTimeout(mclose, timeout);
}

// cancel close timer
function mcancelclosetime()
{
	if(closetimer)
	{
		window.clearTimeout(closetimer);
		closetimer = null;
	}
}


//��Ʒ���ർ��

var menuids=["suckertree1"] //Enter id(s) of SuckerTree UL menus, separated by commas

function buildsubmenus(){
for (var i=0; i<menuids.length; i++){
  var ultags=document.getElementById(menuids[i]).getElementsByTagName("ul")
    for (var t=0; t<ultags.length; t++){
    ultags[t].parentNode.getElementsByTagName("a")[0].className="subfolderstyle"
    ultags[t].parentNode.onmouseover=function(){
    this.getElementsByTagName("ul")[0].style.display="block"
    }
    ultags[t].parentNode.onmouseout=function(){
    this.getElementsByTagName("ul")[0].style.display="none"
    }
    }
  }
}

if (window.addEventListener)
window.addEventListener("load", buildsubmenus, false)
else if (window.attachEvent)
window.attachEvent("onload", buildsubmenus)




// close layer when click-out
document.onclick = mclose; 

function comment_check() {
if ( document.form1.name.value == '' ) {
window.alert('����������^_^');
document.form1.name.focus();
return false;}

if ( document.form1.email.value.length> 0 &&!document.form1.email.value.indexOf('@')==-1|document.form1.email.value.indexOf('.')==-1 ) {
window.alert('��������ȷ��Email��ַ����:webmaster@Junww.com');
document.form1.email.focus();
return false;}

if(document.form1.qq.value.search(/^([0-9]*)([.]?)([0-9]*)$/)   ==   -1)   
      {   
  window.alert("QQֻ��������^_^");   
document.form1.qq.focus();
return false;}

if ( document.form1.content.value == '' ) {
window.alert('����������^_^');
document.form1.content.focus();
return false;}

if ( document.form1.verycode.value == '' ) {
window.alert('��������֤��^_^');
document.form1.verycode.focus();
return false;}

return true;}




function order_check() {
if ( document.form1.ordercount.value == '' ) {
window.alert('�����붩������^_^');
document.form1.ordercount.focus();
return false;}

if(document.form1.ordercount.value.search(/^([0-9]*)([.]?)([0-9]*)$/)   ==   -1)   
      {   
  window.alert("��������ֻ��������^_^");   
document.form1.ordercount.focus();
return false;}


if ( document.form1.name.value == '' ) {
window.alert('��������ϵ��^_^');
document.form1.name.focus();
return false;}

if ( document.form1.address.value == '' ) {
window.alert('��������ϵ��ַ^_^');
document.form1.address.focus();
return false;}

if ( document.form1.tel.value == '' ) {
window.alert('��������ϵ�绰^_^');
document.form1.tel.focus();
return false;}

if ( document.form1.email.value.length> 0 &&!document.form1.email.value.indexOf('@')==-1|document.form1.email.value.indexOf('.')==-1 ) {
window.alert('��������ȷ��Email��ַ����:webmaster@Junww.com');
document.form1.email.focus();
return false;}

if(document.form1.qq.value.search(/^([0-9]*)([.]?)([0-9]*)$/)   ==   -1)   
      {   
  window.alert("QQֻ��������^_^");   
document.form1.qq.focus();
return false;}


if ( document.form1.verycode.value == '' ) {
window.alert('��������֤��^_^');
document.form1.verycode.focus();
return false;}

return true;}




// ��ҳ����ת�� 
  
// -------------- ���²����󲿷ֿ��Ը��� -------------------- 
//s = simplified �������� t = traditional �������� n = normal ������ʾ 
var zh_default = 'n'; //Ĭ�����ԣ��벻Ҫ�ı� 
var zh_choose = 'n'; //��ǰѡ�� 
var zh_expires = 7; //cookie�������� 
var zh_class = 'zh_click'; //���ӵ�class����idΪclass + s/t/n ֮һ 
var zh_style_active = 'font-weight:normal; color:#999999;'; //��ǰѡ�������ʽ�� 
var zh_style_inactive = 'color:#c00108;'; //�ǵ�ǰѡ�������ʽ�� 
var zh_browserLang = ''; //��������� 
var zh_autoLang_t = true; //���������Ϊ����ʱ�Զ����в��� 
var zh_autoLang_s = false; //���������Ϊ����ʱ�Զ����в��� 
var zh_autoLang_alert = true; //�Զ��������Ƿ���ʾ��ʾ��Ϣ 
//�Զ����������ʾ��Ϣ 
var zh_autoLang_msg = '�gӭ����վ,��վ������̨����۵��Ñ�\n1.����UTF-8���H���a,���κ��Z�԰l���������D�a.\n2.�Ԅ��Д෱�w�Ñ�,�@ʾ���w�W�\n3.�ھW����Ϸ����Z���x��,�����[�І��}�r�����ГQ\n4.����Ϣ��cookie��Ч�ڃ�ֻ�@ʾһ��'; 
var zh_autoLang_checked = 0; //�μ�����������,��һ��дcookieΪ1,��ʾ��Ϊ2,��󽫲�����ʾ 
  
  
//�ж���������Ե�����,ieΪСд,ffΪ��д 
var zh_langReg_t = /^zh-tw|zh-hk$/i; 
var zh_langReg_s = /^zh-cn$/i; 
  
//���己������ֱ�,���������滻 
var zh_s = '�������������°Ӱհڰܰ���������������������������������ʱϱбձ߱�����������������������������βϲвѲҲӲԲղֲײ޲����������������������������������������������³ĳųƳͳϳҳճٳ۳ܳݳ�������������������������������������´ǴʴʹϴдѴӴԴմܴ������������������������������������������Ƶ˵еӵݵ޵��������������������������������������ĶƶͶ϶жҶӶԶֶٶ۶�����������������������������������÷ķɷϷѷ׷طܷ߷������������������������������øƸǸɸϸѸӸԸոָٸڸ޸����������������������������ƹ˹йع۹ݹ߹�����������������������������Һ׺غ���������������������������������������������ƻѻӻԻٻ߻���������������������������������������������üƼǼʼ̼ͼмԼռּؼۼݼ߼����������������������������������������������������������������½ýĽŽȽɽʽνϽս׽ھ����������������������ǾɾԾپݾ�����ܽ������������������������������������������ſǿοѿҿٿ��������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������¢£¤¥¦§¨«¬­®¯°±²³¸»¼½¿����������������������������������������������������������������������������������������áèêíóôùûþ����������������������������������ıĶ����������������������������������������šŢťŦŧŨũűŵŷŸŹŻŽ���ӹ���������ƭƮƵƶƻƾ������������������������������ǣǤǥǦǨǩǫǮǯǱǳǴǵǹǺǽǾǿ��������������������������������������������ȣȧȨȰȴȵ������������������������������������ɡɥɧɨɬɱɴɸɹ��������������������������������������ʤʥʦʨʪʫʬʱʴʵʶʻ��������������������������������˧˫˭˰˳˵˶˸˿��������������������������������������̷̸̡̢̧̯̰̱̲̳̾��������������������������������ͭͳͷͼͿ������������������������������ΤΥΧΪΫάέΰαγν������������������������������������������������ϮϰϳϷϸϺϽϿ��������������������������������������������������������ХЫЭЮЯвгдклп��������������������������ѡѢѤѧѫѯѰѱѵѶѷѹѻѼ������������������������������������������������ҡҢңҤҥҩүҳҵҶҽҿ��������������������������������������ӣӤӥӦӧӨөӪӫӬӱӴӵӶӸӻӽӿ������������������������������������ԤԦԧԨԯ԰ԱԲԵԶԸԼԾԿ������������������������������������������������������������������աբթիծձյնշոջս����������������������������������������ְִֽֿ֣֤֡֯��������������������������������������������פרשת׬׮ׯװױ׳״׶׸׹׺׻����������������������������ôΪֻ��׼������������й'; 
var zh_t = '�}�@�K���O�\�W���T�[���C�k�O�ͽ��^�r������U݅ؐ�^�N��v���P�����]߅���H׃�q�p���T�l�I�e�P����K�g�N�a���Q���M�K�N�nœ�}����ȃԜy��Ԍ�v���s�׋�p�P�b�U��L�L���c�S���n܇�؉m��r�ηQ���\�G�V�t�Y�u�X��_�x�����P�I�I�h���N�z�r�A���|̎�����J���N���b�o�~�n�[��ą����Z�e�_���J��������đ���Q�������hʎ�n�v�u�\���I���������f���c�|늝���{��ՙ�B���Vӆ�|�ӗ����Y�٪��xـ�呔྄��ꠌ����D�g�Z�Z�~Ӟ���I�����D�E�l�P�y�m�\�C����؜��L���w�U�M�����^���S�S���h�L���T�p�S�P�wݗ���o�x�}ؓӇ�D�`ԓ�}�w���s���M����䓾V���V怔R���w�t���o���m�ؕ�h�Ϙ�ُ���M��P�^�^�T؞�VҎ���w���|܉Ԏ���F��݁�L假��^��n�h�u�Q�R�M�Z���t����o����W�A����Ԓ�щĚg�h߀���Q�������o�S�e�]�x���V�x���Z���M�d�Lȝ��ⷫ@؛�����C�e���I�u�����O݋���D���E����Ӌӛ�H�^�o�A�v�a�Z⛃r�{���O�Թ{�g�D�}�O�z�A�|���캆���p�]���b�`�vҊ�IŞ���T�u�R���{�Y�����v�u�z���ɔ��q�C�e�_��U�g�I�^�M�A���o�@���i�o�R���d���Q�m���f�x�e��䏑ք��N���ܝ��Y�]�þo�\�H֔�M�x�a�M���G�X�Q�E�^�x܊�E�_�P�w���n��������ѝ�F�K�~���V��r̝�h�Q�����U�Ϟ�D�R��ه�{�ڔr�@�@�m��׎���[���|���E�Ƅڝ����D��I�h�x�Y���Y������[�ўr�`�zɏ�B砑z�i����Ę朑ٟ����Z�����vՏ���|炫C�R���[�C�U�g⏜R�`�X�I�s�����@���\�Ŕn�]�Ǌ䓧�t�J�R�B�]�t���u̔���T�����H���X�H�ҿ|�]�V�G�n���\���y��݆�����S�]Փ�}�_߉茻j��j�����aΛ�R�R���I���u�~�}�m�z�U�M֙؈�^�T�Q���q�]�V�T�����i���i��Ғ�d���R�瑑�}�Q�և�\���c�{�y���X���[�Hā�f����B�m�懙���帔Q���o�~ē���r���Z�W�t���I�a�P�������r���i�_�h�lؚ�O�{�u���H��䁘��VĚ�R�T�M���◉ә���L�T�U�w���t�X�Q���\�l�q���ܠ��N�����@��̃S�N�[�`�J�H�p��A�Ո�c���Fڅ�^�|��x�E����s�o׌���_�@���g�J�x�s�qܛ�J�c�����_�wِ����}�ߝ������Y���W�٠�����p���B�d�z���O�������I�B�K���}���{��Ԋ�ƕr�g���R����ҕԇ�۫F��ݔ���H���g���Q�����p�l����f�T�q�z��Z��A�b�\�K�V�C�m���q�O�p�S�s���i�H��E��؝�c�����TՄ�U���C���l�v�`�R�}�w�ϗl�N�F�d �N�~�y�^�D�T�F�j͑Ó�r�W�E�D�m�����B�f�W�f�`�����H�SȔ���^���^�l���y�����Y��΁�u�C���u���_�oʏ�ǉ]�F���`�a���u��㊑��rݠ�{�b�M�B�v�r�w�y�t��e�@�U�F�I�h�W�w��������lԔ��ʒ�N�ԇ[ϐ�f���y�{�C���a�x�\��d���n�C̓�u��S�w�m܎���x�_�k�W��ԃ���ZӖӍ�d���f������Ӡ鎟��}�����W�������V�����P��ꖰW�B�Ӭ��u���b�G�{ˎ��퓘I�~�t��U�z�x��ρˇ�|���xԄ�h�x�g���[�a��yы������t��Ξ�I��ω�f�ѓ��b�xԁ�����n�]♪q�[�Tݛ�~�O���c�Z�Z�n�R�z�u�A�S�x�Y�@�@�T�A���hs�S耎[�������y���E�\�N�j����s���d����ٝ�E�v菗��^؟��t���\ٛ����܈��l�p�S���ֱK��ݚ�䗣��`���q���~Û�w�U�H�N�@ؑᘂ��\�ꇒ걠�b�����C���̼����S���|�R�K�N�[�\�a�S�����E�i�T�D�T�����A�T�B�v���u�Dٍ���f�b�y�Ѡ��F٘���YՁ��Ɲ�Y�nۙ�C���v�u�{�M荿@��N���b���ʆ���e�Z�N倛�'; 
String.prototype.tran = function() { 
var s1,s2; 
if (zh_choose == 't') { 
   s1 = zh_s; 
   s2 = zh_t; 
}else if(zh_choose == 's') { 
   s1 = zh_t; 
   s2 = zh_s; 
}else { 
   return this; 
} 
var a = ''; 
var l = this.length; 
for(var i=0;i<this.length;i++){ 
        var c = this.charAt(i); 
        var p = s1.indexOf(c) 
        a += p < 0 ? c : s2.charAt(p); 
    } 
return a; 
} 
function setCookie(name, value) { 
var argv = setCookie.arguments; 
var argc = setCookie.arguments.length; 
var expires = (argc > 2) ? argv[2] : null; 
if (expires != null) { 
   var LargeExpDate = new Date (); 
   LargeExpDate.setTime(LargeExpDate.getTime() + (expires*1000*3600*24)); 
} 
document.cookie = name + "=" + escape (value)+((expires == null) ? "" : ("; expires=" +LargeExpDate.toGMTString())); 
} 
function getCookie(Name) { 
var search = Name + "="
if (document.cookie.length > 0) { 
   offset = document.cookie.indexOf(search); 
   if(offset != -1) { 
    offset += search.length; 
    end = document.cookie.indexOf(";", offset); 
    if(end == -1) end = document.cookie.length; 
    return unescape(document.cookie.substring(offset, end)); 
   }else { 
    return ''; 
   } 
} 
} 
  
  
function zh_tranBody(obj) {  
var o = (typeof(obj) == "object") ? obj.childNodes : document.body.childNodes; 
for (var i = 0; i < o.length; i++) { 
   var c = o.item(i); 
   if ('||BR|HR|TEXTAREA|SCRIPT|'.indexOf("|"+c.tagName+"|") > 0) continue; 
   if (c.className == zh_class) { 
    if (c.id == zh_class + '_' + zh_choose) { 
     c.setAttribute('style', zh_style_active); 
     c.style.cssText = zh_style_active; 
    }else { 
     c.setAttribute('style', zh_style_inactive); 
     c.style.cssText = zh_style_inactive; 
    } 
    continue;    
   } 
   if (c.title != '' && c.title != null) c.title = c.title.tran(); 
   if (c.alt != '' && c.alt != null) c.alt = c.alt.tran(); 
   if (c.tagName == "INPUT" && c.value != '' && c.type != 'text' && c.type != 'hidden' && c.type != 'password') c.value = c.value.tran(); 
   if (c.nodeType == 3) { 
    c.data = c.data.tran();   
   }else{ 
    zh_tranBody(c); 
   } 
} 
} 
  
function zh_tran(go) { 
if (go) zh_choose = go; 
setCookie('zh_choose', zh_choose, zh_expires); 
if (go == 'n') { 
   window.location.reload(); 
}else { 
   zh_tranBody(); 
} 
} 
  
function zh_getLang() { 
if (getCookie('zh_choose')) { 
   zh_choose = getCookie('zh_choose'); 
   return true; 
}  
if (!zh_autoLang_t && !zh_autoLang_s) return false; 
if (getCookie('zh_autoLang_checked')) return false; 
if (navigator.language) { 
   zh_browserLang = navigator.language; 
}else if (navigator.browserLanguage) { 
   zh_browserLang = navigator.browserLanguage; 
} 
if (zh_autoLang_t && zh_langReg_t.test(zh_browserLang)) { 
   zh_choose = 't'; 
}else if (zh_autoLang_s && zh_langReg_s.test(zh_browserLang)) { 
   zh_choose = 's'; 
} 
zh_autoLang_checked = 1; 
setCookie('zh_choose', zh_choose, zh_expires); 
if (zh_choose == zh_default) return false; 
return true; 
} 
  
  
function zh_init() { 
zh_getLang();  
c = document.getElementById(zh_class + '_' + zh_choose); 
if (zh_choose != zh_default) { 
   if (window.onload) { 
    window.onload_before_zh_init = window.onload; 
    window.onload = function() { 
     zh_tran(zh_choose); 
     if (getCookie('zh_autoLang_check')) {alert(zh_autoLang_msg);}; 
     window.onload_before_zh_init(); 
     }; 
   }else { 
    window.onload = function() { 
     zh_tran(zh_choose); 
     if (getCookie('zh_autoLang_check')) {alert(zh_autoLang_msg);}; 
     }; 
   } 
} 
  
} 
zh_init();