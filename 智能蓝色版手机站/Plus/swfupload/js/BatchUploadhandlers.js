/* Demo Note:  This demo uses a FileProgress class that handles the UI for displaying the file name and percent complete.
The FileProgress class is not part of SWFUpload.
*/


/* **********************
   Event Handlers
   These are my custom event handlers to make my
   web application behave the way I went when SWFUpload
   completes different tasks.  These aren't part of the SWFUpload
   package.  They are part of my application.  Without these none
   of the actions SWFUpload makes will show up in my application.
   ********************** */
function preLoad() {
	if (!this.support.loading) {
		alert("You need the Flash Player 9.028 or above to use SWFUpload.");
		return false;
	}
}
function loadFailed() {
	alert("Something went wrong while loading SWFUpload. If this were a real application we'd clean up and then give you an alternative");
}
var hasfilename='';
var hasexists=false;
var ffile=null;
var haslimit=false;
var exnum=0;
var hastr=false;
function fileQueued(file) {
	  ffile=file;
    //判断有没有重复选择
	hasexists=false;
	haslimit=false;
	hastr=false;
	if (hasfilename!=''){
		if (hasfilename.indexOf(file.name.toLowerCase()+",")!=-1){
			 hasexists=true;
			 swfu.cancelUpload(file.id); 
			 alert('文件['+file.name+']已存在上传列表中，请不要重复选择！');
			 exnum++;
			return false;
		 }

	}
    if (hasfilename==''){
		 hasfilename=file.name+',';
	}else{
		 hasfilename+=file.name+',';
	}
	
	try {
		var progress = new FileProgress(file, this.customSettings.progressTarget);
		progress.setStatus("<img src='images/sico_03.gif' align='absmiddle'> 等待上传...");
		progress.toggleCancel(true, this);

	} catch (ex) {
		this.debug(ex);
	}

}
function createtr(ffile){
	if (hasexists||haslimit) return;
	if (ffile==null) return;
    totalsize+=ffile.size;
	var ext=ffile.name.substring(ffile.name.lastIndexOf('.')+1,ffile.name.length);
	var str='<tr id="tr'+ffile.id+'"><td class="splittd" nowrap style="padding-left:5px"><img src="../editor/ksplus/FileIcon/'+ext+'.gif" align="absmiddle"/> <strong id="f'+ffile.id+'">'+ffile.name.substring(0,26)+'</strong></td><td class="splittd" style="text-align:center" width="100">'+(ffile.size/1024).toFixed(2)+'KB<span style="display:none" id="size'+ffile.id+'">'+ffile.size+'</span></td><td class="splittd" width="260" id="p'+ffile.id+'">等待上传...</td><td class="splittd" style="text-align:center" id="info'+ffile.id+'"><img style="cursor:hand" sid="'+ffile.id+'" name="cdel" border="0" src="../images/default/close.gif"/></td></tr>';
	jQuery('#t1').append(str);
	
	$("img[name=cdel]").click(function(){
			hasfilename=hasfilename.replace($("#f"+$(this).attr("sid")).html()+',', "");
			realcount--;
			totalsize-=parseFloat($("#size"+$(this).attr("sid")).html());
			UpdateBottom();
			if(realcount==1){  //没内容时只显示一个上传按钮
				$("#table1").hide();
				jQuery("#allnums").hide();
				(o.style||o).height='35px';
			}
			hastr=true;
			swfu.cancelUpload($(this).attr("sid")); 
			$("#tr"+$(this).attr("sid")).remove();
		 })
	
	
}

function fileQueueError(file, errorCode, message) {
	try {
		if (errorCode === SWFUpload.QUEUE_ERROR.QUEUE_LIMIT_EXCEEDED) {
			alert("You have attempted to queue too many files.\n" + (message === 0 ? "You have reached the upload limit." : "You may select " + (message > 1 ? "up to " + message + " files." : "one file.")));
			haslimit=true;
			return;
		}

		var progress = new FileProgress(file, this.customSettings.progressTarget);
		progress.setError();
		progress.toggleCancel(false);

		switch (errorCode) {
		case SWFUpload.QUEUE_ERROR.FILE_EXCEEDS_SIZE_LIMIT:
			progress.setStatus("File is too big.");
			alert("对不起，您选择的文件太大，超过系统限制！");
			this.debug("Error Code: File too big, File name: " + file.name + ", File size: " + file.size + ", Message: " + message);
			break;
		case SWFUpload.QUEUE_ERROR.ZERO_BYTE_FILE:
			progress.setStatus("Cannot upload Zero Byte files.");
			this.debug("Error Code: Zero byte file, File name: " + file.name + ", File size: " + file.size + ", Message: " + message);
			break;
		case SWFUpload.QUEUE_ERROR.INVALID_FILETYPE:
			progress.setStatus("Invalid File Type.");
			this.debug("Error Code: Invalid File Type, File name: " + file.name + ", File size: " + file.size + ", Message: " + message);
			break;
		default:
			if (file !== null) {
				progress.setStatus("Unhandled Error");
			}
			this.debug("Error Code: " + errorCode + ", File name: " + file.name + ", File size: " + file.size + ", Message: " + message);
			break;
		}
	} catch (ex) {
        this.debug(ex);
    }
}

function fileDialogComplete(numFilesSelected, numFilesQueued) {
	if (hasexists) return;
	try {
		if (numFilesSelected > 0) {
			if (hasexists==false){
			realcount=realcount+numFilesQueued;
			}
			UploadFileInput_OnResize();
			//document.getElementById(this.customSettings.cancelButtonId).disabled = false;
		}
		/* I want auto start the upload and I can do that here */
		//this.startUpload();
	} catch (ex)  {
        this.debug(ex);
	}
}

function uploadStart(file) {
	hastr=true;
	this.addPostParam('fileNames',escape(file.name));   
	try {
		/* I don't want to do any file validation or anything,  I'll just update the UI and
		return true to indicate that the upload should start.
		It's important to update the UI here because in Linux no uploadProgress events are called. The best
		we can do is say we are uploading.
		 */
		var progress = new FileProgress(file, this.customSettings.progressTarget);
		progress.setStatus("<img src='images/loading.gif' align='absmiddle'> 正在上传中...");
		progress.toggleCancel(true, this);
	}
	catch (ex) {}
	
	return true;
}

function uploadProgress(file, bytesLoaded, bytesTotal) {
	try {
		var percent = Math.ceil((bytesLoaded / bytesTotal) * 100);

		var progress = new FileProgress(file, this.customSettings.progressTarget);
		progress.setProgress(percent);
		progress.setStatus("<img src='images/loading.gif' align='absmiddle'> 正在上传中...");
	} catch (ex) {
		this.debug(ex);
	}
}

function uploadSuccess(file, serverData) {
	try {
		if (serverData.substring(0, 6) == "error:") {
			alert(unescape(serverData).replace("error:",""));
		 }else{
			var d=unescape(serverData).split('|');
			if (basictype!=99999){parent.InsertFileFromUp(d[0],d[1],d[2],d[3],d[4])}
		 }
		var progress = new FileProgress(file, this.customSettings.progressTarget);
		progress.setComplete();
		if (serverData.substring(0, 6) == "error:") {
		progress.setStatus("<img src='images/wrong.gif' align='absmiddle'> <span style='color:red'>文件没有上传.</span>");
		}else{
		progress.setStatus("<img src='images/accept.png' align='absmiddle'> <span style='color:green'>上传完毕.</span>");
		}
		$("#info"+file.id).html("---");
		progress.toggleCancel(false);

	} catch (ex) {
		this.debug(ex);
	}
}

function uploadError(file, errorCode, message) {
	try {
		var progress = new FileProgress(file, this.customSettings.progressTarget);
		progress.setError();
		progress.toggleCancel(false);

		switch (errorCode) {
		case SWFUpload.UPLOAD_ERROR.HTTP_ERROR:
			progress.setStatus("Upload Error: " + message);
			this.debug("Error Code: HTTP Error, File name: " + file.name + ", Message: " + message);
			break;
		case SWFUpload.UPLOAD_ERROR.UPLOAD_FAILED:
			progress.setStatus("<span style='color:red'>上传出错.</span>");
			this.debug("Error Code: Upload Failed, File name: " + file.name + ", File size: " + file.size + ", Message: " + message);
			break;
		case SWFUpload.UPLOAD_ERROR.IO_ERROR:
			progress.setStatus("Server (IO) Error");
			this.debug("Error Code: IO Error, File name: " + file.name + ", Message: " + message);
			break;
		case SWFUpload.UPLOAD_ERROR.SECURITY_ERROR:
			progress.setStatus("Security Error");
			this.debug("Error Code: Security Error, File name: " + file.name + ", Message: " + message);
			break;
		case SWFUpload.UPLOAD_ERROR.UPLOAD_LIMIT_EXCEEDED:
			progress.setStatus("Upload limit exceeded.");
			this.debug("Error Code: Upload Limit Exceeded, File name: " + file.name + ", File size: " + file.size + ", Message: " + message);
			break;
		case SWFUpload.UPLOAD_ERROR.FILE_VALIDATION_FAILED:
			progress.setStatus("Failed Validation.  Upload skipped.");
			this.debug("Error Code: File Validation Failed, File name: " + file.name + ", File size: " + file.size + ", Message: " + message);
			break;
		case SWFUpload.UPLOAD_ERROR.FILE_CANCELLED:
			// If there aren't any files left (they were all cancelled) disable the cancel button
			if (this.getStats().files_queued === 0) {
				document.getElementById(this.customSettings.cancelButtonId).disabled = true;
			}
			progress.setStatus("Cancelled");
			progress.setCancelled();
			break;
		case SWFUpload.UPLOAD_ERROR.UPLOAD_STOPPED:
			progress.setStatus("Stopped");
			break;
		default:
			progress.setStatus("Unhandled Error: " + errorCode);
			this.debug("Error Code: " + errorCode + ", File name: " + file.name + ", File size: " + file.size + ", Message: " + message);
			break;
		}
	} catch (ex) {
        this.debug(ex);
    }
}

function uploadComplete(file) {
	if (this.getStats().files_queued === 0) {
		//document.getElementById(this.customSettings.cancelButtonId).disabled = true;
	}
}

// This event comes from the Queue Plugin
function queueComplete(numFilesUploaded) {
if (basictype==99999){alert('恭喜，上传成功!');parent.location.href=parent.location.href;}
   
	//var status = document.getElementById("divStatus");
	//status.innerHTML = numFilesUploaded + " file" + (numFilesUploaded === 1 ? "" : "s") + " uploaded.";
}
