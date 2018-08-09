/*
	Queue Plug-in
	
	Features:
		*Adds a cancelQueue() method for cancelling the entire queue.
		*All queued files are uploaded when startUpload() is called.
		*If false is returned from uploadComplete then the queue upload is stopped.
		 If false is not returned (strict comparison) then the queue upload is continued.
		*Adds a QueueComplete event that is fired when all the queued files have finished uploading.
		 Set the event handler with the queue_complete_handler setting.
		
	*/

var SWFUpload;
if (typeof(SWFUpload) === "function") {
	SWFUpload.queue = {};
	
	SWFUpload.prototype.initSettings = (function (oldInitSettings) {
		return function (userSettings) {
			if (typeof(oldInitSettings) === "function") {
				oldInitSettings.call(this, userSettings);
			}
			
			this.queueSettings = {};
			
			this.queueSettings.queue_cancelled_flag = false;
			this.queueSettings.queue_upload_count = 0;
			
			this.queueSettings.user_upload_complete_handler = this.settings.upload_complete_handler;
			this.queueSettings.user_upload_start_handler = this.settings.upload_start_handler;
			this.settings.upload_complete_handler = SWFUpload.queue.uploadCompleteHandler;
			this.settings.upload_start_handler = SWFUpload.queue.uploadStartHandler;
			
			this.settings.queue_complete_handler = userSettings.queue_complete_handler || null;
		};
	})(SWFUpload.prototype.initSettings);

	SWFUpload.prototype.startUpload = function (fileID) {
		this.queueSettings.queue_cancelled_flag = false;
		this.callFlash("StartUpload", [fileID]);
	};

	SWFUpload.prototype.cancelQueue = function () {
		this.queueSettings.queue_cancelled_flag = true;
		this.stopUpload();
		
		var stats = this.getStats();
		while (stats.files_queued > 0) {
			this.cancelUpload();
			stats = this.getStats();
		}
	};
	
	SWFUpload.queue.uploadStartHandler = function (file) {
		var returnValue;
		if (typeof(this.queueSettings.user_upload_start_handler) === "function") {
			returnValue = this.queueSettings.user_upload_start_handler.call(this, file);
		}
		
		// To prevent upload a real "FALSE" value must be returned, otherwise default to a real "TRUE" value.
		returnValue = (returnValue === false) ? false : true;
		
		this.queueSettings.queue_cancelled_flag = !returnValue;

		return returnValue;
	};
	
	SWFUpload.queue.uploadCompleteHandler = function (file) {
		var user_upload_complete_handler = this.queueSettings.user_upload_complete_handler;
		var continueUpload;
		
		if (file.filestatus === SWFUpload.FILE_STATUS.COMPLETE) {
			this.queueSettings.queue_upload_count++;
		}

		if (typeof(user_upload_complete_handler) === "function") {
			continueUpload = (user_upload_complete_handler.call(this, file) === false) ? false : true;
		} else if (file.filestatus === SWFUpload.FILE_STATUS.QUEUED) {
			// If the file was stopped and re-queued don't restart the upload
			continueUpload = false;
		} else {
			continueUpload = true;
		}
		
		if (continueUpload) {
			var stats = this.getStats();
			if (stats.files_queued > 0 && this.queueSettings.queue_cancelled_flag === false) {
				this.startUpload();
			} else if (this.queueSettings.queue_cancelled_flag === false) {
				this.queueEvent("queue_complete_handler", [this.queueSettings.queue_upload_count]);
				this.queueSettings.queue_upload_count = 0;
			} else {
				this.queueSettings.queue_cancelled_flag = false;
				this.queueSettings.queue_upload_count = 0;
			}
		}
	};
}






function preLoad() {
			if (!this.support.loading) {
				alert("You need the Flash Player 9.028 or above to use SWFUpload.");
				return false;
			}
		}
		function loadFailed() {
			alert("Something went wrong while loading SWFUpload. If this were a real application we'd clean up and then give you an alternative");
		}
		
		function fileQueued(file) {
			try {
				this.customSettings.tdFilesQueued.innerHTML = this.getStats().files_queued;
			} catch (ex) {
				this.debug(ex);
			}
		
		}
function fileQueueError(file, errorCode, message) {
	try {
		var imageName = "error.gif";
		var errorName = "";
		if (errorCode === SWFUpload.errorCode_QUEUE_LIMIT_EXCEEDED) {
			errorName = "You have attempted to queue too many files.";
		}

		if (errorName !== "") {
			alert(errorName);
			return;
		}
		switch (errorCode) {
		case SWFUpload.QUEUE_ERROR.ZERO_BYTE_FILE:
			imageName = "zerobyte.gif";
			alert('出错,上传0字节!');
			return;
			break;
		case SWFUpload.QUEUE_ERROR.FILE_EXCEEDS_SIZE_LIMIT:
			imageName = "toobig.gif";
			alert('出错,太大超出限制!');
			return;
			break;
		case SWFUpload.QUEUE_ERROR.ZERO_BYTE_FILE:
		case SWFUpload.QUEUE_ERROR.INVALID_FILETYPE:
		default:
			alert(message);
			break;
		}

		//addImage("images/" + imageName);

	} catch (ex) {
		this.debug(ex);
	}

}
function fileDialogComplete(numFilesSelected, numFilesQueued) {
	      if (numFilesSelected>1){
		   alert('只能选择一个文件!');
		 }else if(numFilesQueued==1){
			 this.startUpload();
		 }
			
		}
		
		function uploadStart(file) {
			this.addPostParam('fileNames',escape(file.name));   
			$(window.parent.document).find("#mesWindow").show();
		
			 //$(parent.document.body).append("<div id='tips' class='floatymessage'>HERRO</div>"); 
			try {
				this.customSettings.progressCount = 0;
				updateDisplay.call(this, file);
			}
			catch (ex) {
				this.debug(ex);
			}
			
		}
		
		function uploadProgress(file, bytesLoaded, bytesTotal) {
			try {
				this.customSettings.progressCount++;
				updateDisplay.call(this, file);
			} catch (ex) {
				this.debug(ex);
			}
			
		}
		
		
		function uploadComplete(file) {
			try{
				//parent.closeWindow(); //关闭弹出的速度提示
				$(window.parent.document).find("#mesWindow").hide();
			}catch(ex){
			}

			this.customSettings.tdFilesQueued.innerHTML = this.getStats().files_queued;
			this.customSettings.tdFilesUploaded.innerHTML = this.getStats().successful_uploads;
			this.customSettings.tdErrors.innerHTML = this.getStats().upload_errors;
		
		}
		
		function updateDisplay(file) {
			this.customSettings.tdCurrentSpeed.innerHTML = SWFUpload.speed.formatBPS(file.currentSpeed);
			this.customSettings.tdAverageSpeed.innerHTML = SWFUpload.speed.formatBPS(file.averageSpeed);
			this.customSettings.tdMovingAverageSpeed.innerHTML = SWFUpload.speed.formatBPS(file.movingAverageSpeed);
			this.customSettings.tdTimeRemaining.innerHTML = SWFUpload.speed.formatTime(file.timeRemaining);
			this.customSettings.tdTimeElapsed.innerHTML = SWFUpload.speed.formatTime(file.timeElapsed);
			this.customSettings.tdPercentUploaded.innerHTML = SWFUpload.speed.formatPercent(file.percentUploaded);
			this.customSettings.tdSizeUploaded.innerHTML = SWFUpload.speed.formatBytes(file.sizeUploaded);
			this.customSettings.tdProgressEventCount.innerHTML = this.customSettings.progressCount;
		
		}

		var swfu,limitnum;
		window.onload = function() {
			if ($(window.parent.document).find("#UploadTips")[0]==undefined){
				var p=new parent.KesionPopup();
				p.MsgBorder=1;
				p.BgColor='#ffffee';
				p.ShowBackground=false;
				p.popup('','<div id="UploadTips" style="padding:5px"><style>#UploadTips span{color:#ff6600;}</style><div style="display:none">Files Queued:<span id="tdFilesQueued"></span>Files Uploaded:	<span id="tdFilesUploaded"></span>Errors:	<span id="tdErrors"></span></div>当前速度:<span id="tdCurrentSpeed">0</span> 平均速度：<span  id="tdAverageSpeed">0</span><br/>已上传：<span id="tdPercentUploaded">0%</span> 大小：<span id="tdSizeUploaded">0</span>&nbsp;剩余时间：<span id="tdTimeRemaining">0</span> 已用时：<span id="tdTimeElapsed">0</span><br/><strong>正在上传中... 请耐心等待!!! 直到该提示框消失。</strong><span style="display:none">Moving Average Speed:<span id="tdMovingAverageSpeed"></span>Progress Event Count:<span id="tdProgressEventCount"></span></span></div></div>',420);
				$(window.parent.document).find("#mesWindow").hide();
				$(window.parent.document).find("#closeBtn").hide();
			}
			if (limitnum==null){ limitnum=0;}

			var settings = {
				flash_url : dir+"plus/swfupload/swfupload/swfupload.swf",
				flash9_url : dir+"plus/swfupload/swfupload/swfupload_fp9.swf",
				upload_url: uploadUrl,
				post_params: post_params,
				file_size_limit : limitSize,
				file_types : fileExt,
				file_types_description : "文件类型",
				file_upload_limit : limitnum,  //限制只能上传一个文件
				file_queue_limit : 0,
				debug: false,
				// Button settings
				//button_image_url : "../plus/swfupload/images/SmallSpyGlassWithTransperancy_17x18d.png",
				button_placeholder_id : "spanButtonPlaceholder",
				button_width: 155,
				button_height: 20,
				button_text : '<span class="button">上传文件(限制'+(limitSize/1024).toFixed(0)+' MB)</span>',
				button_text_style : '.button {'+buttonstyle+'} ',
				button_text_top_padding: 0,
				button_text_left_padding: 0,
				button_window_mode: SWFUpload.WINDOW_MODE.TRANSPARENT,
				button_cursor: SWFUpload.CURSOR.HAND,				
				
				moving_average_history_size: 40,
				// The event handler functions are defined in handlers.js
				swfupload_preload_handler : preLoad,
				swfupload_load_failed_handler : loadFailed,
				file_queue_error_handler : fileQueueError,
				file_queued_handler : fileQueued,
				file_dialog_complete_handler: fileDialogComplete,
				upload_start_handler : uploadStart,
				upload_progress_handler : uploadProgress,
				upload_success_handler : uploadSuccess,
				upload_complete_handler : uploadComplete,
				
				custom_settings : {
					tdFilesQueued : parent.document.getElementById("tdFilesQueued"),
					tdFilesUploaded : parent.document.getElementById("tdFilesUploaded"),
					tdErrors : parent.document.getElementById("tdErrors"),
					tdCurrentSpeed : parent.document.getElementById("tdCurrentSpeed"),
					tdAverageSpeed : parent.document.getElementById("tdAverageSpeed"),
					tdMovingAverageSpeed : parent.document.getElementById("tdMovingAverageSpeed"),
					tdTimeRemaining : parent.document.getElementById("tdTimeRemaining"),
					tdTimeElapsed : parent.document.getElementById("tdTimeElapsed"),
					tdPercentUploaded : parent.document.getElementById("tdPercentUploaded"),
					tdSizeUploaded : parent.document.getElementById("tdSizeUploaded"),
					tdProgressEventCount : parent.document.getElementById("tdProgressEventCount")
				}
			};
			swfu = new SWFUpload(settings);
	     };