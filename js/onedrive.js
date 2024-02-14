


var PWebOneDrive = {
	
	options: {
		editor: "",
		client_id: null,
		task_url: null,
		redirect_url: null,
		spinner_url: null,
		token: false,
		debug: 0 // 0 - disabled, 1 - alerts, 2 - only console
	},
	
	i18n: {
		emergency_exit: "OneDrive emergency exit",
		folder_select_warning: "OneDrive Web picker: select valid file!",
		file_select_warning: "OneDrive Web picker: select valid file!"
	},
	
	initialized: false,
	user: null,
	token: null,
	response: {status: false, message: ""},
	autoreopen: 0,
	
	setOptions: function(options)
	{
		for (var key in options) {
			this.options[key] = options[key];
		}
	},
	
	setI18n: function(i18n)
	{
		for (var key in i18n) {
			this.i18n[key] = i18n[key];
		}
	},
	
	init: function()
	{
		if (!this.initialized) 
		{
			WL.init({
				client_id: this.options.client_id, 
				redirect_uri: this.options.redirect_url, 
				response_type: "code", 
				logging: this.options.debug > 0 
			});
			
			this.createOverlay();
			
			this.initialized = true;
		}
	},

	browseFile: function(options)
	{
		if (typeof options !== "undefined") {
			this.autoreopen = 0;
			this.setOptions(options);
		}
		this.autoreopen++;
				
		this.init();
		this.showOverlay();
		
		WL.login({ "scope": "wl.signin wl.photos wl.skydrive wl.offline_access" }).then(
		    function(response) {
				if (this.options.debug) {
			    	console.log("WL.login response:\n" + JSON.stringify(response));
			    	if (this.options.debug == 1) alert("Logged in successfully.");
			    }
			    			    
				// Browse OneDrive files
				WL.fileDialog({
					mode: 'open',
					select: 'single'
			    }).then(
			        function(response) {
			        	if (this.options.debug) {
			        		console.log("WL.fileDialog response:\n" + JSON.stringify(response.data).replace(/,/g, ",\n"));
			        	}
			        	if (response.data.files && response.data.files.length) {
			        		var file = response.data.files[0];
			        		this.getToken();
			        		if (this.addToStore(file.id)) {
		        				this.insertTag("onedrivefile", file.id);
		        			}
		        			
		        			if (this.options.debug == 1) {
			        			alert("File selected.");
			        		}
			        	}
			        	this.hideOverlay();
			        }.bind(this),
			        function(response) {
						if (this.options.debug) {
		        			console.log("browseFile fileDialog error: "+response.error.code);
		        		}
						if (response.error.code == "invalid_cookie" || response.error.code == "request_token_expired") {
			            	this.resetCookie();
			            	if (this.autoreopen < 5) this.browseFile();
			            }
			            else {
			            	if (response.error.message && this.options.debug == 1) {
					    		alert(response.error.message);
					    	}
					    	if (response.error.code != "request_failed") {
					    		this.hideOverlay();
					    	}
			            }
			        }.bind(this)
			    );
		
		    }.bind(this),
		    function(response) {
				if (this.options.debug) {
		        	console.log("browseFile login error: "+response.error);
				}
				if (response.error == "invalid_cookie" || response.error == "access_denied") {
	            	this.resetCookie();
	            	if (this.autoreopen < 5) this.browseFile();
	            }
	            else {
	            	if (response.error_description) {
			        	alert(response.error_description);
			        }
			        this.hideOverlay();
				}
		    }.bind(this)
		);
	},
	
	browseGallery: function(options)
	{
		if (typeof options !== "undefined") {
			this.autoreopen = 0;
			this.setOptions(options);
		}
		this.autoreopen++;
		
		this.init();
		this.showOverlay();
		
		var that = this;
		
		WL.login({ "scope": "wl.signin wl.photos wl.skydrive wl.offline_access" }).then(
		    function(response) {
		    	if (that.options.debug) {
			    	console.log("WL.login response:\n" + JSON.stringify(response));
			    	if (that.options.debug == 1) alert("Logged in successfully.");
			    }
			    			    
			    //TODO message on first open: select folder with photos, not a file
			    
				// Browse OneDrive files
				WL.fileDialog({
					mode: 'open',
					select: 'multi'
			    }).then(
			        function(response) {			        	
			        	if (that.options.debug) {
			        		console.log("WL.fileDialog response:\n" + JSON.stringify(response.data).replace(/,/g, ",\n"));
			        	}
			        	if (response.data.folders && response.data.folders.length) {
			        		var folders = response.data.folders;
			        		var selected = 0; var inserted = 0;
			        		that.getToken();
			        		for (var i = 0; i < folders.length; i++) {
			        			if (folders[i].type == "album") {
			        				selected++;
			        				if (that.addToStore(folders[i].id)) {
				        				that.insertTag("onedrivegallery", folders[i].id);
				        				inserted++;
				        			}
			        			}
			        		}
			        		if (selected == 0) {
			        			// folder without photos selected instead of album
			        			alert(that.i18n.folder_select_warning);
			        			// reopen OneDrive browser
			        			that.browseGallery();
			        			return;
			        		}
			        		else if (inserted == 0) {
			        			// reopen OneDrive browser
			        			that.browseGallery();
			        			return;
			        		}
			        		else if (that.options.debug == 1) {
			        			alert("Folder selected.");
			        		}
			        	}
			        	else if (response.data.files && response.data.files.length) {
			        		// file selected instead of folder
			        		alert(that.i18n.file_select_warning);
			        		// reopen OneDrive browser
			        		that.browseGallery();
			        		return;
			        	}
			        	
			        	that.hideOverlay();
			        },
			        function(response) {
			            if (that.options.debug) {
		        			console.log("browseGallery fileDialog error: "+response.error.code);
		        		}
						if (response.error.code == "invalid_cookie" || response.error.code == "request_token_expired") {
			            	that.resetCookie();
			            	if (that.autoreopen < 5) that.browseGallery();
			            }
			            else {
			            	if (response.error.message && that.options.debug == 1) {
					    		alert(response.error.message);
					    	}
					    	if (response.error.code != "request_failed") {
					    		that.hideOverlay();
					    	}
			            }
			        }
			    );
		
		    },
		    function(response) {
		        if (that.options.debug) {
		        	console.log("browseGallery login error: "+response.error);
		        }
				if (response.error == "invalid_cookie" || response.error == "access_denied") {
	            	that.resetCookie();
	            	if (that.autoreopen < 5) that.browseGallery();
	            }
	            else {
	            	if (response.error_description) {
			        	alert(response.error_description);
			        }
			        that.hideOverlay();
				}
		    }
		);
	},
	
	resetCookie: function()
	{
    	var exdate = new Date();
    	document.cookie = "wl_auth=;domain="+document.location.host+";path=/;expires="+exdate.toUTCString()-1;
    },
	
	addToStore: function(id)
	{
		if (this.options.debug) {
		    console.log("addToStore "+id);
		}
		
		var that = this,
			requestData = {resource_id: id};
		if (this.token) requestData[this.token] = 1;
		
		// Send request
		jQuery.ajax({
			url: this.options.task_url + "store",
			type: "post",
			async: false,
			data: requestData,
			dataType: "json"
		}).done(function(response) {
			if (that.options.debug) {
		        console.log("addToStore response:\n" + JSON.stringify(response).replace(/,/g, ",\n"));
		    }
			that.response = response;
		});
		
		// Display error message
		if (!this.response.status && this.response.message) 
			alert(this.response.message);
		
		return this.response.status;
	},
	
	getToken: function()
	{
		if (!this.options.token) return;
		
		var that = this;
		
		jQuery.ajax({
			url: this.options.task_url + "token",
			type: "post",
			async: false,
			data: {ajax: true},
			dataType: "text"
		}).done(function(response) {
			if (that.options.debug) {
		        console.log("getToken response: " + response);
		    }
			if (response) that.token = response;
		});
		
		return this.token;
	},
	
	createOverlay: function()
	{
		if (!jQuery('#onedrive-overlay').length) 
		{
			var that = this;
			
			this.overlayLayer = jQuery('<div id="onedrive-overlay" title="'+this.i18n.emergency_exit+'"/>').css({
				display: "none",
				position: "fixed",
				top: 0,
				left: 0,
				zIndex: 65555,
				backgroundColor: "#fff",
				backgroundPosition: "50% 50%",
				backgroundRepeat: "no-repeat",
				backgroundImage: "url("+this.options.spinner_url+")",
				opacity: 0.6,
				cursor: "pointer",
				width: jQuery(window).width(),
				height: jQuery(window).height()
			}).click(function(e){
				e.stopPropagation();
				that.hideOverlay();
			}).appendTo('body');
			
			this.closeButton = jQuery('<div id="onedrive-close">'+this.i18n.emergency_exit+'</div>').css({
				display: "none",
				position: "fixed",
				top: 10,
				right: 10,
				zIndex: 2600001,
				backgroundColor: "#094AB2",
				color: "#fff",
				padding: "10px 20px",
				cursor: "pointer"
			}).click(function(e){
				e.stopPropagation();
				that.hideOverlay();
				
				// Hide OneDrive file picker
				jQuery('iframe[id^=picker]').hide().next().hide();
			}).appendTo('body');
		}
	},
	
	showOverlay: function()
	{
		this.overlayLayer.show();
		this.closeButton.show();
		
		// Show OneDrive file picker if was hidden
		jQuery('iframe[id^=picker]').show().next().show();
	},
	
	hideOverlay: function()
	{
		this.overlayLayer.hide();
		this.closeButton.hide();
	},
	
	insertTag: function(tag, file)
	{
		// abstract method to insert shortcode into editor
	}
};