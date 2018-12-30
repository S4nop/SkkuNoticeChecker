/*=============================================================================
 * Copyright (C) 2005 by SAMSUNG SDS co.,Ltd.
 * All right reserved.
 *
 * SYSTEM     : ACUBE/PORTAL
=============================================================================*/

var DELIM_1 = "|";
var DELIM_2 = "^";
var DELIM_3 = "`";
var isUTF = false;


// [Encode Method] - START //

function jsEncodeURI(sValue)
{
    var sEncValue = isUTF ? encodeURI(sValue) : sValue;

    return sEncValue;
}

function jsEncodeURIComponent(sValue)
{
    var sEncValue = isUTF ? encodeURIComponent(sValue) : sValue;

    return sEncValue;
}

// [Encode Method] - END   //

// [No Mouse R-Click]-START //
function DocContextMenu()
{
    return false;
}

function noContextKey()
{
    if (event.keyCode == 78)
    {
        if (event.ctrlLeft)
        {
            return false;
        }
    }
}

function noClick()
{
    return false;
}
// [No Mouse R-Click]-END //




var acube = new Object();

acube.Window = function()
{
    this.popups = new Array();
}

acube.Window.prototype =
{
    open: function(url, name, options)
    {
        this.setOptions(options);
        if (this.options.center)
        {
            this.options.left = (screen.width - this.options.width) / 2;
            this.options.top = (screen.height - this.options.height) / 2 - 40;
            //alert(this.options.top);
        }
        return this.openWin(url, name);
    },

    setOptions: function(options)
    {
        this.options =
        {
            left: 0,
            top: 0,
            width: 100,
            height: 100,
            fullscreen: 'no',
            resizable: 'no',
            scrollbars: 'yes',
            status: 'no',
            titlebar: 'no',
            toolbar: 'no',
            menubar: 'no',
            location: 'no',
            center: true,
            auto_close: true
        }
        Object.extend(this.options, options || {});
    },

    openWin: function(url, name)
    {
        var winopt =
               "left=" + this.options.left
            + ",top=" + this.options.top
            + ",width=" + this.options.width
            + ",height=" + this.options.height
            + ",fullscreen=" + this.options.fullscreen
            + ",resizable=" + this.options.resizable
            + ",scrollbars=" + this.options.scrollbars
            + ",status=" + this.options.status
            + ",titlebar=" + this.options.titlebar
            + ",toolbar=" + this.options.toolbar
            + ",menubar=" + this.options.menubar
            + ",location=" + this.options.location

        var winObj = window.open(url, name, winopt);
        winObj.focus();
        if (this.options.auto_close)
        {
            this.popups[this.popups.length] = winObj;
        }

        return winObj;
    },

    close: function()
    {
        var count = this.popups.length;
        for (var i=0; i < count; i++)
        {
            try {
                this.popups[i].close();
            }
            catch (e) {}
        }
    }
}

var acubeWindow = new acube.Window();
Event.observe(window, 'unload', _closePage, false);
function _closePage()
{
    acubeWindow.close();
}
 


/**
 * function same as Java's HashMap
 */
var HashMap = function() {
	this.initialize();
}

HashMap.prototype = {
	initialize: function() {
		this.entrySize = 0;
        this.keyEntry = new Array();
        this.valueEntry = new Array();
	},

	indexOf: function(arr, obj) {
		for(i=0; i < arr.length; i++) {
			if(arr[i]==obj) {
				return i;
			}
		}
		return -1;
	},
	
    clear: function() {
    	var stack = new Array();
    	for(var i=0; i < this.keyEntry.length; i++) {
    		stack.push(this.keyEntry[i]);
    	}
    	while(stack.length > 0) {
    		this.remove(stack.pop());
    	}
        this.entrySize = 0;
    },

    put: function(key, value) {
    	var idx = this.indexOf(this.keyEntry,key);
    	if(idx > -1) {
	        this.valueEntry[idx] = value;
    	}
    	else {
	        this.keyEntry[this.keyEntry.length] = key;
	        this.valueEntry[this.valueEntry.length] = value;
	        this.entrySize++;
    	}
    },

    get: function(key) {
        var idx = this.indexOf(this.keyEntry,key);
        if(idx > -1 && this.valueEntry[idx]!=null) {
            return this.valueEntry[idx];
        } else {
            return null;
        }
    },

    remove: function(key) {
        var idx = this.indexOf(this.keyEntry,key);
        if(idx==-1) {
            return null;
        }
        var retValue = this.valueEntry[idx];
        
		this.keyEntry.splice(idx,1);
		this.valueEntry.splice(idx,1);		        
        this.entrySize--;
        return retValue;
    },

    size: function() {
        return this.keyEntry.length;
    },

    containsKey: function(key) {
        var idx = this.indexOf(this.keyEntry,key);
        return idx > -1;
    },

    containsValue: function(value) {
        var idx = this.indexOf(this.valueEntry,key);        
        return idx > -1;
    },
    
    toString: function() {
    	var str = "[";
    	for(var i=0; i < this.keyEntry.length; i++) {
    		var comma = ",";
    		if(i==this.keyEntry.length-1) {
    			comma = "";
    		}
    		str += this.keyEntry[i] + ":" + this.valueEntry[i] + comma; 
    	}
    	str += "]";
    	return str;
    }
}

/*
var mymap = new HashMap();
mymap.put("A","This is A");
mymap.put("B","This is B");
mymap.put("C","This is C");

alert("before :" + mymap.toString() + " (size:"+mymap.size()+")");
mymap.remove("B");
alert("after :" + mymap.toString() + " (size:"+mymap.size()+")");
mymap.clear();
alert("Clear map " + mymap.toString() + " (size:"+mymap.size()+")");
*/











