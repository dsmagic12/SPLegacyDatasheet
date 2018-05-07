var fxOnDatasheetReadyStateChange = function(event){
    try{console.log(event);}catch(err){}
};
var spLDS = {
    arrXHR: [],
    arrH: [],
    bRunOnLoad: true,
    bReplaceQuickEditView: false,
    bBrowserNoSupport: false,
    bEditMode: false,
    instances: [],
    new: function(listGUID, viewGUID, wpSeqID, wpWrapperDivID, wpWrapperDivWebPartID, listDisplayName){
        return {
            listGUID: listGUID,
            listDisplayName: listDisplayName,
            viewGUID: viewGUID, 
            listWeb: _spPageContextInfo.webAbsoluteUrl,
            bAbort: false,
            bReady: false,
            instanceObjectElement: null,
            fxOnReadyStateChange: function(){
                var iBreak = 0;
                var spLDSIndex = this.instanceIndex;
                var dsControl = spLDS.instances[spLDSIndex].instanceObjectElement;
                var intvl = setInterval(function(){
                    iBreak++;
                    if ( iBreak > 1000 ){
                        try{console.log("Tired of waiting for our datasheet to be ready...");}catch(err){}
                        clearInterval(intvl);
                    }
                    else {
                        try{
                            dsControl.DisplayTaskPane = true;
                            dsControl.DisplayTaskPane = false;
                            try{console.log("Datasheet at instance appears to be ready");}catch(err){}
                            spLDS.instances[spLDSIndex].bReady = true;
                            spLDS.instances[spLDSIndex].fxAfterReady(spLDSIndex);
                            try{console.log("Executed this instance's fxAfterReady function");}catch(err){}
                            clearInterval(intvl);
                        }
                        catch(errWaiting){
                            try{console.log(errWaiting);}catch(err){}
                        }
                    }
                }, 500);
            },
            fxAfterReady: function(spLDSIndex){
                try{console.log("Detected that our datasheet at WP sequence |"+ spLDSIndex +"| is ready!");}catch(err){}
                if ( typeof(spLDS.instances[spLDSIndex+1]) === "object" ) {
                    try{console.log("We appear to have another QuickEdit view to replace with a legacy datasheet... kicking off that process!");}catch(err){}
                    spLDS.instances[spLDSIndex+1].replaceQuickEditView();
                }
            },
            replaceWebPart: {
                wpSequenceID: wpSeqID,
                wrapperDivID: wpWrapperDivID,
                wrapperDivWebPartID: wpWrapperDivWebPartID,
                emptyWrapper: function(){
                    var origWebPart = document.getElementById("WebPartWPQ"+this.wpSequenceID);
                    var iBreak = 0;
                    while ( origWebPart.children.length > 0 && iBreak < 20 ){
                        iBreak++;
                        origWebPart.removeChild(origWebPart.children[0])
                    }
                }
            },
            viewSchema: null, 
            listSchema: null,
            listData: null, 
            arrH: [],
            arrXHR: [],
            contentArea: {
                width:0,
                height:0,
                top:0,
                left:0
            },
            instanceIndex: 0,
            getContentArea: function(){
                var origWebPart = document.getElementById("WebPartWPQ"+this.replaceWebPart.wpSequenceID);
                this.contentArea.top = origWebPart.offsetTop;
                this.contentArea.left = origWebPart.offsetLeft;
                this.contentArea.width = origWebPart.scrollWidth
                this.contentArea.height = origWebPart.offsetHeight;
            },
            replaceQuickEditView: function(){
                this.getContentArea();
                if ( !this.listGUID === true ){
                    this.getListGuid(this.listDisplayName, function(listGUID, spLDSIndex){
                        spLDS.instances[spLDSIndex].getAllComponents(spLDSIndex);
                        spLDS.instances[spLDSIndex].replaceWebPart.emptyWrapper();
                    },this.instanceIndex);
                }
                else {
                    this.getAllComponents(this.instanceIndex);
                    this.replaceWebPart.emptyWrapper();
                }
            },
            generateAndAppendDatasheet: function(spLDSIndex){
                /*CLSID for the ActiveX control from https://msdn.microsoft.com/en-us/library/ms416795(v=office.14).aspx*/
                spLDS.instances[spLDSIndex].arrH = ['<object name="STSListControlWPQ'];
                spLDS.instances[spLDSIndex].arrH.push(spLDS.instances[spLDSIndex].replaceWebPart.wpSequenceID)
                spLDS.instances[spLDSIndex].arrH.push('" width="');
                spLDS.instances[spLDSIndex].arrH.push(spLDS.instances[spLDSIndex].contentArea.width);
                spLDS.instances[spLDSIndex].arrH.push('" height="');
                spLDS.instances[spLDSIndex].arrH.push(spLDS.instances[spLDSIndex].contentArea.height);
                spLDS.instances[spLDSIndex].arrH.push('" tabIndex="1" class="ms-dlgDisable" id="STSListControlWPQ');
                spLDS.instances[spLDSIndex].arrH.push(spLDS.instances[spLDSIndex].replaceWebPart.wpSequenceID)
                spLDS.instances[spLDSIndex].arrH.push('" classid="CLSID:65BCBEE4-7728-41A0-97BE-14E1CAE36AAE" onreadystatechange="fxOnDatasheetReadyStateChange()">');
                spLDS.instances[spLDSIndex].arrH.push('<param name="ListName" value="{')
                spLDS.instances[spLDSIndex].arrH.push(spLDS.instances[spLDSIndex].listGUID);
                spLDS.instances[spLDSIndex].arrH.push('}"><param name="ViewGuid" value="{');
                spLDS.instances[spLDSIndex].arrH.push(spLDS.instances[spLDSIndex].viewGUID);
                spLDS.instances[spLDSIndex].arrH.push('}"><param name="ListWeb" value="')
                spLDS.instances[spLDSIndex].arrH.push(spLDS.instances[spLDSIndex].listWeb);
                spLDS.instances[spLDSIndex].arrH.push('/_vti_bin"><param name="ListData" value="');
                spLDS.instances[spLDSIndex].arrH.push(spLDS.instances[spLDSIndex].listData);
                spLDS.instances[spLDSIndex].arrH.push('"><param name="ViewSchema" value="');
                spLDS.instances[spLDSIndex].arrH.push(spLDS.instances[spLDSIndex].viewSchema);
                spLDS.instances[spLDSIndex].arrH.push('"><param name="ListSchema" value="');
                spLDS.instances[spLDSIndex].arrH.push(spLDS.instances[spLDSIndex].listSchema);
                spLDS.instances[spLDSIndex].arrH.push('"><param name="ControlName" value="STSListControlWPQ');
                spLDS.instances[spLDSIndex].arrH.push(spLDS.instances[spLDSIndex].replaceWebPart.wpSequenceID)
                spLDS.instances[spLDSIndex].arrH.push('"><p class="ms-descriptiontext">Ooops your browser\'s ActiveX controls don\'t work<br/><a href="https://www.microsoft.com/en-us/download/details.aspx?id=13255">You may need to install the MS Access database engine from MS Office 2010 (https://www.microsoft.com/en-us/download/details.aspx?id=13255)</a></p></object>');
                var origWebPart = document.getElementById("WebPartWPQ"+spLDS.instances[spLDSIndex].replaceWebPart.wpSequenceID);
                if ( spLDS.instances[spLDSIndex].bAbort === false && spLDS.bBrowserNoSupport === false ){
                    origWebPart.innerHTML = spLDS.instances[spLDSIndex].arrH.join("");
                    spLDS.instances[spLDSIndex].instanceObjectElement = document.getElementById("STSListControlWPQ"+spLDS.instances[spLDSIndex].replaceWebPart.wpSequenceID);
                    spLDS.instances[spLDSIndex].fxOnReadyStateChange();
                }
            },
            waitForAllComponents: function(afterFx, spLDSIndex){
                var doneWithAll = false;
                var iBreak = 0;
                var intvl = setInterval(function(){
                    var bDone = true;
                    for ( var i = 0; i < spLDS.instances[spLDSIndex].arrXHR.length; i++ ){
                        if ( spLDS.instances[spLDSIndex].arrXHR[i].readyState < 4 ) {
                            bDone = false;
                            break;
                        }
                    }
                    iBreak++;
                    if ( iBreak > 1000 ) {
                        clearInterval(intvl);
                    }
                    if ( bDone === true ) {
                        doneWithAll = true;
                        if ( typeof(afterFx) === "function" ){
                            try{afterFx(spLDSIndex);}catch(err){}
                        }
                        clearInterval(intvl);
                    }
                },55);
            },
            getAllComponents: function(spLDSIndex){
                this.getContentArea();
                this.getComponent("listData",this.listWeb + "/_api/web/lists(guid'"+this.listGUID+"')/Items", "xml", function(){}, this.instanceIndex);
                this.getComponent("listSchema",this.listWeb + "/_api/web/lists(guid'"+this.listGUID+"')", "xml", function(){}, this.instanceIndex);
                this.getComponent("viewSchema",this.listWeb + "/_api/web/lists(guid'"+this.listGUID+"')/Views(guid'"+this.viewGUID+"')?$select=ListViewXML","xml", function(){}, this.instanceIndex);
                this.waitForAllComponents(function(spLDSIndex){spLDS.instances[spLDSIndex].generateAndAppendDatasheet(spLDSIndex);}, spLDSIndex);
            },
            getComponent: function(componentName, url, dataType, fxHandleResponse, spLDSIndex){
                var xhr = new XMLHttpRequest();
                xhr.open('GET',url,true);
                xhr.setRequestHeader("X-RequestDigest",document.getElementById("__REQUESTDIGEST").value);
                xhr.onreadystatechange = function(){
                    if ( xhr.readyState === 4 ) {
                        if ( xhr.status !== 200 ){
                            try{console.log("Error retrieving |"+componentName+"| from |"+ url +"|... "+ xhr.status);}catch(err){}
                            try{console.log(xhr.response);}catch(err){}
                            spLDS.instances[spLDSIndex].bAbort = true;
                        }
                        else {
                            if ( dataType === "xml" ) {
                                var iBreak;
                                var resp = xhr.responseText;
                                do {
                                    resp = resp.replace('>','&gt;');
                                    resp = resp.replace('<','&gl;');
                                    iBreak++;
                                    if ( iBreak >= 10000 ){
                                        try{console.log("breaking markup escape loop after 10000 iterations")}catch(err){}
                                        break;
                                    }
                                } while (resp.indexOf('>') >= 0 || resp.indexOf('<') >= 0);
                                spLDS.instances[spLDSIndex][componentName] = resp;
                            }
                        }
                    }
                    
                }
                xhr.send();
                spLDS.instances[spLDSIndex].arrXHR.push(xhr);
            },
            getListGuid: function(listDisplayName, afterFx, spLDSIndex){
                var xhr = new XMLHttpRequest();
                var url = this.listWeb+"/_api/web/lists/GetByTitle('"+ listDisplayName +"')?$select=Id";
                xhr.open('GET',url,true);
                xhr.setRequestHeader("X-RequestDigest",document.getElementById("__REQUESTDIGEST").value);
                xhr.setRequestHeader("accept", "application/json;odata=verbose");
                xhr.setRequestHeader("content-type", "application/json;odata=verbose");
                xhr.onreadystatechange = function(){
                    if ( xhr.readyState === 4 ) {
                        if ( xhr.status !== 200 ){
                            try{console.log("Error retrieving listGUID from display name |"+ listDisplayName +"| from |"+ url +"|... "+ xhr.status);}catch(err){}
                            try{console.log(xhr.response);}catch(err){}
                            spLDS.instances[spLDSIndex].bAbort = true;
                        }
                        else {
                            var resp = JSON.parse(xhr.response);
                            spLDS.instances[spLDSIndex].listGUID = resp.d.Id;
                            spLDS.instances[spLDSIndex].bGotListGUID = true;
                            if ( typeof(afterFx) === "function" ){
                                afterFx(resp.d.Id, spLDSIndex);
                            }
                        }
                    }
                    
                }
                xhr.send();
            }
        };
    },
    getDatasheetViewsOnPage: function(afterFx){
        var coll = document.getElementsByClassName("ms-listviewtable");
        if ( coll.length > 0 ){
            for ( var iViewTable = 0; iViewTable < coll.length; iViewTable++ ){
                var viewTable = coll[iViewTable]; 
                if ( viewTable.className.indexOf("ms-listviewgrid") < 0 ) {
                    var listName = viewTable.summary;
                    if ( viewTable.id.indexOf("}-{") >= 0 ){
                        var listGUID = viewTable.id.split("}-{")[0].substr(1,36);
                        var viewGUID = viewTable.id.split("}-{")[1].substr(0,36);
                    }
                    else {
                        var listGUID = "";
                        var viewGUID = "";
                        for ( var iA = 0; iA < viewTable.attributes.length; iA++ ){
                            if ( viewTable.attributes[iA].name === "o:webquerysourcehref" ){
                                viewGUID = GetUrlKeyValue("View",false,viewTable.attributes[iA].value).substr(1,36);
                            }
                            if ( viewTable.attributes[iA].name === "summary" ){
                                listName = viewTable.attributes[iA].value;
                            }
                        }
                    }
                    var wrapperDivID = "";
                    var wrapperDivWebPartID = "";
                    var wpSequenceID = 0;
                    var wpListData = null;
                    var wpSchemaData = null;
                    var parent = viewTable.parentElement;
                    var iLoopBreak = 0;
                    var bFoundWrapper = false;
                    while (bFoundWrapper === false || iLoopBreak < 10 ) {
                        parent = parent.parentElement;
                        for ( var iP = 0; iP < parent.attributes.length; iP++ ){
                            if ( parent.attributes[iP].name === "webpartid" ){
                                wrapperDivWebPartID = parent.attributes[iP].value;
                                viewGUID = wrapperDivWebPartID;
                                wrapperDivID = parent.id;
                                wpSequenceID = parent.id.replace("WebPartWPQ","");
                                if ( !listName === true ){
                                    listName = document.getElementById("WebPartTitleWPQ"+wpSequenceID).innerText.trim();
                                }
                                bFoundWrapper = true;
                                break;
                            }
                        }
                        iLoopBreak++;
                        if ( iLoopBreak >= 10 ){
                            try{console.log("Breaking loop for wrapper div of ms-listviewtable");}catch(err){}
                        }
                    }
                    if ( bFoundWrapper === true ){
                        spLDS.instances.push(spLDS.new(listGUID, viewGUID, wpSequenceID, wrapperDivID, wrapperDivWebPartID, listName));
                        spLDS.instances[spLDS.instances.length-1].instanceIndex = spLDS.instances.length-1;
                        /*spLDS.instances[spLDS.instances.length-1].replaceQuickEditView();*/
                    }
                }
            }
            if ( typeof(afterFx) === "function" ){
                afterFx();
            }
        }
        else {
            try{console.log("not on a list view page... not calling afterFx")}catch(err){}
        }
    },
    checkBrowserCompatibility: function(){
        var oRet = false;
        var bActiveX = false;
        /* check browser's compatibility with the ListNet control (https://www.microsoft.com/en-us/download/details.aspx?id=13255) */
        try {
            var ActiveXobj = new ActiveXObject('ListNet.ListNet');
            bActiveX = true;
        } catch (err) {}
        if (bActiveX === true) {
            oRet = true;
        } else {
            oRet = false;
        }
        if (oRet === false) {
            document.getElementById("DeltaPlaceHolderMain").innerHTML = document.getElementById("DeltaPlaceHolderMain").innerHTML + '<p class="ms-descriptiontext">Ooops! Your browser\'s ActiveX controls don\'t work<br/><a href="https://www.microsoft.com/en-us/download/details.aspx?id=13255">You may need to install the MS Access database engine from MS Office 2010 (https://www.microsoft.com/en-us/download/details.aspx?id=13255)</a><br/>You must use <a href="https://www.microsoft.com/en-us/download/internet-explorer.aspx" target="_blank">MS Internet Explorer</a></p>';
            spLDS.bBrowserNoSupport = true;
        }
        return oRet;
    },
    init: function(){
        try{console.log("spLDS running after detecting jsgrid.js and core.js as loaded")}catch(err){}
        if ( spLDS.checkBrowserCompatibility() === true ){
            try{console.log("spLDS is looking for all QuickEdit views on the page")}catch(err){}
            spLDS.getDatasheetViewsOnPage(function(){
                /*setTimeout(function(){*/
                    try{console.log("spLDS is replacing the first QuickEdit view on the page with a legacy datasheet")}catch(err){}
                    if ( typeof(spLDS.instances[0]) === "object" ) {
                        spLDS.instances[0].replaceQuickEditView();
                    }    
                /*},567);*/
            });
        }
    },
    isPageInEditMode: function() {
        /*https://sharepoint.stackexchange.com/questions/149096/a-way-to-identify-when-page-is-in-edit-mode-for-javascript-purposes*/
        var result = (window.MSOWebPartPageFormName != undefined) && ((document.forms[window.MSOWebPartPageFormName] && document.forms[window.MSOWebPartPageFormName].MSOLayout_InDesignMode && ("1" == document.forms[window.MSOWebPartPageFormName].MSOLayout_InDesignMode.value)) || (document.forms[window.MSOWebPartPageFormName] && document.forms[window.MSOWebPartPageFormName]._wikiPageMode && ("Edit" == document.forms[window.MSOWebPartPageFormName]._wikiPageMode.value)));
        this.editMode = result || false;
        return result || false;
    },
    onLoad: setTimeout(function(){
        if ( spLDS.bRunOnLoad === true && spLDS.isPageInEditMode() === false ){
            /*
            try{console.log("spLDS waiting for sharepoint to be ready (core.js and jsgrid.js loaded)")}catch(err){}
            SP.SOD.executeFunc('core.js', null, function(){
                try{console.log("spLDS running on load after detecting core.js as loaded")}catch(err){}
                SP.SOD.executeFunc('jsgrid.js', null, function(){
                    spLDS.init();
                });
            });
            */
            try{console.log("spLDS waiting for sharepoint to be ready (core.js and jsgrid.js loaded)")}catch(err){}
            ExecuteOrDelayUntilScriptLoaded(function(){
                try{console.log("spLDS running on load after detecting core.js as loaded")}catch(err){}
                ExecuteOrDelayUntilScriptLoaded(function(){
                    try{console.log("spLDS running on load after detecting jsgrid.js as loaded")}catch(err){}
                    spLDS.init();
                },"jsgrid.js");
            },"core.js");

        }
    },23)
}