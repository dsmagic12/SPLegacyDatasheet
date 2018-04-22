var spLDS = {
    arrXHR: [],
    arrH: [],
    bRunOnLoad: true,
    bReplaceQuickEditView: false,
    bBrowserNoSupport: false,
    listGUID: _spPageContextInfo.pageListId.substr(1,_spPageContextInfo.pageListId.lastIndexOf("}")-1), 
    viewGUID: "", 
    listWeb: _spPageContextInfo.siteAbsoluteUrl,
    viewSchema: null, 
    listSchema: null,
    listData: null, 
    contentArea: {
        width:0,
        height:0,
        top:0,
        left:0
    },
    getViewGUIDFromPage: function(afterFx){
        var coll = document.getElementsByClassName("ms-listviewtable");
        if ( coll.length > 0 ){
            for ( var i = 0; i < coll.length; i++ ){
                if ( coll[i].id.toLowerCase().indexOf(spLDS.listGUID) >= 0 ) {
                    spLDS.viewGUID = coll[i].id.split("}-{")[1].substr(0,coll[i].id.split("}-{")[1].lastIndexOf("}"));
                    if ( typeof(afterFx) === "function" ){
                        afterFx();
                    }
                    coll[i].style.display = "none";
                    break;
                }
            }
        }
        else {
            try{console.log("not on a list view page... not calling afterFx")}catch(err){}
        }
    },
    getContentArea: function(){
        spLDS.contentArea.top = document.getElementById("DeltaPlaceHolderMain").offsetTop;
        spLDS.contentArea.left = document.getElementById("DeltaPlaceHolderMain").offsetLeft;
        spLDS.contentArea.width = document.getElementById("s4-workspace").scrollWidth - spLDS.contentArea.left;
        spLDS.contentArea.height = document.getElementById("s4-workspace").offsetHeight;
    },
    generateAndAppendDatasheet: function(){
        /*CLSID for the ActiveX control from https://msdn.microsoft.com/en-us/library/ms416795(v=office.14).aspx*/
        spLDS.arrH = ['<object name="STSListControlWPQ2" width="'];
        spLDS.arrH.push(spLDS.contentArea.width);
        spLDS.arrH.push('" height="');
        spLDS.arrH.push(spLDS.contentArea.height);
        spLDS.arrH.push('" tabIndex="1" class="ms-dlgDisable" id="STSListControlWPQ2" classid="CLSID:65BCBEE4-7728-41A0-97BE-14E1CAE36AAE">');
        spLDS.arrH.push('<param name="ListName" value="{')
        spLDS.arrH.push(spLDS.listGUID);
        spLDS.arrH.push('}"><param name="ViewGuid" value="{');
        spLDS.arrH.push(spLDS.viewGUID);
        spLDS.arrH.push('}"><param name="ListWeb" value="')
        spLDS.arrH.push(spLDS.listWeb);
        spLDS.arrH.push('/_vti_bin"><param name="ListData" value="');
        spLDS.arrH.push(spLDS.listData);
        spLDS.arrH.push('"><param name="ViewSchema" value="');
        spLDS.arrH.push(spLDS.viewSchema);
        spLDS.arrH.push('"><param name="ListSchema" value="');
        spLDS.arrH.push(spLDS.listSchema);
        spLDS.arrH.push('"><param name="ControlName" value="STSListControlWPQ2"><p class="ms-descriptiontext">Ooops your browser\'s ActiveX controls don\'t work<br/><a href="https://www.microsoft.com/en-us/download/details.aspx?id=13255">You may need to install the MS Access database engine from MS Office 2010 (https://www.microsoft.com/en-us/download/details.aspx?id=13255)</a></p></object>');
        document.getElementById("DeltaPlaceHolderMain").innerHTML = document.getElementById("DeltaPlaceHolderMain").innerHTML + spLDS.arrH.join("");
    },
    waitForAllComponents: function(afterFx){
        var doneWithAll = false;
        var iBreak = 0;
        var intvl = setInterval(function(){
            var bDone = true;
            for ( var i = 0; i < spLDS.arrXHR.length; i++ ){
                if ( spLDS.arrXHR[i].readyState < 4 ) {
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
                    afterFx();
                }
                clearInterval(intvl);
            }
        },55);

    },
    getAllComponents: function(){
        spLDS.getContentArea();
        spLDS.getComponent("listData",_spPageContextInfo.siteAbsoluteUrl + "/_api/web/lists(guid'"+spLDS.listGUID+"')/Items","xml");
        spLDS.getComponent("listSchema",_spPageContextInfo.siteAbsoluteUrl + "/_api/web/lists(guid'"+spLDS.listGUID+"')","xml");
        spLDS.getComponent("viewSchema",_spPageContextInfo.siteAbsoluteUrl + "/_api/web/lists(guid'"+spLDS.listGUID+"')/Views(guid'"+spLDS.viewGUID+"')?$select=ListViewXML","xml");
        spLDS.waitForAllComponents(function(){
            spLDS.generateAndAppendDatasheet();
        });
    },
    getComponent: function(componentName, url, dataType, fxHandleResponse){
        var xhr = new XMLHttpRequest();
        xhr.open('GET',url,true);
        xhr.setRequestHeader("X-RequestDigest",document.getElementById("__REQUESTDIGEST").value);
        xhr.onreadystatechange = function(){
            if ( xhr.readyState === 4 ) {
                if ( dataType === "xml" ) {
                    var iBreak;
                    var resp = xhr.responseText;
                    do {
                        resp = resp.replace('>','&gt;');
                        resp = resp.replace('<','&gl;');
                        iBreak++;
                        if ( iBreak >= 10000 ){
                            try{console.log("breaking markup escape loop")}catch(err){}
                            break;
                        }
                    } while (resp.indexOf('>') >= 0 || resp.indexOf('<') >= 0);
                    spLDS[componentName] = resp;
                }
            }
            
        }
        xhr.send();
        spLDS.arrXHR.push(xhr);
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
        if ( spLDS.checkBrowserCompatibility() === true ){
            try{console.log("spLDS generating calling webservices to generate datasheet view HTML")}catch(err){}
            spLDS.getAllComponents();
        }
    },
    onLoad: setTimeout(function(){
        if ( spLDS.bRunOnLoad === true ){
            try{console.log("spLDS running on load")}catch(err){}
            if ( spLDS.bReplaceQuickEditView === true ){
                spLDS.getViewGUIDFromPage(function(){
                    try{console.log("spLDS replacing list view on page")}catch(err){}
                    spLDS.init();
                });
            }
            else {
                try{console.log("spLDS generating list view from GUIDs in settings")}catch(err){}
                spLDS.init();
            }
        }
    },23)
}