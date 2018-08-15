
var elemScript = document.createElement("SCRIPT");
elemScript.src = "//cdnjs.cloudflare.com/ajax/libs/jquery/1.12.4/jquery.min.js";
document.head.appendChild(elemScript);
jQuery.noConflict();

var stormanAudit = {
    dtNow: new Date(),
    dtNowFormatString: "yyyy-MM-dd hh:mm:ss",
    recordsListName: "storageAudits",
    logObjectSizeLowerThreshold: 1000000,
    failedCreateRecordsList: false,
    failedCapturingRecordItem: false,
    bCaptureChildren: false,
    recordsList: {},
    createdItem: {},
    createdItems: [],
    arrEntries: [],
    arrFailedCapture: [],
    arrUnknown: [],
    currParent: "",
    storManTable: null,
    arrPromises: [],
    arrCreateItemPromises: [],
    arrSubentries: [],
    fieldTypes: {
        /*https://msdn.microsoft.com/en-us/library/microsoft.sharepoint.spfieldtype.aspx*/
        Invalid: 0, Integer: 1, Text: 2, Note: 3, DateTime: 4, Counter: 5, Choice: 6, Lookup: 7, Boolean: 8,
        Number: 9, Currency: 10, URL: 11, Computed: 12, Threading: 13, Guid: 14, MultiChoice: 15, GridChoice: 16,
        Calculated: 17, File: 18, Attachments: 19, User: 20, Recurrence: 21, CrossProjectLink: 22, ModStat: 23,
        Error: 24, ContentTypeId: 25, PageSeparator: 26, ThreadIndex: 27, WorkflowStatus: 28, AllDayEvent: 29,
        WorkflowEventType: 30, Geolocation: null, OutcomeChoice: null, MaxItems: 31
    },
    loopPendingSubEntries: function(){
        try{console.log("function stormanAudit.loopPendingSubEntries fired");}catch(err){}
        if ( stormanAudit.arrSubentries.length > 0 ){
            SP.UI.Notify.addNotification("<p style='color: blue;'>Looping through |"+ stormanAudit.arrSubentries.length +"| pending drill-down pages</p>",false);
            for ( var iPSE = 0; iPSE < stormanAudit.arrSubentries.length; iPSE++ ){
                stormanAudit.getSubEntries(stormanAudit.arrSubentries[iPSE].url, stormanAudit.arrSubentries[iPSE].stormanLinkParent);
                stormanAudit.arrSubentries.splice(iPSE, 1);
            }
            stormanAudit.captureResultsThenGetSubEntries();
        }
        else {
            stormanAudit.captureResultsThenGetSubEntries();
        }
    },
    getSubEntries: function(url, stormanLinkParent){
        try{console.log("function stormanAudit.getSubEntries called on |"+url+"| for stormanLinkParent |"+ stormanLinkParent +"|");}catch(err){}
        var newRoot = stormanAudit.GetUrlKeyValue("root",false,url);
        var promise = jQuery.get(url).done(function(d,s,x){
            jQuery(d).find("#onetidUserRptrTable > TBODY > TR").each(function(i,elm){
                // skip header row
                if ( i > 0 ){
                    var bSkip = false;
                    if ( jQuery(this).children().length < 1 ) {
                        // skip blank rows
                        bSkip = true;
                    }
                    else if ( jQuery(this).children().length < 8 ) {
                        var newObject = {
                            icon: jQuery(this).children("TD").eq(0).children("IMG").eq(0).prop("alt"),
                            name: jQuery(this).children("TD").eq(1).find("A").length > 0 ? jQuery(this).children("TD").eq(1).find("A").text() : jQuery(this).children("TD").eq(1).text(),
                            stormanLink: jQuery(this).children("TD").eq(1).find("A").length > 0 ? jQuery(this).children("TD").eq(1).find("A").prop("href") : newRoot + jQuery(this).children("TD").eq(1).text(),    
                            link: _spPageContextInfo.siteAbsoluteUrl.trim() +"/"+ newRoot +"/"+ jQuery(this).children("TD").eq(1).text().trim(),
                            size: jQuery(this).children("TD").eq(2).text(),
                            percentOfParent: jQuery(this).children("TD").eq(3).text().replace(" %",""),
                            visualization: null,
                            percentOfParentQuota: null,
                            visualizationOfPercentOfParentQuota: null,
                            //lastModified: new Date(jQuery(this).children("TD").eq(5).text().trim()),
                            lastModified: jQuery(this).children("TD").eq(5).text(),
                            parent: newRoot,
                            stormanLinkParent: stormanLinkParent,
                            children: []
                        };
                    }
                    else if ( jQuery(this).children().length > 6 ) {
                        var newObject = {
                            icon: jQuery(this).children("TD").eq(0).children("IMG").eq(0).prop("alt"),
                            name: jQuery(this).children("TD").eq(1).find("A").length > 0 ? jQuery(this).children("TD").eq(1).find("A").text() : jQuery(this).children("TD").eq(1).text(),
                            stormanLink: jQuery(this).children("TD").eq(1).find("A").length > 0 ? jQuery(this).children("TD").eq(1).find("A").prop("href") : newRoot + jQuery(this).children("TD").eq(1).text(),    
                            link: _spPageContextInfo.siteAbsoluteUrl.trim() +"/"+ newRoot +"/"+ jQuery(this).children("TD").eq(1).text().trim(),
                            size: jQuery(this).children("TD").eq(2).text(),
                            percentOfParent: jQuery(this).children("TD").eq(3).text().replace(" %",""),
                            visualization: null,
                            percentOfParentQuota: jQuery(this).children("TD").eq(5).text().replace(" %",""),
                            visualizationOfPercentOfParentQuota: null,
                            //lastModified: new Date(jQuery(this).children("TD").eq(7).text().trim()),
                            lastModified: jQuery(this).children("TD").eq(7).text(),
                            parent: newRoot,
                            stormanLinkParent: stormanLinkParent,
                            children: []
                        };
                    }
                    if ( bSkip === false ){
                        try{newObject.name = newObject.name.replace(/[\n\t]/igm,"");}catch(er){}
                        try{newObject.stormanLink = newObject.stormanLink.replace(/[\n\t]/igm,"");}catch(er){}
                        try{newObject.percentOfParent = newObject.percentOfParent.replace(/[\n\t]/igm,"");}catch(er){}
                        try{
                            var nSize = newObject.size.trim();
                            if ( nSize.indexOf("GB") >= 0 ){
                                nSize = parseFloat(nSize.replace(" GB","").replace("<",""));
                                nSize = nSize * 1000 * 1000 * 1000;
                            }
                            else if ( nSize.indexOf("MB") >= 0 ){
                                nSize = parseFloat(nSize.replace(" MB","").replace("<",""));
                                nSize = nSize * 1000 * 1000;
                            }
                            else if ( nSize.indexOf("KB") >= 0 ){
                                nSize = parseFloat(nSize.replace(" KB","").replace("<",""));
                                nSize = nSize * 1000;
                            }
                            newObject.size = nSize;
                        }
                        catch(err){
                            try{console.log("Failed to convert this to a number of bytes |"+ newObject.size +"|");}catch(e2){}
                            try{console.log(err);}catch(e2){}
                        }
                        var dtModified = new Date(newObject.lastModified.trim());
                        if ( dtModified.toString() === "Invalid Date" ){
                            try{console.log("Failed to convert this to a date |"+ newObject.lastModified.trim() +"|");}catch(e2){}
                        }
                        else {
                            newObject.lastModified = dtModified;
                        }
                        var newLength = stormanAudit.arrEntries.push(newObject);
                        //"http://expertsoverlunch.com/_layouts/15/storman.aspx?root=SiteAssets"
                        //if ( stormanAudit.bCaptureChildren === true ){
                        //    if ( stormanAudit.arrEntries.filter(function(object){return object.stormanLink === stormanLinkParent}).length === 1 ) {
                        //        stormanAudit.arrEntries.filter(function(object){return object.stormanLink === stormanLinkParent})[0].children.push(newLength-1);
                        //    }
                        //}
                        
        
                        //if ( jQuery(this).children("TD").eq(1).find("A").length > 0 ){
                        //    stormanAudit.getSubEntries(stormanAudit.arrEntries[newLength-1].stormanLink, url);
                        //}

                        // only get sub-entries when the parent meets our threshold
                        if ( jQuery(this).children("TD").eq(1).find("A").length > 0 ){
                            if ( newObject.size >= stormanAudit.logObjectSizeLowerThreshold ){
                                //getSubEntries(jQuery(this).children("TD").eq(1).find("A").prop("href"), "arrEntries["+newLength+"]");
                                //stormanAudit.getSubEntries(stormanAudit.arrEntries[newLength-1].stormanLink, stormanAudit.arrEntries[newLength-1].stormanLink);
                                stormanAudit.arrSubentries.push({url: newObject.stormanLink, stormanLinkParent: newObject.stormanLink});
                            }
                            else {
                                //try{console.log("SKIPPING sub-entries");}catch(er){}
                            }
                        }
                        // only get next page when on last data row and it is over our threshold
                        if ( i === jQuery("#onetidUserRptrTable > TBODY > TR").length-2 ) {
                            if (  newObject.size >= stormanAudit.logObjectSizeLowerThreshold ) {
                                if ( jQuery("A IMG[src$='images/next.gif']").length > 0 ) {
                                    try{console.log("getting next page");}catch(er){}
                                    //stormanAudit.getSubEntries(jQuery("A IMG[src$='images/next.gif']").parent().prop("href"), "");
                                    stormanAudit.arrSubentries.push({url: jQuery("A IMG[src$='images/next.gif']").parent().prop("href"), stormanLinkParent: ""});
                                }
                            }
                            else {
                                try{console.log("SKIPPING next page (last item on this page is too small)");}catch(er){}
                            }
                        }
                    }
                }
            });
        });
        stormanAudit.arrPromises.push(promise);
    },
    GetUrlKeyValue: function(c, h, a, g) {
        var e = "";
        if (a == null)
            a = ajaxNavigate.get_href() + "";
        var b;
        b = a.indexOf("#");
        if (b >= 0)
            a = a.substr(0, b);
        var d;
        if (g) {
            c = c.toLowerCase();
            d = a.toLowerCase()
        } else
            d = a;
        b = d.indexOf("&" + c + "=");
        if (b == -1)
            b = d.indexOf("?" + c + "=");
        if (b != -1) {
            var f = a.indexOf("&", b + 1);
            if (f == -1)
                f = a.length;
            e = a.substring(b + c.length + 2, f)
        }
        return h ? e : unescapeProperlyInternal(e)
    },
    addFieldToRecordsList: function(nFieldType, strFieldName, afterAddingFieldFx){
        //stormanAudit.addFieldToRecordsList(stormanAudit.fieldTypes["Text"], "MyTextField");
        var oField = {
            __metadata: { type: 'SP.Field' }, 
            Title: strFieldName, 
            FieldTypeKind: nFieldType, 
            Required: false, 
            EnforceUniqueValues: false, 
            StaticName: strFieldName.replace(/ +/ig,"")
        };
        jQuery.ajax({
            url: _spPageContextInfo.siteAbsoluteUrl +"/_api/web/lists(guid'"+ stormanAudit.recordsList.Id +"')/Fields",
            method: "POST",
            contentType: "application/json;odata=verbose",
            headers: {
                "accept": "application/json;odata=verbose",
                "content-type": "application/json;odata=verbose",
                "X-RequestDigest": jQuery("#__REQUESTDIGEST").val()
            },
            async: false,
            data: JSON.stringify(oField)
        }).done(function(data2,textStatus2,jqXHR2){
            try{console.log("Added field |"+ strFieldName +"| to records list");}catch(er){}
            //SP.UI.Notify.addNotification("<p style='color: green;'>Successfully added field "+ strFieldName +" to records list</p>",false);
            if ( typeof(afterAddingFieldFx) === "function" ){
                afterAddingFieldFx();
            }
        }).fail(function(jqXHR2,textStatus2,errorThrown2){
            SP.UI.Notify.addNotification("<p style='color: red;'>Failed to add field "+ strFieldName +" to records list</p><p>"+ errorThrown2 +"</p>",true);
        });
    },
    createRecordsList: function (afterFx){
        var body = {
            __metadata: { type: 'SP.List' }, 
            BaseTemplate: 100, 
            Description: "stores the results of stormanAudit for your site", 
            Title: stormanAudit.recordsListName
        };
        if ( typeof(afterFx) === "undefined" ) {
            var afterFx = function(){
                try{console.log("Created and configured the records list");}catch(err){}
            };
        }
        jQuery.ajax({
            url: _spPageContextInfo.siteAbsoluteUrl +"/_api/web/lists",
            method: "POST",
            contentType: "application/json;odata=verbose",
            headers: {
                "accept": "application/json;odata=verbose",
                "content-type": "application/json;odata=verbose",
                "X-RequestDigest": jQuery("#__REQUESTDIGEST").val()
            },
            async: false,
            data: JSON.stringify(body)
        }).done(function(data,textStatus,jqXHR){
            //SP.UI.Notify.addNotification("<p style='color: green;'>Successfully created records list</p>",false);
            /*https://docs.microsoft.com/en-us/previous-versions/office/developer/sharepoint-2010/ff411664(v%3doffice.14)*/
            updateStatus(stormanAudit.createRecordsListStatusID,"<p>RecordsList created... adding required fields...</p>!");
            stormanAudit.recordsList = data.d;
            if ( typeof(afterFx) === 'function' ){
                afterFx();
            }
        }).fail(function(jqXHR, textStatus, errorThrown){
            SP.UI.Notify.addNotification("<p style='color: red;'>Failed to create the records list!</p><p>"+ errorThrown +"</p>",true);
            stormanAudit.failedCreateRecordsList = true;
        });
    },
    getRecordsList: function(afterFx){
        /* only try to get the recordsList's definition if we don't already have it */
        if ( typeof(stormanAudit.recordsList.Id) === 'undefined' ){
            jQuery.ajax({
                url: _spPageContextInfo.siteAbsoluteUrl +"/_api/web/lists/GetByTitle('"+ stormanAudit.recordsListName +"')",
                method: "GET",
                contentType: "application/json;odata=verbose",
                headers: {
                    "accept": "application/json;odata=verbose",
                    "content-type": "application/json;odata=verbose",
                    "X-RequestDigest": jQuery("#__REQUESTDIGEST").val()
                }
            }).done(function(data,textStatus,jqXHR){
                //SP.UI.Notify.addNotification("<p style='color: green;'>Successfully found the records list</p>",false);
                stormanAudit.recordsList = data.d;
                if ( typeof(afterFx) === 'function' ){
                    afterFx();
                }
            }).fail(function(jqXHR, textStatus, errorThrown){
                //SP.UI.Notify.addNotification("<p style='color: blue;'>Creating records list |"+ stormanAudit.recordsListName +"|...</p>",false);
                stormanAudit.createRecordsListStatusID = addStatus("stormanAudit","<p>Creating the records list...</p>!");
                setStatusPriColor(stormanAudit.createRecordsListStatusID,"blue");
                if ( stormanAudit.failedCreateRecordsList === false ){
                    //stormanAudit.createRecordsList(afterFx);
                    stormanAudit.createRecordsList(function(){
                        stormanAudit.addFieldToRecordsList(stormanAudit.fieldTypes["Text"], "icon", function(){
                            stormanAudit.addFieldToRecordsList(stormanAudit.fieldTypes["Text"], "link", function(){
                                stormanAudit.addFieldToRecordsList(stormanAudit.fieldTypes["Text"], "name", function(){
                                    stormanAudit.addFieldToRecordsList(stormanAudit.fieldTypes["Text"], "parent", function(){
                                        stormanAudit.addFieldToRecordsList(stormanAudit.fieldTypes["Text"], "percentOfParent", function(){
                                            stormanAudit.addFieldToRecordsList(stormanAudit.fieldTypes["DateTime"], "lastModified", function(){
                                                stormanAudit.addFieldToRecordsList(stormanAudit.fieldTypes["Number"], "size", function(){
                                                    stormanAudit.addFieldToRecordsList(stormanAudit.fieldTypes["Text"], "stormanLink", function(){
                                                        stormanAudit.addFieldToRecordsList(stormanAudit.fieldTypes["Text"], "stormanLinkParent", function(){
                                                            stormanAudit.addFieldToRecordsList(stormanAudit.fieldTypes["Text"], "percentOfParentQuota", function(){
                                                                stormanAudit.addFieldToRecordsList(stormanAudit.fieldTypes["Text"], "scraped", function(){
                                                                    updateStatus(stormanAudit.createRecordsListStatusID,"<p>RecordsList created with required fields...</p>!");
                                                                    setStatusPriColor(stormanAudit.createRecordsListStatusID,"green");
                                                                    setTimeout(function(){
                                                                        /*https://docs.microsoft.com/en-us/previous-versions/office/developer/sharepoint-2010/ff409620(v%3doffice.14)*/
                                                                        removeStatus(stormanAudit.createRecordsListStatusID)
                                                                    },10000);
                                                                });
                                                            });
                                                        });
                                                    });
                                                });
                                            });
                                        });
                                    });
                                });
                            });
                        });
                    });
                }
            });
        }
        else {
            if ( typeof(afterFx) === 'function' ){
                afterFx();
            }
        }
    },
    captureFilteredResults: function(afterFx){
        SP.UI.Notify.addNotification("<p style='color: blue;'>Capturing |"+ stormanAudit.arrEntries.length +"| site objects as list items</p>",false);
        stormanAudit.getRecordsList(function(){
            //var dtNow = new Date();
            //stormanAudit.dtNow.format("yyyy-MM-dd hh:mm:ss");
            //var filteredResults = stormanAudit.arrEntries.filter(function(object){return object.size > stormanAudit.logObjectSizeLowerThreshold});
            var filteredResults = stormanAudit.arrEntries;
            for ( var iFR = 0; iFR < filteredResults.length; iFR++ ){
                // only log site objects at least 1 Mb in size
                //try{console.log(filteredResults[iFR]);}catch(err){}
                /*
                var item = {
                    __metadata: { type: stormanAudit.recordsList.ListItemEntityTypeFullName }, 
                    Title: "Audit |"+ dtNow.format("yyyy-MM-dd hh:mm:ss") +"|",
                    icon: filteredResults[iFR].icon,
                    name: filteredResults[iFR].name.replace(/[\n\t]/igm,""),
                    link: encodeURI(filteredResults[iFR].link.replace(/[\n\t]/igm,"")),
                    stormanLink: filteredResults[iFR].stormanLink.replace(/[\n\t]/igm,""),
                    size: filteredResults[iFR].size,
                    lastModified: filteredResults[iFR].lastModified.toISOString(),
                    parent: filteredResults[iFR].parent.replace(/[\n\t]/igm,""),
                    percentOfParent: filteredResults[iFR].percentOfParent.replace(/[\n\t]/igm,""),
                    stormanLinkParent: filteredResults[iFR].stormanLinkParent.replace(/[\n\t]/igm,"")
                };
                */
                var item = {};
                try{item.__metadata = { type: stormanAudit.recordsList.ListItemEntityTypeFullName };}catch(er){try{console.log("Failed to save item property value |__metadata| to |{type:'"+ stormanAudit.recordsList.ListItemEntityTypeFullName +"'}|");}catch(err2){}}
                try{item.Title =                "Audit |"+ stormanAudit.dtNow +"|";}catch(er){try{console.log("Failed to save item property value |Title| to '"+ "Audit |"+ stormanAudit.dtNow +"|" +"'");}catch(err2){}}
                try{item.icon =                 filteredResults[iFR].icon;}catch(er){try{console.log("Failed to save item property value |icon| to |"+ filteredResults[iFR].icon +"|");}catch(err2){}}
                try{item.name =                 filteredResults[iFR].name.replace(/[\n\t]/igm,"");}catch(er){try{console.log("Failed to save item property value |name| to |"+ filteredResults[iFR].name +"|");}catch(err2){}}
                try{item.link =                 encodeURI(filteredResults[iFR].link.replace(/[\n\t]/igm,""));}catch(er){try{console.log("Failed to save item property value |link| to |"+ filteredResults[iFR].link +"|");}catch(err2){}}
                try{item.stormanLink =          filteredResults[iFR].stormanLink.replace(/[\n\t]/igm,"");}catch(er){try{console.log("Failed to save item property value |stormanLink| to |"+ filteredResults[iFR].stormanLink +"|");}catch(err2){}}
                try{item.size =                 filteredResults[iFR].size;}catch(er){try{console.log("Failed to save item property value |size| to |"+ filteredResults[iFR].size +"|");}catch(err2){}}
                try{item.lastModified =         filteredResults[iFR].lastModified.toISOString();}catch(er){try{console.log("Failed to save item property value |lastModified| to |"+ filteredResults[iFR].lastModified +"|");}catch(err2){}}
                try{item.parent =               filteredResults[iFR].parent.replace(/[\n\t]/igm,"");}catch(er){try{console.log("Failed to save item property value |parent| to |"+ filteredResults[iFR].parent +"|");}catch(err2){}}
                try{item.percentOfParent =      filteredResults[iFR].percentOfParent.replace(/[\n\t]/igm,"");}catch(er){try{console.log("Failed to save item property value |percentOfParent| to |"+ filteredResults[iFR].percentOfParent +"|");}catch(err2){}}
                try{item.stormanLinkParent =    filteredResults[iFR].stormanLinkParent.replace(/[\n\t]/igm,"");}catch(er){try{console.log("Failed to save item property value |stormanLinkParent| to |"+ filteredResults[iFR].stormanLinkParent +"|");}catch(err2){}}
                try{item.percentOfParentQuota = filteredResults[iFR].percentOfParentQuota.replace(/[\n\t]/igm,"");}catch(er){/*try{console.log("Failed to save item property value |percentOfParentQuota| to |"+ filteredResults[iFR].percentOfParentQuota +"|");}catch(err2){}*/}
                stormanAudit.captureRecord(filteredResults[iFR], item);
                /*
                if ( stormanAudit.failedCapturingRecordItem === false ) {
                //if ( bDetectFailed === false && item.size >= stormanAudit.logObjectSizeLowerThreshold ){
                    jQuery.ajax({
                        url: _spPageContextInfo.siteAbsoluteUrl +"/_api/web/lists(guid'"+ stormanAudit.recordsList.Id +"')/Items",
                        data: JSON.stringify(item),
                        contentType: "application/json;odata=verbose",
                        headers: {
                            "IF-MATCH": "*",
                            "accept": "application/json;odata=verbose",
                            "content-type": "application/json;odata=verbose",
                            "X-RequestDigest": jQuery("#__REQUESTDIGEST").val()
                        },
                        dataType: "json",
                        method: "POST",
                        async: false
                    }).done(function(data,textStatus,jqXHR){
                        //stormanAudit.createdItems.push(data.d);
                        //try{console.log("removing arrEntries element at index |"+ iFR +"|");}catch(err){}
                        //stormanAudit.arrEntries.splice(iFR,1);
                        //SP.UI.Notify.addNotification("<p style='color: green;'>Saved results for |"+ item.link +"| to storageAudits list</p>",false);
                    }).fail(function(jqXHR, textStatus, errorThrown){
                        stormanAudit.arrFailedCapture.push(filteredResults[iFR]);
                        SP.UI.Notify.addNotification("<p style='color: red;'>Failed to save item capturing results for |"+ item.link +"|<p>"+ errorThrown +"</p>",true);
                        stormanAudit.failedCapturingRecordItem = true;
                    });
                    if ( stormanAudit.failedCapturingRecordItem === false ){
                        stormanAudit.arrEntries.splice(iFR,1);
                    }
                }
                */
            }
            //SP.UI.Notify.addNotification("<p style='color: green;'>Finished REST calls to capture detailed results for each site object</p>",false);
            if ( stormanAudit.failedCapturingRecordItem === false ){
                if ( typeof(afterFx) === 'function' ){
                    stormanAudit.waitForAllItemsCreated(afterFx);
                    //afterFx();
                }
            }
        });
    },
    captureRecord: function(scrapedElement, spItemData){
        if ( stormanAudit.failedCapturingRecordItem === false ){
            var promise = jQuery.ajax({
                url: _spPageContextInfo.siteAbsoluteUrl +"/_api/web/lists(guid'"+ stormanAudit.recordsList.Id +"')/Items",
                data: JSON.stringify(spItemData),
                contentType: "application/json;odata=verbose",
                headers: {
                    "IF-MATCH": "*",
                    "accept": "application/json;odata=verbose",
                    "content-type": "application/json;odata=verbose",
                    "X-RequestDigest": jQuery("#__REQUESTDIGEST").val()
                },
                dataType: "json",
                method: "POST",
                async: true
            }).done(function(data,textStatus,jqXHR){
                //stormanAudit.createdItems.push(data.d);
                //try{console.log("removing arrEntries element at index |"+ iFR +"|");}catch(err){}
                //stormanAudit.arrEntries.splice(iFR,1);
                if ( stormanAudit.arrEntries.indexOf(scrapedElement) >= 0 ){
                    stormanAudit.arrEntries.splice(stormanAudit.arrEntries.indexOf(scrapedElement),1);
                }
                else {
                    stormanAudit.arrUnknown.push(scrapedElement);
                }
                //SP.UI.Notify.addNotification("<p style='color: green;'>Saved results for |"+ item.link +"| to storageAudits list</p>",false);
            }).fail(function(jqXHR, textStatus, errorThrown){
                stormanAudit.arrFailedCapture.push(scrapedElement);
                SP.UI.Notify.addNotification("<p style='color: red;'>Failed to save item capturing results for |"+ spItemData.link +"|</p><p>"+ errorThrown +"</p>",true);
                stormanAudit.failedCapturingRecordItem = true;
            });
            stormanAudit.arrCreateItemPromises.push(promise);
        }
    },
    waitForAllData: function(afterFx){
        var timeoutCounter = 0;
        var wfIntvl = setInterval(function(){
            if ( timeoutCounter < 1000 ){
                if ( stormanAudit.arrPromises.length > 0 ){
                    SP.UI.Notify.addNotification("<p style='color: blue;'>Waiting for |"+ stormanAudit.arrPromises.length +"| pending scrape data promises</p>",false);
                    for ( var iP = 0; iP < stormanAudit.arrPromises.length; iP++ ) {
                        if ( stormanAudit.arrPromises[iP].readyState < 4 ){
                            timeoutCounter++;
                            // not done
                            break;
                        }
                        else {
                            // this one is done --- remove it
                            stormanAudit.arrPromises.splice(iP,1);
                        }
                    }
                    if ( stormanAudit.arrPromises.length < 1 ){
                        if ( typeof(afterFx) === 'function' ){
                            afterFx();
                        }
                        clearInterval(wfIntvl);
                    }
                }
                else {
                    if ( typeof(afterFx) === 'function' ){
                        afterFx();
                    }
                    clearInterval(wfIntvl);
                }
            }
            else {
                SP.UI.Notify.addNotification("Tired of waiting",true);
                clearInterval(wfIntvl);
            }
        },3210);
    },
    waitForAllItemsCreated: function(afterFx){
        var timeoutCounter = 0;
        var wfIntvl = setInterval(function(){
            if ( timeoutCounter < 1000 ){
                if ( stormanAudit.arrCreateItemPromises.length > 0 ){
                    SP.UI.Notify.addNotification("<p style='color: blue;'>Waiting for |"+ stormanAudit.arrCreateItemPromises.length +"| item creation promises</p>",false);
                    for ( var iP = 0; iP < stormanAudit.arrCreateItemPromises.length; iP++ ) {
                        if ( stormanAudit.arrCreateItemPromises[iP].readyState < 4 ){
                            timeoutCounter++;
                            // not done
                            break;
                        }
                        else {
                            // this one is done --- remove it
                            stormanAudit.arrCreateItemPromises.splice(iP,1);
                        }
                    }
                    if ( stormanAudit.arrCreateItemPromises.length < 1 ){
                        if ( typeof(afterFx) === 'function' ){
                            afterFx();
                        }
                        clearInterval(wfIntvl);
                    }
                }
                else {
                    if ( typeof(afterFx) === 'function' ){
                        afterFx();
                    }
                    clearInterval(wfIntvl);
                }
            }
            else {
                SP.UI.Notify.addNotification("Tired of waiting",true);
                clearInterval(wfIntvl);
            }
        },2109);
    },
    captureResultsThenGetSubEntries: function(afterFx){
        if ( stormanAudit.arrPromises.length > 0 ){
            stormanAudit.waitForAllData(function(){
                stormanAudit.captureFilteredResults(function(){
                    stormanAudit.loopPendingSubEntries();
                    if ( typeof(afterFx) === 'function' ){
                        afterFx();
                    }
                });
            });
        }
        else {
            if ( stormanAudit.arrEntries.length > 0 ) {
                stormanAudit.captureFilteredResults(function(){
                    stormanAudit.loopPendingSubEntries();
                    if ( typeof(afterFx) === 'function' ){
                        afterFx();
                    }
                });
            }
            else {
                //SP.UI.Notify.addNotification("Done!"+ stormanAudit.arrEntries.length,true);
                /*https://docs.microsoft.com/en-us/previous-versions/office/developer/sharepoint-2010/ff410028(v%3doffice.14)*/
                var statusID = addStatus("stormanAudit done!","<H1>Finished scraping and capuring data</H1>!",true);
                /*https://docs.microsoft.com/en-us/previous-versions/office/developer/sharepoint-2010/ff408240(v%3doffice.14)*/
                setStatusPriColor(statusID,"green");
            }
        }
    },
    init: setTimeout(function(){
        stormanAudit.getRecordsList();
        stormanAudit.dtNow = stormanAudit.dtNow.format(stormanAudit.dtNowFormatString);
        stormanAudit.failedCapturingRecordItem = false;
    },123),
    init2: setTimeout(function(){
        stormanAudit.currParent = GetUrlKeyValue("root");
        jQuery("#onetidUserRptrTable > TBODY > TR").each(function(i,elm){
            // skip header row
            if ( i > 0 ){
                var bSkip = false;
                if ( jQuery(this).children().length < 1 ) {
                    // skip blank rows
                    bSkip = true;
                }
                else if ( jQuery(this).children().length < 8 ) {
                    var newObject = {
                        icon: jQuery(this).children("TD").eq(0).children("IMG").eq(0).prop("alt"),
                        name: jQuery(this).children("TD").eq(1).find("A").length > 0 ? jQuery(this).children("TD").eq(1).find("A").text() : jQuery(this).children("TD").eq(1).text(),
                        stormanLink: jQuery(this).children("TD").eq(1).find("A").length > 0 ? jQuery(this).children("TD").eq(1).find("A").prop("href") : jQuery(this).children("TD").eq(1).text(),
                        link: _spPageContextInfo.siteAbsoluteUrl.trim() +"/"+ jQuery(this).children("TD").eq(1).text().trim(),
                        size: jQuery(this).children("TD").eq(2).text(),
                        percentOfParent: jQuery(this).children("TD").eq(3).text().replace(" %",""),
                        visualization: null,
                        percentOfParentQuota: null,
                        visualizationOfPercentOfParentQuota: null,
                        //lastModified: new Date(jQuery(this).children("TD").eq(5).text().trim()),
                        lastModified: jQuery(this).children("TD").eq(5).text(),
                        parent: stormanAudit.currParent,
                        stormanLinkParent: "",
                        children: []
                    };
                }
                else if ( jQuery(this).children().length > 6 ) {
                    var newObject = {
                        icon: jQuery(this).children("TD").eq(0).children("IMG").eq(0).prop("alt"),
                        name: jQuery(this).children("TD").eq(1).find("A").length > 0 ? jQuery(this).children("TD").eq(1).find("A").text() : jQuery(this).children("TD").eq(1).text(),
                        stormanLink: jQuery(this).children("TD").eq(1).find("A").length > 0 ? jQuery(this).children("TD").eq(1).find("A").prop("href") : jQuery(this).children("TD").eq(1).text(), 
                        link: _spPageContextInfo.siteAbsoluteUrl.trim() +"/"+ jQuery(this).children("TD").eq(1).text().trim(),
                        size: jQuery(this).children("TD").eq(2).text(),
                        percentOfParent: jQuery(this).children("TD").eq(3).text().replace(" %",""),
                        visualization: null,
                        percentOfParentQuota: jQuery(this).children("TD").eq(5).text().replace(" %",""),
                        visualizationOfPercentOfParentQuota: null,
                        //lastModified: new Date(jQuery(this).children("TD").eq(7).text().trim()),
                        lastModified: jQuery(this).children("TD").eq(7).text(),
                        parent: stormanAudit.currParent,
                        stormanLinkParent: "",
                        children: []
                    };
                }
                if ( bSkip === false ){
                    try{newObject.name = newObject.name.replace(/[\n\t]/igm,"");}catch(er){}
                    try{newObject.stormanLink = newObject.stormanLink.replace(/[\n\t]/igm,"");}catch(er){}
                    try{newObject.percentOfParent = newObject.percentOfParent.replace(/[\n\t]/igm,"");}catch(er){}
                    try{
                        var nSize = newObject.size.trim();
                        if ( nSize.indexOf("GB") >= 0 ){
                            nSize = parseFloat(nSize.replace(" GB","").replace("<",""));
                            nSize = nSize * 1000 * 1000 * 1000;
                        }
                        else if ( nSize.indexOf("MB") >= 0 ){
                            nSize = parseFloat(nSize.replace(" MB","").replace("<",""));
                            nSize = nSize * 1000 * 1000;
                        }
                        else if ( nSize.indexOf("KB") >= 0 ){
                            nSize = parseFloat(nSize.replace(" KB","").replace("<",""));
                            nSize = nSize * 1000;
                        }
                        newObject.size = nSize;
                    }
                    catch(err){
                        try{console.log("Failed to convert this to a number of bytes |"+ newObject.size +"|");}catch(e2){}
                        try{console.log(err);}catch(e2){}
                    }
                    var dtModified = new Date(newObject.lastModified.trim());
                    if ( dtModified.toString() === "Invalid Date" ){
                        try{console.log("Failed to convert this to a date |"+ newObject.lastModified +"|");}catch(e2){}
                    }
                    else {
                        newObject.lastModified = dtModified;
                    }
                    var newLength = stormanAudit.arrEntries.push(newObject);
                    // only get sub-entries when the parent meets our threshold
                    if ( jQuery(this).children("TD").eq(1).find("A").length > 0 ){
                        if ( newObject.size >= stormanAudit.logObjectSizeLowerThreshold ){
                            //getSubEntries(jQuery(this).children("TD").eq(1).find("A").prop("href"), "arrEntries["+newLength+"]");
                            //stormanAudit.getSubEntries(stormanAudit.arrEntries[newLength-1].stormanLink, stormanAudit.arrEntries[newLength-1].stormanLink);
                            stormanAudit.arrSubentries.push({url: newObject.stormanLink, stormanLinkParent: newObject.stormanLink})
                        }
                        else {
                            //try{console.log("SKIPPING sub-entries");}catch(er){}
                        }
                    }
                    // only get next page when on last data row and it is over our threshold
                    if ( i === jQuery("#onetidUserRptrTable > TBODY > TR").length-2 ) {
                        if (  newObject.size >= stormanAudit.logObjectSizeLowerThreshold ) {
                            if ( jQuery("A IMG[src$='images/next.gif']").length > 0 ) {
                                try{console.log("getting next page");}catch(er){}
                                //stormanAudit.getSubEntries(jQuery("A IMG[src$='images/next.gif']").parent().prop("href"), "");
                                stormanAudit.arrSubentries.push({url: jQuery("A IMG[src$='images/next.gif']").parent().prop("href"), stormanLinkParent: ""});
                            }
                        }
                        else {
                            try{console.log("SKIPPING next page (last item on this page is too small)");}catch(er){}
                        }
                    }
                }
            }
        });
        stormanAudit.captureResultsThenGetSubEntries();
    },321)
}
/*
// find objects over 1 MB in total size
stormanAudit.arrEntries.filter(function(object){return object.size > 1000000})
*/
