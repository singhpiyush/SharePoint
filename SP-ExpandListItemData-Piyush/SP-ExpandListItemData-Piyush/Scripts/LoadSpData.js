var PK = function () {
}

PK.prototype.getQueryStringParameter = function (paramToRetrieve) {
    var params = document.URL.split("?")[1].split("&");
    for (var i = 0; i < params.length; i = i + 1) {
        var singleParam = params[i].split("=");
        if (singleParam[0] == paramToRetrieve)
            return singleParam[1];
    }
};

PK.prototype.items = function () {
    var hostWebUrl = decodeURIComponent(PK.prototype.getQueryStringParameter('SPHostUrl'));
    var appweburl = decodeURIComponent(PK.prototype.getQueryStringParameter('SPAppWebUrl'));
    var executor = new SP.RequestExecutor(appweburl);

    executor.executeAsync({
        url: appweburl + "/_api/Sp.AppContextSite(@target)/Web/Lists(guid'9DED7E87-7BDD-4845-A849-E6C0A67EA635')/items?$select=AssignedTo/ID,AssignedTo/FirstName,AssignedTo/LastName,DueDate,Predecessors/Title,Blog_x0020_State/Term&$expand=AssignedTo,Predecessors&@target='" + hostWebUrl + "'",
        method: "GET",
        scope: this,
        headers: { "Accept": "application/json; odata=verbose" },
        success: PK.prototype.successHandlerItems,
        error: PK.prototype.errorHandlerItems
    });
}

PK.prototype.getitems = function () {
    var hostWebUrl = decodeURIComponent(PK.prototype.getQueryStringParameter('SPHostUrl'));
    var appweburl = decodeURIComponent(PK.prototype.getQueryStringParameter('SPAppWebUrl'));
    var executor = new SP.RequestExecutor(appweburl);

    executor.executeAsync({
        url: appweburl + "/_api/Sp.AppContextSite(@target)/Web/Lists(guid'9DED7E87-7BDD-4845-A849-E6C0A67EA635')/getitems?$select=AssignedTo,DueDate,Predecessors,Blog_x0020_State&@target='" + hostWebUrl + "'",
        method: "POST",
        scope: this,
        headers: {
            "accept": "application/json; odata=verbose",
            "X-RequestDigest": $("#__REQUESTDIGEST").val(),
            "content-type": "application/json; odata=verbose"
        },
        body: JSON.stringify({ "query": { "__metadata": { "type": "SP.CamlQuery" }, "ViewXml": "<View><Query/><ViewFields><FieldRef Name=\"DueDate\" /><FieldRef Name=\"AssignedTo\" /><FieldRef Name=\"Blog_x0020_State\" /><FieldRef Name=\"Predecessors\" /></ViewFields></View>" } }),
        success: PK.prototype.successHandlerGetItems,
        error: PK.prototype.errorHandlerGetItems
    });
}

PK.prototype.errorHandlerItems = function (err) {
    //error handling
}

PK.prototype.successHandlerItems = function (data) {
    //success
    this.scope.createTable(data, "Using $expand", "itemsData");
}

PK.prototype.errorHandlerGetItems = function (err) {
    //error handling
}

PK.prototype.successHandlerGetItems = function (data) {
    //success
    this.scope.createTable(data, "Using CAML", "getItemsData");
}

PK.prototype.createTable = function (data, name, divId) {
    var parseData = function (obj, aryProp, targetProps) {
        var displayText = '';

        var appendStringIdentifier = function (objToTest, identifier) {
            if (objToTest && objToTest.trim().length > 0) {
                objToTest += identifier;
            }

            return objToTest;
        }

        $.each(obj[aryProp].results, function (resIndx, resVal) {
            displayText = appendStringIdentifier(displayText, ', ');

            if (Array.isArray(targetProps)) {
                var totalVal = '';
                $.each(targetProps, function (divIdIndx, divIdVal) {
                    totalVal = appendStringIdentifier(totalVal, ' ');
                    totalVal += resVal[divIdVal];
                })

                displayText += totalVal;
            }
            else {
                displayText += resVal[targetProps];
            }
        });

        return displayText;
    }

    var dataObj = JSON.parse(data.body).d;

    var tbody = document.createElement('tbody');
    var trh = document.createElement('tr');
    var allProp = Object.keys(dataObj.results[0]);
    $.each(allProp, function (_indx, _itm) {
        if (_indx > 0) {
            var th = document.createElement('th');
            th.className = 'tHead';
            th.appendChild(document.createTextNode(PK.prototype.SPDecode(_itm))); //decode hex col name. Ex "Total_x0020_Purchase_x0020_Amoun" will de decoded to "Total Purchase Amoun" 
            trh.appendChild(th);
        }
    })

    tbody.appendChild(trh);

    $.each(dataObj.results, function (indx, itm) {
        var tr = document.createElement('tr');
        $.each(allProp, function (_indx, _itm) {
            if (_indx > 0) {
                var td = document.createElement('td');
                td.className = 'tBody';
                var displayTextCell;
                switch (_itm) {
                    case 'Predecessors':
                        displayTextCell = parseData(itm, _itm, 'Title');
                        break;
                    case 'AssignedTo':
                        displayTextCell = parseData(itm, _itm, ['FirstName', 'LastName']);
                        break;
                    case 'Blog_x0020_State':
                        displayTextCell = parseData(itm, _itm, 'Label');
                        break;
                    default:
                        displayTextCell = new Date(itm[_itm]).toDateString();
                        break;
                }

                td.appendChild(document.createTextNode(displayTextCell));
                tr.appendChild(td);
            }
        })
        tbody.appendChild(tr);
    })

    var table = document.createElement('table');
    table.className = 'tbl';
    table.appendChild(tbody);

    var _h1 = document.createElement('h1');
    _h1.innerText = name;

    var divHd = document.createElement('div');
    divHd.appendChild(_h1);

    document.getElementById(divId).appendChild(divHd);
    document.getElementById(divId).appendChild(table);
}

//decode hex col name. Ex "Total_x0020_Purchase_x0020_Amoun" will de decoded to "Total Purchase Amoun" 
PK.prototype.SPDecode = function (toDecode) {
    var repl1 = new RegExp('_x', 'g');
    var repl2 = new RegExp('_', 'g');

    return unescape(toDecode.replace(repl1, "%u").replace(repl2, ""));
}

var fetchItems = function () {
    $(document).ready(function () {
        PK.prototype.items();
        PK.prototype.getitems();
    })
}

ExecuteOrDelayUntilScriptLoaded(fetchItems, "sp.js");