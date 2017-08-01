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
        url: appweburl + "/_api/Sp.AppContextSite(@target)/Web/Lists(guid'783A6040-E88C-4675-A49F-155A9C37B437')/items?$select=Customer%5Fx0020%5FContinent,Customer%5Fx0020%5FCountry,OrderDate,Total%5Fx0020%5FPurchase%5Fx0020%5FAmoun&@target='" + hostWebUrl + "'",
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
        url: appweburl + "/_api/Sp.AppContextSite(@target)/Web/Lists(guid'783A6040-E88C-4675-A49F-155A9C37B437')/getitems?$select=Customer%5Fx0020%5FContinent,Customer%5Fx0020%5FCountry,OrderDate,Total%5Fx0020%5FPurchase%5Fx0020%5FAmoun&@target='" + hostWebUrl + "'",
        method: "POST",
        scope: this,
        headers: {
            "accept": "application/json; odata=verbose",
            "X-RequestDigest": $("#__REQUESTDIGEST").val(),
            "content-type": "application/json; odata=verbose"
        },
        body: JSON.stringify({ "query": { "__metadata": { "type": "SP.CamlQuery" }, "ViewXml": "<View><Query><OrderBy><FieldRef Name=\"OrderID\" /></OrderBy></Query></View>" } }),
        success: PK.prototype.successHandlerGetItems,
        error: PK.prototype.errorHandlerGetItems
    });
}

PK.prototype.errorHandlerItems = function (err) {
    //error handling
}

PK.prototype.successHandlerItems = function (data) {
    //success
    this.scope.createTable(data, "Items", "itemsData");
}

PK.prototype.errorHandlerGetItems = function (err) {
    //error handling
}

PK.prototype.successHandlerGetItems = function (data) {
    //success
    this.scope.createTable(data, "GetItems", "getItemsData");
}

PK.prototype.createTable = function (data, name, divId) {
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
                td.appendChild(document.createTextNode(itm[_itm]));
                tr.appendChild(td);
            }
        })
        tbody.appendChild(tr);
    })

    var table = document.createElement('table');
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