class PK {
    constructor() {
    };

    getQueryStringParameter(paramToRetrieve) {
        var params = document.URL.split("?")[1].split("&"),
            i = 0;

        for (; i < params.length; i = i + 1) {
            var singleParam = params[i].split("=");
            if (singleParam[0] == paramToRetrieve)
                return singleParam[1];
        }
    };

    getitems() {
        var hostWebUrl = decodeURIComponent(PK.prototype.getQueryStringParameter('SPHostUrl')),
            appweburl = decodeURIComponent(PK.prototype.getQueryStringParameter('SPAppWebUrl')),
            executor = new SP.RequestExecutor(appweburl);

        executor.executeAsync({
            url: appweburl + "/_api/Sp.AppContextSite(@target)/Web/Lists(guid'9DED7E87-7BDD-4845-A849-E6C0A67EA635')/getitems?$select=FieldValuesAsText/AssignedTo,FieldValuesAsText/DueDate,FieldValuesAsText/Predecessors,FieldValuesAsText/Blog_x0020_State&$expand=FieldValuesAsText&@target='" + hostWebUrl + "'",
            method: "POST",
            scope: this,
            body: "{ 'query' : {'__metadata': { 'type': 'SP.CamlQuery' }, 'ViewXml': '' } }",
            headers: {
                "Accept": "application/json; odata=verbose",
                "content-type": "application/json; odata=verbose"
            },
            success: PK.prototype.successHandler,
            error: PK.prototype.errorHandler
        });
    };

    successHandler(data) {
        this.scope.createTable(data, 'FieldValuesAsText');
    };

    createTable() {
    };

    errorHandler(err) {
        //error handling
    };

    createTable(data, name) {
        var divId = 'divContainer',
            _this = PK.prototype,

            dataObj = JSON.parse(data.body).d,

            tbody = document.createElement('tbody'),
            trh = document.createElement('tr'),

            allProp = ['Predecessors', 'AssignedTo', 'Blog_x0020_State', 'DueDate'];

        $.each(allProp, function (_indx, _itm) {
            var th = document.createElement('th');

            th.className = 'tHead';
            th.appendChild(document.createTextNode(_this.sPDecode(_itm))); //decode hex col name. Ex "Total_x0020_Purchase_x0020_Amoun" will de decoded to "Total Purchase Amoun" 
            trh.appendChild(th);
        })

        tbody.appendChild(trh);

        $.each(dataObj.results, function (indx, itm) {
            var tr = document.createElement('tr'),
                fieldtextValues = itm.FieldValuesAsText;

            $.each(allProp, function (_indx, _itm) {
                //if (_indx > 0) {
                var td = document.createElement('td'),
                    displayTextCell;

                td.className = 'tBody';
                
                displayTextCell = fieldtextValues[_this.xmlEncodeUnderScore(_itm)];

                td.appendChild(document.createTextNode(displayTextCell));
                tr.appendChild(td);
            });
            tbody.appendChild(tr);
        });

        var table = document.createElement('table'),
            _h1 = document.createElement('h1'),
            divHd = document.createElement('div');

        table.className = 'tbl';
        table.appendChild(tbody);

        _h1.innerText = name;

        divHd.appendChild(_h1);

        document.getElementById(divId).appendChild(divHd);
        document.getElementById(divId).appendChild(table);
    };

    //decode hex col name. Ex "Total_x0020_Purchase_x0020_Amoun" will de decoded to "Total Purchase Amoun" 
    sPDecode(toDecode) {
        var repl1 = new RegExp('_x', 'g'),
            repl2 = new RegExp('_', 'g');

        return unescape(toDecode.replace(repl1, "%u").replace(repl2, ""));
    };

    //replace any '_' with '_x005f-'
    xmlEncodeUnderScore(_columnName) {
        var nameParts = _columnName.split('_');
        if (nameParts.length > 1) {
            _columnName = nameParts.join('_x005f_');
        }

        return _columnName;
    };
};

ExecuteOrDelayUntilScriptLoaded(() => {
    new PK().getitems();
}, "sp.js");
