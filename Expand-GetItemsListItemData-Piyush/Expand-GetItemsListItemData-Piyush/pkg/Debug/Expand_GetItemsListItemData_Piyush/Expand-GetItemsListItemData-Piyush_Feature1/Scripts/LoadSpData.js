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
            //url: appweburl + "/_api/Sp.AppContextSite(@target)/Web/Lists(guid'9DED7E87-7BDD-4845-A849-E6C0A67EA635')/getitems?$select=AssignedTo,DueDate,Predecessors,Blog_x0020_State&@target='" + hostWebUrl + "'",
            url: appweburl + "/_api/Sp.AppContextSite(@target)/Web/Lists(guid'9DED7E87-7BDD-4845-A849-E6C0A67EA635')/getitems?@target='" + hostWebUrl + "'",
            method: "POST",
            scope: this,
            body: "{ 'query' : {'__metadata': { 'type': 'SP.CamlQuery' }, 'ViewXml': '<View><Query/><ViewFields><FieldRef Name=\"DueDate\" /><FieldRef Name=\"AssignedTo\" /><FieldRef Name=\"Blog_x0020_State\" /><FieldRef Name=\"Predecessors\" /></ViewFields></View>' } }",
            headers: {
                "Accept": "application/json; odata=verbose",
                "content-type": "application/json; odata=verbose"
            },
            success: PK.prototype.successHandler,
            //success: successHandler,
            error: PK.prototype.errorHandler
        });
    };

    successHandler(data) {
        debugger;
        this.scope.createTable(data, 'CAML');
    };

    createTable() {
    };

    errorHandler(err) {
        //error handling
    };

    createTable(data, name) {
        var divId = 'divContainer',

            parseData = function (obj, aryProp, targetProps) {
                var displayText = '',

                    appendStringIdentifier = function (objToTest, identifier) {
                        if (objToTest && objToTest.trim().length > 0) {
                            objToTest += identifier;
                        }

                        return objToTest;
                    },

                    getResults = function () {
                        aryProp += ($.isEmptyObject(obj[aryProp]) ? 'Id' : '');
                        return obj[aryProp].results;
                    };

                $.each(getResults(), function (resIndx, resVal) {
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
            },

            dataObj = JSON.parse(data.body).d,

            tbody = document.createElement('tbody'),
            trh = document.createElement('tr'),
            //allProp = Object.keys(dataObj.results[0]);
            allProp = ['Predecessors', 'AssignedTo', 'Blog_x0020_State', 'DueDate'];

        $.each(allProp, function (_indx, _itm) {
            //if (_indx > 0) {
                var th = document.createElement('th');
                th.className = 'tHead';
                th.appendChild(document.createTextNode(PK.prototype.SPDecode(_itm))); //decode hex col name. Ex "Total_x0020_Purchase_x0020_Amoun" will de decoded to "Total Purchase Amoun" 
                trh.appendChild(th);
            //}
        })

        tbody.appendChild(trh);

        $.each(dataObj.results, function (indx, itm) {
            var tr = document.createElement('tr');
            $.each(allProp, function (_indx, _itm) {
                //if (_indx > 0) {
                var td = document.createElement('td'),
                    displayTextCell;

                td.className = 'tBody';

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
                //}
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
    SPDecode(toDecode) {
        var repl1 = new RegExp('_x', 'g'),
            repl2 = new RegExp('_', 'g');

        return unescape(toDecode.replace(repl1, "%u").replace(repl2, ""));
    };
};







//var fetchItems = function () {
//    $(document).ready(function () {
//        //PK.prototype.getitems();
//        new PK().getitems();
//    })
//};

ExecuteOrDelayUntilScriptLoaded(() => {
    //$(document).ready(function () {
    //PK.prototype.getitems();
    new PK().getitems();
    //});
}, "sp.js");
