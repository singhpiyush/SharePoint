// ===========================================
// Copyright (c) 2017. All rights reserved.
// Author:: Piyush Kumar Singh
// Purpose:: SharePoint REST calls from C#
// ===========================================
// Change History
// 01/01/2017		=> First version
// ===========================================


using System.Threading.Tasks;

namespace SharePointRestCallsCSharp
{
    class ProgramPk
    {
        static void Main(string[] args)
        {
            //replace this with your credential
            ConnectPk connectToSp = new SharePointRestCallsCSharp.ConnectPk("https://domain.sharepoint.com", "piyush@domain.onmicrosoft.com", "yourPassword");
            Task[] tasks = new Task[] {
                connectToSp.GetListItems("{0}/_api/web/lists/GetByTitle('ODataList_Manual')/items?$top=1"),
                connectToSp.GetList("{0}/_api/web/lists/GetByTitle('ODataList_Manual')"),
                connectToSp.GetWeb("{0}/_api/web")
            };

            Task.WaitAll(tasks);
        }
    }
}
