using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Microsoft.Office.Interop.Outlook;

namespace OutlookHassBridge
{

    public partial class ThisAddIn
    {
        private NameSpace outlookNameSpace;
        private MAPIFolder inbox;
        private Items items;
        private OutlookStatus LastStatus { get; set; } = new OutlookStatus(null);

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            outlookNameSpace = this.Application.GetNamespace("MAPI");
            inbox = outlookNameSpace.GetDefaultFolder(OlDefaultFolders.olFolderInbox);

            items = inbox.Items;
            items.ItemAdd += items_ItemAdd;
            items.ItemChange += items_ItemChanged;
            Task.Run(UpdateHass).Wait();
        }

        private async void items_ItemAdd(object item)
        {
            await UpdateHass();
        }

        private async void items_ItemChanged(object item)
        {
            await UpdateHass();
        }

        private int GetUnread()
        {
            return inbox.Items.Restrict("[Unread]=true").Count;
        }

        private bool IsUnread()
        {
            return GetUnread() > 0;
        }

        private async Task UpdateHass()
        {
            if (string.IsNullOrWhiteSpace(Properties.Settings.Default.Url))
                return;

            var status = new OutlookStatus(IsUnread());

            if (!status.Equals(LastStatus))
            {
                LastStatus = status;

                using (var client = new HttpClient())
                {
                    var json = JsonConvert.SerializeObject(status);
                    var stringContent = new StringContent(json, Encoding.UTF8, "application/json");
                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                    
                    await client.PutAsync(Properties.Settings.Default.Url, stringContent);
                }

            }
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
        }
        
        #endregion
    }
}
