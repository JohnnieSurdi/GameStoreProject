using Sitecore.Shell.Framework.Commands;
using Sitecore.Web.UI.Sheer;
using Sitecore.Data.Items;
using Sitecore.SecurityModel;
using GemBox.Spreadsheet;
using System.Data;
using Sitecore.Data;
using Sitecore.Resources.Media;
using System.IO;
using GameStoreProject.Foundation.Configuration.Helpers;

namespace GameStoreProject.Foundation.Configuration.Commands
{
    public class UploadData : Command
    {
        public override void Execute(CommandContext context)
        {
            string excelItemPath = "/sitecore/media library/Files/games";
            Database masterDatabase = Sitecore.Configuration.Factory.GetDatabase("master");
            Item excelItem = masterDatabase.GetItem(excelItemPath);

            if (excelItem != null && excelItem.Paths.Path == excelItemPath)
            {
                FileProcessor processor = new FileProcessor();
                processor.ExcelFileProcessor(excelItem);
            }
            else
            {
                SheerResponse.Alert("Error (Excel item access)");
            }
        }
    }
}

