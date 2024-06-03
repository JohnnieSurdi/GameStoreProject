using Sitecore.Shell.Framework.Commands;
using Sitecore.Web.UI.Sheer;
using Sitecore.Data.Items;
using Sitecore.SecurityModel;
using GemBox.Spreadsheet;
using System.Data;
using Sitecore.Data;
using Sitecore.Resources.Media;
using System.IO;

namespace GameStoreProject.Foundation.Configuration.Helpers
{
    public class FileProcessor
    {
        public void ExcelFileProcessor(Item excelItem)
        {
            var media = MediaManager.GetMedia(excelItem);
            using (var stream = media.GetStream().Stream)
            {
                if (stream != null)
                {
                    ProcessExcelFile(stream);
                }
                else
                {
                    SheerResponse.Alert("Error (stream)");
                }
            }
        }

        private void ProcessExcelFile(Stream excelFileStream)
        {
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
            var workbook = ExcelFile.Load(excelFileStream);
            var worksheet = workbook.Worksheets[0];

            DataTable dataTable = worksheet.CreateDataTable(new CreateDataTableOptions()
            {
                ColumnHeaders = true,
                StartRow = 0,
                NumberOfColumns = 5,
                NumberOfRows = worksheet.Rows.Count - 1,
                Resolution = ColumnTypeResolution.AutoPreferStringCurrentCulture
            });

            foreach (DataRow row in dataTable.Rows)
            {
                string productId = row[0]?.ToString();
                string name = row[1]?.ToString();
                string description = row[2]?.ToString();
                decimal price = System.Convert.ToDecimal(row[3]);

                CreateItemFromExcelData(productId, name, description, price);
            }
        }

        private void CreateItemFromExcelData(string productId, string name, string description, decimal price)
        {
            Sitecore.Data.Database master = Sitecore.Configuration.Factory.GetDatabase("master");
            Item home = master.GetItem("/sitecore/content/GlobalDatasource");
            TemplateItem templateItem = master.Templates["{779C9389-B8F4-4E79-9686-D2456F8F6006}"];

            if (templateItem != null)
            {
                Item parentItem = home;

                using (new SecurityDisabler())
                {
                    if (parentItem.Children[name] != null)
                    {
                        return;
                    }
                    Item newItem = parentItem.Add(name, new BranchItem(templateItem));

                    if (newItem != null)
                    {
                        newItem.Editing.BeginEdit();
                        try
                        {
                            newItem["ProductId"] = productId;
                            newItem["Name"] = name;
                            newItem["Description"] = description;
                            newItem["Price"] = price.ToString();
                        }
                        finally
                        {
                            newItem.Editing.EndEdit();
                        }
                    }
                    else
                    {
                        SheerResponse.Alert("Failed to create item under parentItem.");
                    }
                }
            }
            else
            {
                SheerResponse.Alert("Template item not found or is not a template.");
            }
        }
    }
}
