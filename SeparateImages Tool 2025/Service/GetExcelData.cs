using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SeparateImages_Tool_2025.MVVM.Model;
using Syncfusion.XlsIO;

namespace SeparateImages_Tool_2025.Service
{
    class GetExcelData : Abstract.ViewBaseModel
    {
        public List<ImageListModel> ExtractImageNames(string excelFilePath, string imageNameColumn = "ImageName")
        {
                List<ImageListModel> imageNames = new List<ImageListModel>();
            try
            {


                using (ExcelEngine excelEngine = new ExcelEngine())
                {
                    IApplication application = excelEngine.Excel;
                    application.DefaultVersion = ExcelVersion.Xlsx;

                    IWorkbook workbook = application.Workbooks.Open(excelFilePath);
                    IWorksheet worksheet = workbook.Worksheets[0];

                    int lastRow = worksheet.UsedRange.LastRow;
                    int lastColumn = worksheet.UsedRange.LastColumn;

                    int imageNameColumnIndex = -1;

                    // Find the column index of the desired header (e.g. "ImageName")
                    for (int col = 1; col <= lastColumn; col++)
                    {
                        string header = worksheet[1, col].DisplayText.Trim();
                        if (string.Equals(header, imageNameColumn, StringComparison.OrdinalIgnoreCase))
                        {
                            imageNameColumnIndex = col;
                            break;
                        }
                    }

                    if (imageNameColumnIndex == -1)
                        throw new Exception($"Column '{imageNameColumn}' not found in Excel file.");

                    // Read the values under the image name column
                    for (int row = 2; row <= lastRow; row++)
                    {
                        string imageName = worksheet[row, imageNameColumnIndex].DisplayText?.Trim();
                        if (!string.IsNullOrEmpty(imageName))
                        {
                            imageNames.Add(new ImageListModel
                            {
                                FileName = imageName,
                                IsSelected = true
                            });
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                ErrorMessage(ex);
            }
                return imageNames;
        }
    }
}
