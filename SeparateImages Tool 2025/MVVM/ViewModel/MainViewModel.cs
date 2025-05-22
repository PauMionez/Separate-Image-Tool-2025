using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.Mvvm;
using SeparateImages_Tool_2025.MVVM.Model;
using SeparateImages_Tool_2025.Service;
using Syncfusion.XlsIO;

namespace SeparateImages_Tool_2025.MVVM.ViewModel
{
    class MainViewModel: Abstract.ViewBaseModel
    {
        public DelegateCommand SelectImageFileCommand { get; set; }
        public DelegateCommand SelectExcelListCommand { get; set; }
        public AsyncCommand StartProcessCommand { get; set; }

        #region properties
        private string _imageOutputFolder;

        public string ImageOutputFolder
        {
            get { return _imageOutputFolder; }
            set { _imageOutputFolder = value; RaisePropertiesChanged(nameof(ImageOutputFolder)); }
        }

        private string _imageListFolderPath;

        public string ImageListFolderPath
        {
            get { return _imageListFolderPath; }
            set { _imageListFolderPath = value; RaisePropertiesChanged(nameof(ImageListFolderPath)); }
        }


        private string _excelListFilePath;

        public string ExcelListFilePath
        {
            get { return _excelListFilePath; }
            set { _excelListFilePath = value; RaisePropertiesChanged(nameof(ExcelListFilePath)); }
        }

        private string _imageExcelHeader;

        public string ImageExcelHeader
        {
            get { return _imageExcelHeader; }
            set { _imageExcelHeader = value; RaisePropertiesChanged(nameof(ImageExcelHeader)); }
        }


        private double _progressValue;
        public double ProgressValue
        {
            get { return _progressValue; }
            set { _progressValue = value; RaisePropertiesChanged(nameof(ProgressValue)); }
        }


        private string _currentFileName;
        public string CurrentFileName
        {
            get { return _currentFileName; }
            set { _currentFileName = value; RaisePropertiesChanged(nameof(CurrentFileName)); }
        }

        #endregion

        #region fields
        private static readonly string[] EXCEL_EXTENSION = { "*.xlsx", "*.xls", "*.xlsm", "*.xltx", "*.xltm", "*.xlsb" };
        private List<ImageListModel> imageNameList;
        private readonly GetExcelData ExcelHelper;
        #endregion

        public MainViewModel()
        {
            SelectImageFileCommand = new DelegateCommand(OnSelectImageFile);
            SelectExcelListCommand = new DelegateCommand(OnSelectExcelList);
            StartProcessCommand = new AsyncCommand(StartProcess);
            imageNameList = new List<ImageListModel>();
            ExcelHelper = new GetExcelData();
        }

        private void OnSelectExcelList()
        {
            try
            {
                string filterExtensions = string.Join(";", EXCEL_EXTENSION.Select(ext => $"*.{ext}"));
                string excelListFilePath = GetFilePath("Excel File (*.xlsx)|*.xlsx", filterExtensions, "Select Old Output Excel File");

                if (!string.IsNullOrWhiteSpace(excelListFilePath)) { ExcelListFilePath = excelListFilePath; }

                imageNameList = ExcelHelper.ExtractImageNames(ExcelListFilePath, ImageExcelHeader);
            }
            catch (Exception ex)
            {
                ErrorMessage(ex);
            }
        }

        private void OnSelectImageFile()
        {
            try
            {
                string selectedImageFolder = GetFolderPath("Select Input Image Folder");
                if (string.IsNullOrEmpty(selectedImageFolder)){ return; }

                ImageListFolderPath = selectedImageFolder;
                ImageOutputFolder = Path.Combine(selectedImageFolder, $"{Path.GetFileName(selectedImageFolder)}_Output");
            }
            catch (Exception ex)
            {
                ErrorMessage(ex);
            }
        }


        private async Task StartProcess()
        {
            try
            {
                if (!Directory.Exists(ImageOutputFolder))
                    Directory.CreateDirectory(ImageOutputFolder);

                List<ImageListModel> selectedItems = imageNameList.Where(i => i.IsSelected).ToList();
                int total = selectedItems.Count;
                int current = 0;

                ProgressValue = 0;

                await Task.Run(() =>
                {
                    foreach (var item in selectedItems)
                    {
                        string imagePath = Path.Combine(ImageListFolderPath, item.FileName);
                        string outputPath = Path.Combine(ImageOutputFolder, item.FileName);

                        CurrentFileName = item.FileName;

                        if (File.Exists(imagePath))
                        {
                            File.Copy(imagePath, outputPath, true);
                        }

                        current++;
                        ProgressValue = (double)current / total * 100;
                    }
                });

                InformationMessage("Selected images copied.", "Success");
            }
            catch (Exception ex)
            {
                ErrorMessage(ex);
            }
            finally
            {
                CurrentFileName = string.Empty;
            }
        }





    }
}
