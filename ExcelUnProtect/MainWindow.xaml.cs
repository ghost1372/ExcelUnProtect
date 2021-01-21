using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using HandyControl.Data;
using System.Windows;
using System.Windows.Forms;
using System.Xml.Linq;
using Application = System.Windows.Application;
using Button = System.Windows.Controls.Button;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;

namespace ExcelUnProtect
{
    public partial class MainWindow
    {
        private static readonly string Path = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "UnProtectTemp");

        private readonly string _archivePath = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "UnProtectedArchives");
        private readonly string _pathWorkBook = System.IO.Path.Combine(Path, @"xl\workbook.xml");
        private readonly string _pathSheets = System.IO.Path.Combine(Path, @"xl\worksheets");

        public MainWindow()
        {
            InitializeComponent();
        }

        #region Change Skin
        private void ButtonConfig_OnClick(object sender, RoutedEventArgs e) => PopupConfig.IsOpen = true;

        private void ButtonSkins_OnClick(object sender, RoutedEventArgs e)
        {
            if (e.OriginalSource is Button button && button.Tag is SkinType tag)
            {
                PopupConfig.IsOpen = false;
                ((App)Application.Current).UpdateSkin(tag);
            }
        }
        #endregion

        private void ButtonBase_OnClick(object sender, RoutedEventArgs e)
        {
            if (Directory.Exists(_archivePath))
            {
                Directory.Delete(_archivePath, true);
            }

            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Multiselect = true, Filter = "Excel Files|*.xlsx;*.xls"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                foreach (var file in openFileDialog.FileNames)
                {
                    txtLog.Text += System.IO.Path.GetFileName(file) + " Loaded" + Environment.NewLine;
                    UnProtectExcel(file);
                }

                SaveUnProtectedExcel();
            }
        }

        private void UnProtectExcel(string fileName)
        {
            if (Directory.Exists(Path))
            {
                Directory.Delete(Path, true);
            }

            ZipFile.ExtractToDirectory(fileName, Path);

            txtLog.Text += $"UnProtecting {System.IO.Path.GetFileNameWithoutExtension(fileName)} WorkBook" + Environment.NewLine;

            UnProtect(_pathWorkBook, "workbookProtection");

            UnProtectionSheets();

            if (!Directory.Exists(_archivePath))
            {
                Directory.CreateDirectory(_archivePath);
            }
            ZipFile.CreateFromDirectory(Path, System.IO.Path.Combine(_archivePath, System.IO.Path.Combine(_archivePath, System.IO.Path.GetFileName(fileName))));

            txtLog.Text += "Done!" + Environment.NewLine;

        }

        private void SaveUnProtectedExcel()
        {
            using var fbd = new FolderBrowserDialog();
            fbd.Description = "Save UnProtected Excels";
            fbd.UseDescriptionForTitle = true;
            fbd.ShowNewFolderButton = true;
            DialogResult result = fbd.ShowDialog();

            if (result == System.Windows.Forms.DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
            {
                var files = Directory.GetFiles(_archivePath);
                foreach (var file in files)
                {
                    File.Move(file, System.IO.Path.Combine(fbd.SelectedPath, System.IO.Path.GetFileName(file)), true);
                }
            }
        }

        private void UnProtect(string path, string tagname)
        {
            XNamespace xn = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            XDocument doc = XDocument.Load(path);
            var q = from node in doc.Descendants(xn + tagname)
                select node;
            q.ToList().ForEach(x => x.Remove());
            doc.Save(path);
        }

        private void UnProtectionSheets()
        {
            var files = Directory.EnumerateFiles(_pathSheets);
            foreach (var file in files)
            {
                txtLog.Text += $"UnProtecting {System.IO.Path.GetFileNameWithoutExtension(file)}" + Environment.NewLine;

                UnProtect(file, "sheetProtection");
            }
        }
    }
}
