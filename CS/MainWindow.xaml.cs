using System;
using System.Diagnostics;
using System.Drawing.Imaging;
using System.IO;
using System.Windows;

using Microsoft.Win32;

using DevExpress.Xpf.Core;
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.Export;
using DevExpress.XtraRichEdit.Export.Html;
using DevExpress.XtraRichEdit.Services;
using DevExpress.XtraRichEdit.Utils;
using DevExpress.Office.Utils;

namespace GetTextMethodsExample {
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window {
        string fileName = String.Empty;

        public MainWindow() {
            InitializeComponent();
            richEditControl1.LoadDocument("sample_document.rtf", DocumentFormat.Rtf); 
        }

        private void btnSaveAsMht_Click(object sender, RoutedEventArgs e) {
            this.fileName = GetFileName("Web Archive|*.mht");
            if (String.IsNullOrEmpty(fileName))
                return;
            try {
                string mhtText = this.richEditControl1.Document.GetMhtText(this.richEditControl1.Document.Range);
                SaveFile(this.fileName, mhtText);
                OpenFile(this.fileName);
            }
            finally {
                this.fileName = String.Empty;
            }
        }

        private void btnSaveAsHtml_Click(object sender, RoutedEventArgs e) {
            this.fileName = GetFileName("Hypertext Markup Language|*.html");
            if (String.IsNullOrEmpty(fileName))
                return;
            try {
                CustomUriProvider uriProvider = new CustomUriProvider(System.IO.Path.GetDirectoryName(fileName));
                string htmlText = this.richEditControl1.Document.GetHtmlText(this.richEditControl1.Document.Range, uriProvider);
                SaveFile(this.fileName, htmlText);
                OpenFile(this.fileName);
            }
            finally {
                this.fileName = String.Empty;
            }
        }

        private void btnSaveAsDocx_Click(object sender, RoutedEventArgs e) {
            this.fileName = GetFileName("Word Document|*.docx");
            if (String.IsNullOrEmpty(fileName))
                return;
            try {
                byte[] bytes = this.richEditControl1.Document.GetOpenXmlBytes(this.richEditControl1.Document.Range);
                SaveFile(this.fileName, bytes);
                OpenFile(this.fileName);
            }
            finally {
                this.fileName = String.Empty;
            }
        }

        private void btnSaveAsRtf_Click(object sender, RoutedEventArgs e) {
            this.fileName = GetFileName("Rich Text Format|*.rtf");
            if (String.IsNullOrEmpty(fileName))
                return;
            try {
                string rtfText = this.richEditControl1.Document.GetRtfText(this.richEditControl1.Document.Range);
                SaveFile(this.fileName, rtfText);
                OpenFile(this.fileName);
            }
            finally {
                this.fileName = String.Empty;
            }
        }

        #region #beforeexport
        private void richEditControl1_BeforeExport(object sender, BeforeExportEventArgs e) {
            HtmlDocumentExporterOptions options = e.Options as HtmlDocumentExporterOptions;
            if (options != null) {
                options.CssPropertiesExportType = CssPropertiesExportType.Link;
                options.HtmlNumberingListExportFormat = HtmlNumberingListExportFormat.HtmlFormat;
                options.TargetUri = System.IO.Path.GetFileNameWithoutExtension(this.fileName);
            }
        }
        #endregion #beforeexport

        private void SaveFile(string fileName, string value) {
            using (FileStream stream = new FileStream(fileName, FileMode.Create, FileAccess.Write)) {
                using (StreamWriter writer = new StreamWriter(stream)) {
                    writer.Write(value);
                }
            }
        }
        private void SaveFile(string fileName, byte[] bytes) {
            using (FileStream stream = new FileStream(fileName, FileMode.Create, FileAccess.Write)) {
                stream.Write(bytes, 0, bytes.Length);
            }
        }
        private string GetFileName(string filter) {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = filter;
            saveFileDialog.RestoreDirectory = true;
            saveFileDialog.CheckFileExists = false;
            saveFileDialog.CheckPathExists = true;
            saveFileDialog.OverwritePrompt = true;
            saveFileDialog.DereferenceLinks = true;
            saveFileDialog.ValidateNames = true;
            if (saveFileDialog.ShowDialog(this) == true)
                return saveFileDialog.FileName;
            return String.Empty;
        }
        private void OpenFile(string fileName) {
            if (DXMessageBox.Show("Do you want to open this file?", "Html Example", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes) {
                Process process = new Process();
                try {
                    process.StartInfo.FileName = fileName;
                    process.Start();
                }
                catch {
                }
            }
        }

    }
    #region #uriprovider
    public class CustomUriProvider : IUriProvider {
        string rootDirecory;
        public CustomUriProvider(string rootDirectory) {
            if (String.IsNullOrEmpty(rootDirectory))
                Exceptions.ThrowArgumentException("rootDirectory", rootDirectory);
            this.rootDirecory = rootDirectory;
        }

        public string CreateCssUri(string rootUri, string styleText, string relativeUri) {
            string cssDir = String.Format("{0}\\{1}", this.rootDirecory, rootUri.Trim('/'));
            if (!Directory.Exists(cssDir))
                Directory.CreateDirectory(cssDir);
            string cssFileName = String.Format("{0}\\style.css", cssDir);
            File.AppendAllText(cssFileName, styleText);
            return GetRelativePath(cssFileName);
        }
        public string CreateImageUri(string rootUri, DevExpress.Office.Utils.OfficeImage image, string relativeUri) {
            string imagesDir = String.Format("{0}\\{1}", this.rootDirecory, rootUri.Trim('/'));
            if (!Directory.Exists(imagesDir))
                Directory.CreateDirectory(imagesDir);
            string imageName = String.Format("{0}\\{1}.png", imagesDir, Guid.NewGuid());
            image.NativeImage.Save(imageName, ImageFormat.Png);
            return GetRelativePath(imageName);
        }
        string GetRelativePath(string path) {
            string substring = path.Substring(this.rootDirecory.Length);
            return substring.Replace("\\", "/").Trim('/');
        }
    }
    #endregion #uriprovider
}
