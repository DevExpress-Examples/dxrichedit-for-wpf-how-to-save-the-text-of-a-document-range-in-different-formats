Imports System
Imports System.Diagnostics
Imports System.Drawing.Imaging
Imports System.IO
Imports System.Windows
Imports Microsoft.Win32
Imports DevExpress.Xpf.Core
Imports DevExpress.XtraRichEdit
Imports DevExpress.XtraRichEdit.Export
Imports DevExpress.XtraRichEdit.Export.Html
Imports DevExpress.Office.Services
Imports DevExpress.Office.Utils

Namespace GetTextMethodsExample

    ''' <summary>
    ''' Interaction logic for MainWindow.xaml
    ''' </summary>
    Public Partial Class MainWindow
        Inherits Window

        Private fileName As String = String.Empty

        Public Sub New()
            Me.InitializeComponent()
            Me.richEditControl1.LoadDocument("sample_document.rtf", DocumentFormat.Rtf)
        End Sub

        Private Sub btnSaveAsMht_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            fileName = GetFileName("Web Archive|*.mht")
            If String.IsNullOrEmpty(fileName) Then Return
            Try
                Dim mhtText As String = Me.richEditControl1.Document.GetMhtText(Me.richEditControl1.Document.Range)
                SaveFile(fileName, mhtText)
                OpenFile(fileName)
            Finally
                fileName = String.Empty
            End Try
        End Sub

        Private Sub btnSaveAsHtml_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            fileName = GetFileName("Hypertext Markup Language|*.html")
            If String.IsNullOrEmpty(fileName) Then Return
            Try
                Dim uriProvider As CustomUriProvider = New CustomUriProvider(Path.GetDirectoryName(fileName))
                Dim htmlText As String = Me.richEditControl1.Document.GetHtmlText(Me.richEditControl1.Document.Range, uriProvider)
                SaveFile(fileName, htmlText)
                OpenFile(fileName)
            Finally
                fileName = String.Empty
            End Try
        End Sub

        Private Sub btnSaveAsDocx_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            fileName = GetFileName("Word Document|*.docx")
            If String.IsNullOrEmpty(fileName) Then Return
            Try
                Dim bytes As Byte() = Me.richEditControl1.Document.GetOpenXmlBytes(Me.richEditControl1.Document.Range)
                SaveFile(fileName, bytes)
                OpenFile(fileName)
            Finally
                fileName = String.Empty
            End Try
        End Sub

        Private Sub btnSaveAsRtf_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            fileName = GetFileName("Rich Text Format|*.rtf")
            If String.IsNullOrEmpty(fileName) Then Return
            Try
                Dim rtfText As String = Me.richEditControl1.Document.GetRtfText(Me.richEditControl1.Document.Range)
                SaveFile(fileName, rtfText)
                OpenFile(fileName)
            Finally
                fileName = String.Empty
            End Try
        End Sub

'#Region "#beforeexport"
        Private Sub richEditControl1_BeforeExport(ByVal sender As Object, ByVal e As BeforeExportEventArgs)
            Dim options As HtmlDocumentExporterOptions = TryCast(e.Options, HtmlDocumentExporterOptions)
            If options IsNot Nothing Then
                options.CssPropertiesExportType = CssPropertiesExportType.Link
                options.HtmlNumberingListExportFormat = HtmlNumberingListExportFormat.HtmlFormat
                options.TargetUri = Path.GetFileNameWithoutExtension(fileName)
            End If
        End Sub

'#End Region  ' #beforeexport
        Private Sub SaveFile(ByVal fileName As String, ByVal value As String)
            Using stream As FileStream = New FileStream(fileName, FileMode.Create, FileAccess.Write)
                Using writer As StreamWriter = New StreamWriter(stream)
                    writer.Write(value)
                End Using
            End Using
        End Sub

        Private Sub SaveFile(ByVal fileName As String, ByVal bytes As Byte())
            Using stream As FileStream = New FileStream(fileName, FileMode.Create, FileAccess.Write)
                stream.Write(bytes, 0, bytes.Length)
            End Using
        End Sub

        Private Function GetFileName(ByVal filter As String) As String
            Dim saveFileDialog As SaveFileDialog = New SaveFileDialog()
            saveFileDialog.Filter = filter
            saveFileDialog.RestoreDirectory = True
            saveFileDialog.CheckFileExists = False
            saveFileDialog.CheckPathExists = True
            saveFileDialog.OverwritePrompt = True
            saveFileDialog.DereferenceLinks = True
            saveFileDialog.ValidateNames = True
            If saveFileDialog.ShowDialog(Me) = True Then Return saveFileDialog.FileName
            Return String.Empty
        End Function

        Private Sub OpenFile(ByVal fileName As String)
            If DXMessageBox.Show("Do you want to open this file?", "Html Example", MessageBoxButton.YesNo, MessageBoxImage.Question) = MessageBoxResult.Yes Then
                Dim process As Process = New Process()
                Try
                    process.StartInfo.FileName = fileName
                    process.Start()
                Catch
                End Try
            End If
        End Sub
    End Class

'#Region "#uriprovider"
    Public Class CustomUriProvider
        Implements IUriProvider

        Private rootDirecory As String

        Public Sub New(ByVal rootDirectory As String)
            If String.IsNullOrEmpty(rootDirectory) Then Exceptions.ThrowArgumentException("rootDirectory", rootDirectory)
            rootDirecory = rootDirectory
        End Sub

        Public Function CreateCssUri(ByVal rootUri As String, ByVal styleText As String, ByVal relativeUri As String) As String Implements IUriProvider.CreateCssUri
            Dim cssDir As String = String.Format("{0}\{1}", rootDirecory, rootUri.Trim("/"c))
            If Not Directory.Exists(cssDir) Then Directory.CreateDirectory(cssDir)
            Dim cssFileName As String = String.Format("{0}\style.css", cssDir)
            File.AppendAllText(cssFileName, styleText)
            Return GetRelativePath(cssFileName)
        End Function

        Public Function CreateImageUri(ByVal rootUri As String, ByVal image As OfficeImage, ByVal relativeUri As String) As String Implements IUriProvider.CreateImageUri
            Dim imagesDir As String = String.Format("{0}\{1}", rootDirecory, rootUri.Trim("/"c))
            If Not Directory.Exists(imagesDir) Then Directory.CreateDirectory(imagesDir)
            Dim imageName As String = String.Format("{0}\{1}.png", imagesDir, Guid.NewGuid())
            image.NativeImage.Save(imageName, ImageFormat.Png)
            Return GetRelativePath(imageName)
        End Function

        Private Function GetRelativePath(ByVal path As String) As String
            Dim substring As String = path.Substring(rootDirecory.Length)
            Return substring.Replace("\", "/").Trim("/"c)
        End Function
    End Class
'#End Region  ' #uriprovider
End Namespace
