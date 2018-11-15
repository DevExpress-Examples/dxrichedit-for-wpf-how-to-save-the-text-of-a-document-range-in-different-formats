<!-- default file list -->
*Files to look at*:

* [MainWindow.xaml](./CS/MainWindow.xaml) (VB: [MainWindow.xaml.vb](./VB/MainWindow.xaml.vb))
* [MainWindow.xaml.cs](./CS/MainWindow.xaml.cs) (VB: [MainWindow.xaml.vb](./VB/MainWindow.xaml.vb))
<!-- default file list end -->
# DXRichEdit for WPF: How to save the text of a document range in different formats


<p>This example illustrates API methods used to get the text of the <a href="http://documentation.devexpress.com/#WindowsForms/clsDevExpressXtraRichEditAPINativeDocumentRangetopic"><u>document range</u></a> in different formats - RTF, HTML, MHT, DOCX.<br />
Although the preferable technique to save the document in different formats is the <a href="http://documentation.devexpress.com/#WindowsForms/DevExpressXtraRichEditAPINativeDocument_SaveDocumenttopic"><u>SaveDocument</u></a> and the <a href="http://documentation.devexpress.com/#WindowsForms/DevExpressXtraRichEditRichEditControl_SaveDocumentAstopic"><u>SaveDocumentAs</u></a> methods, several methods allow obtaining text of the specified range in different formats. Current example provides code snippets which use these methods. Note the implementation of the <a href="http://documentation.devexpress.com/#WindowsForms/clsDevExpressXtraRichEditServicesIUriProvidertopic"><u>IUriProvider</u></a> interface required for HTML export.</p><br />


<br/>


