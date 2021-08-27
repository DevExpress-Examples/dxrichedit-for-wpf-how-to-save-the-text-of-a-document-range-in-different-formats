<!-- default badges list -->
![](https://img.shields.io/endpoint?url=https://codecentral.devexpress.com/api/v1/VersionRange/128607157/12.2.4%2B)
[![](https://img.shields.io/badge/Open_in_DevExpress_Support_Center-FF7200?style=flat-square&logo=DevExpress&logoColor=white)](https://supportcenter.devexpress.com/ticket/details/E3267)
[![](https://img.shields.io/badge/📖_How_to_use_DevExpress_Examples-e9f6fc?style=flat-square)](https://docs.devexpress.com/GeneralInformation/403183)
<!-- default badges end -->
<!-- default file list -->
*Files to look at*:

* [MainWindow.xaml](./CS/MainWindow.xaml) (VB: [MainWindow.xaml](./VB/MainWindow.xaml))
* [MainWindow.xaml.cs](./CS/MainWindow.xaml.cs) (VB: [MainWindow.xaml.vb](./VB/MainWindow.xaml.vb))
<!-- default file list end -->
# DXRichEdit for WPF: How to save the text of a document range in different formats


<p>This example illustrates API methods used to get the text of the <a href="http://documentation.devexpress.com/#WindowsForms/clsDevExpressXtraRichEditAPINativeDocumentRangetopic"><u>document range</u></a> in different formats - RTF, HTML, MHT, DOCX.<br />
Although the preferable technique to save the document in different formats is the <a href="http://documentation.devexpress.com/#WindowsForms/DevExpressXtraRichEditAPINativeDocument_SaveDocumenttopic"><u>SaveDocument</u></a> and the <a href="http://documentation.devexpress.com/#WindowsForms/DevExpressXtraRichEditRichEditControl_SaveDocumentAstopic"><u>SaveDocumentAs</u></a> methods, several methods allow obtaining text of the specified range in different formats. Current example provides code snippets which use these methods. Note the implementation of the <a href="http://documentation.devexpress.com/#WindowsForms/clsDevExpressXtraRichEditServicesIUriProvidertopic"><u>IUriProvider</u></a> interface required for HTML export.</p><br />


<br/>


