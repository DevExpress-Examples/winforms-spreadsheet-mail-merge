<!-- default badges list -->
![](https://img.shields.io/endpoint?url=https://codecentral.devexpress.com/api/v1/VersionRange/128614039/14.2.3%2B)
[![](https://img.shields.io/badge/Open_in_DevExpress_Support_Center-FF7200?style=flat-square&logo=DevExpress&logoColor=white)](https://supportcenter.devexpress.com/ticket/details/T109352)
[![](https://img.shields.io/badge/📖_How_to_use_DevExpress_Examples-e9f6fc?style=flat-square)](https://docs.devexpress.com/GeneralInformation/403183)
[![](https://img.shields.io/badge/💬_Leave_Feedback-feecdd?style=flat-square)](#does-this-example-address-your-development-requirementsobjectives)
<!-- default badges end -->
<!-- default file list -->
*Files to look at*:

* [Form1.cs](./CS/DXApplication1/Form1.cs) (VB: [Form1.vb](./VB/DXApplication1/Form1.vb))
<!-- default file list end -->
# Spreadsheet Mail Merge - Getting Started


<p>This example demonstrates how to use the <a href="http://help.devexpress.com/#WindowsForms/CustomDocument16257">Spreadsheet Mail Merge</a> functionality to automatically generate a document based on data retrieved from a specified data source (the <strong>Categories</strong> table of the <strong>Northwind</strong> database). The <em>nwind.mdb</em> database file is included in the project. This file ships with the XtraSpreadsheet Suite installation.</p>
<p>The application includes the following controls.</p>
<p>1. <strong>SpreadsheetControl</strong></p>
<p>A ready-to-use <a href="http://help.devexpress.com/#WindowsForms/CustomDocument17018">mail merge template</a> (the <em>MailMergeTemplate.xlsx</em> file) is automatically loaded into the SpreadsheetControl when invoking the application. This template is bound to the <strong>Categories</strong> table of the <strong>Northwind</strong> database via the <a href="http://help.devexpress.com/#CoreLibraries/DevExpressSpreadsheetIWorkbook_MailMergeDataSourcetopic">IWorkbook.MailMergeDataSource</a> and <a href="http://help.devexpress.com/#CoreLibraries/DevExpressSpreadsheetIWorkbook_MailMergeDataMembertopic">IWorkbook.MailMergeDataMember</a> properties in code.</p>
<p>2. <strong>Field List</strong></p>
<p>The Field List panel shows the structure of the mail merge data source at runtime. You can add mail merge fields to template cells by drag-and-drop or by double-clicking the corresponding items in this panel.</p>
<p>3. <strong>RibbonControl</strong> with the <strong>Mail Merge</strong> tab<br>The <strong>Mail Merge</strong> page contains various command buttons that you can use to modify the mail merge template. For example, you can do the following.<br>- Enable the <strong>Mail Merge Design View</strong>. This is a specific mode to display worksheets intended for preparing mail merge templates. It is recommended that you activate this view before creating or modifying a template.<br>- Select one of the available modes to specify whether data records should be merged into separate workbooks, separate worksheets in a single workbook, or a single worksheet. By default, the <strong>Single Sheet</strong> mode is used.<br>- Set the vertical or horizontal orientation for the resulting worksheet.<br>- Specify the <strong>Detail</strong>, <strong>Header</strong> and <strong>Footer</strong> ranges in the template.<br>- <a href="http://help.devexpress.com/#WindowsForms/CustomDocument16975">Sort</a>, <a href="http://help.devexpress.com/#WindowsForms/CustomDocument16986">group</a> and <a href="http://help.devexpress.com/#WindowsForms/CustomDocument16995">filter</a> the data to appear in the resulting document. <br>- Preview the merged document.<br>In this example, the <strong>Result</strong> button is added to the <strong>Mail Merge</strong> page. It calls the <a href="http://help.devexpress.com/#CoreLibraries/DevExpressSpreadsheetIWorkbook_GenerateMailMergeDocumentstopic">IWorkbook.GenerateMailMergeDocuments</a> method, which returns the resulting workbook and saves it to a file. If the mail merge mode is set to <strong>Multiple Documents</strong>, each workbook returned by the <strong>GenerateMailMergeDocuments</strong> method is saved to a separate file.</p>

<br/>


<!-- feedback -->
## Does this example address your development requirements/objectives?

[<img src="https://www.devexpress.com/support/examples/i/yes-button.svg"/>](https://www.devexpress.com/support/examples/survey.xml?utm_source=github&utm_campaign=winforms-spreadsheet-mail-merge&~~~was_helpful=yes) [<img src="https://www.devexpress.com/support/examples/i/no-button.svg"/>](https://www.devexpress.com/support/examples/survey.xml?utm_source=github&utm_campaign=winforms-spreadsheet-mail-merge&~~~was_helpful=no)

(you will be redirected to DevExpress.com to submit your response)
<!-- feedback end -->
