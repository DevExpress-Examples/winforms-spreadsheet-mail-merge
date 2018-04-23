Imports DevExpress.Spreadsheet
Imports DevExpress.XtraBars.Ribbon
Imports System.Collections.Generic

Namespace DXApplication1
    Partial Public Class Form1
        Inherits RibbonForm
#Region "#mailmerge1"
        Private dataSet As nwindDataSet
        Private adapter As nwindDataSetTableAdapters.CategoriesTableAdapter
        Private template As IWorkbook

        Public Sub New()
            InitializeComponent()

            dataSet = New nwindDataSet()
            adapter = New nwindDataSetTableAdapters.CategoriesTableAdapter()
            adapter.Fill(dataSet.Categories)
        End Sub

        Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
            spreadsheetControl1.LoadDocument("Data\MailMergeTemplate.xlsx")
            template = spreadsheetControl1.Document
            template.MailMergeDataSource = dataSet
            template.MailMergeDataMember = "Categories"
        End Sub
#End Region ' #mailmerge1
        #Region "#mailmerge2"
        Private Sub barButtonItem1_ItemClick(ByVal sender As Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles barButtonItem1.ItemClick
            Dim resultWorkbooks As IList(Of IWorkbook) = spreadsheetControl1.Document.GenerateMailMergeDocuments()
            Dim index As Integer = 0
            For Each workbook As IWorkbook In resultWorkbooks
                Dim fileName As String = String.Format("Data\SavedDocument{0}" & ".xlsx", index)
                index += 1
                workbook.SaveDocument(fileName, DocumentFormat.OpenXml)
            Next workbook
            System.Diagnostics.Process.Start("explorer.exe ", Application.StartupPath + "\Data")
        End Sub
        #End Region ' #mailmerge2
    End Class
End Namespace
