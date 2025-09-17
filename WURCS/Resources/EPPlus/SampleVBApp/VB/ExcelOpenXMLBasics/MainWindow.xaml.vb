Option Explicit On
Public Class MainWindow

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

#If EN_US_CULTURE Then
        System.Threading.Thread.CurrentThread.CurrentUICulture = System.Threading.Thread.CurrentThread.CurrentCulture
        System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("en-US")
#End If

    End Sub

    Public Sub btnCreateBasicWorkbook_Click(sender As System.Object, e As System.Windows.RoutedEventArgs)
        Me.CreateBasicWorkbook("BasicWorkbook.xlsx", True)
    End Sub

    Private Sub btnCreate10000SharedStrings_Click(sender As System.Object, e As System.Windows.RoutedEventArgs)
        Me.CreateStringWorkbook("SharedStrings10000.xlsx", 10000, True)
    End Sub

    Private Sub btnCreate10000Strings_Click(sender As System.Object, e As System.Windows.RoutedEventArgs)
        Me.CreateStringWorkbook("Strings10000.xlsx", 10000, False)
    End Sub

    Private Sub btnCreateBasicWorkbookPredefinedStyles_Click(sender As System.Object, e As System.Windows.RoutedEventArgs)
        Me.CreateBasicWorkbook("BasicWorkbookPredefinedStyles.xlsx", False)
    End Sub

    ''' <summary>
    ''' Creates a workbook with specified amount of strings
    ''' </summary>
    ''' <param name="workbookName">Name of the workbook</param>
    ''' <param name="stringCount">Number of strings to add</param>
    ''' <param name="useShared">Use shared strings?</param>
    ''' <returns>True if succesful</returns>
    Private Function CreateStringWorkbook(workbookName As String, stringCount As Int32, useShared As Boolean) As Boolean
        Dim spreadsheet As DocumentFormat.OpenXml.Packaging.SpreadsheetDocument
        Dim worksheet As DocumentFormat.OpenXml.Spreadsheet.Worksheet

        spreadsheet = Excel.CreateWorkbook(workbookName)
        If (spreadsheet Is Nothing) Then
            Return False
        End If

        Excel.AddBasicStyles(spreadsheet)
        Excel.AddWorksheet(spreadsheet, "Strings")
        worksheet = spreadsheet.WorkbookPart.WorksheetParts.First().Worksheet

        ' Add shared strings
        For counter As UInt32 = 0 To stringCount - 1 Step 1
            Excel.SetStringCellValue(spreadsheet, worksheet, 1, counter + 1, "Some string", useShared, False)
        Next
        ' Set column widths
        Excel.SetColumnWidth(worksheet, 1, 15)

        worksheet.Save()
        spreadsheet.Close()

        System.Diagnostics.Process.Start(workbookName)
        Return True
    End Function

    ''' <summary>
    ''' Creates a basic workbook
    ''' </summary>
    ''' <param name="workbookName">Name of the workbook</param>
    ''' <param name="createStylesInCode">Create the styles in code?</param>
    Private Sub CreateBasicWorkbook(workbookName As String, createStylesInCode As Boolean)
        Dim spreadsheet As DocumentFormat.OpenXml.Packaging.SpreadsheetDocument
        Dim worksheet As DocumentFormat.OpenXml.Spreadsheet.Worksheet
        Dim styleXml As String

        spreadsheet = Excel.CreateWorkbook(workbookName)
        If (spreadsheet Is Nothing) Then
            Return
        End If

        If (createStylesInCode) Then
            Excel.AddBasicStyles(spreadsheet)
        Else
            Using styleXmlReader As System.IO.StreamReader = New System.IO.StreamReader("PredefinedStyles.xml")
                styleXml = styleXmlReader.ReadToEnd()
                Excel.AddPredefinedStyles(spreadsheet, styleXml)
            End Using
        End If

        Excel.AddSharedString(spreadsheet, "Shared string")
        Excel.AddWorksheet(spreadsheet, "Test 1")
        Excel.AddWorksheet(spreadsheet, "Test 2")
        worksheet = spreadsheet.WorkbookPart.WorksheetParts.First().Worksheet

        ' Add shared strings
        Excel.SetStringCellValue(spreadsheet, worksheet, 1, 1, "Shared string", True)
        Excel.SetStringCellValue(spreadsheet, worksheet, 1, 2, "Shared string", True)
        Excel.SetStringCellValue(spreadsheet, worksheet, 1, 3, "Shared string", True)

        ' Add a string
        Excel.SetStringCellValue(spreadsheet, worksheet, 1, 5, "Number", False, False)
        ' Add a decimal number
        Excel.SetDoubleCellValue(spreadsheet, worksheet, 2, 5, 1.23, Nothing, True)

        ' Add a string
        Excel.SetStringCellValue(spreadsheet, worksheet, 1, 6, "Integer", False, False)
        ' Add an integer number
        Excel.SetDoubleCellValue(spreadsheet, worksheet, 2, 6, 1, Nothing, True)

        ' Add a string
        Excel.SetStringCellValue(spreadsheet, worksheet, 1, 7, "Currency", False, False)
        ' Add currency
        Excel.SetDoubleCellValue(spreadsheet, worksheet, 2, 7, 1.23, 2, True)

        ' Add a string
        Excel.SetStringCellValue(spreadsheet, worksheet, 1, 8, "Date", False, False)
        ' Add date
        Excel.SetDateCellValue(spreadsheet, worksheet, 2, 8, System.DateTime.Now, 1, True)

        ' Add a string
        Excel.SetStringCellValue(spreadsheet, worksheet, 1, 9, "Percentage", False, False)
        ' Add percentage
        Excel.SetDoubleCellValue(spreadsheet, worksheet, 2, 9, 0.123, 3, True)

        ' Add a string
        Excel.SetStringCellValue(spreadsheet, worksheet, 1, 10, "Boolean", False, False)
        ' Add boolean
        Excel.SetBooleanCellValue(spreadsheet, worksheet, 2, 10, True, Nothing, True)

        ' Set column widths
        Excel.SetColumnWidth(worksheet, 1, 15)
        Excel.SetColumnWidth(worksheet, 2, 20)

        worksheet.Save()
        spreadsheet.Close()

        System.Diagnostics.Process.Start(workbookName)
    End Sub

End Class
