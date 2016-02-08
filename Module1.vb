Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.IO

Module Module1
    Public conn As New OleDbConnection()
    Public Filename As String
    Public chkexcel As Boolean
    Public oexcel As Excel.Application
    Public obook As Excel.Workbook
    Public osheet As Excel.Worksheet
    Public R As Integer

    Sub Main()
        Try
            Dbopen()
            'File name and path, here i used abc file to be stored in Bin directory in the sloution directory
            Filename = AppDomain.CurrentDomain.BaseDirectory & "abc.xls"
            'check if file already exists then delete it to create a new file
            If File.Exists(Filename) Then
                File.Delete(Filename)
            End If
            If Not File.Exists(Filename) Then
                chkexcel = False
                'create new excel application
                oexcel = CreateObject("Excel.Application")
                'add a new workbook
                obook = oexcel.Workbooks.Add
                'set the application alerts not to be displayed for confirmation
                oexcel.Application.DisplayAlerts = True
                'check total sheets in workboob
                Dim S As Integer = oexcel.Application.Sheets.Count()
                'leaving first sheet delete all the remaining sheets
                If S > 1 Then
                    oexcel.Application.DisplayAlerts = False
                    Dim J As Integer = S
                    Do While J > 1
                        oexcel.Application.Sheets(J).delete()
                        J = oexcel.Application.Sheets.Count()
                    Loop
                End If
                'to check the session of excel application
                chkexcel = True


                oexcel.Visible = True
                'this procedure populate the sheet
                Generate_Sheet()
                'save excel file
                obook.SaveAs(Filename)
                'end application object and session
                osheet = Nothing
                oexcel.Application.DisplayAlerts = False
                obook.Close()
                oexcel.Application.DisplayAlerts = True
                obook = Nothing
                oexcel.Quit()
                oexcel = Nothing
                chkexcel = False
                'mail excel file as an attachment
                automail("Mongkol.Me@glsict.com", "Auto Excel File", "any message", Filename)
            End If
        Catch ex As Exception
            'mail error message
            automail("b5209194@gmail.com", "Error Message", ex.Message, "")
        Finally
            Dbclose()
        End Try
    End Sub

    Public Sub automail(ByVal mail_to As String, ByVal subject As String, ByVal msg As String, ByVal filename As String)
        Dim myOutlook As New Outlook.Application()
        Dim myMailItem, attach As Object
        myMailItem = myOutlook.CreateItem(Outlook.OlItemType.olMailItem)
        myMailItem.Body = msg
        If File.Exists(filename) Then
            attach = myMailItem.Attachments
            attach.Add(filename)
        End If
        If Trim(mail_to) <> "" Then
            myMailItem.to = Trim(mail_to)
        End If
        myMailItem.SUBJECT = subject
        myMailItem.send()
        myMailItem = Nothing
        myOutlook = Nothing
    End Sub

    Public Sub Dbopen()
        'open connection for db.mdb stroed in the base directory
        conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source='" & AppDomain.CurrentDomain.BaseDirectory & "db.mdb'"
        conn.Open()
    End Sub
    Public Sub Dbclose()
        'check and close db connection
        If conn.State = ConnectionState.Open Then
            conn.Close()
            conn.Dispose()
            conn = Nothing
        End If
        'check and close excel application
        If chkexcel = True Then
            osheet = Nothing
            oexcel.Application.DisplayAlerts = False
            obook.Close()
            oexcel.Application.DisplayAlerts = True
            obook = Nothing
            oexcel.Quit()
            oexcel = Nothing
        End If
        End
    End Sub
    Sub Generate_Sheet()
        Console.WriteLine("Generating Auto Report")
        osheet = oexcel.Worksheets(1)
        'rename the sheet
        osheet.Name = "Excel Charts"
        osheet.Range("A1:AZ400").Interior.ColorIndex = 2
        osheet.Range("A1").Font.Size = 12
        osheet.Range("A1").Font.Bold = True
        osheet.Range("A1:I1").Merge()
        osheet.Range("A1").Value = "Excel Automation With Charts"
        osheet.Range("A1").EntireColumn.AutoFit()
        'format headings
        osheet.Range("A3:C3").Font.Color = RGB(255, 255, 255)
        osheet.Range("A3:C3").Interior.ColorIndex = 5
        osheet.Range("A3:C3").Font.Bold = True
        osheet.Range("A3:C3").Font.Size = 10
        'columns heading
        osheet.Range("A3").Value = "Item"
        osheet.Range("A3").BorderAround(8)
        osheet.Range("B3").Value = "Sale"
        osheet.Range("B3").BorderAround(8)
        osheet.Range("C3").Value = "Income"
        osheet.Range("C3").BorderAround(8)
        'populate data from DB
        Dim SQlQuery As String = "select * from Sales"
        Dim SQLCommand As New OleDbCommand(SQlQuery, conn)
        Dim SQlReader As OleDbDataReader = SQLCommand.ExecuteReader
        Dim R As Integer = 3
        While SQlReader.Read
            R = R + 1
            osheet.Range("A" & R).Value = SQlReader.GetValue(0).ToString
            osheet.Range("A" & R).BorderAround(8)
            osheet.Range("B" & R).Value = SQlReader.GetValue(1).ToString
            osheet.Range("B" & R).BorderAround(8)
            osheet.Range("C" & R).Value = SQlReader.GetValue(2).ToString
            osheet.Range("C" & R).BorderAround(8)
        End While
        SQlReader.Close()
        SQlReader = Nothing
        'create chart objects
        Dim oChart As Excel.Chart
        Dim MyCharts As Excel.ChartObjects
        Dim MyCharts1 As Excel.ChartObject
        MyCharts = osheet.ChartObjects
        'set chart location
        MyCharts1 = MyCharts.Add(150, 30, 400, 250)
        oChart = MyCharts1.Chart
        'use the follwoing line if u want to draw chart on the default location
        'ochart.Location(Excel.XlChartLocation.xlLocationAsObject, osheet.Name)
        With oChart
            'set data range for chart
            Dim chartRange As Excel.Range
            chartRange = osheet.Range("A3", "C" & R)
            .SetSourceData(chartRange)
            'set how you want to draw chart i.e column wise or row wise
            .PlotBy = Excel.XlRowCol.xlColumns
            'set data lables for bars
            .ApplyDataLabels(Excel.XlDataLabelsType.xlDataLabelsShowNone)
            'set legend to be displayed or not
            .HasLegend = True
            'set legend location
            .Legend.Position = Excel.XlLegendPosition.xlLegendPositionRight
            'select chart type
            '.ChartType = Excel.XlChartType.xl3DBarClustered
            'chart title
            .HasTitle = True
            .ChartTitle.Text = "Sale/Income Bar Chart"
            'set titles for Axis values and categories
            Dim xlAxisCategory, xlAxisValue As Excel.Axes
            xlAxisCategory = CType(oChart.Axes(, Excel.XlAxisGroup.xlPrimary), Excel.Axes)
            xlAxisCategory.Item(Excel.XlAxisType.xlCategory).HasTitle = True
            xlAxisCategory.Item(Excel.XlAxisType.xlCategory).AxisTitle.Characters.Text = "Items"
            xlAxisValue = CType(oChart.Axes(, Excel.XlAxisGroup.xlPrimary), Excel.Axes)
            xlAxisValue.Item(Excel.XlAxisType.xlValue).HasTitle = True
            xlAxisValue.Item(Excel.XlAxisType.xlValue).AxisTitle.Characters.Text = "Sale/Income"
        End With

        'set style to show the totals
        R = R + 1
        osheet.Range("A" & R & ":C" & R).Font.Bold = True
        osheet.Range("A" & R & ":C" & R).Font.Color = RGB(255, 255, 255)
        osheet.Range("A" & R).Value = "Total"
        osheet.Range("A" & R & ":C" & R).Interior.ColorIndex = 5
        osheet.Range("A" & R & ":C" & R).BorderAround(8)
        'sum the values from column 2 to 3
        Dim columnno = 2
        For columnno = 2 To 3
            Dim Htotal As String = 0
            Dim RowCount As Integer = 4
            Do While RowCount <= R
                Htotal = Htotal + osheet.Cells(RowCount, columnno).value
                osheet.Cells(RowCount, columnno).borderaround(8)
                RowCount = RowCount + 1
            Loop
            'display value
            osheet.Cells(R, columnno).Value = Htotal
            'format colums
            With DirectCast(osheet.Columns(columnno), Excel.Range)
                .AutoFit()
                .NumberFormat = "0,00"
            End With
        Next
        'add a pie chart for total comparison
        MyCharts = osheet.ChartObjects
        MyCharts1 = MyCharts.Add(150, 290, 400, 250)
        oChart = MyCharts1.Chart
        With oChart
            Dim chartRange As Excel.Range
            chartRange = osheet.Range("A" & R, "C" & R)
            .SetSourceData(chartRange)
            .PlotBy = Excel.XlRowCol.xlRows
            .ChartType = Excel.XlChartType.xl3DPie

            .ApplyDataLabels(Excel.XlDataLabelsType.xlDataLabelsShowPercent)
            .HasLegend = False
            .HasTitle = True
            .ChartTitle.Text = "Sale/Income Pie Chart"
            .ChartTitle.Font.Bold = True
        End With
    End Sub
End Module
