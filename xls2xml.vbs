Dim fso,oExcel,wb,ws
Dim str1,str2,str3,str4
Dim getBase

'Skill Calling
'Shell("scriptPath xlsPath)
'(ei. shell("D:/xls2xml.vbs D:/TEST.xls"))
'(ei. shell("CMD /c D:/xls2xml.vbs D:/TEST.xls"))'

'Create Scripting filessystemobject'
Set fso = WScript.CreateObject("Scripting.FileSystemObject")

'Create Excel application object
Set oExcel = WScript.CreateObject("Excel.Application")
oExcel.Visible = True
strPath = WScript.Arguments.Item(0)

' Work_path = fso.GetAbsolutePathName("")
getBase = fso.getbasename(strPath)


'Create a new Xml Type text file on script path.
Set XmlFile = fso.CreateTextFile(getBase & ".xml",True,True)

'Style'
str1 = "<Borders>" & vbNewLine & "<Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""" & vbNewLine & "ss:Color=""#000000""/>"
str2 = "<Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1"" ss:Color=""#000000""/>"
str3 = "<Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1"" ss:Color=""#000000""/>"
str4 = "<Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1"" ss:Color=""#000000""/>" & vbNewLine & "</Borders>"

'Write some xml basic syntax into the created xml file
XmlFile.Write("<?xml version=""1.0""?>"&vbNewLine)
XmlFile.Write("<?mso-application progid=""Excel.Sheet""?>"&vbNewLine)
XmlFile.Write("<Workbook xmlns=""urn:schemas-microsoft-com:office:spreadsheet"""&vbNewLine)
XmlFile.Write("xmlns:o=""urn:schemas-microsoft-com:office:office"""&vbNewLine)
XmlFile.Write("xmlns:x=""urn:schemas-microsoft-com:office:excel"""&vbNewLine)
XmlFile.Write("xmlns:ss=""urn:schemas-microsoft-com:office:spreadsheet"""&vbNewLine)
XmlFile.Write("xmlns:html=""http://www.w3.org/TR/REC-html40"">"&vbNewLine)

XmlFile.Write("<DocumentProperties xmlns=""urn:schemas-microsoft-com:office:office"">"&vbNewLine)
XmlFile.Write("<Author></Author>"&vbNewLine)
XmlFile.Write("<LastAuthor></LastAuthor>"&vbNewLine)
XmlFile.Write("<Revision>1</Revision>"&vbNewLine)
XmlFile.Write("<TotalTime>0</TotalTime>"&vbNewLine)
XmlFile.Write("<Created>2016-06-22T14:58:23Z</Created>"&vbNewLine)
XmlFile.Write("<LastSaved>2016-06-23T03:40:09Z</LastSaved>"&vbNewLine)
XmlFile.Write("<Version>16.00</Version>"&vbNewLine)
XmlFile.Write("</DocumentProperties>"&vbNewLine)

XmlFile.Write("<OfficeDocumentSettings xmlns=""urn:schemas-microsoft-com:office:office"">"& vbNewLine & "<AllowPNG/>" & vbNewLine &  "</OfficeDocumentSettings>")
XmlFile.Write("<ExcelWorkbook xmlns=""urn:schemas-microsoft-com:office:excel"">"& vbNewLine & "<WindowHeight>6825</WindowHeight>" & vbNewLine & "<WindowWidth>13815</WindowWidth>")
XmlFile.Write("<WindowTopX>0</WindowTopX>" & vbNewLine & "<WindowTopY>0</WindowTopY>" & vbNewLine & "<ProtectStructure>False</ProtectStructure>" & vbNewLine & "<ProtectWindows>False</ProtectWindows>")
XmlFile.Write("</ExcelWorkbook>" & vbNewLine)
XmlFile.Write("<Styles>" & vbNewLine)

XmlFile.Write("<Style ss:ID=""Default"" ss:Name=""Normal"">" & vbNewLine & "<Alignment ss:Vertical=""Bottom""/>" & vbNewLine & "<Borders/>"&vbNewLine)
XmlFile.Write("<Font ss:FontName=""Liberation Sans"" ss:Size=""11"" ss:Color=""#000000""/>")
XmlFile.Write("<Interior/>" & vbNewLine & "<NumberFormat/>" & vbNewLine & "<Protection/>" & vbNewLine & "</Style>")
XmlFile.Write("<Style ss:ID=""s68"" ss:Name=""一般 10 19"">" & vbNewLine & "<Alignment ss:Vertical=""Center""/>" & vbNewLine & "<Borders/>"&vbNewLine)
XmlFile.Write("<Font ss:FontName=""新細明體"" x:Family=""Swiss"" ss:Size=""12"" ss:Color=""#000000""/>")
XmlFile.Write("<Interior/>" & vbNewLine & "<NumberFormat/>" & vbNewLine & "<Protection/>" & vbNewLine & "</Style>")
XmlFile.Write("<Style ss:ID=""s69"" ss:Parent=""s68"">" & vbNewLine & "<Alignment ss:Horizontal=""Center"" ss:Vertical=""Center""/>"&vbNewLine)
XmlFile.Write("<Borders>"&vbNewLine)
XmlFile.Write("<Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1"" ss:Color=""#000000""/>"&vbNewLine)
XmlFile.Write("<Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1"" ss:Color=""#000000""/>"&vbNewLine)
XmlFile.Write("<Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1"" ss:Color=""#000000""/>")
XmlFile.Write("</Borders>" & vbNewLine & "<Font ss:FontName=""Calibri"" x:Family=""Swiss"" ss:Color=""#000000""/>" & vbNewLine & "<Interior ss:Color=""#FFC000"" ss:Pattern=""Solid""/>" & vbNewLine & "</Style>")
XmlFile.Write("<Style ss:ID=""s70"" ss:Parent=""s68"">" & vbNewLine & "<Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:WrapText=""1""/>" & vbNewLine & str1 & vbNewLine & str2 & vbNewLine & str3 & vbNewLine & str4)
XmlFile.Write("<Font ss:FontName=""Calibri"" x:Family=""Swiss"" ss:Color=""#000000""/>" & vbNewLine & "<Interior ss:Color=""#FFC000"" ss:Pattern=""Solid""/>" & vbNewLine & "</Style>"&vbNewLine)
XmlFile.Write("<Style ss:ID=""s71"" ss:Parent=""s68"">" & vbNewLine & "<Alignment ss:Horizontal=""Center"" ss:Vertical=""Center""/>" & vbNewLine & str1 & vbNewLine & str2 & vbNewLine & str3 & vbNewLine & str4)
XmlFile.Write("<Font ss:FontName=""Calibri"" x:Family=""Swiss"" ss:Color=""#000000""/>" & vbNewLine & "<Interior ss:Color=""#FFC000"" ss:Pattern=""Solid""/>" & vbNewLine & "</Style>"&vbNewLine)
XmlFile.Write("<Style ss:ID=""s72"" ss:Parent=""s68"">" & vbNewLine & "<Alignment ss:Horizontal=""Center"" ss:Vertical=""Center""/>" & vbNewLine & str1 & vbNewLine & str2 & vbNewLine & str3 & vbNewLine & str4)
XmlFile.Write("<Font ss:FontName=""Calibri"" x:Family=""Swiss"" ss:Color=""#000000""/>" & vbNewLine & "<Interior/>" & vbNewLine & "</Style>")
XmlFile.Write("<Style ss:ID=""s73"" ss:Parent=""s68"">" & vbNewLine & "<Alignment ss:Horizontal=""Center"" ss:Vertical=""Top""/>" & vbNewLine & str1 & vbNewLine & str2 & vbNewLine & str3 & vbNewLine & str4)
XmlFile.Write("<Font ss:FontName=""Calibri"" x:Family=""Swiss"" ss:Color=""#000000""/>" & vbNewLine & "<Interior ss:Color=""#FFC000"" ss:Pattern=""Solid""/>" & vbNewLine & "</Style>"&vbNewLine)
XmlFile.Write("<Style ss:ID=""s74"" ss:Parent=""s68"">" & vbNewLine & "<Alignment ss:Horizontal=""Center"" ss:Vertical=""Top"" ss:WrapText=""1""/>" & vbNewLine & str1 & vbNewLine & str2 & vbNewLine & str3 & vbNewLine & str4)
XmlFile.Write("<Font ss:FontName=""Calibri"" x:Family=""Swiss"" ss:Size=""14"" ss:Color=""#FF0000""/>" & vbNewLine & "<Interior ss:Color=""#FFC000"" ss:Pattern=""Solid""/>" & vbNewLine & "</Style>"&vbNewLine)
XmlFile.Write("<Style ss:ID=""s75"" ss:Parent=""s68"">" & vbNewLine & "<Alignment ss:Horizontal=""Center"" ss:Vertical=""Top"" ss:WrapText=""1""/>" & vbNewLine & str1 & vbNewLine & str2 & vbNewLine & str3 & vbNewLine & str4)
XmlFile.Write("<Font ss:FontName=""Calibri"" x:Family=""Swiss"" ss:Size=""14"" ss:Color=""#558ED5""/>" & vbNewLine & "<Interior ss:Color=""#FFFFFF"" ss:Pattern=""Solid""/>" & vbNewLine & "</Style>"&vbNewLine)

XmlFile.Write("<Style ss:ID=""s76"" ss:Parent=""s68"">" & vbNewLine & "<Alignment ss:Horizontal=""Center"" ss:Vertical=""Center""/>" & vbNewLine & str1 & vbNewLine & str2 & vbNewLine & str3 & vbNewLine & str4)
XmlFile.Write("<Font ss:FontName=""Calibri"" x:Family=""Swiss"" ss:Color=""#000000""/>" & vbNewLine & "<Interior ss:Color=""#FFFFFF"" ss:Pattern=""Solid""/>" & vbNewLine & "</Style>"&vbNewLine)
XmlFile.Write("</Styles>"&vbNewLine)

XmlFile.Write("<Worksheet ss:Name=""Sheet1"">"&vbNewLine)
XmlFile.Write("<Table ss:ExpandedColumnCount=""12"" ss:ExpandedRowCount=""8"" x:FullColumns=""1"" x:FullRows=""1"" ss:DefaultColumnWidth=""54"" ss:DefaultRowHeight=""14.25"">"&vbNewLine)
XmlFile.Write("<Column ss:AutoFitWidth=""0"" ss:Width=""75""/>" & vbNewLine & "<Column ss:AutoFitWidth=""0"" ss:Width=""63.75"" ss:Span=""1""/>" & vbNewLine & "<Column ss:Index=""4"" ss:AutoFitWidth=""0"" ss:Width=""72""/>")
XmlFile.Write("<Column ss:AutoFitWidth=""0"" ss:Width=""63.75"" ss:Span=""7""/>"&vbNewLine)



'Open the excel file from specified path
Set wb = oExcel.Workbooks.Open(strPath)
Set ws = wb.Sheets(1)

Row_count = ws.UsedRange.Rows.Count 'Given the used rows count of sheet'
Column_count = ws.UsedRange.Columns.Count 	'Given the used column count of sheet'

For i = 1 To Row_count
	if i = 1 Then
		XmlFile.Write("<Row ss:Height=""25.5"">"&vbNewLine)
	Else
		XmlFile.Write("<Row ss:Height=""18.75"">"&vbNewLine)
	End if

	For j = 1 To Column_count
		if i = 1 Then
			XmlFile.Write("<Cell ss:StyleID=""s69""><Data ss:Type=""String"">"&ws.cells(i,j).value&"</Data></Cell>"&vbNewLine)
		Else
			XmlFile.Write("<Cell ss:StyleID=""s72""><Data ss:Type=""String"">"&ws.cells(i,j).value&"</Data></Cell>"&vbNewLine)
		End if
	Next

	XmlFile.Write("</Row>")
Next

XmlFile.Write("</Table>"&vbNewLine)
XmlFile.Write("<WorksheetOptions xmlns=""urn:schemas-microsoft-com:office:excel"">"&vbNewLine)
XmlFile.Write("<PageSetup>" & vbNewLine & "<Header x:Margin=""0"" x:Data=""&amp;C&amp;A""/>" & vbNewLine & "<Footer x:Margin=""0"" x:Data=""&amp;CPage &amp;P""/>"&vbNewLine)
XmlFile.Write("<PageMargins x:Bottom=""0.39370078740157483"" x:Left=""0"" x:Right=""0"" x:Top=""0.39370078740157483""/>" & vbNewLine & "</PageSetup>"&vbNewLine)
XmlFile.Write("<Selected/>"&vbNewLine)
XmlFile.Write("<ProtectObjects>False</ProtectObjects>"&vbNewLine)
XmlFile.Write("<ProtectScenarios>False</ProtectScenarios>"&vbNewLine)
XmlFile.Write("</WorksheetOptions>"&vbNewLine)
XmlFile.Write("<ConditionalFormatting xmlns=""urn:schemas-microsoft-com:office:excel"">"&vbNewLine)
XmlFile.Write("<Range>R1C1:R2C2,R1C6,R1C8,R1C10,R1C12,R3C2:R4C2,R5C1:R5C2,R6C2:R8C2</Range>"&vbNewLine)
XmlFile.Write("<Condition>"&vbNewLine)
XmlFile.Write("<Qualifier>Equal</Qualifier>"&vbNewLine)
XmlFile.Write("<Value1>&quot;TBD&quot;</Value1>"&vbNewLine)
XmlFile.Write("<Format Style='color:#C00000;font-weight:700'/>"&vbNewLine)
XmlFile.Write("</Condition>"&vbNewLine)
XmlFile.Write("</ConditionalFormatting>"&vbNewLine)
XmlFile.Write("</Worksheet>"&vbNewLine)
XmlFile.Write("</Workbook>"&vbNewLine)

wb.Close		'Close the opened workbook'
oExcel.Quit		'Quit excel application object and realease excel object'
' MsgBox("Done")