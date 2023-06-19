<h1>Payroll Engine Document</h1>

Document library for the Payroll Engine, including the data transfomration for exchange and report.

## Report
Reports based on [FastReport Open Source](https://github.com/FastReports/FastReport):
- MIT license
- Report template format: *.frx (XML)
- Visual designer: [FastReport Designer Community Edition](https://github.com/FastReports/FastReport/releases/latest)

## FastReport Report Template
Use *TableDataSource* as data source for a data set table

```xml
1: <?xml version="1.0" encoding="utf-8"?>
2: <Report ScriptLanguage="CSharp" ReportInfo.Created="06/20/2009 22:40:42" ReportInfo.Modified="04/15/2023 09:13:19" ReportInfo.CreatorVersion="2023.2.0.0">
3:   <Dictionary>
4:     <TableDataSource Name="Users" ReferenceName="Data.Users" DataType="System.Int32" Enabled="true">
5:       <Column Name="Id" DataType="System.Int32"/>
6:       <Column Name="Status" DataType="System.String"/>
7:       <Column Name="Created" DataType="System.DateTime"/>
8:       <Column Name="Updated" DataType="System.DateTime"/>
9:       <Column Name="Identifier" DataType="System.String"/>
10:       <Column Name="FirstName" DataType="System.String"/>
11:       <Column Name="LastName" DataType="System.String"/>
12:       <Column Name="Culture" DataType="System.String"/>
13:       <Column Name="Language" DataType="System.String"/>
14:       <Column Name="Attributes" DataType="System.String"/>
15:     </TableDataSource>
16:   </Dictionary>
17: </Report>
```

## MS Build
Supported runtime environment variables:
- *PayrollEnginePackageDir* - the NuGet package target direcotry (optional)
