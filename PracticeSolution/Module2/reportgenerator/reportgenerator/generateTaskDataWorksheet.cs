using System;
using DocumentFormat.OpenXml.Packaging;
using Wetp = DocumentFormat.OpenXml.Office2013.WebExtentionPane;
using DocumentFormat.OpenXml;
using We = DocumentFormat.OpenXml.Office2013.WebExtension;
using A = DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Spreadsheet;
using X15ac = DocumentFormat.OpenXml.Office2013.ExcelAc;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using C14 = DocumentFormat.OpenXml.Office2010.Drawing.Charts;
using Cs = DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;
using C15 = DocumentFormat.OpenXml.Office2013.Drawing.Chart;
using X15 = DocumentFormat.OpenXml.Office2013.Excel;
using Thm15 = DocumentFormat.OpenXml.Office2013.Theme;
using Op = DocumentFormat.OpenXml.CustomProperties;
using Vt = DocumentFormat.OpenXml.VariantTypes;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Collections.Generic;
using System.Diagnostics;

namespace reportgenerator
{
    partial class Generator
    {
        private List<string> progress = new List<string>() { "Not started", "In Progress", "Completed" };
        private List<string> rowArry = new List<string>() { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J" };
        private List<string> headers = new List<string>()
                { "Task", "Owner", "Email", "Bucket", "Progress", "Due Date", "Completed Date", "Completed By", "Created Date", "Task Id" };
        private reportDataObject reportData;
        private void generateTaskDataWorksheet(WorksheetPart worksheetPart4)
        {
            foreach (KeyValuePair<string, userDetailsObject> person in reportData.people)
            {
                string displayName = person.Value.displayName;
                string userid = person.Key;
                Console.WriteLine(displayName);
                Console.WriteLine(userid);
            }
            Worksheet worksheet4 = new Worksheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            worksheet4.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            worksheet4.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            worksheet4.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
            //worksheet4.AddNamespaceDeclaration("xr", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision");
            //worksheet4.AddNamespaceDeclaration("xr2", "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2");
            //worksheet4.AddNamespaceDeclaration("xr3", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3");
            //worksheet4.SetAttribute(new OpenXmlAttribute("xr", "uid", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision", genStrGuid()));
            SheetDimension sheetDimension4 = new SheetDimension() { Reference = "A1:J" + (reportData.tasks.value.Count + 1).ToString() };

            SheetViews sheetViews4 = new SheetViews();

            SheetView sheetView4 = new SheetView() { WorkbookViewId = (UInt32Value)0U };
            Selection selection2 = new Selection() { ActiveCell = "G23", SequenceOfReferences = new ListValue<StringValue>() { InnerText = "G23" } };

            sheetView4.Append(selection2);

            sheetViews4.Append(sheetView4);
            SheetFormatProperties sheetFormatProperties4 = new SheetFormatProperties() { DefaultRowHeight = 15D, DyDescent = 0.25D };

            Columns columns4 = new Columns();
            Column column18 = new Column() { Min = (UInt32Value)1U, Max = (UInt32Value)1U, Width = 26.85546875D, BestFit = true, CustomWidth = true };
            Column column19 = new Column() { Min = (UInt32Value)2U, Max = (UInt32Value)2U, Width = 12.85546875D, BestFit = true, CustomWidth = true };
            Column column20 = new Column() { Min = (UInt32Value)3U, Max = (UInt32Value)3U, Width = 34.42578125D, BestFit = true, CustomWidth = true };
            Column column21 = new Column() { Min = (UInt32Value)4U, Max = (UInt32Value)4U, Width = 21.7109375D, BestFit = true, CustomWidth = true };
            Column column22 = new Column() { Min = (UInt32Value)5U, Max = (UInt32Value)5U, Width = 11.140625D, BestFit = true, CustomWidth = true };
            Column column23 = new Column() { Min = (UInt32Value)6U, Max = (UInt32Value)6U, Width = 15.7109375D, BestFit = true, CustomWidth = true };
            Column column24 = new Column() { Min = (UInt32Value)7U, Max = (UInt32Value)7U, Width = 18D, BestFit = true, CustomWidth = true };
            Column column25 = new Column() { Min = (UInt32Value)8U, Max = (UInt32Value)8U, Width = 16D, BestFit = true, CustomWidth = true };
            Column column26 = new Column() { Min = (UInt32Value)9U, Max = (UInt32Value)9U, Width = 15.5703125D, BestFit = true, CustomWidth = true };
            Column column27 = new Column() { Min = (UInt32Value)10U, Max = (UInt32Value)10U, Width = 35.85546875D, BestFit = true, CustomWidth = true };

            columns4.Append(column18);
            columns4.Append(column19);
            columns4.Append(column20);
            columns4.Append(column21);
            columns4.Append(column22);
            columns4.Append(column23);
            columns4.Append(column24);
            columns4.Append(column25);
            columns4.Append(column26);
            columns4.Append(column27);

            SheetData sheetData4 = new SheetData();

            Row headerRow = new Row() { RowIndex = (UInt32Value)1U, Spans = new ListValue<StringValue>() { InnerText = "1:10" }, DyDescent = 0.25D };

            for (int i = 1; i <= headers.Count; i++)
            {
                Console.WriteLine("Header value: {0} ", headers[i - 1]);
                Console.WriteLine("Header cell ref: {0} ", rowArry[i - 1].ToString() + "1");


                Cell cellHeader = new Cell() { CellReference = rowArry[i - 1].ToString() + (1).ToString(), DataType = CellValues.InlineString };
                InlineString cellHeaderValue = new InlineString();

                cellHeaderValue.Text = new Text(headers[i - 1]);

                cellHeader.Append(cellHeaderValue);

                headerRow.Append(cellHeader);
            }

            sheetData4.Append(headerRow);

            for (int tIndex = 1; tIndex <= reportData.tasks.value.Count; tIndex++)
            {

                int iCol = 0;
                uint iRow = (uint)tIndex + 1;

                Row dataRow = new Row() { RowIndex = new UInt32Value((uint)iRow), Spans = new ListValue<StringValue>() { InnerText = "1:10" }, DyDescent = 0.25D };

                taskObject task = reportData.tasks.value[tIndex - 1];

                Cell cellData = new Cell() { CellReference = rowArry[iCol++].ToString() + (iRow).ToString(), DataType = CellValues.InlineString };
                //CellValue cellValue = new CellValue
                InlineString cellValue = new InlineString(new Text(task.title));

                cellData.Append(cellValue);
                dataRow.Append(cellData);

                cellData = new Cell() { CellReference = rowArry[iCol++].ToString() + (iRow).ToString(), DataType = CellValues.InlineString };
                cellValue = new InlineString();
                string userid = "";
                foreach (KeyValuePair<string, assignmentObject> pair in task.assignments)
                {
                    userid = pair.Key;
                    cellValue.Text = new Text(reportData.people[pair.Key].displayName);
                    break;
                }

                cellData.Append(cellValue);
                dataRow.Append(cellData);

                string email = " ";
                if (userid != "")
                    email = reportData.people[userid].mail;
                cellData = new Cell() { CellReference = rowArry[iCol++].ToString() + (iRow).ToString(), DataType = CellValues.InlineString };
                cellValue = new InlineString(new Text(email));

                cellData.Append(cellValue);
                dataRow.Append(cellData);

                string bucketName = "";
                foreach (bucketObject b in reportData.buckets)
                {
                    if (b.id == task.bucketId)
                        bucketName = b.name;
                }

                cellData = new Cell() { CellReference = rowArry[iCol++].ToString() + (iRow).ToString(), DataType = CellValues.InlineString };
                cellValue = new InlineString(new Text(bucketName));

                cellData.Append(cellValue);
                dataRow.Append(cellData);

                cellData = new Cell() { CellReference = rowArry[iCol++].ToString() + (iRow).ToString(), DataType = CellValues.InlineString };
                cellValue = new InlineString();
                if (task.percentComplete == 0)
                    cellValue.Text = new Text("Not started");
                else
                    cellValue.Text = new Text(progress[100 / task.percentComplete]);

                cellData.Append(cellValue);
                dataRow.Append(cellData);

                cellData = new Cell() { CellReference = rowArry[iCol++].ToString() + (iRow).ToString(), DataType = CellValues.InlineString };
                cellValue = new InlineString(new Text(getStringFromDate(task.dueDateTime)));

                cellData.Append(cellValue);
                dataRow.Append(cellData);

                cellData = new Cell() { CellReference = rowArry[iCol++].ToString() + (iRow).ToString(), DataType = CellValues.InlineString };
                cellValue = new InlineString(new Text(getStringFromDate(task.completedDateTime)));

                cellData.Append(cellValue);
                dataRow.Append(cellData);

                string completedByUser = "";
                cellData = new Cell() { CellReference = rowArry[iCol++].ToString() + (iRow).ToString(), DataType = CellValues.InlineString };
                if (task.completedBy != null)
                    completedByUser = reportData.people[task.completedBy].displayName;
                cellValue = new InlineString(new Text(completedByUser));

                cellData.Append(cellValue);
                dataRow.Append(cellData);

                cellData = new Cell() { CellReference = rowArry[iCol++].ToString() + (iRow).ToString(), DataType = CellValues.InlineString };
                cellValue = new InlineString(new Text(getStringFromDate(task.createdDateTime)));

                cellData.Append(cellValue);
                dataRow.Append(cellData);

                cellData = new Cell() { CellReference = rowArry[iCol++].ToString() + (iRow).ToString(), DataType = CellValues.InlineString };
                cellValue = new InlineString(new Text(task.id));

                cellData.Append(cellValue);
                dataRow.Append(cellData);
                sheetData4.Append(dataRow);
                //Debugger.Break();
            }
            PageMargins pageMargins6 = new PageMargins() { Left = 0.7D, Right = 0.7D, Top = 0.75D, Bottom = 0.75D, Header = 0.3D, Footer = 0.3D };
            PageSetup pageSetup6 = new PageSetup() { Orientation = OrientationValues.Portrait, Id = "rId1" };

            TableParts tableParts1 = new TableParts() { Count = (UInt32Value)1U };
            TablePart tablePart1 = new TablePart() { Id = "rId1" };

            tableParts1.Append(tablePart1);

            worksheet4.Append(sheetDimension4);
            worksheet4.Append(sheetViews4);
            worksheet4.Append(sheetFormatProperties4);
            worksheet4.Append(columns4);
            worksheet4.Append(sheetData4);
            worksheet4.Append(pageMargins6);
            worksheet4.Append(pageSetup6);
            worksheet4.Append(tableParts1);

            worksheetPart4.Worksheet = worksheet4;
        }


        private void generateTaskTableContent(TableDefinitionPart tableDefinitionPart)
        {
            Table table = new Table() { Id = (UInt32Value)4U, Name = "WorkItemsTable", DisplayName = "WorkItemsTable", Reference = "A1:J" + (reportData.tasks.value.Count+1).ToString(), TotalsRowShown = false };
            table.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            //table1.AddNamespaceDeclaration("xr", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision");
            //table1.AddNamespaceDeclaration("xr3", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3");
            //table1.SetAttribute(new OpenXmlAttribute("xr", "uid", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision", "{5EFA48B8-3293-4748-AA3B-548E4EF8F199}"));

            AutoFilter autoFilter1 = new AutoFilter() { Reference = "A1:J" + (reportData.tasks.value.Count + 1).ToString() };
            //autoFilter1.SetAttribute(new OpenXmlAttribute("xr", "uid", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision", "{EDA87F4E-B12B-41C9-8B20-4850D8E9E195}"));

            TableColumns tableColumns = new TableColumns() { Count = (uint)headers.Count };

            for (int i = 1; i <= headers.Count; i++)
            {

                TableColumn tableColumn = new TableColumn() { Id = (uint)i, Name = headers[i - 1] };
                //tableColumn.SetAttribute(new OpenXmlAttribute("xr3", "uid", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3", "{D2D20073-D29B-4500-9950-65429DC6D749}"));
                Console.WriteLine("Header value: {0} ", headers[i - 1]);
                Console.WriteLine("Header cell ref: {0} ", rowArry[i - 1].ToString() + "1");

                tableColumns.Append(tableColumn);
            }

            TableStyleInfo tableStyleInfo = new TableStyleInfo() { Name = "TableStyleMedium2", ShowFirstColumn = false, ShowLastColumn = false, ShowRowStripes = true, ShowColumnStripes = false };

            table.Append(autoFilter1);
            table.Append(tableColumns);
            table.Append(tableStyleInfo);

            tableDefinitionPart.Table = table;
        }

        public static string genStrGuid()
        {
            return "{" + Guid.NewGuid() + "}";
        }

        public string getStringFromDate(string dateStr)
        {
            DateTime date1970 = new DateTime(1969, 12, 31, 23, 59, 59);
            string retDate = "";
            if (dateStr != null)
            {
                DateTime date = DateTime.Parse(dateStr);
                if (date.Year > date1970.Year)
                    retDate = date.ToString("d");
            }
            return retDate;
        }

        public string testPeopleJsonData = @"{""people"":{""0bc80813-850c-4b00-af5b-f2ca56c8fb70"":{""@odata.context"":""https://graph.microsoft.com/v1.0/$metadata#users/$entity"",""businessPhones"":[""503-985-2078""],""displayName"":""Tom Jebo"",""givenName"":""Tom"",""jobTitle"":""Escalation Engineer"",""mail"":""tomjebo@jebosoft.onmicrosoft.com"",""mobilePhone"":""+1 5039852078"",""officeLocation"":null,""preferredLanguage"":""en-US"",""surname"":""Jebo"",""userPrincipalName"":""tomjebo@jebosoft.onmicrosoft.com"",""id"":""0bc80813-850c-4b00-af5b-f2ca56c8fb70""},""84c93f67-4762-4d52-aae4-cabd195ad45b"":{""@odata.context"":""https://graph.microsoft.com/v1.0/$metadata#users/$entity"",""businessPhones"":[],""displayName"":""Will Gregg"",""givenName"":""Will"",""jobTitle"":""Manager"",""mail"":""grjoh@jebosoft.onmicrosoft.com"",""mobilePhone"":""+1 702-202-9610"",""officeLocation"":null,""preferredLanguage"":""en-US"",""surname"":""Gregg"",""userPrincipalName"":""grjoh@jebosoft.onmicrosoft.com"",""id"":""84c93f67-4762-4d52-aae4-cabd195ad45b""},""48d31887-5fad-4d73-a9f5-3c356e68a038"" : {""@odata.context"":""https://graph.microsoft.com/v1.0/$metadata#users/$entity"", ""businessPhones"": [ ""+1 412 555 0109"" ], ""displayName"": ""Megan Bowen"", ""givenName"": ""Megan"", ""jobTitle"": ""Auditor"", ""mail"": ""MeganB@M365x214355.onmicrosoft.com"", ""mobilePhone"": null, ""officeLocation"": ""12/1110"", ""preferredLanguage"": ""en-US"", ""surname"": ""Bowen"", ""userPrincipalName"": ""MeganB@M365x214355.onmicrosoft.com"", ""id"": ""48d31887-5fad-4d73-a9f5-3c356e68a038"" }}";
        public string testBucketJsonData = @"""buckets"": [{ ""@odata.etag"": ""W/\""JzEtQnVja2V0QEBAQEBAQEBAQEBAQEBAYCc=\"""", ""name"": ""To do"", ""planId"": ""CONGZUWfGUu4msTgNP66e2UAAySi"", ""orderHint"": ""8586962670142056623P]"", ""id"": ""XfAo8wVDl0-Tw6tlzmALKmUAHMrn"" } ]";
        public string testMiscJsonData3 = @"""planId"":""0wnpDRazhkmmFqbDTGyrj2UACGYj"",""bucketId"":""bN2qlVl-f0Kxo8IzpYxZp2UAMvGr"",""planTitle"":""Focus Plan"",""graphToken"":""eyJ0eXAiOiJKV1QiLCJub25jZSI6InZrYXVmN3k4aTV6TW5aOVowR3dMcG9pS2xwNFo1YlYycDVzY3hSM3BYSmMiLCJhbGciOiJSUzI1NiIsIng1dCI6ImFQY3R3X29kdlJPb0VOZzNWb09sSWgydGlFcyIsImtpZCI6ImFQY3R3X29kdlJPb0VOZzNWb09sSWgydGlFcyJ9.eyJhdWQiOiJodHRwczovL2dyYXBoLm1pY3Jvc29mdC5jb20iLCJpc3MiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC8wZjZjZjYyOC1kYjU5LTRmZTMtOWYzNi1jYTNiOThkMDYyMTMvIiwiaWF0IjoxNTcwNjc4OTg2LCJuYmYiOjE1NzA2Nzg5ODYsImV4cCI6MTU3MDY4Mjg4NiwiYWNjdCI6MCwiYWNyIjoiMSIsImFpbyI6IkFTUUEyLzhOQUFBQXBHRnR6NW1waFZHanZhS2F1R2FRUkNQRzdkSFZQRGtRZS8rSzgyemlJRXM9IiwiYW1yIjpbInB3ZCJdLCJhcHBfZGlzcGxheW5hbWUiOiJGb2N1c0plYm9zb2Z0IEFwcCIsImFwcGlkIjoiMjZjNTQ0NjgtNzFjOS00YzEzLWEyMGUtZWM2ODRjZmY5OTI4IiwiYXBwaWRhY3IiOiIwIiwiZmFtaWx5X25hbWUiOiJKZWJvIiwiZ2l2ZW5fbmFtZSI6IlRvbSIsImlwYWRkciI6Ijk3LjEyNi42NC4yMDIiLCJuYW1lIjoiVG9tIEplYm8iLCJvaWQiOiIwYmM4MDgxMy04NTBjLTRiMDAtYWY1Yi1mMmNhNTZjOGZiNzAiLCJwbGF0ZiI6IjMiLCJwdWlkIjoiMTAwMzAwMDA5NzU2MTM5QSIsInNjcCI6IkRpcmVjdG9yeS5BY2Nlc3NBc1VzZXIuQWxsIERpcmVjdG9yeS5SZWFkLkFsbCBEaXJlY3RvcnkuUmVhZFdyaXRlLkFsbCBGaWxlcy5SZWFkV3JpdGUuQWxsIEdyb3VwLlJlYWQuQWxsIEdyb3VwLlJlYWRXcml0ZS5BbGwgTWFpbC5TZW5kIG9wZW5pZCBPcmdhbml6YXRpb24uUmVhZC5BbGwgT3JnYW5pemF0aW9uLlJlYWRXcml0ZS5BbGwgUGVvcGxlLlJlYWQgUGVvcGxlLlJlYWQuQWxsIFRhc2tzLlJlYWRXcml0ZSBUYXNrcy5SZWFkV3JpdGUuU2hhcmVkIFVzZXIuUmVhZCBVc2VyLlJlYWQuQWxsIFVzZXIuUmVhZFdyaXRlLkFsbCIsInNpZ25pbl9zdGF0ZSI6WyJrbXNpIl0sInN1YiI6ImVpUzFLMXlPdEswbm1fbW9CWmw4UFhVVHpIZmlyTVM5UnhpYzZLeE1POW8iLCJ0aWQiOiIwZjZjZjYyOC1kYjU5LTRmZTMtOWYzNi1jYTNiOThkMDYyMTMiLCJ1bmlxdWVfbmFtZSI6InRvbWplYm9AamVib3NvZnQub25taWNyb3NvZnQuY29tIiwidXBuIjoidG9tamVib0BqZWJvc29mdC5vbm1pY3Jvc29mdC5jb20iLCJ1dGkiOiJQckFnTnNSaGRVeWtnRUpiNmxjS0FBIiwidmVyIjoiMS4wIiwieG1zX3RjZHQiOjE0NjA0OTkwNDd9.oh3sOoxByJtiI7xGt3q51wMkEVtQZOESyALobIhiWoMmFiH_5sEZhY642JGXLXFPDbiGAxthcZdBzyyzWSL795dNcPWuCELton4V_v11pSKvbbSLWoZZO3PFLg1xwCXUnR4u6TdZONadYNxXIxgALCYaHvTq6zd6XrScp6K5HCsyNig5c3zXU5GezSs12C8ygYbwIw1iJ7cAAajk1KbiGPsnb_DHwfAz2CSUPHjcR8PUPDeC6gI6baAHZhar4SmW_CJv00M6WiahQ7fQoiRh_PZqY_OlxANKy2-3zlilWxIQYYo0Dm8VbiO1TaBkF-1HVzI50mQq-aEak4H5PtfBkA"",""userDetails"":{""@odata.context"":""https://graph.microsoft.com/v1.0/$metadata#users/$entity"",""businessPhones"":[""503-985-2078""],""displayName"":""Tom Jebo"",""givenName"":""Tom"",""jobTitle"":""Escalation Engineer"",""mail"":""tomjebo@jebosoft.onmicrosoft.com"",""mobilePhone"":""+1 5039852078"",""officeLocation"":null,""preferredLanguage"":""en-US"",""surname"":""Jebo"",""userPrincipalName"":""tomjebo@jebosoft.onmicrosoft.com"",""id"":""0bc80813-850c-4b00-af5b-f2ca56c8fb70""},""snappedImageUrl"":null,""imageUrl"":null}";
        //
        // Example data from Microsoft Graph 
        //""tasks"":{""@odata.context"":""https://graph.microsoft.com/v1.0/$metadata#Collection(microsoft.graph.plannerTask)"",""@odata.count"":4,""value"":[{""@odata.etag"":""W/\""JzEtVGFzayAgQEBAQEBAQEBAQEBAQEBASCc=\"""",""planId"":""0wnpDRazhkmmFqbDTGyrj2UACGYj"",""bucketId"":""bN2qlVl-f0Kxo8IzpYxZp2UAMvGr"",""title"":""Painting broken"",""orderHint"":""8586312100425054321"",""assigneePriority"":""8586312100425054321"",""percentComplete"":0,""startDateTime"":null,""createdDateTime"":""2019-10-06T21:20:42.9721486Z"",""dueDateTime"":""2019-10-07T07:00:00Z"",""hasDescription"":true,""previewType"":""noPreview"",""completedDateTime"":null,""completedBy"":null,""referenceCount"":0,""checklistItemCount"":0,""activeChecklistItemCount"":0,""conversationThreadId"":null,""id"":""KpZ1JQ7pBUK9_Vc5vkNaOGUALtPI"",""createdBy"":{""user"":{""displayName"":null,""id"":""0bc80813-850c-4b00-af5b-f2ca56c8fb70""}},""appliedCategories"":{},""assignments"":{""0bc80813-850c-4b00-af5b-f2ca56c8fb70"":{""@odata.type"":""#microsoft.graph.plannerAssignment"",""assignedDateTime"":""2019-10-06T21:20:42.9721486Z"",""orderHint"":""8586312101026304547P-"",""assignedBy"":{""user"":{""displayName"":null,""id"":""0bc80813-850c-4b00-af5b-f2ca56c8fb70""}}}}},{""@odata.etag"":""W/\""JzEtVGFzayAgQEBAQEBAQEBAQEBAQEBAUCc=\"""",""planId"":""0wnpDRazhkmmFqbDTGyrj2UACGYj"",""bucketId"":""bN2qlVl-f0Kxo8IzpYxZp2UAMvGr"",""title"":""Painting problem"",""orderHint"":""8586312656582923755"",""assigneePriority"":""8586312656582923755"",""percentComplete"":50,""startDateTime"":null,""createdDateTime"":""2019-10-06T05:53:47.1852052Z"",""dueDateTime"":""2019-10-09T07:00:00Z"",""hasDescription"":true,""previewType"":""noPreview"",""completedDateTime"":null,""completedBy"":null,""referenceCount"":0,""checklistItemCount"":0,""activeChecklistItemCount"":0,""conversationThreadId"":null,""id"":""AwFR1Ni2mEynSonFD0-3G2UAOfd2"",""createdBy"":{""user"":{""displayName"":null,""id"":""0bc80813-850c-4b00-af5b-f2ca56c8fb70""}},""appliedCategories"":{},""assignments"":{""0bc80813-850c-4b00-af5b-f2ca56c8fb70"":{""@odata.type"":""#microsoft.graph.plannerAssignment"",""assignedDateTime"":""2019-10-06T05:53:47.1852052Z"",""orderHint"":""8586312656788092627PE"",""assignedBy"":{""user"":{""displayName"":null,""id"":""0bc80813-850c-4b00-af5b-f2ca56c8fb70""}}}}},{""@odata.etag"":""W/\""JzEtVGFzayAgQEBAQEBAQEBAQEBAQEBATCc=\"""",""planId"":""0wnpDRazhkmmFqbDTGyrj2UACGYj"",""bucketId"":""bN2qlVl-f0Kxo8IzpYxZp2UAMvGr"",""title"":""New issue"",""orderHint"":""8586313939922466308"",""assigneePriority"":""8586313939922466308"",""percentComplete"":50,""startDateTime"":null,""createdDateTime"":""2019-10-04T18:14:53.2309499Z"",""dueDateTime"":""2019-10-11T07:00:00Z"",""hasDescription"":true,""previewType"":""noPreview"",""completedDateTime"":null,""completedBy"":null,""referenceCount"":0,""checklistItemCount"":0,""activeChecklistItemCount"":0,""conversationThreadId"":null,""id"":""6W6L9MEr6kmCZtYIcYsb2mUAJfOu"",""createdBy"":{""user"":{""displayName"":null,""id"":""0bc80813-850c-4b00-af5b-f2ca56c8fb70""}},""appliedCategories"":{},""assignments"":{""0bc80813-850c-4b00-af5b-f2ca56c8fb70"":{""@odata.type"":""#microsoft.graph.plannerAssignment"",""assignedDateTime"":""2019-10-04T18:14:53.2309499Z"",""orderHint"":""8586313940522935318PW"",""assignedBy"":{""user"":{""displayName"":null,""id"":""0bc80813-850c-4b00-af5b-f2ca56c8fb70""}}}}},{""@odata.etag"":""W/\""JzEtVGFzayAgQEBAQEBAQEBAQEBAQEBATCc=\"""",""planId"":""0wnpDRazhkmmFqbDTGyrj2UACGYj"",""bucketId"":""bN2qlVl-f0Kxo8IzpYxZp2UAMvGr"",""title"":""Jean Reno is cool"",""orderHint"":""8586314605823363866"",""assigneePriority"":"""",""percentComplete"":0,""startDateTime"":null,""createdDateTime"":""2019-10-03T23:45:03.1411941Z"",""dueDateTime"":""2019-10-09T07:00:00Z"",""hasDescription"":true,""previewType"":""noPreview"",""completedDateTime"":null,""completedBy"":null,""referenceCount"":0,""checklistItemCount"":0,""activeChecklistItemCount"":0,""conversationThreadId"":null,""id"":""rabpSxRAbES1Mp6Ysl1c4WUAIxzx"",""createdBy"":{""user"":{""displayName"":null,""id"":""0bc80813-850c-4b00-af5b-f2ca56c8fb70""}},""appliedCategories"":{},""assignments"":{""84c93f67-4762-4d52-aae4-cabd195ad45b"":{""@odata.type"":""#microsoft.graph.plannerAssignment"",""assignedDateTime"":""2019-10-03T23:45:26.4496714Z"",""orderHint"":""8586314606192153915P0"",""assignedBy"":{""user"":{""displayName"":null,""id"":""0bc80813-850c-4b00-af5b-f2ca56c8fb70""}}}}}]},
        //{
        //    "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#Collection(microsoft.graph.plannerTask)",
        //    "@odata.count": 2,
        //    "value": [
        //        {
        //            "@odata.etag": "W/\"JzEtVGFzayAgQEBAQEBAQEBAQEBAQEBAYCc=\"",
        //            "planId": "CONGZUWfGUu4msTgNP66e2UAAySi",
        //            "bucketId": "XfAo8wVDl0-Tw6tlzmALKmUAHMrn",
        //            "title": "Take inventory ",
        //            "orderHint": "8586962669214907176",
        //            "assigneePriority": "",
        //            "percentComplete": 0,
        //            "startDateTime": null,
        //            "createdDateTime": "2017-09-13T21:59:23.9868631Z",
        //            "dueDateTime": null,
        //            "hasDescription": false,
        //            "previewType": "automatic",
        //            "completedDateTime": null,
        //            "completedBy": null,
        //            "referenceCount": 0,
        //            "checklistItemCount": 0,
        //            "activeChecklistItemCount": 0,
        //            "conversationThreadId": null,
        //            "id": "jXZ0-ND1XU2yASUnu3iVkGUACNks",
        //            "createdBy": {
        //                "user": {
        //                    "displayName": null,
        //                    "id": "48d31887-5fad-4d73-a9f5-3c356e68a038"
        //                }
        //            },
        //            "appliedCategories": {},
        //            "assignments": {}
        //        },
        //        {
        //            "@odata.etag": "W/\"JzEtVGFzayAgQEBAQEBAQEBAQEBAQEBAXCc=\"",
        //            "planId": "CONGZUWfGUu4msTgNP66e2UAAySi",
        //            "bucketId": "XfAo8wVDl0-Tw6tlzmALKmUAHMrn",
        //            "title": "Research new trends",
        //            "orderHint": "8586962669337441659",
        //            "assigneePriority": "",
        //            "percentComplete": 0,
        //            "startDateTime": null,
        //            "createdDateTime": "2017-09-13T21:59:11.7334148Z",
        //            "dueDateTime": "2017-08-31T12:00:00Z",
        //            "hasDescription": false,
        //            "previewType": "automatic",
        //            "completedDateTime": null,
        //            "completedBy": null,
        //            "referenceCount": 0,
        //            "checklistItemCount": 0,
        //            "activeChecklistItemCount": 0,
        //            "conversationThreadId": null,
        //            "id": "TKlyl6DaVk6_R4dEG3LGgGUAKQif",
        //            "createdBy": {
        //                "user": {
        //                    "displayName": null,
        //                    "id": "48d31887-5fad-4d73-a9f5-3c356e68a038"
        //                }
        //            },
        //            "appliedCategories": {},
        //            "assignments": {}
        //        }
        //    ]
        //}
    }
}
