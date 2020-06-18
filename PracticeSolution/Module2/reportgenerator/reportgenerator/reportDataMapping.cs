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

namespace reportgenerator
{
    public class reportDataObject
    {
        public Dictionary<string, userDetailsObject> people { get; set; }
        public tasksObject tasks { get; set; }
        public IList<bucketObject> buckets { get; set; }
        public string planId { get; set; }
        public string bucketId { get; set; }
        public string planTitle { get; set; }
        public string graphToken { get; set; }
        public userDetailsObject userDetails { get; set; }
        public string snappedImageUrl { get; set; }
        public string imageUrl { get; set; }
        // TODO: add group id to this collection. Sent in json from focus frontend.
        // TODO: add channel name to this collection. Sent in json from focus frontend.
    }
    public class userDetailsObject
    {
        [JsonPropertyName("@odata.context")]
        public string odataContext { get; set; }
        public IList<String> businessPhone { get; set; }
        public string displayName { get; set; }
        public string givenName { get; set; }
        public string jobTitle { get; set; }
        public string mail { get; set; }
        public string mobilePhone { get; set; }
        public string officeLocation { get; set; }
        public string preferredLanguage { get; set; }
        public string surname { get; set; }
        public string userPrincipalName { get; set; }
        public string id { get; set; }
    }
    public class tasksObject
    {
        [JsonPropertyName("@odata.context")]
        public string odataContext { get; set; }
        [JsonPropertyName("@odata.count")]
        public int dataCount { get; set; }
        public IList<taskObject> value { get; set; }

    }
    public class taskObject
    {
        [JsonPropertyName("@odata.etag")]
        public string odataEtag { get; set; }

        public string planId { get; set; }      //: "0wnpDRazhkmmFqbDTGyrj2UACGYj",
        public string bucketId { get; set; }      //: "bN2qlVl-f0Kxo8IzpYxZp2UAMvGr",
        public string title { get; set; }      //: "Painting broken",
        public string orderHint { get; set; }      //: "8586312100425054321",
        public string assigneePriority { get; set; }      //: "8586312100425054321",
        public int percentComplete { get; set; }     //: 0,
        public string startDateTime { get; set; }      //: null,
        public string createdDateTime { get; set; }      //: "2019-10-06T21:20:42.9721486Z",
        public string dueDateTime { get; set; }      //: "2019-10-07T07:00:00Z",
        public bool hasDescription { get; set; }      //: true,
        public string previewType { get; set; }      //: "noPreview",
        public string completedDateTime { get; set; }      //: null,
        public string completedBy { get; set; }      //: null,
        public int referenceCount { get; set; }     //: 0,
        public int checklistItemCount { get; set; }     //: 0,
        public int activeChecklistItemCount { get; set; }     //: 0,
        public string conversationThreadId { get; set; }      //: null,
        public string id { get; set; }      //: "KpZ1JQ7pBUK9_Vc5vkNaOGUALtPI",
        public createdByObject createdBy { get; set; }
        public appliedCategoriesObject appliedCategories { get; set; }

        public Dictionary<string, assignmentObject> assignments { get; set; }
    }
    public class assignmentObject
    {
        [JsonPropertyName("@odata.type")]
        public string odataType { get; set; }
        public string assignmentDateTime { get; set; }
        public string orderHint { get; set; }
        public userSimpleObject assignedBy { get; set; }
    }

    public class createdByObject
    {
        public userSimpleObject user { get; set; }
    }

    public class userSimpleObject
    {
        public string displayname { get; set; }
        public string id { get; set; }
    }

    public class appliedCategoriesObject
    {

    }

    public class bucketObject
    {
        [JsonPropertyName("@odata.etag")]
        public string odataEtag { get; set; }
        public string name { get; set; }
        public string planId { get; set; }
        public string orderHint { get; set; }
        public string id { get; set; }
    }
    public class WeatherForecast
    {
        public DateTimeOffset Date { get; set; }
        public int TemperatureC { get; set; }
        public string Summary { get; set; }
        public IList<DateTimeOffset> DatesAvailable { get; set; }
        public Dictionary<string, HighLowTemperatures> TemperatureRanges { get; set; }
        public string[] SummaryWords { get; set; }
    }

    public class HighLowTemperatures
    {
        public Temperature High { get; set; }
        public Temperature Low { get; set; }
    }

    public class Temperature
    {
        public int DegreesCelsius { get; set; }
    }
}
