/* =========================================================================================
   POWER BI POWER QUERY (M) SCRIPT - SUPERSTORE SALES DASHBOARD
   =========================================================================================
   Instructions:
   1. In Power BI Desktop, click "Transform Data" to open Power Query Editor.
   2. Go to "Advanced Editor" or create New Blank Queries for each table defined below.
   3. Update #"DatasetPath" if your file is in a custom folder.
   ========================================================================================= */

// -----------------------------------------------------------------------------------------
// PARAMETER: File Path to Sample_Superstore.csv
// -----------------------------------------------------------------------------------------
let
    DatasetPath = "C:\Users\aksha\Downloads\power bi\Sample_Superstore.csv" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true]
in
    DatasetPath


// -----------------------------------------------------------------------------------------
// QUERY 1: Dim_Date (Dynamic Enterprise Calendar Table)
// -----------------------------------------------------------------------------------------
let
    StartDate = #date(2021, 1, 1),
    EndDate = #date(2024, 12, 31),
    TotalDays = Duration.Days(Duration.From(EndDate - StartDate)) + 1,
    DatesList = List.Dates(StartDate, TotalDays, #duration(1, 0, 0, 0)),
    TableFromList = Table.FromList(DatesList, Splitter.SplitByNothing(), {"Date"}, null, ExtraValues.Error),
    SetDateType = Table.TransformColumnTypes(TableFromList, {{"Date", type date}}),
    
    // Add Calendar Attributes
    AddYear = Table.AddColumn(SetDateType, "Year", each Date.Year([Date]), Int64.Type),
    AddQuarterNo = Table.AddColumn(AddYear, "Quarter Number", each Date.QuarterOfYear([Date]), Int64.Type),
    AddQuarter = Table.AddColumn(AddQuarterNo, "Quarter", each "Q" & Text.From([Quarter Number]), type text),
    AddYearQuarter = Table.AddColumn(AddQuarter, "Year-Quarter", each Text.From([Year]) & "-Q" & Text.From([Quarter Number]), type text),
    AddMonthNo = Table.AddColumn(AddYearQuarter, "Month Number", each Date.Month([Date]), Int64.Type),
    AddMonthName = Table.AddColumn(AddMonthNo, "Month Name", each Date.MonthName([Date]), type text),
    AddMonthShort = Table.AddColumn(AddMonthName, "Month Short", each Text.Start([Month Name], 3), type text),
    AddYearMonth = Table.AddColumn(AddMonthShort, "Year-Month", each Text.From([Year]) & "-" & Text.PadStart(Text.From([Month Number]), 2, "0"), type text),
    AddDayOfMonth = Table.AddColumn(AddYearMonth, "Day of Month", each Date.Day([Date]), Int64.Type),
    AddDayOfWeekNo = Table.AddColumn(AddDayOfMonth, "Day of Week Number", each Date.DayOfWeek([Date], Day.Monday) + 1, Int64.Type),
    AddDayOfWeekName = Table.AddColumn(AddDayOfWeekNo, "Day of Week", each Date.DayOfWeekName([Date]), type text),
    AddIsWeekend = Table.AddColumn(AddDayOfWeekName, "Is Weekend", each if [Day of Week Number] >= 6 then "Weekend" else "Weekday", type text),
    AddFiscalYear = Table.AddColumn(AddIsWeekend, "Fiscal Year", each if [Month Number] >= 7 then "FY" & Text.End(Text.From([Year] + 1), 2) else "FY" & Text.End(Text.From([Year]), 2), type text)
in
    AddFiscalYear


// -----------------------------------------------------------------------------------------
// QUERY 2: Staging_Superstore (Raw Data Cleaning & Type Casting)
// -----------------------------------------------------------------------------------------
let
    Source = Csv.Document(File.Contents("C:\Users\aksha\Downloads\power bi\Sample_Superstore.csv"), [Delimiter=",", Columns=17, Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
    PromoteHeaders = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    
    // Remove blank rows and exact duplicates
    RemoveBlanks = Table.SelectRows(PromoteHeaders, each [Order ID] <> null and [Order ID] <> ""),
    RemoveDuplicates = Table.Distinct(RemoveBlanks),
    
    // Proper Type Casting
    TypedTable = Table.TransformColumnTypes(RemoveDuplicates, {
        {"Order ID", type text},
        {"Order Date", type date},
        {"Ship Date", type date},
        {"Customer ID", type text},
        {"Customer Name", type text},
        {"Segment", type text},
        {"Country", type text},
        {"City", type text},
        {"State", type text},
        {"Region", type text},
        {"Category", type text},
        {"Sub-Category", type text},
        {"Product Name", type text},
        {"Sales", type number},
        {"Quantity", Int64.Type},
        {"Discount", type number},
        {"Profit", type number}
    }),

    // Handle any missing values
    CleanNulls = Table.ReplaceValue(TypedTable, null, 0, Replacer.ReplaceValue, {"Sales", "Quantity", "Discount", "Profit"})
in
    CleanNulls


// -----------------------------------------------------------------------------------------
// QUERY 3: Dim_Customer (Star Schema Dimension)
// -----------------------------------------------------------------------------------------
let
    Source = Staging_Superstore,
    SelectColumns = Table.SelectColumns(Source, {"Customer ID", "Customer Name", "Segment"}),
    RemoveDuplicates = Table.Distinct(SelectColumns, {"Customer ID"})
in
    RemoveDuplicates


// -----------------------------------------------------------------------------------------
// QUERY 4: Dim_Product (Star Schema Dimension)
// -----------------------------------------------------------------------------------------
let
    Source = Staging_Superstore,
    SelectColumns = Table.SelectColumns(Source, {"Product Name", "Category", "Sub-Category"}),
    RemoveDuplicates = Table.Distinct(SelectColumns, {"Product Name"}),
    AddProductID = Table.AddIndexColumn(RemoveDuplicates, "ProductKey", 1, 1, Int64.Type)
in
    AddProductID


// -----------------------------------------------------------------------------------------
// QUERY 5: Dim_Geography (Star Schema Dimension)
// -----------------------------------------------------------------------------------------
let
    Source = Staging_Superstore,
    SelectColumns = Table.SelectColumns(Source, {"Country", "State", "City", "Region"}),
    RemoveDuplicates = Table.Distinct(SelectColumns, {"Country", "State", "City"}),
    AddGeoKey = Table.AddIndexColumn(RemoveDuplicates, "GeoKey", 1, 1, Int64.Type)
in
    AddGeoKey


// -----------------------------------------------------------------------------------------
// QUERY 6: Fact_Sales (Star Schema Fact Table)
// -----------------------------------------------------------------------------------------
let
    Source = Staging_Superstore,
    
    // Merge Geography Key
    MergeGeo = Table.NestedJoin(Source, {"Country", "State", "City"}, Dim_Geography, {"Country", "State", "City"}, "GeoDim", JoinKind.LeftOuter),
    ExpandGeo = Table.ExpandTableColumn(MergeGeo, "GeoDim", {"GeoKey"}, {"GeoKey"}),
    
    // Merge Product Key
    MergeProd = Table.NestedJoin(ExpandGeo, {"Product Name"}, Dim_Product, {"Product Name"}, "ProdDim", JoinKind.LeftOuter),
    ExpandProd = Table.ExpandTableColumn(MergeProd, "ProdDim", {"ProductKey"}, {"ProductKey"}),
    
    // Select Fact Attributes & Foreign Keys
    KeepColumns = Table.SelectColumns(ExpandProd, {
        "Order ID",
        "Order Date",
        "Ship Date",
        "Customer ID",
        "ProductKey",
        "GeoKey",
        "Sales",
        "Quantity",
        "Discount",
        "Profit"
    })
in
    KeepColumns
