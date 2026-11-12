// Power Query (M): paste each query into Power BI Desktop > Transform data > New Source > Blank Query > Advanced Editor.
// First create a Text parameter named DataFolder whose value is the full path of the "data" folder (for example C:\CNG_Safety_Risk_Intelligence\Dashboard\PowerBI_Import_Pack\data\).
// All records are SIMULATED. Columns ending "(calc)" are dropped on load; the DAX file recomputes every score, category and status from the inputs.

// ---------- Hazards
let
    Source = Csv.Document(File.Contents(DataFolder & "Hazards.csv"), [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
    Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    Typed = Table.TransformColumnTypes(Promoted, {{"Initial Likelihood", Int64.Type}, {"Initial Severity", Int64.Type}, {"Residual Likelihood", Int64.Type}, {"Residual Severity", Int64.Type}, {"Assessment Date (simulated)", type date}}),
    Dropped = Table.RemoveColumns(Typed, List.Select(Table.ColumnNames(Typed), each Text.EndsWith(_, "(calc)")), MissingField.Ignore)
in
    Dropped

// ---------- Actions
let
    Source = Csv.Document(File.Contents(DataFolder & "Actions.csv"), [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
    Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    Typed = Table.TransformColumnTypes(Promoted, {{"Date Identified", type date}, {"Target Completion Date", type date}, {"Completion Date", type date}, {"Verification Date", type date}}),
    Blanks = Table.ReplaceValue(Typed, "", null, Replacer.ReplaceValue, {"Evidence of Closure", "Verified By", "Related Record ID"}),
    Dropped = Table.RemoveColumns(Blanks, List.Select(Table.ColumnNames(Blanks), each Text.EndsWith(_, "(calc)")), MissingField.Ignore)
in
    Dropped

// ---------- Records (incident and near-miss log)
let
    Source = Csv.Document(File.Contents(DataFolder & "Records.csv"), [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
    Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    Typed = Table.TransformColumnTypes(Promoted, {{"Date (simulated)", type date}, {"Potential Severity (1-5)", Int64.Type}, {"Actual Severity (0=no harm)", Int64.Type}})
in
    Typed

// ---------- Checklist
let
    Source = Csv.Document(File.Contents(DataFolder & "Checklist.csv"), [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
    Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars=true])
in
    Promoted

// ---------- DimRiskBand
let
    Source = Csv.Document(File.Contents(DataFolder & "DimRiskBand.csv"), [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
    Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    Typed = Table.TransformColumnTypes(Promoted, {{"MinScore", Int64.Type}, {"MaxScore", Int64.Type}, {"SortOrder", Int64.Type}})
in
    Typed

// ---------- DimLikelihood
let
    Source = Csv.Document(File.Contents(DataFolder & "DimLikelihood.csv"), [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
    Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    Typed = Table.TransformColumnTypes(Promoted, {{"Level", Int64.Type}})
in
    Typed

// ---------- DimSeverity
let
    Source = Csv.Document(File.Contents(DataFolder & "DimSeverity.csv"), [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
    Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    Typed = Table.TransformColumnTypes(Promoted, {{"Level", Int64.Type}})
in
    Typed

// ---------- Settings
let
    Source = Csv.Document(File.Contents(DataFolder & "Settings.csv"), [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
    Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    Typed = Table.TransformColumnTypes(Promoted, {{"StatusDate", type date}})
in
    Typed
