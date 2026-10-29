// Create a function query named fnLoadCsv. Each dataset has its own strict schema.
(fileName as text, schema as list, primaryKey as nullable text) as table =>
let
    Source = Csv.Document(File.Contents(pFolder & "\" & fileName),
        [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
    Headers = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    Expected = List.Transform(schema, each _{0}),
    Checked = if Table.ColumnNames(Headers) = Expected then Headers
        else error Error.Record("Schema mismatch", fileName, [Expected=Expected, Actual=Table.ColumnNames(Headers)]),
    // CSV booleans are PostgreSQL t/f; empty CSV fields represent null.
    Nulls = Table.ReplaceValue(Checked, "", null, Replacer.ReplaceValue, Expected),
    Booleans = List.Select(schema, each _{1} = type logical),
    Prepared = Table.TransformColumns(Nulls, List.Transform(Booleans, each
        {_{0}, (v) => if v = null then null else if v = "t" then true else if v = "f" then false else error "Invalid boolean", type logical})),
    Typed = Table.TransformColumnTypes(Prepared, schema, "en-US"),
    Clean = if Table.RowCount(Table.SelectRowsWithErrors(Typed)) = 0 then Typed else error "Invalid source type",
    ValidKey = if primaryKey = null then true else
        List.NonNullCount(Table.Column(Clean, primaryKey)) = Table.RowCount(Clean)
        and Table.RowCount(Table.Distinct(Table.SelectColumns(Clean, {primaryKey}))) = Table.RowCount(Clean),
    Result = if ValidKey then Clean else error Error.Record("Invalid primary key", fileName, primaryKey)
in
    Result
