using System.Drawing;
using System.IO.Compression;
using OfficeOpenXml;
using OfficeOpenXml.DataValidation;
using OfficeOpenXml.Drawing.Controls;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;

ExcelPackage.License.SetNonCommercialPersonal("Sofoste93");

var root = FindRepositoryRoot();
var outputPath = args.Length > 0
    ? Path.GetFullPath(args[0])
    : Path.Combine(root, "artifacts", "TerminPlanner.xlsm");

Directory.CreateDirectory(Path.GetDirectoryName(outputPath)!);
if (File.Exists(outputPath)) File.Delete(outputPath);

using (var package = new ExcelPackage())
{
    package.Workbook.Properties.Title = "TerminPlanner";
    package.Workbook.Properties.Subject = "A local appointment planner built with Excel and VBA";
    package.Workbook.Properties.Author = "Sofoste93";
    package.Workbook.Properties.Company = "Sofoste";
    package.Workbook.Properties.Comments = "TerminPlanner 1.0.0 · Professional Blue";
    package.Workbook.CalcMode = ExcelCalcMode.Automatic;

    var dashboard = package.Workbook.Worksheets.Add("Dashboard");
    var appointments = package.Workbook.Worksheets.Add("Appointments");
    var lists = package.Workbook.Worksheets.Add("Lists");
    var guide = package.Workbook.Worksheets.Add("Guide");

    BuildLists(lists, package);
    BuildAppointments(appointments);
    BuildDashboard(dashboard, package);
    BuildGuide(guide);
    AddVbaProject(package, root);

    lists.Hidden = eWorkSheetHidden.VeryHidden;
    dashboard.Select("C12");
    package.SaveAs(new FileInfo(outputPath));
}

ValidateWorkbook(outputPath);
Console.WriteLine($"TerminPlanner workbook created: {outputPath}");

static string FindRepositoryRoot()
{
    var current = new DirectoryInfo(Environment.CurrentDirectory);
    while (current is not null)
    {
        if (File.Exists(Path.Combine(current.FullName, "README.md")) &&
            Directory.Exists(Path.Combine(current.FullName, "src", "VBA")))
            return current.FullName;
        current = current.Parent;
    }
    throw new InvalidOperationException("Run the builder inside the TerminPlanner repository.");
}

static void BuildLists(ExcelWorksheet sheet, ExcelPackage package)
{
    string[] categories = ["Meeting", "Consultation", "Personal", "Administration", "Follow-up", "Other"];
    string[] statuses = ["Planned", "Confirmed", "Completed", "Cancelled"];
    string[] reminders = ["None", "5 minutes", "15 minutes", "30 minutes", "1 hour", "1 day"];

    sheet.Cells["A1"].Value = "Categories";
    sheet.Cells["B1"].Value = "Statuses";
    sheet.Cells["C1"].Value = "Reminders";
    for (var index = 0; index < categories.Length; index++) sheet.Cells[index + 2, 1].Value = categories[index];
    for (var index = 0; index < statuses.Length; index++) sheet.Cells[index + 2, 2].Value = statuses[index];
    for (var index = 0; index < reminders.Length; index++) sheet.Cells[index + 2, 3].Value = reminders[index];

    package.Workbook.Names.Add("Categories", sheet.Cells[$"A2:A{categories.Length + 1}"]);
    package.Workbook.Names.Add("Statuses", sheet.Cells[$"B2:B{statuses.Length + 1}"]);
    package.Workbook.Names.Add("Reminders", sheet.Cells[$"C2:C{reminders.Length + 1}"]);
}

static void BuildAppointments(ExcelWorksheet sheet)
{
    sheet.View.ShowGridLines = false;
    sheet.View.FreezePanes(5, 1);
    sheet.Cells["A1:N1"].Merge = true;
    sheet.Cells["A1"].Value = "APPOINTMENT REGISTER";
    StyleTitle(sheet.Cells["A1:N1"]);
    sheet.Cells["A2"].Value = "Select a row before using Delete selected. Filters in the header remain available.";
    sheet.Cells["A2:N2"].Merge = true;
    sheet.Cells["A2"].Style.Font.Color.SetColor(Color.FromArgb(85, 109, 126));

    string[] headers = ["ID", "Date", "Start", "End", "Subject", "Client", "Category", "Status",
        "Location", "Contact", "Notes", "Reminder", "Created", "Updated"];
    for (var column = 0; column < headers.Length; column++) sheet.Cells[4, column + 1].Value = headers[column];
    var table = sheet.Tables.Add(sheet.Cells["A4:N5"], "tblAppointments");
    table.TableStyle = TableStyles.Medium2;
    table.ShowFilter = true;

    sheet.Column(1).Width = 22;
    sheet.Column(2).Width = 13;
    sheet.Column(3).Width = 10;
    sheet.Column(4).Width = 10;
    sheet.Column(5).Width = 28;
    sheet.Column(6).Width = 22;
    sheet.Column(7).Width = 16;
    sheet.Column(8).Width = 15;
    sheet.Column(9).Width = 22;
    sheet.Column(10).Width = 24;
    sheet.Column(11).Width = 35;
    sheet.Column(12).Width = 15;
    sheet.Column(13).Width = 19;
    sheet.Column(14).Width = 19;
    sheet.Cells["B5:B1004"].Style.Numberformat.Format = "yyyy-mm-dd";
    sheet.Cells["C5:D1004"].Style.Numberformat.Format = "hh:mm";
    sheet.Cells["M5:N1004"].Style.Numberformat.Format = "yyyy-mm-dd hh:mm";
    sheet.Cells["A:N"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

    AddButton(sheet, "SearchButton", "SEARCH", "SearchAppointments", 1, 8, 110);
    AddButton(sheet, "ClearSearchButton", "CLEAR SEARCH", "ClearSearch", 1, 10, 120);
    AddButton(sheet, "TodayButton", "GO TO TODAY", "GoToToday", 1, 12, 120);
    AddButton(sheet, "DeleteButton", "DELETE SELECTED", "DeleteSelectedAppointment", 2, 12, 140, danger: true);
}

static void BuildDashboard(ExcelWorksheet sheet, ExcelPackage package)
{
    sheet.View.ShowGridLines = false;
    sheet.View.FreezePanes(3, 1);
    sheet.Column(1).Width = 3;
    for (var column = 2; column <= 14; column++) sheet.Column(column).Width = column is 3 or 6 or 9 ? 15 : 12;
    sheet.Row(1).Height = 30;
    sheet.Row(2).Height = 20;

    sheet.Cells["A1:N2"].Merge = true;
    sheet.Cells["A1"].Value = "TERMINPLANNER  ·  PROFESSIONAL BLUE";
    StyleHero(sheet.Cells["A1:N2"]);

    CreateCard(sheet, "B4:D6", "TODAY", "=COUNTIFS(tblAppointments[Date],TODAY())");
    CreateCard(sheet, "F4:H6", "UPCOMING", "=COUNTIFS(tblAppointments[Date],\">=\"&TODAY(),tblAppointments[Status],\"<>Cancelled\")");
    CreateCard(sheet, "J4:L6", "COMPLETED", "=COUNTIFS(tblAppointments[Status],\"Completed\")");

    sheet.Cells["B8:N8"].Merge = true;
    sheet.Cells["B8"].Value = "NEW APPOINTMENT";
    StyleSection(sheet.Cells["B8:N8"]);

    Label(sheet, "B10", "DATE *");
    Input(sheet, "C10:D10", DateTime.Today, "yyyy-mm-dd");
    Label(sheet, "F10", "START *");
    Input(sheet, "G10", new TimeSpan(9, 0, 0), "hh:mm");
    Label(sheet, "I10", "END *");
    Input(sheet, "J10", new TimeSpan(9, 30, 0), "hh:mm");
    Label(sheet, "B12", "SUBJECT *");
    Input(sheet, "C12:G12", "");
    Label(sheet, "I12", "CATEGORY");
    Input(sheet, "J12:L12", "Meeting");
    Label(sheet, "B14", "CLIENT *");
    Input(sheet, "C14:D14", "");
    Label(sheet, "F14", "CONTACT");
    Input(sheet, "G14:H14", "");
    Label(sheet, "I14", "STATUS");
    Input(sheet, "J14:L14", "Planned");
    Label(sheet, "B16", "LOCATION");
    Input(sheet, "C16:G16", "");
    Label(sheet, "I16", "REMINDER");
    Input(sheet, "J16:L16", "15 minutes");
    Label(sheet, "B18", "NOTES");
    Input(sheet, "C18:L20", "");
    sheet.Cells["C18:L20"].Style.WrapText = true;

    AddListValidation(sheet, "J12", "Categories");
    AddListValidation(sheet, "J14", "Statuses");
    AddListValidation(sheet, "J16", "Reminders");
    var dateValidation = sheet.DataValidations.AddDateTimeValidation("C10");
    dateValidation.Operator = ExcelDataValidationOperator.greaterThanOrEqual;
    dateValidation.Formula.Value = DateTime.Today.AddYears(-1);

    package.Workbook.Names.Add("FormDate", sheet.Cells["C10"]);
    package.Workbook.Names.Add("FormStart", sheet.Cells["G10"]);
    package.Workbook.Names.Add("FormEnd", sheet.Cells["J10"]);
    package.Workbook.Names.Add("FormSubject", sheet.Cells["C12"]);
    package.Workbook.Names.Add("FormCategory", sheet.Cells["J12"]);
    package.Workbook.Names.Add("FormClient", sheet.Cells["C14"]);
    package.Workbook.Names.Add("FormContact", sheet.Cells["G14"]);
    package.Workbook.Names.Add("FormStatus", sheet.Cells["J14"]);
    package.Workbook.Names.Add("FormLocation", sheet.Cells["C16"]);
    package.Workbook.Names.Add("FormReminder", sheet.Cells["J16"]);
    package.Workbook.Names.Add("FormNotes", sheet.Cells["C18"]);

    AddButton(sheet, "SaveButton", "SAVE APPOINTMENT", "AddAppointment", 22, 1, 170);
    AddButton(sheet, "ClearButton", "CLEAR FORM", "ClearForm", 22, 4, 120);
    AddButton(sheet, "RefreshButton", "REFRESH", "RefreshPlanner", 22, 7, 110);
    AddButton(sheet, "ExportButton", "EXPORT UPCOMING", "ExportUpcomingCsv", 22, 9, 160);
    AddButton(sheet, "GuideButton", "HELP", "OpenGuide", 22, 12, 90);

    sheet.Cells["B25:H25"].Merge = true;
    sheet.Cells["B25"].Value = "NEXT APPOINTMENTS";
    StyleSection(sheet.Cells["B25:H25"]);
    string[] upcomingHeaders = ["Date", "Start", "Subject", "Client", "Category", "Status", "Location"];
    for (var column = 0; column < upcomingHeaders.Length; column++)
    {
        sheet.Cells[26, column + 2].Value = upcomingHeaders[column];
        StyleHeader(sheet.Cells[26, column + 2]);
    }
    sheet.Cells["B27:B36"].Style.Numberformat.Format = "yyyy-mm-dd";
    sheet.Cells["C27:C36"].Style.Numberformat.Format = "hh:mm";
    sheet.Cells["B27:H36"].Style.Fill.PatternType = ExcelFillStyle.Solid;
    sheet.Cells["B27:H36"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(245, 249, 252));
    sheet.Cells["B27:H36"].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
    sheet.Cells["B27:H36"].Style.Border.Bottom.Color.SetColor(Color.FromArgb(209, 224, 234));
    sheet.Cells["B4:L36"].Style.Font.Name = "Aptos";
    sheet.Protection.IsProtected = false;
}

static void BuildGuide(ExcelWorksheet sheet)
{
    sheet.View.ShowGridLines = false;
    sheet.Column(1).Width = 4;
    sheet.Column(2).Width = 28;
    sheet.Column(3).Width = 80;
    sheet.Cells["A1:H2"].Merge = true;
    sheet.Cells["A1"].Value = "TERMINPLANNER  ·  QUICK START";
    StyleHero(sheet.Cells["A1:H2"]);

    var steps = new (string Title, string Body)[]
    {
        ("1 · Enable macros", "Open the downloaded .xlsm in desktop Excel, inspect its source if desired, then enable macros for this trusted local file."),
        ("2 · Add an appointment", "Complete the required fields on Dashboard and select SAVE APPOINTMENT. Date, time and required text are validated before storage."),
        ("3 · Work with the register", "Use Appointments to filter, search, review and delete records. Select a cell in a row before choosing DELETE SELECTED."),
        ("4 · Export", "EXPORT UPCOMING creates a UTF-friendly CSV beside the saved workbook. The workbook never uploads appointments."),
        ("5 · Back up", "The workbook is the database. Save it regularly and keep a backup copy in a location you control.")
    };
    var row = 5;
    foreach (var step in steps)
    {
        sheet.Cells[row, 2].Value = step.Title;
        sheet.Cells[row, 2].Style.Font.Bold = true;
        sheet.Cells[row, 2].Style.Font.Color.SetColor(Color.FromArgb(7, 91, 150));
        sheet.Cells[row, 3].Value = step.Body;
        sheet.Cells[row, 3].Style.WrapText = true;
        sheet.Row(row).Height = 42;
        row += 2;
    }
    sheet.Cells[row + 1, 2, row + 1, 3].Merge = true;
    sheet.Cells[row + 1, 2].Value = "PRIVACY · TerminPlanner stores data only inside this workbook and optional local CSV exports. It does not encrypt files or provide multi-user conflict handling.";
    sheet.Cells[row + 1, 2].Style.WrapText = true;
    sheet.Cells[row + 1, 2].Style.Font.Color.SetColor(Color.FromArgb(91, 72, 15));
    sheet.Cells[row + 1, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
    sheet.Cells[row + 1, 2].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 247, 214));
    sheet.Row(row + 1).Height = 52;
}

static void AddVbaProject(ExcelPackage package, string root)
{
    package.Workbook.CreateVBAProject();
    package.Workbook.VbaProject.Name = "TerminPlanner";
    package.Workbook.CodeModule.Code = ReadVbaCode(Path.Combine(root, "src", "VBA", "ThisWorkbook.cls"));
    var module = package.Workbook.VbaProject.Modules.AddModule("modTerminPlanner");
    module.Code = ReadVbaCode(Path.Combine(root, "src", "VBA", "modTerminPlanner.bas"));
}

static string ReadVbaCode(string path)
{
    var source = File.ReadAllText(path);
    var start = source.IndexOf("Option Explicit", StringComparison.Ordinal);
    if (start < 0) throw new InvalidDataException($"Option Explicit is missing from {path}");
    return source[start..]
        .Replace("\r\n", "\n")
        .Replace("\r", "\n")
        .Replace("\n", "\r\n");
}

static void AddListValidation(ExcelWorksheet sheet, string address, string name)
{
    var validation = sheet.DataValidations.AddListValidation(address);
    validation.Formula.ExcelFormula = name;
    validation.ShowErrorMessage = true;
    validation.ErrorTitle = "Choose a listed value";
    validation.Error = "Select a value from the drop-down list.";
}

static void AddButton(ExcelWorksheet sheet, string name, string text, string macro,
    int row, int column, int width, bool danger = false)
{
    ExcelControlButton button = sheet.Drawings.AddButtonControl(name);
    button.Text = text;
    button.Macro = macro;
    button.SetPosition(row, 3, column, 3);
    button.SetSize(width, 28);
}

static void Input(ExcelWorksheet sheet, string address, object value, string? format = null)
{
    var cells = sheet.Cells[address];
    if (address.Contains(':')) cells.Merge = true;
    cells.Value = value;
    cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
    cells.Style.Fill.BackgroundColor.SetColor(Color.White);
    cells.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.FromArgb(176, 202, 217));
    cells.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
    if (format is not null) cells.Style.Numberformat.Format = format;
}

static void Label(ExcelWorksheet sheet, string address, string value)
{
    var cell = sheet.Cells[address];
    cell.Value = value;
    cell.Style.Font.Bold = true;
    cell.Style.Font.Size = 9;
    cell.Style.Font.Color.SetColor(Color.FromArgb(72, 101, 119));
}

static void CreateCard(ExcelWorksheet sheet, string address, string label, string formula)
{
    var range = sheet.Cells[address];
    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
    range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(238, 247, 252));
    range.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.FromArgb(182, 218, 235));
    var labelRange = sheet.Cells[range.Start.Row, range.Start.Column, range.Start.Row, range.End.Column];
    labelRange.Merge = true;
    labelRange.Value = label;
    labelRange.Style.Font.Size = 9;
    labelRange.Style.Font.Bold = true;
    labelRange.Style.Font.Color.SetColor(Color.FromArgb(72, 101, 119));
    labelRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
    var valueRange = sheet.Cells[range.Start.Row + 1, range.Start.Column, range.End.Row, range.End.Column];
    valueRange.Merge = true;
    valueRange.Formula = formula;
    valueRange.Style.Font.Size = 24;
    valueRange.Style.Font.Bold = true;
    valueRange.Style.Font.Color.SetColor(Color.FromArgb(7, 91, 150));
    valueRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
    valueRange.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
}

static void StyleHero(ExcelRange range)
{
    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
    range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(5, 56, 92));
    range.Style.Font.Name = "Aptos Display";
    range.Style.Font.Size = 20;
    range.Style.Font.Bold = true;
    range.Style.Font.Color.SetColor(Color.White);
    range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
    range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
    range.Style.Indent = 1;
}

static void StyleTitle(ExcelRange range) => StyleHero(range);

static void StyleSection(ExcelRange range)
{
    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
    range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(7, 91, 150));
    range.Style.Font.Bold = true;
    range.Style.Font.Color.SetColor(Color.White);
    range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
    range.Style.Indent = 1;
}

static void StyleHeader(ExcelRange range)
{
    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
    range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(219, 237, 247));
    range.Style.Font.Bold = true;
    range.Style.Font.Color.SetColor(Color.FromArgb(5, 56, 92));
    range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
    range.Style.Border.Bottom.Color.SetColor(Color.FromArgb(126, 174, 199));
}

static void ValidateWorkbook(string path)
{
    if (!File.Exists(path) || new FileInfo(path).Length < 10_000)
        throw new InvalidDataException("The generated workbook is missing or unexpectedly small.");

    using var archive = ZipFile.OpenRead(path);
    string[] required = ["[Content_Types].xml", "xl/workbook.xml", "xl/vbaProject.bin"];
    foreach (var entry in required)
    {
        if (archive.GetEntry(entry) is null) throw new InvalidDataException($"Missing workbook part: {entry}");
    }

    using var package = new ExcelPackage(new FileInfo(path));
    string[] sheets = ["Dashboard", "Appointments", "Lists", "Guide"];
    foreach (var sheet in sheets)
    {
        if (package.Workbook.Worksheets[sheet] is null) throw new InvalidDataException($"Missing sheet: {sheet}");
    }
    if (package.Workbook.VbaProject?.Modules["modTerminPlanner"] is null)
        throw new InvalidDataException("The VBA module was not embedded.");

    var moduleCode = package.Workbook.VbaProject.Modules["modTerminPlanner"].Code;
    string[] procedures = ["AddAppointment", "RefreshPlanner", "SearchAppointments", "ExportUpcomingCsv"];
    foreach (var procedure in procedures)
    {
        if (!moduleCode.Contains($"Public Sub {procedure}", StringComparison.Ordinal))
            throw new InvalidDataException($"Missing VBA procedure: {procedure}");
    }
    if (moduleCode.Contains("\r\r\n", StringComparison.Ordinal))
        throw new InvalidDataException("The embedded VBA source contains invalid doubled carriage returns.");

    if (package.Workbook.Worksheets["Appointments"].Tables["tblAppointments"] is null)
        throw new InvalidDataException("The appointment table was not created.");
    string[] rangeNames = ["FormDate", "FormStart", "FormEnd", "FormSubject", "FormClient"];
    foreach (var name in rangeNames)
    {
        if (package.Workbook.Names[name] is null) throw new InvalidDataException($"Missing form range: {name}");
    }
    var assignedMacros = package.Workbook.Worksheets["Dashboard"].Drawings
        .OfType<ExcelControlButton>()
        .Select(button => button.Macro)
        .ToHashSet(StringComparer.Ordinal);
    foreach (var procedure in new[] { "AddAppointment", "ClearForm", "RefreshPlanner", "ExportUpcomingCsv", "OpenGuide" })
    {
        if (!assignedMacros.Contains(procedure))
            throw new InvalidDataException($"No Dashboard button is assigned to {procedure}.");
    }
}
