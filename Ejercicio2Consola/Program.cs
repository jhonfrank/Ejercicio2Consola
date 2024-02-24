// See https://aka.ms/new-console-template for more information
using Ejercicio2Consola;
using Microsoft.Extensions.Configuration;
using System.Net.Mail;
using System.Text.Json;
using System.Net;
using System.Collections.Generic;
using System.Collections;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Runtime.Intrinsics.X86;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2010.ExcelAc;
using System.Data;

// configuration
var builder = new ConfigurationBuilder()
    .AddJsonFile($"appsettings.json", true, true)
    .AddEnvironmentVariables();
var configuration = builder.Build();

// Client
HttpClient clientAPI = new HttpClient()
{
    BaseAddress = new Uri(configuration.GetSection("APIbaseURL").Value)
};

var url = string.Format("/api/Employee");
var result = new List<Employee>();
var response = await clientAPI.GetAsync(url);
if (response.IsSuccessStatusCode)
{
    var stringResponse = await response.Content.ReadAsStringAsync();

    result = JsonSerializer.Deserialize<List<Employee>>(stringResponse,
        new JsonSerializerOptions() { PropertyNamingPolicy = JsonNamingPolicy.CamelCase });
}
else
{
    throw new HttpRequestException(response.ReasonPhrase);
}

// Filter admission_date > 2020
result = result.Where(x => x.admission_date.Year > 2020).ToList();

var fi = DateTime.Parse(configuration.GetSection("consulta_fecha_inicio").Value);
var ff = DateTime.Parse(configuration.GetSection("consulta_fecha_final").Value);
result = result
            .Where(x => DateTime.Compare(x.admission_date, fi)>0 & DateTime.Compare(x.admission_date,ff)<0)
            .OrderBy(x => x.admission_date)
            .ToList();


// Create excel
MemoryStream memoryStream = new MemoryStream();
SpreadsheetDocument myWorkbook = SpreadsheetDocument.Create(memoryStream, SpreadsheetDocumentType.Workbook);
// workbook Part
WorkbookPart workbookPart = myWorkbook.AddWorkbookPart();
var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
string relId = workbookPart.GetIdOfPart(worksheetPart);

// file Version
var fileVersion = new FileVersion { ApplicationName = "Microsoft Office Excel" };

// sheets               
var sheets = new Sheets();
var sheet = new Sheet { Name = "Hoja1", SheetId = 1, Id = relId };
sheets.Append(sheet);

// data
SheetData sheetData = new SheetData(CreateSheetData(result));

// add the parts to the workbook and save
var workbook = new Workbook();
workbook.Append(fileVersion);
workbook.Append(sheets);
var worksheet = new Worksheet();
worksheet.Append(sheetData);
worksheetPart.Worksheet = worksheet;
worksheetPart.Worksheet.Save();
myWorkbook.WorkbookPart.Workbook = workbook;
myWorkbook.WorkbookPart.Workbook.Save();



// Send mail
var host = "smtp.ethereal.email";
var port = 587;
var username = "liliane.mertz64@ethereal.email";
var password = "HrZkESRFWF11pFna72";
var enable = true;

var smtpClient = new SmtpClient(host)
{
    Port = port,
    EnableSsl = enable,
    Credentials = new NetworkCredential(username, password)
};

var mailMessage = new MailMessage("liliane.mertz64@ethereal.email", configuration.GetSection("email").Value);
mailMessage.Subject = "Reporte Empleados - Examen Técnico Oechsle";
mailMessage.Body = "";
mailMessage.IsBodyHtml = true;

mailMessage.Attachments.Add(new Attachment(memoryStream, "Report.xlsx"));

smtpClient.Send(mailMessage);


List<Row> CreateSheetData(List<Employee> list)
{
    List<Row> elements = new List<Row>();

    // row header
    var rowHeader = new Row();
    Cell[] cellsHeader = new Cell[7];
    string[] cellsHeaderName = { "id", "name", "document_number", "salary", "age", "profile", "admission_date" };

    for (int i = 0; i < 7; i++)
    {
        cellsHeader[i] = new Cell();
        cellsHeader[i].DataType = CellValues.String;
        cellsHeader[i].CellValue = new CellValue(cellsHeaderName[i]);
    }
    rowHeader.Append(cellsHeader);
    elements.Add(rowHeader);

    // rows data
    foreach (Employee item in list)
    {
        var row = new Row();
        Cell[] cells = new Cell[7];
        object[] cellsHeaderValue = { item.id, item.name, item.document_number, item.salary, item.age, item.profile, item.admission_date };
        for (int i = 0; i < 7; i++)
        {
            cells[i] = new Cell();
            cells[i].DataType = CellValues.String;
            cells[i].CellValue = new CellValue(cellsHeaderValue[i].ToString());
        }
        row.Append(cells);
        elements.Add(row);
    }
    return elements;
}