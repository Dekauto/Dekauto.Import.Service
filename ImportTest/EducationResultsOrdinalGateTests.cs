using System.Reflection;
using Dekauto.Import.Service.Domain.Services;
using OfficeOpenXml;

namespace ImportTest;

[TestClass]
public class EducationResultsOrdinalGateTests
{
    [DataTestMethod]
    [DataRow(1, true)]
    [DataRow(1.0, true)]
    [DataRow("1,0", true)]
    [DataRow("№ п/п", false)]
    [DataRow(null, false)]
    public void EducationResultsOrdinalLooksLikeImportsService_recognizes_row_numbers(object? value, bool expected)
    {
        var method = typeof(ImportsService).GetMethod(
            "EducationResultsOrdinalLooksLikeImportsService",
            BindingFlags.Static | BindingFlags.NonPublic);
        Assert.IsNotNull(method);

        var actual = (bool)method!.Invoke(null, new object?[] { value })!;

        Assert.AreEqual(expected, actual);
    }

    [TestMethod]
    public void ParseEducationResultsWorksheet_skips_two_row_header_before_body()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        using var package = new ExcelPackage();
        var sheet = package.Workbook.Worksheets.Add("Ро1");

        sheet.Cells[5, 2].Value = "№ п/п";
        sheet.Cells[5, 3].Value = "Наименование дисциплины";
        sheet.Cells[6, 2].Value = "№ п/п";
        sheet.Cells[6, 3].Value = "Наименование дисциплины (модуль)";
        sheet.Cells[7, 2].Value = 1;
        sheet.Cells[7, 3].Value = "Математика";
        sheet.Cells[7, 4].Value = 3;
        sheet.Cells[7, 5].Value = 108;
        sheet.Cells[7, 6].Value = 72;
        sheet.Cells[7, 8].Value = "экзамен";
        sheet.Cells[7, 9].Value = 12;

        var parseMethod = typeof(ImportsService).GetMethod(
            "ParseEducationResultsWorksheet",
            BindingFlags.Instance | BindingFlags.NonPublic);
        Assert.IsNotNull(parseMethod);

        var service = CreateImportsService();
        var results = (List<Dekauto.Import.Service.Domain.Entities.StudentDisciplineResult>)parseMethod!.Invoke(
            service,
            new object[] { sheet })!;

        Assert.IsTrue(results.Any(r => r.DisciplineName == "Математика"));
        Assert.IsFalse(results.Any(r => r.DisciplineName.Contains("Наименование", StringComparison.OrdinalIgnoreCase)));
    }

    private static ImportsService CreateImportsService()
    {
        var config = new Microsoft.Extensions.Configuration.ConfigurationBuilder().Build();
        return new ImportsService(config, Microsoft.Extensions.Logging.Abstractions.NullLogger<ImportsService>.Instance);
    }
}
