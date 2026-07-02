using Dekauto.Import.Service.Domain.Services;

namespace ImportTest
{
    [TestClass]
    public class SupplementCourseOfTrainingTests
    {
        [TestMethod]
        public void MergeSupplementSpecialtyCodeWithName_prepends_code_when_name_only()
        {
            var merged = ImportsService.MergeSupplementSpecialtyCodeWithName(
                "09.03.03 Прикладная информатика",
                "Прикладная информатика");
            Assert.AreEqual("09.03.03 Прикладная информатика", merged);
        }

        [DataTestMethod]
        [DataRow("09.03.03 Прикладная информатика", "09.03.03 Прикладная информатика")]
        [DataRow("Прикладная информатика", "Прикладная информатика")]
        public void ResolveSupplementCourseOfTraining_keeps_full_or_name_without_reference(string input, string? expected)
        {
            using var package = new OfficeOpenXml.ExcelPackage();
            package.Workbook.Worksheets.Add("ОбщСведения");
            var actual = ImportsService.ResolveSupplementCourseOfTraining(package.Workbook.Worksheets, input);
            Assert.AreEqual(expected, actual);
        }

        [TestMethod]
        public void ResolveSupplementCourseOfTraining_uses_spravochnik_when_c78_name_only()
        {
            using var package = new OfficeOpenXml.ExcelPackage();
            var refSheet = package.Workbook.Worksheets.Add("Справочник");
            refSheet.Cells[3, 3].Value = "09.03.03 Прикладная информатика";

            var actual = ImportsService.ResolveSupplementCourseOfTraining(
                package.Workbook.Worksheets,
                "Прикладная информатика");

            Assert.AreEqual("09.03.03 Прикладная информатика", actual);
        }
    }
}
