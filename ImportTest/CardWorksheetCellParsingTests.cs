using Dekauto.Import.Service.Domain.Services;

namespace ImportTest
{
    [TestClass]
    public class CardWorksheetCellParsingTests
    {
        [DataTestMethod]
        [DataRow("X", null)]
        [DataRow("x", null)]
        [DataRow(null, null)]
        [DataRow("", null)]
        [DataRow("5", 5d)]
        [DataRow("5,5", 5.5d)]
        [DataRow(3.0, 3d)]
        [DataRow(7, 7d)]
        public void TryGetCardCellDouble_parses_or_returns_null(object? value, double? expected)
        {
            var actual = ImportsService.TryGetCardCellDouble(value);
            if (expected is null)
                Assert.IsNull(actual);
            else
                Assert.AreEqual(expected.Value, actual!.Value, 0.0001);
        }
    }
}
