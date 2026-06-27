using Dekauto.Import.Service.Domain.Entities;
using Dekauto.Import.Service.Domain.Services;

namespace ImportTest
{
    [TestClass]
    public class StatementImportMatchingTests
    {
        [TestMethod]
        public void NormalizeGroupName_MatchesWithDifferentCaseAndHyphen()
        {
            var a = StatementImportMatching.NormalizeGroupNameForComparison("22ИТ-ПИ");
            var b = StatementImportMatching.NormalizeGroupNameForComparison("22ит-пи");
            Assert.AreEqual(a, b);
        }

        [TestMethod]
        public void NormalizeGroupName_MatchesHomoglyphs()
        {
            var latin = StatementImportMatching.NormalizeGroupNameForComparison("22IT-PI");
            var cyrillic = StatementImportMatching.NormalizeGroupNameForComparison("22ИТ-ПИ");
            Assert.AreEqual(latin, cyrillic);
        }

        [TestMethod]
        public void TryExtractGroupFromStatementRow_FindsExtendedGroupNearLabel()
        {
            string CellText(int row, int col) => row switch
            {
                4 when col == 2 => "Группа",
                4 when col == 3 => "22ИТ-ПИ(б/о)ПИП-1",
                _ => string.Empty
            };

            var ok = StatementImportMatching.TryExtractGroupFromStatementRow(4, 5, CellText, out var groupRaw);

            Assert.IsTrue(ok);
            Assert.AreEqual("22ИТ-ПИ(б/о)ПИП-1", groupRaw);
        }

        [TestMethod]
        public void GroupNamesMatch_ExtendedJournalGroupAndShortSheetPrefix()
        {
            Assert.IsTrue(StatementImportMatching.GroupNamesMatchForStatement(
                "22ИТ-ПИ",
                "22ИТ-ПИ(б/о)ПИП-1"));
            Assert.IsTrue(StatementImportMatching.GroupNamesMatchForStatement(
                "22ИТ-ПИ(б/о)ПИП-1",
                "22ИТ-ПИ(б/о)ПИП-1"));
        }

        [TestMethod]
        public void TryExtractGroupFromStatementRow_FindsGroupNearLabel()
        {
            string CellText(int row, int col) => row switch
            {
                4 when col == 2 => "Группа",
                4 when col == 3 => "22ИТ-ПИ",
                _ => string.Empty
            };

            var ok = StatementImportMatching.TryExtractGroupFromStatementRow(4, 5, CellText, out var groupRaw);

            Assert.IsTrue(ok);
            Assert.AreEqual("22ИТ-ПИ", groupRaw);
        }

        [TestMethod]
        public void FioMatch_ExactKeyFip_FindsSingleCandidate()
        {
            var student = new Student
            {
                Surname = "Иванов",
                Name = "Иван",
                Patronymic = "Петрович"
            };

            var (keyFi, keyFip) = StatementImportMatching.BuildStatementStudentKeys(student);
            Assert.AreEqual("иванови", keyFi);
            Assert.AreEqual("ивановип", keyFip);

            string CellText(int row, int col) => row == 8 && col == 2 ? "Иванов И.П." : string.Empty;

            var candidates = StatementImportMatching.FindStatementFioCandidateRows(CellText, 2, 8, 10, keyFi, keyFip);

            Assert.AreEqual(1, candidates.Count);
            Assert.AreEqual(8, candidates[0]);
        }

        [TestMethod]
        public void FioMatch_TwoSimilarRows_IsAmbiguous()
        {
            var student = new Student
            {
                Surname = "Иванов",
                Name = "Иван",
                Patronymic = string.Empty
            };

            var (keyFi, keyFip) = StatementImportMatching.BuildStatementStudentKeys(student);

            string CellText(int row, int col)
            {
                if (col != 2)
                    return string.Empty;
                return row switch
                {
                    8 => "Иванов И.",
                    12 => "Иванов И.",
                    _ => string.Empty
                };
            }

            var candidates = StatementImportMatching.FindStatementFioCandidateRows(CellText, 2, 8, 15, keyFi, keyFip);

            Assert.AreEqual(2, candidates.Count);
            CollectionAssert.Contains(candidates, 8);
            CollectionAssert.Contains(candidates, 12);
        }

        [TestMethod]
        public void GroupFilter_IsolatesStudentsByGroup()
        {
            var sheetGroupKey = StatementImportMatching.NormalizeGroupNameForComparison("22ИТ-ПИ");
            var students = new List<Student>
            {
                new() { Surname = "Иванов", Name = "Иван", GroupName = "22ИТ-ПИ" },
                new() { Surname = "Петров", Name = "Пётр", GroupName = "21ИТ-ПИ" }
            };

            var onSheet = students
                .Where(s => StatementImportMatching.NormalizeGroupNameForComparison(s.GroupName) == sheetGroupKey)
                .ToList();

            Assert.AreEqual(1, onSheet.Count);
            Assert.AreEqual("Иванов", onSheet[0].Surname);
        }

        [TestMethod]
        public void FioMatch_KeyFiWithoutPatronymicInCell_MatchesStudentWithPatronymic()
        {
            var student = new Student
            {
                Surname = "Иванов",
                Name = "Иван",
                Patronymic = "Петрович"
            };

            var (keyFi, keyFip) = StatementImportMatching.BuildStatementStudentKeys(student);
            string CellText(int row, int col) => row == 9 && col == 1 ? "Иванов И." : string.Empty;

            var candidates = StatementImportMatching.FindStatementFioCandidateRows(CellText, 1, 8, 10, keyFi, keyFip);

            Assert.AreEqual(1, candidates.Count);
        }

        [TestMethod]
        public void FioMatch_FullName_FindsSingleCandidate()
        {
            var student = new Student
            {
                Surname = "Иванов",
                Name = "Иван",
                Patronymic = "Петрович"
            };

            var (keyFi, keyFip) = StatementImportMatching.BuildStatementStudentKeys(student);
            string CellText(int row, int col) =>
                row == 10 && col == 1 ? "Иванов Иван Петрович" : string.Empty;

            var candidates = StatementImportMatching.FindStatementFioCandidateRows(CellText, 1, 8, 12, keyFi, keyFip);

            Assert.AreEqual(1, candidates.Count);
        }

        [TestMethod]
        public void TryExtractGroupFromStatementSheet_UsesRow4Only()
        {
            string CellText(int row, int col) => row switch
            {
                4 when col == 4 => "Группа",
                4 when col == 5 => "22 ИТ-ПИ",
                5 when col == 4 => "Группа",
                5 when col == 5 => "99ИТ-ХХ",
                _ => string.Empty
            };

            var ok = StatementImportMatching.TryExtractGroupFromStatementSheet(8, CellText, out var groupRaw);

            Assert.IsTrue(ok);
            Assert.AreEqual("22ИТ-ПИ", groupRaw);
        }

        [TestMethod]
        public void TryResolveSheetGroupRaw_DoesNotUseWorksheetNameFallback()
        {
            string CellText(int row, int col) => string.Empty;
            var students = new List<Student>
            {
                new() { GroupName = "22ИТ-ПИ(б/о)ПИП-1" }
            };

            var ok = StatementImportMatching.TryResolveSheetGroupRaw(
                5, CellText, "22ИТ-ПИ", students, out var groupRaw, out var usedFallback);

            Assert.IsFalse(ok);
            Assert.IsTrue(string.IsNullOrEmpty(groupRaw));
            Assert.IsFalse(usedFallback);
        }

        [TestMethod]
        public void HasProbableStatementStudentRows_DetectsAbbreviatedFioColumn()
        {
            string CellText(int row, int col) => row switch
            {
                >= 8 and <= 10 when col == 2 => "Иванов И.П.",
                _ => string.Empty
            };

            Assert.IsTrue(StatementImportMatching.HasProbableStatementStudentRows(5, 12, CellText));
        }

        [TestMethod]
        public void IsStatementGradeSheet_DetectsFioHeader()
        {
            var headers = new List<string> { "№ п/п", "ФИО студента", "Математика" };

            Assert.IsTrue(StatementImportMatching.IsStatementGradeSheet(headers, h =>
                !string.IsNullOrWhiteSpace(h) && h.Contains("фио", StringComparison.OrdinalIgnoreCase)));
        }

        [TestMethod]
        public void IsStatementFioHeader_DetectsDottedAbbreviation()
        {
            Assert.IsTrue(StatementImportMatching.IsStatementFioHeader("Ф.И.О."));
            Assert.IsTrue(StatementImportMatching.IsStatementFioHeader("Ф.И.О. студента"));
        }

        [TestMethod]
        public void FioMatch_AbbreviatedIoFormat_UsesCompactNormalization()
        {
            var student = new Student
            {
                Surname = "Иванов",
                Name = "Иван",
                Patronymic = "Петрович"
            };

            var (keyFi, keyFip) = StatementImportMatching.BuildStatementStudentKeys(student);
            Assert.IsTrue(StatementImportMatching.StatementFioCellMatchesStudent("Иванов И.П.", keyFi, keyFip));
            Assert.IsTrue(StatementImportMatching.StatementFioCellMatchesStudent("Иванов И. П.", keyFi, keyFip));
        }

        [TestMethod]
        public void FioMatch_DoesNotUseContains()
        {
            var student = new Student { Surname = "Иванов", Name = "Иван", Patronymic = "Петрович" };
            var (keyFi, keyFip) = StatementImportMatching.BuildStatementStudentKeys(student);

            Assert.IsFalse(StatementImportMatching.StatementFioCellMatchesStudent("Сидоров И.П.", keyFi, keyFip));
            Assert.IsFalse(StatementImportMatching.StatementFioCellMatchesStudent("Иванова И.П.", keyFi, keyFip));
        }
    }
}
