using Dekauto.Import.Service.Domain.Entities;
using Dekauto.Import.Service.Domain.Services;

namespace ImportTest
{
    [TestClass]
    public class StatementPlanReconciliationTests
    {
        [TestMethod]
        public void ExactMatch_UsesPlanName_WithoutReview()
        {
            var catalog = new List<PlanDisciplineCatalogEntry>
            {
                new()
                {
                    Semester = 1,
                    PlanName = "Математический анализ",
                    DisplayName = "Математический анализ",
                    ControlType = "экзамен",
                    TotalHours = 144,
                    AudHours = 72
                }
            };

            var snapshot = new StatementDisciplineSnapshot
            {
                StatementName = "Математический анализ",
                Score = "5",
                Semester = 1,
                Year = 2024
            };

            var result = StatementPlanReconciliation.ReconcileSnapshot(snapshot, catalog);

            Assert.AreEqual("Математический анализ", result.DisciplineName);
            Assert.AreEqual("экзамен", result.ControlType);
            Assert.AreEqual(144, result.TotalHours);
            Assert.IsFalse(result.RequiresManualValidation);
        }

        [TestMethod]
        public void FuzzyMatch_Typo_UsesPlanName_WithReview()
        {
            var catalog = new List<PlanDisciplineCatalogEntry>
            {
                new()
                {
                    Semester = 3,
                    PlanName = "Проектный практикум",
                    DisplayName = "Проектный практикум",
                    ControlType = "зачёт",
                    TotalHours = 36,
                    AudHours = 36
                }
            };

            var snapshot = new StatementDisciplineSnapshot
            {
                StatementName = "Проектный пракикум",
                Score = "4",
                Semester = 3
            };

            var result = StatementPlanReconciliation.ReconcileSnapshot(snapshot, catalog);

            Assert.AreEqual("Проектный практикум", result.DisciplineName);
            Assert.AreEqual(36, result.TotalHours);
            Assert.IsTrue(result.RequiresManualValidation);
        }

        [TestMethod]
        public void NoMatch_KeepsStatementName_WithReview_NoHours()
        {
            var catalog = new List<PlanDisciplineCatalogEntry>
            {
                new()
                {
                    Semester = 1,
                    PlanName = "Физика",
                    DisplayName = "Физика",
                    ControlType = "экзамен",
                    TotalHours = 108
                }
            };

            var snapshot = new StatementDisciplineSnapshot
            {
                StatementName = "Совсем другая дисциплина",
                Score = "3",
                Semester = 1
            };

            var result = StatementPlanReconciliation.ReconcileSnapshot(snapshot, catalog);

            Assert.AreEqual("Совсем другая дисциплина", result.DisciplineName);
            Assert.IsNull(result.TotalHours);
            Assert.IsNull(result.AudHours);
            Assert.IsTrue(result.RequiresManualValidation);
        }

        [TestMethod]
        public void PracticeMatch_UsesPlanDisplayName_AlwaysReview()
        {
            var displayName = StatementPlanReconciliation.FormatPracticeDisplayName(
                "Производственная практика",
                "практика по получению профессиональных компетенций");

            var catalog = new List<PlanDisciplineCatalogEntry>
            {
                new()
                {
                    Semester = 4,
                    PlanName = "практика по получению профессиональных компетенций",
                    DisplayName = displayName,
                    IsPractice = true,
                    ControlType = "практика",
                    TotalHours = 180
                }
            };

            var snapshot = new StatementDisciplineSnapshot
            {
                StatementName = "Производственная практика, практика по получению профессиональных компетенций",
                Score = "5",
                Semester = 4,
                IsPractice = true
            };

            var result = StatementPlanReconciliation.ReconcileSnapshot(snapshot, catalog);

            Assert.AreEqual(displayName, result.DisciplineName);
            Assert.AreEqual("практика", result.ControlType);
            Assert.AreEqual(180, result.TotalHours);
            Assert.IsTrue(result.RequiresManualValidation);
        }

        [TestMethod]
        public void UpsertSnapshot_LastFileWins()
        {
            var list = new List<StatementDisciplineSnapshot>();

            StatementPlanReconciliation.UpsertSnapshot(list, new StatementDisciplineSnapshot
            {
                StatementName = "Физика",
                Score = "3",
                Semester = 1
            });

            StatementPlanReconciliation.UpsertSnapshot(list, new StatementDisciplineSnapshot
            {
                StatementName = "Физика",
                Score = "5",
                Semester = 1,
                SourceFileName = "ved2.xlsx"
            });

            Assert.AreEqual(1, list.Count);
            Assert.AreEqual("5", list[0].Score);
            Assert.AreEqual("ved2.xlsx", list[0].SourceFileName);
        }

        [TestMethod]
        public void ReconcileStudentsWithPlan_BuildsDisciplineResults()
        {
            var student = new Student
            {
                PendingStatementDisciplines =
                {
                    new StatementDisciplineSnapshot
                    {
                        StatementName = "История",
                        Score = "4",
                        Semester = 2
                    }
                }
            };

            var catalog = new List<PlanDisciplineCatalogEntry>
            {
                new()
                {
                    Semester = 2,
                    PlanName = "История",
                    DisplayName = "История",
                    ControlType = "зачёт"
                }
            };

            StatementPlanReconciliation.ReconcileStudentsWithPlan(new[] { student }, catalog);

            Assert.AreEqual(0, student.PendingStatementDisciplines.Count);
            Assert.AreEqual(1, student.DisciplineResults.Count);
            Assert.AreEqual("История", student.DisciplineResults[0].DisciplineName);
        }

        [TestMethod]
        public void RegularDiscipline_MatchesPlanRowMarkedAsPracticeInCatalog()
        {
            var catalog = new List<PlanDisciplineCatalogEntry>
            {
                new()
                {
                    Semester = 3,
                    PlanName = "Теория вероятностей",
                    DisplayName = "Учебная практика, теория вероятностей",
                    IsPractice = true,
                    ControlType = "экзамен",
                    TotalHours = 72,
                    AudHours = 36
                }
            };

            var snapshot = new StatementDisciplineSnapshot
            {
                StatementName = "Теория вероятностей",
                Score = "4",
                Semester = 3
            };

            var result = StatementPlanReconciliation.ReconcileSnapshot(snapshot, catalog);

            Assert.AreEqual("Теория вероятностей", result.DisciplineName);
            Assert.AreEqual("экзамен", result.ControlType);
            Assert.AreEqual(72, result.TotalHours);
            Assert.IsFalse(result.RequiresManualValidation);
        }

        [TestMethod]
        public void IsStatementPracticeDiscipline_DetectsPracticaPrefix_NotPracticum()
        {
            Assert.IsTrue(StatementPlanReconciliation.IsStatementPracticeDiscipline(
                "Практика по получению навыков"));
            Assert.IsFalse(StatementPlanReconciliation.IsStatementPracticeDiscipline(
                "Проектный практикум"));
        }
    }
}
