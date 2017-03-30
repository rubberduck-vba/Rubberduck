using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Inspections;
using RubberduckTests.Mocks;
using Rubberduck.Settings;
using System.Threading;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Inspections.Rubberduck.Inspections;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace RubberduckTests.Inspections
{
    [TestClass]
    public class ObsoleteCallStatementInspectionTests
    {
        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteCallStatement_ReturnsResult()
        {
            const string inputCode =
@"Sub Foo()
    Call Foo
End Sub";

            var settings = new Mock<IGeneralConfigService>();
            var config = GetTestConfig();
            settings.Setup(x => x.LoadConfiguration()).Returns(config);

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteCallStatementInspection(state);
            var inspector = new Inspector(settings.Object, new IInspection[] { inspection });

            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteCallStatement_DoesNotReturnResult()
        {
            const string inputCode =
@"Sub Foo()
    Foo
End Sub";

            var settings = new Mock<IGeneralConfigService>();
            var config = GetTestConfig();
            settings.Setup(x => x.LoadConfiguration()).Returns(config);

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteCallStatementInspection(state);
            var inspector = new Inspector(settings.Object, new IInspection[] { inspection });

            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteCallStatement_DoesNotReturnResult_InstructionSeparator()
        {
            const string inputCode =
@"Sub Foo()
    Call Foo: Foo
End Sub";

            var settings = new Mock<IGeneralConfigService>();
            var config = GetTestConfig();
            settings.Setup(x => x.LoadConfiguration()).Returns(config);

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteCallStatementInspection(state);
            var inspector = new Inspector(settings.Object, new IInspection[] { inspection });

            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteCallStatement_ReturnsResult_ColonInComment()
        {
            const string inputCode =
@"Sub Foo()
    Call Foo ' I''ve got a colon: see?
End Sub";

            var settings = new Mock<IGeneralConfigService>();
            var config = GetTestConfig();
            settings.Setup(x => x.LoadConfiguration()).Returns(config);

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteCallStatementInspection(state);
            var inspector = new Inspector(settings.Object, new IInspection[] { inspection });

            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteCallStatement_ReturnsResult_ColonInStringLiteral()
        {
            const string inputCode =
@"Sub Foo(ByVal str As String)
    Call Foo("":"")
End Sub";

            var settings = new Mock<IGeneralConfigService>();
            var config = GetTestConfig();
            settings.Setup(x => x.LoadConfiguration()).Returns(config);

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteCallStatementInspection(state);
            var inspector = new Inspector(settings.Object, new IInspection[] { inspection });

            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteCallStatement_ReturnsMultipleResults()
        {
            const string inputCode =
@"Sub Foo()
    Call Goo(1, ""test"")
End Sub

Sub Goo(arg1 As Integer, arg1 As String)
    Call Foo
End Sub";

            var settings = new Mock<IGeneralConfigService>();
            var config = GetTestConfig();
            settings.Setup(x => x.LoadConfiguration()).Returns(config);

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteCallStatementInspection(state);
            var inspector = new Inspector(settings.Object, new IInspection[] { inspection });

            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            Assert.AreEqual(2, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteCallStatement_ReturnsResults_SomeObsoleteCallStatements()
        {
            const string inputCode =
@"Sub Foo()
    Call Goo(1, ""test"")
End Sub

Sub Goo(arg1 As Integer, arg1 As String)
    Foo
End Sub";

            var settings = new Mock<IGeneralConfigService>();
            var config = GetTestConfig();
            settings.Setup(x => x.LoadConfiguration()).Returns(config);

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteCallStatementInspection(state);
            var inspector = new Inspector(settings.Object, new IInspection[] { inspection });

            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteCallStatement_Ignored_DoesNotReturnResult()
        {
            const string inputCode =
@"Sub Foo()
    '@Ignore ObsoleteCallStatement
    Call Foo
End Sub";

            var settings = new Mock<IGeneralConfigService>();
            var config = GetTestConfig();
            settings.Setup(x => x.LoadConfiguration()).Returns(config);

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteCallStatementInspection(state);
            var inspector = new Inspector(settings.Object, new IInspection[] { inspection });

            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            Assert.IsFalse(inspectionResults.Any());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteCallStatement_QuickFixWorks_RemoveCallStatement()
        {
            const string inputCode =
@"Sub Foo()
    Call Goo(1, ""test"")
End Sub

Sub Goo(arg1 As Integer, arg1 As String)
    Call Foo
End Sub";

            const string expectedCode =
@"Sub Foo()
    Goo 1, ""test""
End Sub

Sub Goo(arg1 As Integer, arg1 As String)
    Foo
End Sub";

            var settings = new Mock<IGeneralConfigService>();
            var config = GetTestConfig();
            settings.Setup(x => x.LoadConfiguration()).Returns(config);

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteCallStatementInspection(state);
            var inspector = new Inspector(settings.Object, new IInspection[] { inspection });

            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            foreach (var inspectionResult in inspectionResults)
            {
                inspectionResult.QuickFixes.First().Fix();
            }

            Assert.AreEqual(expectedCode, component.CodeModule.Content());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteCallStatement_IgnoreQuickFixWorks()
        {
            const string inputCode =
@"Sub Foo()
    Call Goo(1, ""test"")
End Sub

Sub Goo(arg1 As Integer, arg1 As String)
    Call Foo
End Sub";

            const string expectedCode =
@"Sub Foo()
'@Ignore ObsoleteCallStatement
    Call Goo(1, ""test"")
End Sub

Sub Goo(arg1 As Integer, arg1 As String)
'@Ignore ObsoleteCallStatement
    Call Foo
End Sub";

            var settings = new Mock<IGeneralConfigService>();
            var config = GetTestConfig();
            settings.Setup(x => x.LoadConfiguration()).Returns(config);

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteCallStatementInspection(state);
            var inspector = new Inspector(settings.Object, new IInspection[] { inspection });

            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            foreach (var inspectionResult in inspectionResults)
            {
                inspectionResult.QuickFixes.Single(s => s is IgnoreOnceQuickFix).Fix();
            }

            Assert.AreEqual(expectedCode, component.CodeModule.Content());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionType()
        {
            var inspection = new ObsoleteCallStatementInspection(null);
            Assert.AreEqual(CodeInspectionType.LanguageOpportunities, inspection.InspectionType);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionName()
        {
            const string inspectionName = "ObsoleteCallStatementInspection";
            var inspection = new ObsoleteCallStatementInspection(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }

        private Configuration GetTestConfig()
        {
            var settings = new CodeInspectionSettings();
            settings.CodeInspections.Add(new CodeInspectionSetting
            {
                Description = new ObsoleteCallStatementInspection(null).Description,
                Severity = CodeInspectionSeverity.Suggestion
            });
            return new Configuration
            {
                UserSettings = new UserSettings(null, null, null, settings, null, null, null)
            };
        }
    }
}
