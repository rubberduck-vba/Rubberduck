using System.Linq;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Inspections;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete.Rubberduck.Inspections;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;
using Rubberduck.Settings;
using Rubberduck.VBEditor.Application;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestClass]
    public class ObsoleteCommentSyntaxInspectionTests
    {
        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteCommentSyntax_ReturnsResult()
        {
            const string inputCode = @"Rem test";

            var settings = new Mock<IGeneralConfigService>();
            var config = GetTestConfig();
            settings.Setup(x => x.LoadConfiguration()).Returns(config);

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteCommentSyntaxInspection(state);
            var inspector = new Inspector(settings.Object, new IInspection[] { inspection });

            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;
            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteCommentSyntax_DoesNotReturnResult()
        {
            const string inputCode = @"' test";

            var settings = new Mock<IGeneralConfigService>();
            var config = GetTestConfig();
            settings.Setup(x => x.LoadConfiguration()).Returns(config);

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteCommentSyntaxInspection(state);
            var inspector = new Inspector(settings.Object, new IInspection[] { inspection });

            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;
            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteCommentSyntax_DoesNotReturnResult_RemInStringLiteral()
        {
            const string inputCode = 
@"Sub Foo()
    Dim bar As String
    bar = ""iejo rem oernp"" ' test
End Sub";

            var settings = new Mock<IGeneralConfigService>();
            var config = GetTestConfig();
            settings.Setup(x => x.LoadConfiguration()).Returns(config);

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteCommentSyntaxInspection(state);
            var inspector = new Inspector(settings.Object, new IInspection[] { inspection });

            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;
            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteCommentSyntax_ReturnsMultipleResults()
        {
            const string inputCode =
@"Rem test1
Rem test2";

            var settings = new Mock<IGeneralConfigService>();
            var config = GetTestConfig();
            settings.Setup(x => x.LoadConfiguration()).Returns(config);

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteCommentSyntaxInspection(state);
            var inspector = new Inspector(settings.Object, new IInspection[] { inspection });

            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;
            Assert.AreEqual(2, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteCommentSyntax_ReturnsResults_SomeObsoleteCommentSyntax()
        {
            const string inputCode =
@"Rem test1
' test2";

            var settings = new Mock<IGeneralConfigService>();
            var config = GetTestConfig();
            settings.Setup(x => x.LoadConfiguration()).Returns(config);

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteCommentSyntaxInspection(state);
            var inspector = new Inspector(settings.Object, new IInspection[] { inspection });

            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;
            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteCommentSyntax_Ignored_DoesNotReturnResult()
        {
            const string inputCode = @"
'@Ignore ObsoleteCommentSyntax
Rem test";

            var settings = new Mock<IGeneralConfigService>();
            var config = GetTestConfig();
            settings.Setup(x => x.LoadConfiguration()).Returns(config);

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteCommentSyntaxInspection(state);
            var inspector = new Inspector(settings.Object, new IInspection[] { inspection });

            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;
            Assert.IsFalse(inspectionResults.Any());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteCommentSyntax_QuickFixWorks_UpdateComment()
        {
            const string inputCode =
@"Rem test1";

            const string expectedCode =
@"' test1";

            var settings = new Mock<IGeneralConfigService>();
            var config = GetTestConfig();
            settings.Setup(x => x.LoadConfiguration()).Returns(config);

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteCommentSyntaxInspection(state);
            var inspector = new Inspector(settings.Object, new IInspection[] { inspection });

            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;
            inspectionResults.First().QuickFixes.Single(s => s is ReplaceObsoleteCommentMarkerQuickFix).Fix();

            Assert.AreEqual(expectedCode, component.CodeModule.Content());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteCommentSyntax_QuickFixWorks_UpdateCommentHasContinuation()
        {
            const string inputCode =
@"Rem this is _
a comment";

            const string expectedCode =
@"' this is _
a comment";

            var settings = new Mock<IGeneralConfigService>();
            var config = GetTestConfig();
            settings.Setup(x => x.LoadConfiguration()).Returns(config);

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteCommentSyntaxInspection(state);
            var inspector = new Inspector(settings.Object, new IInspection[] { inspection });

            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;
            inspectionResults.First().QuickFixes.Single(s => s is ReplaceObsoleteCommentMarkerQuickFix).Fix();

            Assert.AreEqual(expectedCode, component.CodeModule.Content());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteCommentSyntax_QuickFixWorks_RemoveComment()
        {
            const string inputCode =
@"Rem test1";

            const string expectedCode =
@"";

            var settings = new Mock<IGeneralConfigService>();
            var config = GetTestConfig();
            settings.Setup(x => x.LoadConfiguration()).Returns(config);

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteCommentSyntaxInspection(state);
            var inspector = new Inspector(settings.Object, new IInspection[] { inspection });

            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;
            inspectionResults.First().QuickFixes.Single(s => s is RemoveCommentQuickFix).Fix();

            Assert.AreEqual(expectedCode, component.CodeModule.Content());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteCommentSyntax_QuickFixWorks_RemoveCommentHasContinuation()
        {
            const string inputCode =
@"Rem test1 _
continued";

            const string expectedCode =
@"";

            var settings = new Mock<IGeneralConfigService>();
            var config = GetTestConfig();
            settings.Setup(x => x.LoadConfiguration()).Returns(config);

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteCommentSyntaxInspection(state);
            var inspector = new Inspector(settings.Object, new IInspection[] { inspection });

            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;
            inspectionResults.First().QuickFixes.Single(s => s is RemoveCommentQuickFix).Fix();

            Assert.AreEqual(expectedCode, component.CodeModule.Content());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteCommentSyntax_QuickFixWorks_UpdateComment_LineHasCode()
        {
            const string inputCode =
@"Dim Foo As Integer: Rem This is a comment";

            const string expectedCode =
@"Dim Foo As Integer: ' This is a comment";

            var settings = new Mock<IGeneralConfigService>();
            var config = GetTestConfig();
            settings.Setup(x => x.LoadConfiguration()).Returns(config);

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteCommentSyntaxInspection(state);
            var inspector = new Inspector(settings.Object, new IInspection[] { inspection });

            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;
            inspectionResults.First().QuickFixes.Single(s => s is ReplaceObsoleteCommentMarkerQuickFix).Fix();

            Assert.AreEqual(expectedCode, component.CodeModule.Content());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteCommentSyntax_QuickFixWorks_UpdateComment_LineHasCodeAndContinuation()
        {
            const string inputCode =
@"Dim Foo As Integer: Rem This is _
a comment";

            const string expectedCode =
@"Dim Foo As Integer: ' This is _
a comment";

            var settings = new Mock<IGeneralConfigService>();
            var config = GetTestConfig();
            settings.Setup(x => x.LoadConfiguration()).Returns(config);

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteCommentSyntaxInspection(state);
            var inspector = new Inspector(settings.Object, new IInspection[] { inspection });

            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;
            inspectionResults.First().QuickFixes.Single(s => s is ReplaceObsoleteCommentMarkerQuickFix).Fix();

            Assert.AreEqual(expectedCode, component.CodeModule.Content());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteCommentSyntax_QuickFixWorks_RemoveComment_LineHasCode()
        {
            const string inputCode =
@"Dim Foo As Integer: Rem This is a comment";

            const string expectedCode =
@"Dim Foo As Integer:";

            var settings = new Mock<IGeneralConfigService>();
            var config = GetTestConfig();
            settings.Setup(x => x.LoadConfiguration()).Returns(config);

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteCommentSyntaxInspection(state);
            var inspector = new Inspector(settings.Object, new IInspection[] { inspection });

            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;
            inspectionResults.First().QuickFixes.Single(s => s is RemoveCommentQuickFix).Fix();

            Assert.AreEqual(expectedCode, component.CodeModule.Content());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteCommentSyntax_QuickFixWorks_RemoveComment_LineHasCodeAndContinuation()
        {
            const string inputCode =
@"Dim Foo As Integer: Rem This is _
a comment";

            const string expectedCode =
@"Dim Foo As Integer:";

            var settings = new Mock<IGeneralConfigService>();
            var config = GetTestConfig();
            settings.Setup(x => x.LoadConfiguration()).Returns(config);

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteCommentSyntaxInspection(state);
            var inspector = new Inspector(settings.Object, new IInspection[] { inspection });

            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;
            inspectionResults.First().QuickFixes.Single(s => s is RemoveCommentQuickFix).Fix();

            Assert.AreEqual(expectedCode, component.CodeModule.Content());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteCommentSyntax_IgnoreQuickFixWorks()
        {
            const string inputCode =
@"Rem test1";

            const string expectedCode =
@"'@Ignore ObsoleteCommentSyntax
Rem test1";

            var settings = new Mock<IGeneralConfigService>();
            var config = GetTestConfig();
            settings.Setup(x => x.LoadConfiguration()).Returns(config);

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteCommentSyntaxInspection(state);
            var inspector = new Inspector(settings.Object, new IInspection[] { inspection });

            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            inspectionResults.First().QuickFixes.Single(s => s is IgnoreOnceQuickFix).Fix();
            Assert.AreEqual(expectedCode, component.CodeModule.Content());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionType()
        {
            var inspection = new ObsoleteCommentSyntaxInspection(null);
            Assert.AreEqual(CodeInspectionType.LanguageOpportunities, inspection.InspectionType);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionName()
        {
            const string inspectionName = "ObsoleteCommentSyntaxInspection";
            var inspection = new ObsoleteCommentSyntaxInspection(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }

        private Configuration GetTestConfig()
        {
            var settings = new CodeInspectionSettings();
            settings.CodeInspections.Add(new CodeInspectionSetting
            {
                Description = new ObsoleteCommentSyntaxInspection(null).Description,
                Severity = CodeInspectionSeverity.Suggestion
            });
            return new Configuration
            {
                UserSettings = new UserSettings(null, null, null, settings, null, null, null)
            };
        }
    }
}
