using System.Linq;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;
using Rubberduck.Settings;
using Rubberduck.SettingsProvider;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestClass]
    public class UseMeaningfulNameInspectionTests
    {
        [TestMethod]
        [TestCategory("Inspections")]
        public void UseMeaningfulName_ReturnsResult_NameWithAllTheSameLetters()
        {
            const string inputCode =
@"
Private aaa As String
Private bbb As String 
Private ccc As String
Private ddd As String
Private eee As String
Private iii As String
Private ooo As String
Private uuu As String

Sub Eeeeee()
Dim a2z as String       'This is the only declaration that should pass
Dim gGGG as String
End Sub";

            AssertVbaFragmentYieldsExpectedInspectionResultCount(inputCode, 10);
        }


        [TestMethod]
        [TestCategory("Inspections")]
        public void UseMeaningfulName_ReturnsResult_NameWithoutVowels()
        {
            const string inputCode =
@"Sub Ffffff()
End Sub";
            AssertVbaFragmentYieldsExpectedInspectionResultCount(inputCode, 1);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UseMeaningfulName_ReturnsResult_NameUnderThreeLetters()
        {
            const string inputCode =
@"Sub Oo()
End Sub";
            AssertVbaFragmentYieldsExpectedInspectionResultCount(inputCode, 1);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UseMeaningfulName_ReturnsResult_NameEndsWithDigit()
        {
            const string inputCode =
@"Sub Foo1()
End Sub";

            AssertVbaFragmentYieldsExpectedInspectionResultCount(inputCode, 1);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UseMeaningfulName_DoesNotReturnsResult_GoodName_LowerCaseVowels()
        {
            const string inputCode =
@"Sub FooBar()
End Sub";

            AssertVbaFragmentYieldsExpectedInspectionResultCount(inputCode, 0);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UseMeaningfulName_DoesNotReturnsResult_GoodName_UpperCaseVowels()
        {
            const string inputCode =
@"Sub FOOBAR()
End Sub";

            AssertVbaFragmentYieldsExpectedInspectionResultCount(inputCode, 0);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UseMeaningfulName_DoesNotReturnsResult_OptionBase()
        {
            const string inputCode =
@"Option Base 1";

            AssertVbaFragmentYieldsExpectedInspectionResultCount(inputCode, 0);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UseMeaningfulName_DoesNotReturnResult_NameWithoutVowels_NameIsInWhitelist()
        {
            const string inputCode =
@"Sub sss()
End Sub";

            AssertVbaFragmentYieldsExpectedInspectionResultCount(inputCode, 0);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UseMeaningfulName_Ignored_DoesNotReturnResult()
        {
            const string inputCode =
@"'@Ignore UseMeaningfulName
Sub Ffffff()
End Sub";

            AssertVbaFragmentYieldsExpectedInspectionResultCount(inputCode, 0);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UseMeaningfulName_IgnoreQuickFixWorks()
        {
            const string inputCode =
@"Sub Ffffff()
End Sub";

            const string expectedCode =
@"'@Ignore UseMeaningfulName
Sub Ffffff()
End Sub";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("VBAProject", ProjectProtection.Unprotected)
                .AddComponent("MyClass", ComponentType.ClassModule, inputCode)
                .Build();
            var component = project.Object.VBComponents[0];
            var vbe = builder.AddProject(project).Build();

            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new UseMeaningfulNameInspection(parser.State, GetInspectionSettings().Object);
            var inspectionResults = inspection.GetInspectionResults();
            
            new IgnoreOnceQuickFix(parser.State, new[] {inspection}).Fix(inspectionResults.First());
            Assert.AreEqual(expectedCode, parser.State.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionType()
        {
            var inspection = new UseMeaningfulNameInspection(null, null);
            Assert.AreEqual(CodeInspectionType.MaintainabilityAndReadabilityIssues, inspection.InspectionType);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionName()
        {
            const string inspectionName = "UseMeaningfulNameInspection";
            var inspection = new UseMeaningfulNameInspection(null, null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }

        private void AssertVbaFragmentYieldsExpectedInspectionResultCount(string inputCode, int expectedCount)
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("VBAProject", ProjectProtection.Unprotected)
                .AddComponent("MyClass", ComponentType.ClassModule, inputCode)
                .Build();
            var vbe = builder.AddProject(project).Build();

            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new UseMeaningfulNameInspection(parser.State, GetInspectionSettings().Object);
            var inspectionResults = inspection.GetInspectionResults();
            Assert.AreEqual(expectedCount, inspectionResults.Count());
        }

        internal static Mock<IPersistanceService<CodeInspectionSettings>> GetInspectionSettings()
        {
            var settings = new Mock<IPersistanceService<CodeInspectionSettings>>();
            settings.Setup(s => s.Load(It.IsAny<CodeInspectionSettings>()))
                .Returns(new CodeInspectionSettings(null, new[]
                {
                    new WhitelistedIdentifierSetting("sss"),
                    new WhitelistedIdentifierSetting("oRange")
                }, true));

            return settings;
        }
    }
}
