using System.Linq;
using System.Threading;
using NUnit.Framework;
using Moq;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Settings;
using Rubberduck.SettingsProvider;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class UseMeaningfulNameInspectionTests
    {
        [Test]
        [Category("Inspections")]
        public void UseMeaningfulName_NoResultForLineNumberLabels()
        {
            const string inputCode = @"
Sub DoSomething()
10 Debug.Print 42
End Sub
";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);

            using(var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new UseMeaningfulNameInspection(state, GetInspectionSettings().Object);
                var inspectionResults = inspection.GetInspectionResults().Where(i => i.Target.DeclarationType == DeclarationType.LineLabel);

                Assert.IsFalse(inspectionResults.Any());
            }
        }

        [Test]
        [Category("Inspections")]
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


        [Test]
        [Category("Inspections")]
        public void UseMeaningfulName_ReturnsResult_NameWithoutVowels()
        {
            const string inputCode =
@"Sub Ffffff()
End Sub";
            AssertVbaFragmentYieldsExpectedInspectionResultCount(inputCode, 1);
        }

        [Test]
        [Category("Inspections")]
        public void UseMeaningfulName_ReturnsResult_NameUnderThreeLetters()
        {
            const string inputCode =
@"Sub Oo()
End Sub";
            AssertVbaFragmentYieldsExpectedInspectionResultCount(inputCode, 1);
        }

        [Test]
        [Category("Inspections")]
        public void UseMeaningfulName_ReturnsResult_NameEndsWithDigit()
        {
            const string inputCode =
@"Sub Foo1()
End Sub";

            AssertVbaFragmentYieldsExpectedInspectionResultCount(inputCode, 1);
        }

        [Test]
        [Category("Inspections")]
        public void UseMeaningfulName_DoesNotReturnsResult_GoodName_LowerCaseVowels()
        {
            const string inputCode =
@"Sub FooBar()
End Sub";

            AssertVbaFragmentYieldsExpectedInspectionResultCount(inputCode, 0);
        }

        [Test]
        [Category("Inspections")]
        public void UseMeaningfulName_DoesNotReturnsResult_GoodName_UpperCaseVowels()
        {
            const string inputCode =
@"Sub FOOBAR()
End Sub";

            AssertVbaFragmentYieldsExpectedInspectionResultCount(inputCode, 0);
        }

        [Test]
        [Category("Inspections")]
        public void UseMeaningfulName_DoesNotReturnsResult_OptionBase()
        {
            const string inputCode =
@"Option Base 1";

            AssertVbaFragmentYieldsExpectedInspectionResultCount(inputCode, 0);
        }

        [Test]
        [Category("Inspections")]
        public void UseMeaningfulName_DoesNotReturnResult_NameWithoutVowels_NameIsInWhitelist()
        {
            const string inputCode =
@"Sub sss()
End Sub";

            AssertVbaFragmentYieldsExpectedInspectionResultCount(inputCode, 0);
        }

        [Test]
        [Category("Inspections")]
        public void UseMeaningfulName_Ignored_DoesNotReturnResult()
        {
            const string inputCode =
@"'@Ignore UseMeaningfulName
Sub Ffffff()
End Sub";

            AssertVbaFragmentYieldsExpectedInspectionResultCount(inputCode, 0);
        }

        [Test]
        [Category("Inspections")]
        public void InspectionType()
        {
            var inspection = new UseMeaningfulNameInspection(null, null);
            Assert.AreEqual(CodeInspectionType.MaintainabilityAndReadabilityIssues, inspection.InspectionType);
        }

        [Test]
        [Category("Inspections")]
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

            using(var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new UseMeaningfulNameInspection(state, GetInspectionSettings().Object);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(expectedCount, inspectionResults.Count());
            }
        }

        internal static Mock<IPersistanceService<CodeInspectionSettings>> GetInspectionSettings()
        {
            var settings = new Mock<IPersistanceService<CodeInspectionSettings>>();
            settings.Setup(s => s.Load(It.IsAny<CodeInspectionSettings>()))
                .Returns(new CodeInspectionSettings(Enumerable.Empty<CodeInspectionSetting>(), new[]
                {
                    new WhitelistedIdentifierSetting("sss"),
                    new WhitelistedIdentifierSetting("oRange")
                }, true));

            return settings;
        }
    }
}
