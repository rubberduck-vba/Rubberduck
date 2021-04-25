using System.Linq;
using NUnit.Framework;
using Moq;
using Rubberduck.SettingsProvider;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;
using Rubberduck.CodeAnalysis.Settings;
using Rubberduck.Parsing.VBA;
using System.Collections.Generic;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class UseMeaningfulNameInspectionTests : InspectionTestsBase
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
            Assert.AreEqual(0, InspectionResultsForModules(("TestModule", inputCode, ComponentType.StandardModule)).Count());
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

            Assert.AreEqual(10, InspectionResultsForModules(("MyClass", inputCode, ComponentType.ClassModule)).Count());
        }


        [Test]
        [Category("Inspections")]
        public void UseMeaningfulName_ReturnsResult_NameWithoutVowels()
        {
            const string inputCode =
@"Sub Ffffff()
End Sub";
            Assert.AreEqual(1, InspectionResultsForModules(("MyClass", inputCode, ComponentType.ClassModule)).Count());
        }

        [Test]
        [Category("Inspections")]
        public void UseMeaningfulName_ReturnsResult_NameUnderThreeLetters()
        {
            const string inputCode =
@"Sub Oo()
End Sub";
            Assert.AreEqual(1, InspectionResultsForModules(("MyClass", inputCode, ComponentType.ClassModule)).Count());
        }

        [Test]
        [Category("Inspections")]
        public void UseMeaningfulName_ReturnsResult_NameEndsWithDigit()
        {
            const string inputCode =
@"Sub Foo1()
End Sub";

            Assert.AreEqual(1, InspectionResultsForModules(("MyClass", inputCode, ComponentType.ClassModule)).Count());
        }

        [Test]
        [Category("Inspections")]
        public void UseMeaningfulName_DoesNotReturnsResult_GoodName_LowerCaseVowels()
        {
            const string inputCode =
@"Sub FooBar()
End Sub";

            Assert.AreEqual(0, InspectionResultsForModules(("MyClass", inputCode, ComponentType.ClassModule)).Count());
        }

        [Test]
        [Category("Inspections")]
        public void UseMeaningfulName_DoesNotReturnsResult_GoodName_UpperCaseVowels()
        {
            const string inputCode =
@"Sub FOOBAR()
End Sub";

            Assert.AreEqual(0, InspectionResultsForModules(("MyClass", inputCode, ComponentType.ClassModule)).Count());
        }

        [Test]
        [Category("Inspections")]
        public void UseMeaningfulName_DoesNotReturnsResult_OptionBase()
        {
            const string inputCode =
@"Option Base 1";

            Assert.AreEqual(0, InspectionResultsForModules(("MyClass", inputCode, ComponentType.ClassModule)).Count());
        }

        [Test]
        [Category("Inspections")]
        public void UseMeaningfulName_DoesNotReturnResult_NameWithoutVowels_NameIsInWhitelist()
        {
            const string inputCode =
@"Sub sss()
End Sub";

            Assert.AreEqual(0, InspectionResultsForModules(("MyClass", inputCode, ComponentType.ClassModule)).Count());
        }

        [Test]
        [Category("Inspections")]
        public void UseMeaningfulName_Ignored_DoesNotReturnResult()
        {
            const string inputCode =
@"'@Ignore UseMeaningfulName
Sub Ffffff()
End Sub";

            Assert.AreEqual(0, InspectionResultsForModules(("MyClass", inputCode, ComponentType.ClassModule)).Count());
        }

        [Test]
        [Category("Inspections")]
        public void InspectionName()
        {
            var inspection = new UseMeaningfulNameInspection(null, null);

            Assert.AreEqual(nameof(UseMeaningfulNameInspection), inspection.Name);
        }

        internal static Mock<IConfigurationService<CodeInspectionSettings>> GetInspectionSettings()
        {
            var settings = new Mock<IConfigurationService<CodeInspectionSettings>>();
            settings.Setup(s => s.Read())
                .Returns(new CodeInspectionSettings(Enumerable.Empty<CodeInspectionSetting>(), new[]
                {
                    new WhitelistedIdentifierSetting("sss"),
                    new WhitelistedIdentifierSetting("oRange")
                }, true));

            return settings;
        }

        private new IEnumerable<IInspectionResult> InspectionResultsForModules(params (string name, string content, ComponentType componentType)[] modules)
        {
            var vbe = MockVbeBuilder.BuildFromModules("TestProject", modules, Enumerable.Empty<ReferenceLibrary>());
            return InspectionResults(vbe.Object);
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new UseMeaningfulNameInspection(state, GetInspectionSettings().Object);
        }
    }
}
