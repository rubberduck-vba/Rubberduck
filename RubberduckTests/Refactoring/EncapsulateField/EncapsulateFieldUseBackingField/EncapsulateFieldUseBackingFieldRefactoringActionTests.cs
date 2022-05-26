using NUnit.Framework;
using Rubberduck.Common;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.EncapsulateField;
using Rubberduck.Refactorings.EncapsulateFieldUseBackingField;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;
using RubberduckTests.Settings;
using System;
using System.Collections.Generic;
using System.Linq;

namespace RubberduckTests.Refactoring.EncapsulateField.EncapsulateFieldUseBackingField
{
    [TestFixture]
    public class EncapsulateFieldUseBackingFieldRefactoringActionTests
    {
        private EncapsulateFieldTestSupport Support { get; } = new EncapsulateFieldTestSupport();

        [TestCase(false, "Name")]
        [TestCase(true, "Name")]
        [TestCase(false, null)]
        [TestCase(true, null)]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(EncapsulateFieldUseBackingFieldRefactoringAction))]
        public void EncapsulatePublicField(bool isReadOnly, string propertyIdentifier)
        {
            var target = "fizz";
            var inputCode = $"Public {target} As Integer";

            EncapsulateFieldUseBackingFieldModel modelBuilder(RubberduckParserState state, EncapsulateFieldTestsResolver resolver)
            {
                var field = state.DeclarationFinder.MatchName(target).Single();
                var fieldModel = new FieldEncapsulationModel(field as VariableDeclaration, isReadOnly, propertyIdentifier);

                var modelFactory = resolver.Resolve<IEncapsulateFieldUseBackingFieldModelFactory>();
                return modelFactory.Create( new List<FieldEncapsulationModel>() { fieldModel });
            }

            var refactoredCode = RefactoredCode(inputCode, modelBuilder);

            var resultPropertyIdentifier = propertyIdentifier ?? target.CapitalizeFirstLetter();

            var backingField = propertyIdentifier != null
                ? target
                : $"{target}1";

            StringAssert.Contains($"Public Property Get {resultPropertyIdentifier}()", refactoredCode);
            StringAssert.Contains($"{resultPropertyIdentifier} = {backingField}", refactoredCode);

            if (isReadOnly)
            {
                StringAssert.DoesNotContain($"Public Property Let {resultPropertyIdentifier}(", refactoredCode);
                StringAssert.DoesNotContain($"{backingField} = ", refactoredCode);
            }
            else
            {
                StringAssert.Contains($"Public Property Let {resultPropertyIdentifier}(", refactoredCode);
                StringAssert.Contains($"{backingField} = ", refactoredCode);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(EncapsulateFieldUseBackingFieldRefactoringAction))]
        public void EmptyTargetSet()
        {
            var target = "fizz";
            var inputCode = $"Public {target} As Integer";

            EncapsulateFieldUseBackingFieldModel modelBuilder(RubberduckParserState state, EncapsulateFieldTestsResolver resolver)
            {
                var modelFactory = resolver.Resolve<IEncapsulateFieldUseBackingFieldModelFactory>();
                return modelFactory.Create(Enumerable.Empty<FieldEncapsulationModel>());
            }

            var refactoredCode = RefactoredCode(inputCode, modelBuilder);
            Assert.AreEqual(refactoredCode, inputCode);
        }

        [TestCase(true)]
        [TestCase(false)]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(EncapsulateFieldUseBackingFieldRefactoringAction))]
        public void RespectsGroupRelatedPropertiesIndenterSetting(bool groupRelatedProperties)
        {
            var inputCode =
@"
Public mTestField As Long
Public mTestField1 As Long
Public mTestField2 As Long
";

            EncapsulateFieldUseBackingFieldModel modelBuilder(RubberduckParserState state, EncapsulateFieldTestsResolver resolver)
            {
                var field = state.AllUserDeclarations.Single(d => d.IdentifierName.Equals("mTestField"));
                var fieldModel = new FieldEncapsulationModel(field as VariableDeclaration, false);

                var modelFactory = resolver.Resolve<IEncapsulateFieldUseBackingFieldModelFactory>();
                var model = modelFactory.Create(new List<FieldEncapsulationModel>() { fieldModel });
                foreach (var candidate in model.EncapsulationCandidates)
                {
                    candidate.EncapsulateFlag = true;
                }
                return model;
            }

            var testIndenter = new Indenter(null, () =>
            {
                var s = IndenterSettingsTests.GetMockIndenterSettings();
                s.VerticallySpaceProcedures = true;
                s.LinesBetweenProcedures = 1;
                s.GroupRelatedProperties = groupRelatedProperties;
                return s;
            });

            var refactoredCode = RefactoredCode(inputCode, modelBuilder, testIndenter);

            var lines = refactoredCode.Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
            var expectedGrouped = new[]
            {
                "Public Property Get TestField() As Long",
                "TestField = mTestField",
                "End Property",
                "Public Property Let TestField(ByVal RHS As Long)",
                "mTestField = RHS",
                "End Property",
                "",
                "Public Property Get TestField1() As Long",
                "TestField1 = mTestField1",
                "End Property",
                "Public Property Let TestField1(ByVal RHS As Long)",
                "mTestField1 = RHS",
                "End Property",
                "",
                "Public Property Get TestField2() As Long",
                "TestField2 = mTestField2",
                "End Property",
                "Public Property Let TestField2(ByVal RHS As Long)",
                "mTestField2 = RHS",
                "End Property",
                "",
            };

            var expectedNotGrouped = new[]
            {
                "Public Property Get TestField() As Long",
                "TestField = mTestField",
                "End Property",
                "",
                "Public Property Let TestField(ByVal RHS As Long)",
                "mTestField = RHS",
                "End Property",
                "",
                "Public Property Get TestField1() As Long",
                "TestField1 = mTestField1",
                "End Property",
                "",
                "Public Property Let TestField1(ByVal RHS As Long)",
                "mTestField1 = RHS",
                "End Property",
                "",
                "Public Property Get TestField2() As Long",
                "TestField2 = mTestField2",
                "End Property",
                "",
                "Public Property Let TestField2(ByVal RHS As Long)",
                "mTestField2 = RHS",
                "End Property",
                "",
            };

            var idx = 0;

            IList<string> expected = groupRelatedProperties
                ? expectedGrouped.ToList()
                : expectedNotGrouped.ToList();

            var refactoredLinesOfInterest = lines.SkipWhile(rl => !rl.Contains(expected[0]));

            Assert.IsTrue(refactoredLinesOfInterest.Any());

            foreach (var line in refactoredLinesOfInterest)
            {
                if (!line.Contains("="))
                {
                    StringAssert.AreEqualIgnoringCase(expected[idx], line);
                }
                idx++;
            }
        }

        private string RefactoredCode(string inputCode, Func<RubberduckParserState, EncapsulateFieldTestsResolver, EncapsulateFieldUseBackingFieldModel> modelBuilder, IIndenter indenter = null)
        {
            return Support.RefactoredCode<EncapsulateFieldUseBackingFieldRefactoringAction, EncapsulateFieldUseBackingFieldModel>(inputCode, modelBuilder, indenter);
        }
    }
}
