using NUnit.Framework;
using Rubberduck.Common;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.EncapsulateField;
using Rubberduck.Refactorings.EncapsulateFieldUseBackingUDTMember;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;
using System;
using System.Collections.Generic;
using System.Linq;

namespace RubberduckTests.Refactoring.EncapsulateField.EncapsulateFieldUseBackingUDTMember
{
    [TestFixture]
    public class EncapsulateFieldUseBackingUDTMemberRefactoringActionTests
    {
        private EncapsulateFieldTestSupport Support { get; } = new EncapsulateFieldTestSupport();

        [TestCase(false, "Name")]
        [TestCase(true, "Name")]
        [TestCase(false, null)]
        [TestCase(true, null)]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(EncapsulateFieldUseBackingUDTMemberRefactoringAction))]
        public void EncapsulatePublicFields(bool isReadOnly, string propertyIdentifier)
        {
            var target = "fizz";
            var inputCode = $"Public {target} As Integer";

            EncapsulateFieldUseBackingUDTMemberModel modelBuilder(RubberduckParserState state, EncapsulateFieldTestsResolver resolver)
            {
                var field = state.DeclarationFinder.MatchName(target).Single();
                var fieldModel = new FieldEncapsulationModel(field as VariableDeclaration, isReadOnly);

                var modelFactory = resolver.Resolve<IEncapsulateFieldUseBackingUDTMemberModelFactory>();
                return modelFactory.Create(new List<FieldEncapsulationModel>() { fieldModel });
            }

            var refactoredCode = RefactoredCode(inputCode, modelBuilder);

            var resultPropertyIdentifier = target.CapitalizeFirstLetter();

            var backingFieldexpression = propertyIdentifier != null
                ? $"this.{resultPropertyIdentifier}"
                : $"this.{resultPropertyIdentifier}";

            StringAssert.Contains($"T{MockVbeBuilder.TestModuleName}", refactoredCode);
            StringAssert.Contains($"Public Property Get {resultPropertyIdentifier}()", refactoredCode);
            StringAssert.Contains($"{resultPropertyIdentifier} = {backingFieldexpression}", refactoredCode);

            if (isReadOnly)
            {
                StringAssert.DoesNotContain($"Public Property Let {resultPropertyIdentifier}(", refactoredCode);
                StringAssert.DoesNotContain($"{backingFieldexpression} = ", refactoredCode);
            }
            else
            {
                StringAssert.Contains($"Public Property Let {resultPropertyIdentifier}(", refactoredCode);
                StringAssert.Contains($"{backingFieldexpression} = ", refactoredCode);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(EncapsulateFieldUseBackingUDTMemberRefactoringAction))]
        public void EncapsulatePublicFields_ExistingObjectStateUDT()
        {
            var inputCode =
$@"
Option Explicit

Private Type T{MockVbeBuilder.TestModuleName}
    FirstValue As Long
    SecondValue As String
End Type

Private this As T{MockVbeBuilder.TestModuleName}

Public thirdValue As Integer

Public bazz As String";

            EncapsulateFieldUseBackingUDTMemberModel modelBuilder(RubberduckParserState state, EncapsulateFieldTestsResolver resolver)
            {
                var firstValueField = state.DeclarationFinder.MatchName("thirdValue").Single(d => d.DeclarationType.HasFlag(DeclarationType.Variable));
                var bazzField = state.DeclarationFinder.MatchName("bazz").Single();
                var fieldModelfirstValueField = new FieldEncapsulationModel(firstValueField as VariableDeclaration);
                var fieldModelfirstbazzField = new FieldEncapsulationModel(bazzField as VariableDeclaration);
                var inputList = new List<FieldEncapsulationModel>() { fieldModelfirstValueField, fieldModelfirstbazzField };

                var modelFactory = resolver.Resolve<IEncapsulateFieldUseBackingUDTMemberModelFactory>();
                return modelFactory.Create(inputList);
            }

            var refactoredCode = RefactoredCode(inputCode, modelBuilder);

            StringAssert.Contains($" ThirdValue As Integer", refactoredCode);
            StringAssert.Contains($"Property Get ThirdValue", refactoredCode);
            StringAssert.Contains($" ThirdValue = this.ThirdValue", refactoredCode);

            StringAssert.Contains($"Property Let ThirdValue", refactoredCode);
            StringAssert.Contains($" this.ThirdValue =", refactoredCode);

            StringAssert.Contains($" Bazz As String", refactoredCode);
            StringAssert.Contains($"Property Get Bazz", refactoredCode);
            StringAssert.Contains($" Bazz = this.Bazz", refactoredCode);

            StringAssert.Contains($"Property Let Bazz", refactoredCode);
            StringAssert.Contains($" this.Bazz =", refactoredCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(EncapsulateFieldUseBackingUDTMemberRefactoringAction))]
        public void EncapsulatePublicFields_ExistingUDT()
        {
            var inputCode =
$@"
Option Explicit

Private Type TestType
    FirstValue As Long
    SecondValue As String
End Type

Private this As TestType

Public thirdValue As Integer

Public bazz As String";

            EncapsulateFieldUseBackingUDTMemberModel modelBuilder(RubberduckParserState state, EncapsulateFieldTestsResolver resolver)
            {
                var thirdValueField = state.DeclarationFinder.MatchName("thirdValue").Single(d => d.DeclarationType.HasFlag(DeclarationType.Variable));
                var bazzField = state.DeclarationFinder.MatchName("bazz").Single();
                var fieldModelThirdValueField = new FieldEncapsulationModel(thirdValueField as VariableDeclaration);
                var fieldModelBazzField = new FieldEncapsulationModel(bazzField as VariableDeclaration);

                var inputList = new List<FieldEncapsulationModel>() { fieldModelThirdValueField, fieldModelBazzField };

                var targetUDT = state.DeclarationFinder.MatchName("this").Single(d => d.DeclarationType.HasFlag(DeclarationType.Variable));

                var modelFactory = resolver.Resolve<IEncapsulateFieldUseBackingUDTMemberModelFactory>();
                return modelFactory.Create(inputList, targetUDT);
            }

            var refactoredCode = RefactoredCode(inputCode, modelBuilder);

            StringAssert.DoesNotContain($"T{ MockVbeBuilder.TestModuleName}", refactoredCode);

            StringAssert.Contains($" ThirdValue As Integer", refactoredCode);
            StringAssert.Contains($"Property Get ThirdValue", refactoredCode);
            StringAssert.Contains($" ThirdValue = this.ThirdValue", refactoredCode);

            StringAssert.Contains($"Property Let ThirdValue", refactoredCode);
            StringAssert.Contains($" this.ThirdValue =", refactoredCode);

            StringAssert.Contains($" Bazz As String", refactoredCode);
            StringAssert.Contains($"Property Get Bazz", refactoredCode);
            StringAssert.Contains($" Bazz = this.Bazz", refactoredCode);

            StringAssert.Contains($"Property Let Bazz", refactoredCode);
            StringAssert.Contains($" this.Bazz =", refactoredCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(EncapsulateFieldUseBackingUDTMemberRefactoringAction))]
        public void EncapsulatePublicFields_NestedPathForPrivateUDTField()
        {
            var inputCode =
$@"
Option Explicit

Private Type TVehicle
    Wheels As Integer
End Type

Private mVehicle As TVehicle
";

            EncapsulateFieldUseBackingUDTMemberModel modelBuilder(RubberduckParserState state, EncapsulateFieldTestsResolver resolver)
            {
                var mVehicleField = state.DeclarationFinder.UserDeclarations(DeclarationType.Variable).Single(d => d.IdentifierName.Equals("mVehicle"));
                var fieldModelMVehicleField = new FieldEncapsulationModel(mVehicleField as VariableDeclaration, false, "Vehicle");

                var inputList = new List<FieldEncapsulationModel>() { fieldModelMVehicleField };

                var modelFactory = resolver.Resolve<IEncapsulateFieldUseBackingUDTMemberModelFactory>();
                return modelFactory.Create(inputList);
            }

            var refactoredCode = RefactoredCode(inputCode, modelBuilder);

            StringAssert.Contains($"T{ MockVbeBuilder.TestModuleName}", refactoredCode);

            StringAssert.Contains($" Vehicle As TVehicle", refactoredCode);
            StringAssert.Contains($"Property Get Wheels", refactoredCode);
            StringAssert.Contains($" Wheels = this.Vehicle.Wheels", refactoredCode);

            StringAssert.Contains($"Property Let Wheels", refactoredCode);
            StringAssert.Contains($" this.Vehicle.Wheels =", refactoredCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(EncapsulateFieldUseBackingUDTMemberRefactoringAction))]
        public void EncapsulatePublicFields_DifferentLevelForNestedProperties()
        {
            var inputCode =
$@"
Option Explicit

Private Type FirstType
    FirstValue As Integer
End Type

Private Type SecondType
    SecondValue As Integer
    FirstTypeValue As FirstType
End Type

Private Type ThirdType
    ThirdValue As SecondType
End Type

Private mTest As ThirdType
";

            EncapsulateFieldUseBackingUDTMemberModel modelBuilder(RubberduckParserState state, EncapsulateFieldTestsResolver resolver)
            {
                var mTestField = state.DeclarationFinder.UserDeclarations(DeclarationType.Variable).Single(d => d.IdentifierName.Equals("mTest"));
                var fieldModelMTest = new FieldEncapsulationModel(mTestField as VariableDeclaration, false);

                var inputList = new List<FieldEncapsulationModel>() { fieldModelMTest };

                var modelFactory = resolver.Resolve<IEncapsulateFieldUseBackingUDTMemberModelFactory>();
                return modelFactory.Create(inputList);
            }

            var refactoredCode = RefactoredCode(inputCode, modelBuilder);

            StringAssert.Contains($"T{ MockVbeBuilder.TestModuleName}", refactoredCode);

            StringAssert.Contains($" Test As ThirdType", refactoredCode);
            StringAssert.Contains($"Property Get FirstValue", refactoredCode);
            StringAssert.Contains($"Property Get SecondValue", refactoredCode);

            StringAssert.Contains($" this.Test.ThirdValue.FirstTypeValue.FirstValue =", refactoredCode);
            StringAssert.Contains($" this.Test.ThirdValue.SecondValue =", refactoredCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(EncapsulateFieldUseBackingUDTMemberRefactoringAction))]
        public void EmptyTargetSet_Throws()
        {
            var inputCode = $"Public fizz As Integer";

            EncapsulateFieldUseBackingUDTMemberModel modelBuilder(RubberduckParserState state, EncapsulateFieldTestsResolver resolver)
            {
                var modelFactory = resolver.Resolve<IEncapsulateFieldUseBackingUDTMemberModelFactory>();
                return modelFactory.Create(Enumerable.Empty<FieldEncapsulationModel>());
            }

            Assert.Throws<System.ArgumentException>(() => RefactoredCode(inputCode, modelBuilder));
        }

        [TestCase("notAUserDefinedTypeField")]
        [TestCase("notAnOption")]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(EncapsulateFieldUseBackingUDTMemberRefactoringAction))]
        public void InvalidObjectStateTarget_Throws(string objectStateTargetIdentifier)
        {
            var inputCode =
$@"
Option Explicit

Public Type CannotUseThis
    FirstValue As Long
    SecondValue As String
End Type

Private Type TestType
    FirstValue As Long
    SecondValue As String
End Type

Private this As TestType

Public notAnOption As CannotUseThis

Public notAUserDefinedTypeField As String";

            EncapsulateFieldUseBackingUDTMemberModel modelBuilder(RubberduckParserState state, EncapsulateFieldTestsResolver resolver)
            {
                var invalidTarget = state.DeclarationFinder.MatchName(objectStateTargetIdentifier).Single(d => d.DeclarationType.HasFlag(DeclarationType.Variable));
                var fieldModel = new FieldEncapsulationModel(invalidTarget as VariableDeclaration);

                var modelFactory = resolver.Resolve<IEncapsulateFieldUseBackingUDTMemberModelFactory>();
                return modelFactory.Create(new List<FieldEncapsulationModel>() { fieldModel }, invalidTarget);
            }

            Assert.Throws<ArgumentException>(() => RefactoredCode(inputCode, modelBuilder));
        }

        private string RefactoredCode(string inputCode, Func<RubberduckParserState, EncapsulateFieldTestsResolver, EncapsulateFieldUseBackingUDTMemberModel> modelBuilder)
        {
            return Support.RefactoredCode<EncapsulateFieldUseBackingUDTMemberRefactoringAction, EncapsulateFieldUseBackingUDTMemberModel>(inputCode, modelBuilder);
        }
    }
}
