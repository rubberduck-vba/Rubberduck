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

        [TestCase(true)]
        [TestCase(false)]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(EncapsulateFieldUseBackingUDTMemberRefactoringAction))]
        //Forces the 'Read Only" flag to false if there are external write references.
        //If invoked via the refactoring dialog, the user will not given the option to choose a Readonly implementation
        //if there are external write references.
        public void ExternallyReferencedField_IgnoresReadOnlyFlagIfExternalWriteReferencesExist(bool isReadOnly)
        {
            var propertyIdentifier = "Name";
            var codeClass1 =
@"Public fizz As Integer

Sub Foo()
    fizz = 1
End Sub";
            var codeClass2 =
@"Sub Foo()
    Dim theClass As Class1
    Set theClass = new Class1
    theClass.fizz = 0
    Bar theClass.fizz
End Sub

Sub Bar(ByVal v As Integer)
End Sub";
            (string name, string content, ComponentType compType) class1 = ("Class1", codeClass1, ComponentType.ClassModule);
            (string name, string content, ComponentType compType) class2 = ("Class2", codeClass2, ComponentType.ClassModule);

            EncapsulateFieldUseBackingUDTMemberModel modelBuilder(RubberduckParserState state, EncapsulateFieldTestsResolver resolver)
            {

                var field = state.DeclarationFinder.MatchName("fizz").Single();
                var fieldModel = new FieldEncapsulationModel(field as VariableDeclaration, isReadOnly, propertyIdentifier);

                var modelFactory = resolver.Resolve<IEncapsulateFieldUseBackingUDTMemberModelFactory>();
                return modelFactory.Create(new List<FieldEncapsulationModel>() { fieldModel });
            }

            var vbe = MockVbeBuilder.BuildFromModules((class1.name, class1.content, class1.compType), (class2.name, class2.content, class2.compType)).Object;
            var refactoredCode = Support.RefactoredCode<EncapsulateFieldUseBackingUDTMemberRefactoringAction, EncapsulateFieldUseBackingUDTMemberModel>(vbe, modelBuilder);

            StringAssert.Contains("this.Name = RHS", refactoredCode["Class1"]);
            StringAssert.Contains("Name = 1", refactoredCode["Class1"]);
            StringAssert.Contains($"theClass.{propertyIdentifier} = 0", refactoredCode["Class2"]);
            StringAssert.Contains($"Bar theClass.{propertyIdentifier}", refactoredCode["Class2"]);
            StringAssert.DoesNotContain("fizz", refactoredCode["Class1"]);
            StringAssert.DoesNotContain("fizz", refactoredCode["Class2"]);
        }

        private string RefactoredCode(string inputCode, Func<RubberduckParserState, EncapsulateFieldTestsResolver, EncapsulateFieldUseBackingUDTMemberModel> modelBuilder)
        {
            return Support.RefactoredCode<EncapsulateFieldUseBackingUDTMemberRefactoringAction, EncapsulateFieldUseBackingUDTMemberModel>(inputCode, modelBuilder);
        }
    }
}
