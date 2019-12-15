using NUnit.Framework;
using Moq;
using Rubberduck.Refactorings.EncapsulateField;
using RubberduckTests.Mocks;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using System.Linq;
using System.Collections.Generic;
using Rubberduck.VBEditor;
using Rubberduck.Refactorings;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.Utility;
using Rubberduck.Parsing.Symbols;
using System;

namespace RubberduckTests.Refactoring.EncapsulateField
{
    [TestFixture]
    public class EncapsulateFieldValidatorTests : InteractiveRefactoringTestBase<IEncapsulateFieldPresenter, EncapsulateFieldModel>
    {
        private EncapsulateFieldTestSupport Support { get; } = new EncapsulateFieldTestSupport();

        [TestCase("fizz", "_Fizz", false)]
        [TestCase("fizz", "FizzProp", true)]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void VBAIdentifier_Property(string originalFieldName, string newPropertyName, bool expectedResult)
        {
            string inputCode =
$@"Public {originalFieldName} As String";

            var encapsulatedField = Support.RetrieveEncapsulatedField(inputCode, originalFieldName);

            encapsulatedField.PropertyName = newPropertyName;
            var field = encapsulatedField as IEncapsulateFieldCandidateValidations;
            Assert.AreEqual(expectedResult, field.HasVBACompliantPropertyIdentifier);
        }

        [TestCase("fizz", "_Fizz", false)]
        [TestCase("fizz", "FizzProp", true)]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void VBAIdentifier_FieldName(string originalFieldName, string newFieldName, bool expectedResult)
        {
            string inputCode =
$@"Public {originalFieldName} As String";

            var encapsulatedField = Support.RetrieveEncapsulatedField(inputCode, originalFieldName);

            encapsulatedField.FieldIdentifier = newFieldName;
            var field = encapsulatedField as IEncapsulateFieldCandidateValidations;
            Assert.AreEqual(expectedResult, field.HasVBACompliantFieldIdentifier);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulatePrivateField_ReadOnlyRequiresSet()
        {
            const string inputCode =
                @"|Private fizz As Collection";

            const string expectedCode =
                @"Private fizz As Collection

Public Property Get Name() As Collection
    Set Name = fizz
End Property
";
            var presenterAction = Support.SetParametersForSingleTarget("fizz", "Name", isReadonly: true);
            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            Assert.AreEqual(expectedCode.Trim(), actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void PropertyNameNotDuplicated()
        {
            const string inputCode =
                @"Public var|iable As Integer, variable1 As Long, variable2 As Integer";

            var userInput = new UserInputDataObject()
                .AddAttributeSet("variable")
                .AddAttributeSet("variable1")
                .AddAttributeSet("variable2");

            var presenterAction = Support.SetParameters(userInput);
            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            StringAssert.Contains("Public Property Get Variable() As Integer", actualCode);
            StringAssert.Contains("Variable = variable_1", actualCode);
            StringAssert.Contains("Public Property Get Variable1() As Long", actualCode);
            StringAssert.Contains("Variable1 = variable1_1", actualCode);
            StringAssert.Contains("Public Property Get Variable2() As Integer", actualCode);
            StringAssert.Contains("Variable2 = variable2_1", actualCode);
            StringAssert.DoesNotContain("Public Property Get Variable3() As Integer", actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void UDTMemberPropertyDefaultsToValidValue()
        {
            string inputCode =
$@"
Private Type TBar
    First As String
    Second As Long
End Type

Public myBar As TBar

Private Function First() As String
    First = myBar.First
End Function";

            var encapsulatedField = Support.RetrieveEncapsulatedField(inputCode, "First", DeclarationType.UserDefinedTypeMember);
            var validation = encapsulatedField as IEncapsulateFieldCandidateValidations;
            var result = validation.HasConflictingPropertyIdentifier;
            Assert.AreEqual(true, validation.HasConflictingPropertyIdentifier);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void FieldNameDefaultsToNonConflictName()
        {
            string inputCode =
$@"Public fizz As String

            Private fizzle As String

            'fizz1 is the initial default name for encapsulating 'fizz'            
            Public Property Get Fizz_1() As String
                Fizz_1 = fizzle
            End Property

            Public Property Let Fizz_1(ByVal value As String)
                fizzle = value
            End Property
            ";
            var encapsulatedField = Support.RetrieveEncapsulatedField(inputCode, "fizz");
            Assert.IsTrue(encapsulatedField.HasValidEncapsulationAttributes);
        }

        [TestCase("Name")]
        [TestCase("mName")]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void UserEntersConflictingName(string userModifiedPropertyName)
        {
            string inputCode =
$@"Public fizz As String

            Private mName As String

            Public Property Get Name() As String
                Name = mName
            End Property

            Public Property Let Name(ByVal value As String)
                mName = value
            End Property
            ";

            var fieldUT = "fizz";
            var userInput = new UserInputDataObject("fizz", userModifiedPropertyName, true);

            var presenterAction = Support.SetParameters(userInput);

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _).Object;
            var model = Support.RetrieveUserModifiedModelPriorToRefactoring(vbe, fieldUT, DeclarationType.Variable, presenterAction);

            Assert.IsFalse(model["fizz"].HasValidEncapsulationAttributes);
        }

        [TestCase("Number", "Bazzle", true, true)]
        [TestCase("Number", "Number", false, false)]
        [TestCase("Test", "Number", false, true)]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void UserModificationIsExistingPropertyNameConflicts(string fizz_modifiedPropertyName, string bazz_modifiedPropertyName, bool fizz_expectedResult, bool bazz_expectedResult)
        {
            string inputCode =
$@"Public fizz As Integer
Public bazz As Integer
Public buzz As Integer

Private mTest As Integer

Public Property Get Test() As Integer
    Test = mTest
End Property";

            var fieldUT = "fizz";
            var userInput = new UserInputDataObject()
                .AddAttributeSet(fieldUT, fizz_modifiedPropertyName, true)
                .AddAttributeSet("bazz", bazz_modifiedPropertyName, true);


            var presenterAction = Support.SetParameters(userInput);

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _).Object;
            var model = Support.RetrieveUserModifiedModelPriorToRefactoring(vbe, fieldUT, DeclarationType.Variable, presenterAction);

            Assert.AreEqual(fizz_expectedResult, model["fizz"].HasValidEncapsulationAttributes, "fizz failed");
            Assert.AreEqual(bazz_expectedResult, model["bazz"].HasValidEncapsulationAttributes, "bazz failed");
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulateMultipleFields_UDTConflicts()
        {
            string inputCode =
$@"
Private Type TBar
    First As Long
    Second As String
End Type

Public this As TBar

Public that As TBar
";
            var fieldUT = "this";
            var userInput = new UserInputDataObject()
                .AddAttributeSet(fieldUT, "That", true);

            var presenterAction = Support.SetParameters(userInput);

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _).Object;
            var model = Support.RetrieveUserModifiedModelPriorToRefactoring(vbe, fieldUT, DeclarationType.Variable, presenterAction);

            Assert.AreEqual(false, model[fieldUT].HasValidEncapsulationAttributes, $"{fieldUT} failed");
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void PropertyNameConflictsWithModuleVariable()
        {
            string inputCode =
$@"
Public longValue As Long

Public wholeNumber As String
";
            var fieldUT = "longValue";
            var userInput = new UserInputDataObject()
                .AddAttributeSet(fieldUT, "WholeNumber", true);

            var presenterAction = Support.SetParameters(userInput);

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _).Object;
            var model = Support.RetrieveUserModifiedModelPriorToRefactoring(vbe, fieldUT, DeclarationType.Variable, presenterAction);

            Assert.AreEqual(false, model[fieldUT].HasValidEncapsulationAttributes, $"{fieldUT} failed");
        }

        [TestCase("Dim test As String", "arg")] //Local variable
        [TestCase(@"Const test As String = ""Foo""", "arg")] //Local constant
        [TestCase(@"Const localTest As String = ""Foo""", "test")] //parameter
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void TargetNameUsedForLimitedScopeDeclarations(string localDeclaration, string parameter)
        {
            string inputCode =
                $@"

Private te|st As Long

Private Function Foo({parameter} As String) As String
    {localDeclaration}
    test = test & ""Foo""
    Foo = test
End Function
";
            var presenterAction = Support.UserAcceptsDefaults();
            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            StringAssert.Contains("Test", actualCode);
            StringAssert.Contains("test_1", actualCode);
            StringAssert.DoesNotContain("Test_1", actualCode);
        }

//        [Test]
//        [Category("Refactorings")]
//        [Category("Encapsulate Field")]
//        public void TargetNameDuplicatedAsLocalVariable_Array()
//        {
//            const string inputCode =
//                @"

//Private Bou|ndedArray(1 To 100) As Byte

//Private Function Foo() As Variant
//    Dim BoundedArray(1 To 3) As Integer
//    BoundedArray(1) = 7
//    BoundedArray(2) = 8
//    BoundedArray(3) = 9
//    Foo = BoundedArray
//End Function
//";
//            var presenterAction = Support.UserAcceptsDefaults();
//            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
//            StringAssert.Contains("BoundedArray", actualCode);
//            StringAssert.Contains("boundedArray_1", actualCode);
//            StringAssert.DoesNotContain("BoundedArray_1", actualCode);
//        }

        protected override IRefactoring TestRefactoring(IRewritingManager rewritingManager, RubberduckParserState state, IRefactoringPresenterFactory factory, ISelectionService selectionService)
        {
            return Support.SupportTestRefactoring(rewritingManager, state, factory, selectionService);
        }
    }
}
