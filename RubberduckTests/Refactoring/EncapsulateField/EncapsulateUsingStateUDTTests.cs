using NUnit.Framework;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.EncapsulateField;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.Utility;
using RubberduckTests.Mocks;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RubberduckTests.Refactoring.EncapsulateField
{
    [TestFixture]
    public class EncapsulateUsingStateUDTTests : InteractiveRefactoringTestBase<IEncapsulateFieldPresenter, EncapsulateFieldModel>
    {
        private EncapsulateFieldTestSupport Support { get; } = new EncapsulateFieldTestSupport();

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulatePrivateFieldAsUDT()
        {
            const string inputCode =
                @"|Private fizz As Integer";

            var presenterAction = Support.SetParametersForSingleTarget("fizz", "Name", asUDT: true);
            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            StringAssert.Contains("Name As Integer", actualCode);
            StringAssert.Contains("this.Name = value", actualCode);
        }

        [TestCase("Public")]
        [TestCase("Private")]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void UserDefinedType_UserAcceptsDefaults_ConflictWithStateUDT(string accessibility)
        {
            string inputCode =
$@"
Private Type TBar
    First As String
    Second As Long
End Type

{accessibility} my|Bar As TBar

Private this As Long";


            var presenterAction = Support.UserAcceptsDefaults(convertFieldToUDTMember: true);
            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            StringAssert.Contains("Private this As Long", actualCode);
            StringAssert.Contains($"Private this_1 As {Support.StateUDTDefaultType}", actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void UserDefinedTypeMembers_OnlyEncapsulateUDTMembers()
        {
            string inputCode =
$@"
Private Type TBar
    First As String
    Second As Long
End Type

Private my|Bar As TBar";


            var userInput = new UserInputDataObject()
                .UserSelectsField("myBar");

            userInput.EncapsulateUsingUDTField();

            var presenterAction = Support.SetParameters(userInput);

            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            StringAssert.Contains("this.MyBar.First = value", actualCode);
            StringAssert.Contains($"First = this.MyBar.First", actualCode);
            StringAssert.Contains("this.MyBar.Second = value", actualCode);
            StringAssert.Contains($"Second = this.MyBar.Second", actualCode);
            StringAssert.Contains($"MyBar As TBar", actualCode);
            StringAssert.Contains($"MyBar As TBar", actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void UserDefinedTypeMembers_UDTFieldReferences()
        {
            string inputCode =
$@"
Private Type TBar
    First As String
    Second As Long
End Type

Private my|Bar As TBar

Public Sub Foo(newValue As String)
    myBar.First = newValue
End Sub";


            var userInput = new UserInputDataObject()
                .UserSelectsField("myBar");

            userInput.EncapsulateUsingUDTField();

            var presenterAction = Support.SetParameters(userInput);

            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            StringAssert.Contains("  First = newValue", actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void LoadsExistingUDT()
        {
            string inputCode =
$@"
Private Type TBar
    First As String
    Second As Long
End Type

Private my|Bar As TBar

Public foo As Long
Public bar As String
Public foobar As Byte
";

            var userInput = new UserInputDataObject()
                .UserSelectsField("foo")
                .UserSelectsField("bar")
                .UserSelectsField("foobar");

            userInput.EncapsulateUsingUDTField("myBar");
            //userInput.ObjectStateUDTTargetID = "myBar";

            var presenterAction = Support.SetParameters(userInput);
            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            StringAssert.DoesNotContain($"Private this As {Support.StateUDTDefaultType}", actualCode);
            //StringAssert.Contains($"Private Type {Support.StateUDTDefaultType}", actualCode);
            StringAssert.Contains("Foo As Long", actualCode);
            StringAssert.DoesNotContain("Public foo As Long", actualCode);
            StringAssert.Contains("Bar As String", actualCode);
            StringAssert.DoesNotContain("Public bar As Long", actualCode);
            StringAssert.Contains("Foobar As Byte", actualCode);
            StringAssert.DoesNotContain("Public foobar As Long", actualCode);
            StringAssert.DoesNotContain("MyBar As TBar", actualCode);
            StringAssert.DoesNotContain("Private this As TBar", actualCode);
            StringAssert.Contains("First As String", actualCode);
            StringAssert.Contains("Second As Long", actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void MultipleFields()
        {
            string inputCode =
$@"
Public fo|o As Long
Public bar As String
Public foobar As Byte
";

            var userInput = new UserInputDataObject()
                .UserSelectsField("foo")
                .UserSelectsField("bar")
                .UserSelectsField("foobar");

            userInput.EncapsulateUsingUDTField();

            var presenterAction = Support.SetParameters(userInput);
            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            StringAssert.Contains($"Private this As {Support.StateUDTDefaultType}", actualCode);
            StringAssert.Contains($"Private Type {Support.StateUDTDefaultType}", actualCode);
            StringAssert.Contains("Foo As Long", actualCode);
            StringAssert.Contains("Bar As String", actualCode);
            StringAssert.Contains("Foobar As Byte", actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void UserDefinedType_MultipleFieldsWithUDT()
        {
            string inputCode =
$@"

Private Type TBar
    First As Long
    Second As String
End Type

Public fo|o As Long
Public myBar As TBar
";

            var userInput = new UserInputDataObject()
                .UserSelectsField("myBar");

            userInput.EncapsulateUsingUDTField();

            var presenterAction = Support.SetParameters(userInput);
            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            StringAssert.Contains("this.MyBar.First = value", actualCode);
            StringAssert.Contains("First = this.MyBar.First", actualCode);
            StringAssert.Contains("this.MyBar.Second = value", actualCode);
            StringAssert.Contains("Second = this.MyBar.Second", actualCode);
            var index = actualCode.IndexOf("Get Second", StringComparison.InvariantCultureIgnoreCase);
            var indexLast = actualCode.LastIndexOf("Get Second", StringComparison.InvariantCultureIgnoreCase);
            Assert.AreEqual(index, indexLast);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void UserDefinedType_MultipleFieldsOfSameUDT()
        {
            string inputCode =
$@"

Private Type TBar
    First As Long
    Second As String
End Type

Public fooBar As TBar
Public myBar As TBar
";

            var userInput = new UserInputDataObject()
                .UserSelectsField("myBar");

            userInput.EncapsulateUsingUDTField();

            var presenterAction = Support.SetParameters(userInput);
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _).Object;
            var model = Support.RetrieveUserModifiedModelPriorToRefactoring(vbe, "myBar", DeclarationType.Variable, presenterAction);

            Assert.AreEqual(1, model.ObjectStateUDTCandidates.Count());
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void UserDefinedType_PrivateEnumField()
        {
            const string inputCode =
@"
Private Enum NumberTypes 
    Whole = -1 
    Integral = 0 
    Rational = 1 
End Enum

Public numberT|ype As NumberTypes
";


            var userInput = new UserInputDataObject()
                .UserSelectsField("numberType");

            userInput.EncapsulateUsingUDTField();

            var presenterAction = Support.SetParameters(userInput);
            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            StringAssert.Contains("Property Get NumberType() As Long", actualCode);
            StringAssert.Contains("NumberType = this.NumberType", actualCode);
            StringAssert.Contains(" NumberType As NumberTypes", actualCode);
        }

        [TestCase("anArray", "5")]
        [TestCase("anArray", "1 To 100")]
        [TestCase("anArray", "")]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void UserDefinedType_BoundedArrayField(string arrayIdentifier, string dimensions)
        {
            var selectedInput = arrayIdentifier.Replace("n", "n|");
            string inputCode =
$@"
Public {selectedInput}({dimensions}) As String
";

            var userInput = new UserInputDataObject()
                .UserSelectsField(arrayIdentifier);

            userInput.EncapsulateUsingUDTField();

            var presenterAction = Support.SetParameters(userInput);
            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            StringAssert.Contains("Property Get AnArray() As Variant", actualCode);
            StringAssert.Contains("AnArray = this.AnArray", actualCode);
            StringAssert.Contains($" AnArray({dimensions}) As String", actualCode);
        }

        [TestCase("anArray", "5")]
        [TestCase("anArray", "1 To 100")]
        [TestCase("anArray", "")]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void UserDefinedType_LocallyReferencedArray(string arrayIdentifier, string dimensions)
        {
            var selectedInput = arrayIdentifier.Replace("n", "n|");
            string inputCode =
$@"
Public {selectedInput}({dimensions}) As String

Public Property Get AnArrayTest() As Variant
    AnArrayTest = anArray
End Property

";

            var userInput = new UserInputDataObject()
                .UserSelectsField(arrayIdentifier);

            userInput.EncapsulateUsingUDTField();

            var presenterAction = Support.SetParameters(userInput);
            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            StringAssert.Contains("Property Get AnArray() As Variant", actualCode);
            StringAssert.Contains("AnArray = this.AnArray", actualCode);
            StringAssert.Contains("AnArrayTest = this.AnArray", actualCode);
            StringAssert.Contains($" AnArray({dimensions}) As String", actualCode);
        }

        [TestCase("Public Sub This_Type()\r\nEnd Sub", "This_Type_1")]
        [TestCase("Private This_Type As Long\r\nPrivate This_Type_1 As String", "This_Type_2")]
        [TestCase("Public Property Get This_Type() As Long\r\nEnd Property", "This_Type_1")]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void UserDefinedTypeDefaultNameHasConflicts(string declaration, string expectedIdentifier)
        {
            declaration = declaration.Replace("This_Type", $"{Support.StateUDTDefaultType}");
            expectedIdentifier = expectedIdentifier.Replace("This_Type", $"{Support.StateUDTDefaultType}");
            string inputCode =
$@"

Private Type TBar
    First As Long
    Second As String
End Type

Public fo|o As Long
Public myBar As TBar

{declaration}
";

            var userInput = new UserInputDataObject()
                .UserSelectsField("myBar");


            userInput.EncapsulateUsingUDTField();

            var presenterAction = Support.SetParameters(userInput);
            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            StringAssert.Contains($"Private Type {expectedIdentifier}", actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void StateObjectCandidatesContent()
        {
            string inputCode =
$@"
Private Type TBar
    First As String
    Second As Long
End Type

Public mFoo As String
Public mBar As Long
Private mFizz

Public myBar As TBar";

            var userInput = new UserInputDataObject()
                .UserSelectsField("mFizz");

            userInput.EncapsulateUsingUDTField();

            var presenterAction = Support.SetParameters(userInput);

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _).Object;
            var model = Support.RetrieveUserModifiedModelPriorToRefactoring(vbe, "mFizz", DeclarationType.Variable, presenterAction);
            var test = model.ObjectStateUDTCandidates;

            Assert.AreEqual(2, model.ObjectStateUDTCandidates.Count());
        }

        protected override IRefactoring TestRefactoring(IRewritingManager rewritingManager, RubberduckParserState state, IRefactoringPresenterFactory factory, ISelectionService selectionService)
        {
            return Support.SupportTestRefactoring(rewritingManager, state, factory, selectionService);
        }
    }
}
