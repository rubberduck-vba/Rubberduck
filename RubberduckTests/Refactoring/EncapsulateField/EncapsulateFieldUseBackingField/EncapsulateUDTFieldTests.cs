using NUnit.Framework;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.EncapsulateField;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.Utility;
using RubberduckTests.Mocks;

namespace RubberduckTests.Refactoring.EncapsulateField
{
    [TestFixture]
    public class EncapsulateUDTFieldTests : EncapsulateFieldInteractiveRefactoringTest
    {
        private EncapsulateFieldTestSupport Support { get; } = new EncapsulateFieldTestSupport();

        [TestCase("Public")]
        [TestCase("Private")]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void UserDefinedType_UserAcceptsDefaults(string accessibility)
        {
            var inputCode =
$@"
Private Type TBar
    First As String
    Second As Long
End Type

{accessibility} th|is As TBar";

            var presenterAction = Support.UserAcceptsDefaults();

            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            StringAssert.Contains("Private this As TBar", actualCode);
            StringAssert.DoesNotContain("this = value", actualCode);
            StringAssert.DoesNotContain($"This = this", actualCode);
            StringAssert.Contains($"Public Property Get First", actualCode);
            StringAssert.Contains($"Public Property Get Second", actualCode);
            StringAssert.Contains($"this.First = {Support.RHSIdentifier}", actualCode);
            StringAssert.Contains($"First = this.First", actualCode);
            StringAssert.Contains($"this.Second = {Support.RHSIdentifier}", actualCode);
            StringAssert.Contains($"Second = this.Second", actualCode);
        }

        [TestCase(true, true)]
        [TestCase(false, false)]
        [TestCase(true, false)]
        [TestCase(false, true)]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void UserDefinedType_TwoFields(bool encapsulateThis, bool encapsulateThat)
        {
            var inputCode =
$@"
Private Type TBar
    First As String
    Second As Long
End Type

Public th|is As TBar
Public that As TBar";

            var expectedThis = new EncapsulationIdentifiers("this");
            var expectedThat = new EncapsulationIdentifiers("that");

            var userInput = new UserInputDataObject()
                .AddUserInputSet(expectedThis.TargetFieldName, encapsulationFlag: encapsulateThis)
                .AddUserInputSet(expectedThat.TargetFieldName, encapsulationFlag: encapsulateThat);

            var presenterAction = Support.SetParameters(userInput);

            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            if (encapsulateThis && encapsulateThat)
            {
                StringAssert.Contains($"Private {expectedThis.TargetFieldName} As TBar", actualCode);
                StringAssert.Contains($"First = {expectedThis.TargetFieldName}.First", actualCode);
                StringAssert.Contains($"Second = {expectedThis.TargetFieldName}.Second", actualCode);
                StringAssert.Contains($"{expectedThis.TargetFieldName}.First = {Support.RHSIdentifier}", actualCode);
                StringAssert.Contains($"{expectedThis.TargetFieldName}.Second = {Support.RHSIdentifier}", actualCode);
                StringAssert.Contains($"Property Get First", actualCode);
                StringAssert.Contains($"Property Get Second", actualCode);

                StringAssert.Contains($"Private {expectedThat.TargetFieldName} As TBar", actualCode);
                StringAssert.Contains($"First1 = {expectedThat.TargetFieldName}.First", actualCode);
                StringAssert.Contains($"Second1 = {expectedThat.TargetFieldName}.Second", actualCode);
                StringAssert.Contains($"{expectedThat.TargetFieldName}.First = {Support.RHSIdentifier}", actualCode);
                StringAssert.Contains($"{expectedThat.TargetFieldName}.Second = {Support.RHSIdentifier}", actualCode);
                StringAssert.Contains($"Property Get First1", actualCode);
                StringAssert.Contains($"Property Get Second1", actualCode);

                StringAssert.Contains($"Private {expectedThis.TargetFieldName} As TBar", actualCode);
                StringAssert.Contains($"Private {expectedThat.TargetFieldName} As TBar", actualCode);
            }
            else if (encapsulateThis && !encapsulateThat)
            {
                StringAssert.Contains($"First = {expectedThis.TargetFieldName}.First", actualCode);
                StringAssert.Contains($"Second = {expectedThis.TargetFieldName}.Second", actualCode);
                StringAssert.Contains($"{expectedThis.TargetFieldName}.First = {Support.RHSIdentifier}", actualCode);
                StringAssert.Contains($"{expectedThis.TargetFieldName}.Second = {Support.RHSIdentifier}", actualCode);
                StringAssert.Contains($"Property Get First", actualCode);
                StringAssert.Contains($"Property Get Second", actualCode);

                StringAssert.Contains($"Private {expectedThis.TargetFieldName} As TBar", actualCode);
                StringAssert.Contains($"Public {expectedThat.TargetFieldName} As TBar", actualCode);
            }
            else if (!encapsulateThis && encapsulateThat)
            {
                StringAssert.Contains($"First = {expectedThat.TargetFieldName}.First", actualCode);
                StringAssert.Contains($"Second = {expectedThat.TargetFieldName}.Second", actualCode);
                StringAssert.Contains($"{expectedThat.TargetFieldName}.First = {Support.RHSIdentifier}", actualCode);
                StringAssert.Contains($"{expectedThat.TargetFieldName}.Second = {Support.RHSIdentifier}", actualCode);
                StringAssert.Contains($"Property Get First", actualCode);
                StringAssert.Contains($"Property Get Second", actualCode);

                StringAssert.Contains($"Public {expectedThis.TargetFieldName} As TBar", actualCode);
                StringAssert.Contains($"Private {expectedThat.TargetFieldName} As TBar", actualCode);
            }
            else
            {
                StringAssert.Contains($"Public {expectedThis.TargetFieldName} As TBar", actualCode);
                StringAssert.Contains($"Public {expectedThat.TargetFieldName} As TBar", actualCode);
            }
        }

        [TestCase("Public")]
        [TestCase("Private")]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void UserDefinedTypeMembersAndFields(string accessibility)
        {
            var inputCode =
$@"
Private Type TBar
    First As String
    Second As Long
End Type

Public mFoo As String
Public mBar As Long
Private mFizz

{accessibility} t|his As TBar";

            var userInput = new UserInputDataObject()
            .UserSelectsField("this", "MyType")
            .UserSelectsField("mFoo", "Foo")
            .UserSelectsField("mBar", "Bar")
            .UserSelectsField("mFizz", "Fizz");

            var presenterAction = Support.SetParameters(userInput);

            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);

            StringAssert.Contains("Private this As TBar", actualCode);
            StringAssert.DoesNotContain("this = value", actualCode);
            StringAssert.DoesNotContain("MyType = this", actualCode);
            StringAssert.Contains($"this.First = {Support.RHSIdentifier}", actualCode);
            StringAssert.Contains($"First = this.First", actualCode);
            StringAssert.Contains($"this.Second = {Support.RHSIdentifier}", actualCode);
            StringAssert.Contains($"Second = this.Second", actualCode);
            StringAssert.DoesNotContain($"Second = Second", actualCode);
            StringAssert.Contains($"Private mFoo As String", actualCode);
            StringAssert.Contains($"Public Property Get Foo(", actualCode);
            StringAssert.Contains($"Public Property Let Foo(", actualCode);
            StringAssert.Contains($"Private mBar As Long", actualCode);
            StringAssert.Contains($"Public Property Get Bar(", actualCode);
            StringAssert.Contains($"Public Property Let Bar(", actualCode);
            StringAssert.Contains($"Private mFizz As Variant", actualCode);
            StringAssert.Contains($"Public Property Get Fizz() As Variant", actualCode);
            StringAssert.Contains($"Public Property Let Fizz(", actualCode);
        }

        [TestCase("Public")]
        [TestCase("Private")]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void UserDefinedTypeMember_ContainsObjects(string accessibility)
        {
            var inputCode =
$@"
Private Type TBar
    First As Class1
    Second As Long
End Type

{accessibility} th|is As TBar";

            var class1Code =
@"Option Explicit

Public Sub Foo()
End Sub
";

            var userInput = new UserInputDataObject()
                .UserSelectsField("this", "MyType");

            var presenterAction = Support.SetParameters(userInput);

            var actualModuleCode = Support.RefactoredCode(presenterAction, inputCode.ToCodeString(),
                ("Class1", class1Code, ComponentType.ClassModule));

            var actualCode = actualModuleCode[MockVbeBuilder.TestModuleName];

            StringAssert.Contains("Private this As TBar", actualCode);
            StringAssert.DoesNotContain($"this = {Support.RHSIdentifier}", actualCode);
            StringAssert.DoesNotContain("MyType = this", actualCode);
            StringAssert.Contains($"Property Set First(ByVal {Support.RHSIdentifier} As Class1)", actualCode);
            StringAssert.Contains("Property Get First() As Class1", actualCode);
            StringAssert.Contains($"Set this.First = {Support.RHSIdentifier}", actualCode);
            StringAssert.Contains($"Set First = this.First", actualCode);
            StringAssert.Contains($"this.Second = {Support.RHSIdentifier}", actualCode);
            StringAssert.Contains($"Second = this.Second", actualCode);
            StringAssert.DoesNotContain($"Second = Second", actualCode);
            Assert.AreEqual(actualCode.IndexOf($"this.First = {Support.RHSIdentifier}"), actualCode.LastIndexOf($"this.First = {Support.RHSIdentifier}"));
        }

        [TestCase("Public")]
        [TestCase("Private")]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void UserDefinedTypeMember_ContainsVariant(string accessibility)
        {
            var inputCode =
$@"
Private Type TBar
    First As String
    Second As Variant
End Type

{accessibility} th|is As TBar";

            var userInput = new UserInputDataObject()
                .UserSelectsField("this", "MyType");

            var presenterAction = Support.SetParameters(userInput);

            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            StringAssert.Contains("Private this As TBar", actualCode);
            StringAssert.DoesNotContain($"this = {Support.RHSIdentifier}", actualCode);
            StringAssert.DoesNotContain("MyType = this", actualCode);
            StringAssert.Contains($"this.First = {Support.RHSIdentifier}", actualCode);
            StringAssert.Contains($"First = this.First", actualCode);
            StringAssert.Contains($"IsObject", actualCode);
            StringAssert.Contains($"this.Second = {Support.RHSIdentifier}", actualCode);
            StringAssert.Contains($"Second = this.Second", actualCode);
            StringAssert.DoesNotContain($"Second = Second", actualCode);
        }

        [TestCase("Public")]
        [TestCase("Private")]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void UserDefinedTypeMember_ContainsArrays(string accessibility)
        {
            var inputCode =
$@"
Private Type TBar
    First(5) As String
    Second(1 to 100) As Long
    Third() As Double
End Type

{accessibility} t|his As TBar";

            var userInput = new UserInputDataObject()
                .UserSelectsField("this", "MyType");

            var presenterAction = Support.SetParameters(userInput);

            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            StringAssert.Contains($"Private this As TBar", actualCode);
            StringAssert.DoesNotContain($"this = {Support.RHSIdentifier}", actualCode);
            StringAssert.DoesNotContain($"this.First = {Support.RHSIdentifier}", actualCode);
            StringAssert.Contains($"Property Get First() As Variant", actualCode);
            StringAssert.Contains($"First = this.First", actualCode);
            StringAssert.DoesNotContain($"this.Second = {Support.RHSIdentifier}", actualCode);
            StringAssert.Contains($"Second = this.Second", actualCode);
            StringAssert.Contains($"Property Get Second() As Variant", actualCode);
            StringAssert.Contains($"Third = this.Third", actualCode);
            StringAssert.Contains($"Property Get Third() As Variant", actualCode);
        }

        [TestCase("Public")]
        [TestCase("Private")]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void UserDefinedTypeMembers_ExternallyDefinedType(string accessibility)
        {
            var inputCode =
$@"
Option Explicit

{accessibility} th|is As TBar";

            var typeDefinition =
$@"
Public Type TBar
    First As String
    Second As Long
End Type
";

            var userInput = new UserInputDataObject()
                .UserSelectsField("this", "MyType");

            var presenterAction = Support.SetParameters(userInput);

            var actualModuleCode = Support.RefactoredCode(presenterAction,
                ("Class1", inputCode.ToCodeString(), ComponentType.ClassModule),
                ("Module1", typeDefinition, ComponentType.StandardModule));

            Assert.AreEqual(typeDefinition, actualModuleCode["Module1"]);

            var actualCode = actualModuleCode["Class1"];
            StringAssert.Contains("Private this As TBar", actualCode);

            if (accessibility == "Public")
            {
                StringAssert.Contains($"this = {Support.RHSIdentifier}", actualCode);
                StringAssert.Contains("MyType = this", actualCode);
                StringAssert.Contains($"Public Property Get MyType", actualCode);
                Assert.AreEqual(actualCode.IndexOf("MyType = this"), actualCode.LastIndexOf("MyType = this"));
                StringAssert.DoesNotContain($"this.First = {Support.RHSIdentifier}", actualCode);
                StringAssert.DoesNotContain($"this.Second = {Support.RHSIdentifier}", actualCode);
                StringAssert.DoesNotContain($"Second = Second", actualCode);
                StringAssert.Contains($"Public Property Let MyType(ByRef {Support.RHSIdentifier} As TBar", actualCode);
            }
        }

        [TestCase("Public")]
        [TestCase("Private")]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void UserDefinedTypeMembers_ObjectField(string accessibility)
        {
            var inputCode =
$@"
Option Explicit

{accessibility} mThe|Class As Class1";

            var classContent =
$@"
Option Explicit

Private Sub Class_Initialize()
End Sub
";

            var presenterAction = Support.UserAcceptsDefaults();

            var actualModuleCode = Support.RefactoredCode(presenterAction, inputCode.ToCodeString(),
                ("Class1", classContent, ComponentType.ClassModule));

            var actualCode = actualModuleCode[MockVbeBuilder.TestModuleName];

            StringAssert.Contains($"Private mTheClass As Class1", actualCode);
            StringAssert.Contains($"Set mTheClass = {Support.RHSIdentifier}", actualCode);
            StringAssert.Contains($"Set TheClass = mTheClass", actualCode);
            StringAssert.Contains($"Public Property Set TheClass", actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void UserDefinedTypeUserUpdatesToBeReadOnly()
        {
            var inputCode =
$@"
Private Type TBar
    First As String
    Second As Long
End Type

Private t|his As TBar";

            var userInput = new UserInputDataObject()
                .UserSelectsField("this", isReadOnly: true);

            var presenterAction = Support.SetParameters(userInput);

            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);

            StringAssert.DoesNotContain("Public Property Let First", actualCode);
            StringAssert.DoesNotContain("Public Property Let Second", actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void MultipleUserDefinedTypesOfSameName()
        {
            var inputCode =
$@"
Option Explicit

Private Type TBar
    FirstValue As Long
    SecondValue As String
End Type

Public mF|oo As TBar
";

            var module2Content =
$@"
Public Type TBar
    FirstVal As Long
    SecondVal As String
End Type
";

            var presenterAction = Support.UserAcceptsDefaults();

            var actualModuleCode = Support.RefactoredCode(presenterAction, inputCode.ToCodeString(),
                ("Module2", module2Content, ComponentType.StandardModule));
            StringAssert.Contains($"Public Property Let FirstValue(", actualModuleCode[MockVbeBuilder.TestModuleName]);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void UDTMemberPropertyConflictsWithExistingFunction()
        {
            var inputCode =
$@"
Private Type TBar
    First As String
    Second As Long
End Type

Public myB|ar As TBar

Private Function First() As String
    First = myBar.First
End Function";

            var presenterAction = Support.UserAcceptsDefaults();

            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);

            StringAssert.Contains("Public Property Let First1", actualCode);
            StringAssert.Contains("Public Property Let Second", actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void UDTMemberIsPrivateUDT()
        {
            var inputCode =
$@"

Private Type TFoo
    Foo As Integer
    Bar As Byte
End Type

Private Type TBar
    FooBar As TFoo
End Type

Private my|Bar As TBar
";

            var presenterAction = Support.UserAcceptsDefaults();

            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);

            StringAssert.Contains("Public Property Let Foo(", actualCode);
            StringAssert.Contains("Public Property Let Bar(", actualCode);
            StringAssert.Contains($"myBar.FooBar.Foo = {Support.RHSIdentifier}", actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void UDTMemberIsPrivateUDT_RepeatedType()
        {
            var inputCode =
$@"

Private Type TFoo
    Foo As Integer
    Bar As Byte
End Type

Private Type TBar
    FooBar As TFoo
    ReBar As TFoo
End Type

Private my|Bar As TBar
";

            var presenterAction = Support.UserAcceptsDefaults();

            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);

            StringAssert.Contains("Public Property Let Foo(", actualCode);
            StringAssert.Contains("Public Property Let Bar(", actualCode);
            StringAssert.Contains("Public Property Let Foo1(", actualCode);
            StringAssert.Contains("Public Property Let Bar1(", actualCode);
            StringAssert.Contains($"myBar.FooBar.Foo = {Support.RHSIdentifier}", actualCode);
            StringAssert.Contains($"myBar.ReBar.Foo = {Support.RHSIdentifier}", actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void UDTMemberIsPublicUDT()
        {
            var inputCode =
$@"

Public Type TFoo
    Foo As Integer
    Bar As Byte
End Type

Private Type TBar
    FooBar As TFoo
End Type

Private my|Bar As TBar
";

            var presenterAction = Support.UserAcceptsDefaults();

            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);

            StringAssert.Contains("Public Property Let FooBar(", actualCode);
            StringAssert.Contains($"myBar.FooBar = {Support.RHSIdentifier}", actualCode);
        }

        protected override IRefactoring TestRefactoring(
            IRewritingManager rewritingManager,
            RubberduckParserState state,
            RefactoringUserInteraction<IEncapsulateFieldPresenter, EncapsulateFieldModel> userInteraction,
            ISelectionService selectionService)
        {
            return Support.SupportTestRefactoring(rewritingManager, state, userInteraction, selectionService);
        }
    }
}
