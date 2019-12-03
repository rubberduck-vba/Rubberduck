using NUnit.Framework;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.EncapsulateField;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.Utility;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RubberduckTests.Refactoring.EncapsulateField
{
    [TestFixture]
    public class EncapsulateAsUDTTests : InteractiveRefactoringTestBase<IEncapsulateFieldPresenter, EncapsulateFieldModel>
    {
        private EncapsulateFieldTestSupport Support { get; } = new EncapsulateFieldTestSupport();

        [TestCase("Public")]
        [TestCase("Private")]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void UserDefinedType_UserAcceptsDefaults(string accessibility)
        {
            string inputCode =
$@"
Private Type TBar
    First As String
    Second As Long
End Type

{accessibility} th|is As TBar";


            var presenterAction = Support.UserAcceptsDefaults();
            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            StringAssert.Contains("Private this1 As TBar", actualCode);
            StringAssert.Contains("this1 = value", actualCode);
            StringAssert.Contains($"This = this1", actualCode);
            StringAssert.DoesNotContain($"this1.First1 = value", actualCode);
        }


        [TestCase("Public")]
        [TestCase("Private")]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void UserDefinedType_UserAcceptsDefaults_TwoUDTInstances(string accessibility)
        {
            string selectionModule = "ClassInput";
            string inputCode =
$@"
Private Type TBar
    First As String
    Second As Long
End Type

{accessibility} th|is As TBar

Private that As TBar";

            string otherModule = "OtherModule";
            string moduleCode =
$@"
Private Type TBar
    First As String
    Second As Long
End Type

{accessibility} this As TBar

Private that As TBar";


            var presenterAdjustment = Support.UserAcceptsDefaults();
            var codeString = inputCode.ToCodeString();
            var actualCode = RefactoredCode(selectionModule, codeString.CaretPosition.ToOneBased(), presenterAdjustment, null, false, (selectionModule, codeString.Code, ComponentType.ClassModule), (otherModule, moduleCode, ComponentType.StandardModule));
            StringAssert.Contains("Private this1 As TBar", actualCode[selectionModule]);
            StringAssert.Contains("this1 = value", actualCode[selectionModule]);
            StringAssert.Contains($"This = this1", actualCode[selectionModule]);
            StringAssert.DoesNotContain($"this1.First1 = value", actualCode[selectionModule]);
        }

        [TestCase("Public")]
        [TestCase("Private")]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void UserDefinedType_UserAcceptsDefaults_ConflictWithEncapsulationUDT(string accessibility)
        {
            string inputCode =
$@"
Private Type TBar
    First As String
    Second As Long
End Type

{accessibility} my|Bar As TBar

Private this As Long";


            var presenterAction = Support.UserAcceptsDefaults(asUDT: true);
            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            StringAssert.Contains("Private this As Long", actualCode);
            StringAssert.Contains("Private this1 As This_Type", actualCode);
            //StringAssert.DoesNotContain($"Private this1 As This_Type", actualCode);
        }

        [TestCase("Public", true, true)]
        [TestCase("Private", true, true)]
        [TestCase("Public", false, false)]
        [TestCase("Private", false, false)]
        [TestCase("Public", true, false)]
        [TestCase("Private", true, false)]
        [TestCase("Public", false, true)]
        [TestCase("Private", false, true)]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void UserDefinedType_TwoInstanceVariables(string accessibility, bool encapsulateThis, bool encapsulateThat)
        {
            string inputCode =
$@"
Private Type TBar
    First As String
    Second As Long
End Type

{accessibility} th|is As TBar
{accessibility} that As TBar";

            var userInput = new UserInputDataObject("this", "MyType", encapsulateThis);
            userInput.AddAttributeSet("that", "MyOtherType", encapsulateThat);
            userInput.AddUDTMemberNameFlagPairs(("this", "First", true));
            userInput.AddUDTMemberNameFlagPairs(("this", "Second", true));
            userInput.AddUDTMemberNameFlagPairs(("that", "First", true));
            userInput.AddUDTMemberNameFlagPairs(("that", "Second", true));

            var expectedThis = new EncapsulationIdentifiers("this") { Property = "MyType" };
            var expectedThat = new EncapsulationIdentifiers("that") { Property = "MyOtherType" };

            var presenterAction = Support.SetParameters(userInput);
            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            if (encapsulateThis)
            {
                StringAssert.Contains($"Private {expectedThis.Field} As TBar", actualCode);
                StringAssert.Contains($"MyType = {expectedThis.Field}", actualCode);
                StringAssert.Contains($"{expectedThis.Field} = value", actualCode);
                StringAssert.Contains($"{expectedThis.Field}.First = value", actualCode);
                StringAssert.Contains($"{expectedThis.Field}.First = value", actualCode);
                StringAssert.Contains($"{expectedThis.Field}.Second = value", actualCode);
            }
            else
            {
                StringAssert.Contains($"{accessibility} this As TBar", actualCode);
                StringAssert.Contains($"this.First = value", actualCode);
                StringAssert.Contains($"this.Second = value", actualCode);
            }

            if (encapsulateThat)
            {
                StringAssert.Contains($"Private {expectedThat.Field} As TBar", actualCode);
                StringAssert.Contains($"MyOtherType = {expectedThat.Field}", actualCode);
                StringAssert.Contains($"{expectedThat.Field} = value", actualCode);
                StringAssert.Contains($"{expectedThat.Field}.First = value", actualCode);
                StringAssert.Contains($"{expectedThat.Field}.Second = value", actualCode);
            }
            else
            {
                StringAssert.Contains($"{accessibility} that As TBar", actualCode);
                StringAssert.Contains($"that.First = value", actualCode);
                StringAssert.Contains($"that.Second = value", actualCode);
            }

            StringAssert.Contains($"Property Get This_First", actualCode);
            StringAssert.Contains($"Property Get That_First", actualCode);
            StringAssert.Contains($"Property Get This_Second", actualCode);
            StringAssert.Contains($"Property Get That_Second", actualCode);
        }


        [TestCase("First")]
        [TestCase("Second")]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void UserDefinedTypeMembers_Subset(string memberToEncapsulate)
        {
            string inputCode =
$@"
Private Type TBar
    First As String
    Second As Long
End Type

Private th|is As TBar";


            var userInput = new UserInputDataObject("this", "MyType", true);
            userInput.AddUDTMemberNameFlagPairs(("this", memberToEncapsulate, true));

            var presenterAction = Support.SetParameters(userInput);

            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            StringAssert.Contains($"this.{memberToEncapsulate} = value", actualCode);
            StringAssert.Contains($"{memberToEncapsulate} = this.{memberToEncapsulate}", actualCode);
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

Private th|is As TBar";


            var userInput = new UserInputDataObject("this");
            userInput["this"].EncapsulateFlag = false;
            userInput.AddUDTMemberNameFlagPairs(("this", "First", true));
            userInput.AddUDTMemberNameFlagPairs(("this", "Second", true));

            var presenterAction = Support.SetParameters(userInput);

            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            StringAssert.Contains("this.First = value", actualCode);
            StringAssert.Contains($"First = this.First", actualCode);
            StringAssert.Contains("this.Second = value", actualCode);
            StringAssert.Contains($"Second = this.Second", actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void UserDefinedTypeMembers_OnlyEncapsulateUDTMembersUsingUDT()
        {
            string inputCode =
$@"
Private Type TBar
    First As String
    Second As Long
End Type

Private my|Bar As TBar";


            var userInput = new UserInputDataObject("myBar");
            userInput["myBar"].EncapsulateFlag = false;
            userInput.AddUDTMemberNameFlagPairs(("myBar", "First", true));
            userInput.AddUDTMemberNameFlagPairs(("myBar", "Second", true));
            userInput.EncapsulateAsUDT = true;

            var presenterAction = Support.SetParameters(userInput);

            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            StringAssert.Contains("this.MyBar.First = value", actualCode);
            StringAssert.Contains($"First = this.MyBar.First", actualCode);
            StringAssert.Contains("this.MyBar.Second = value", actualCode);
            StringAssert.Contains($"Second = this.MyBar.Second", actualCode);
            StringAssert.Contains($"MyBar As TBar", actualCode);
            StringAssert.DoesNotContain($"myBar As TBar", actualCode);
        }

        [TestCase("Public")]
        [TestCase("Private")]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void UserDefinedTypeMembersAndFields(string accessibility)
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

{accessibility} t|his As TBar";


            var userInput = new UserInputDataObject("this", "MyType", true);
            userInput.AddUDTMemberNameFlagPairs(("this", "First", true), ("this", "Second", true));
            userInput.AddAttributeSet("mFoo", "Foo", true);
            userInput.AddAttributeSet("mBar", "Bar", true);
            userInput.AddAttributeSet("mFizz", "Fizz", true);

            var presenterAction = Support.SetParameters(userInput);

            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);

            StringAssert.Contains("Private this As TBar", actualCode);
            StringAssert.Contains("this = value", actualCode);
            StringAssert.Contains("MyType = this", actualCode);
            Assert.AreEqual(actualCode.IndexOf("MyType = this"), actualCode.LastIndexOf("MyType = this"));
            StringAssert.Contains($"this.First = value", actualCode);
            StringAssert.Contains($"First = this.First", actualCode);
            StringAssert.Contains($"this.Second = value", actualCode);
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
            string inputCode =
$@"
Private Type TBar
    First As Class1
    Second As Long
End Type

{accessibility} th|is As TBar";

            string class1Code =
@"Option Explicit

Public Sub Foo()
End Sub
";

            var userInput = new UserInputDataObject("this", "MyType", true);

            userInput.AddUDTMemberNameFlagPairs(("this", "First", true), ("this", "Second", true));

            var presenterAction = Support.SetParameters(userInput);

            var codeString = inputCode.ToCodeString();
            var actualModuleCode = RefactoredCode(
                "Module1",
                codeString.CaretPosition.ToOneBased(),
                presenterAction,
                null,
                false,
                ("Class1", class1Code, ComponentType.ClassModule),
                ("Module1", codeString.Code, ComponentType.StandardModule));

            var actualCode = actualModuleCode["Module1"];

            StringAssert.Contains("Private this As TBar", actualCode);
            StringAssert.Contains("this = value", actualCode);
            StringAssert.Contains("MyType = this", actualCode);
            StringAssert.Contains("Property Set First(ByVal value As Class1)", actualCode);
            StringAssert.Contains("Property Get First() As Class1", actualCode);
            StringAssert.Contains($"Set this.First = value", actualCode);
            StringAssert.Contains($"Set First = this.First", actualCode);
            StringAssert.Contains($"this.Second = value", actualCode);
            StringAssert.Contains($"Second = this.Second", actualCode);
            StringAssert.DoesNotContain($"Second = Second", actualCode);
        }

        [TestCase("Public")]
        [TestCase("Private")]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void UserDefinedTypeMember_ContainsVariant(string accessibility)
        {
            string inputCode =
$@"
Private Type TBar
    First As String
    Second As Variant
End Type

{accessibility} th|is As TBar";

            var userInput = new UserInputDataObject("this", "MyType", true);
            userInput.AddUDTMemberNameFlagPairs(("this", "First", true), ("this", "Second", true));

            var presenterAction = Support.SetParameters(userInput);

            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            StringAssert.Contains("Private this As TBar", actualCode);
            StringAssert.Contains("this = value", actualCode);
            StringAssert.Contains("MyType = this", actualCode);
            Assert.AreEqual(actualCode.IndexOf("MyType = this"), actualCode.LastIndexOf("MyType = this"));
            StringAssert.Contains($"this.First = value", actualCode);
            StringAssert.Contains($"First = this.First", actualCode);
            StringAssert.Contains($"IsObject", actualCode);
            StringAssert.Contains($"this.Second = value", actualCode);
            StringAssert.Contains($"Second = this.Second", actualCode);
            StringAssert.DoesNotContain($"Second = Second", actualCode);
        }

        [TestCase("Public")]
        [TestCase("Private")]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void UserDefinedTypeMember_ContainsArrays(string accessibility)
        {
            string inputCode =
$@"
Private Type TBar
    First(5) As String
    Second(1 to 100) As Long
    Third() As Double
End Type

{accessibility} t|his As TBar";

            var userInput = new UserInputDataObject("this", "MyType", false);
            userInput.AddUDTMemberNameFlagPairs(("this", "First", true), ("this", "Second", true), ("this", "Third", true));

            var presenterAction = Support.SetParameters(userInput);

            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            StringAssert.Contains($"{accessibility} this As TBar", actualCode);
            StringAssert.DoesNotContain("this = value", actualCode);
            StringAssert.DoesNotContain($"this.First = value", actualCode);
            StringAssert.Contains($"Property Get First() As Variant", actualCode);
            StringAssert.Contains($"First = this.First", actualCode);
            StringAssert.DoesNotContain($"this.Second = value", actualCode);
            StringAssert.Contains($"Second = this.Second", actualCode);
            StringAssert.Contains($"Property Get Second() As Variant", actualCode);
            StringAssert.Contains($"Third = this.Third", actualCode);
            StringAssert.Contains($"Property Get Third() As Variant", actualCode);
        }

        [TestCase("Public", "Public")]
        [TestCase("Private", "Private")]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void UserDefinedTypeMembersOnly(string accessibility, string expectedAccessibility)
        {
            string inputCode =
$@"
Private Type TBar
    First As String
    Second As Long
End Type

{accessibility} th|is As TBar";

            var userInput = new UserInputDataObject("this", "MyType", false);
            userInput.AddUDTMemberNameFlagPairs(("this", "First", true), ("this", "Second", true));

            var presenterAction = Support.SetParameters(userInput);

            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            StringAssert.Contains($"{expectedAccessibility} this As TBar", actualCode);
            StringAssert.DoesNotContain("this = value", actualCode);
            StringAssert.DoesNotContain("MyType = this", actualCode);
            StringAssert.Contains($"this.First = value", actualCode);
            StringAssert.Contains($"First = this.First", actualCode);
            StringAssert.Contains($"this.Second = value", actualCode);
            StringAssert.Contains($"Second = this.Second", actualCode);
            StringAssert.DoesNotContain($"Second = Second", actualCode);
        }

        [TestCase("Public")]
        [TestCase("Private")]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void UserDefinedTypeMembers_ExternallyDefinedType(string accessibility)
        {
            string inputCode =
$@"
Option Explicit

{accessibility} th|is As TBar";

            string typeDefinition =
$@"
Public Type TBar
    First As String
    Second As Long
End Type
";

            var userInput = new UserInputDataObject("this", "MyType", true);
            userInput.AddUDTMemberNameFlagPairs(("this", "First", true), ("this", "Second", true));

            var presenterAction = Support.SetParameters(userInput);

            var codeString = inputCode.ToCodeString();
            var actualModuleCode = RefactoredCode(
                "Class1",
                codeString.CaretPosition.ToOneBased(),
                presenterAction,
                null,
                false,
                ("Class1", codeString.Code, ComponentType.ClassModule),
                ("Module1", typeDefinition, ComponentType.StandardModule));

            Assert.AreEqual(typeDefinition, actualModuleCode["Module1"]);


            var actualCode = actualModuleCode["Class1"];
            StringAssert.Contains("Private this As TBar", actualCode);
            StringAssert.Contains("this = value", actualCode);
            StringAssert.Contains("MyType = this", actualCode);
            Assert.AreEqual(actualCode.IndexOf("MyType = this"), actualCode.LastIndexOf("MyType = this"));
            StringAssert.Contains($"this.First = value", actualCode);
            StringAssert.Contains($"First = this.First", actualCode);
            StringAssert.Contains($"this.Second = value", actualCode);
            StringAssert.Contains($"Second = this.Second", actualCode);
            StringAssert.DoesNotContain($"Second = Second", actualCode);
        }

        [TestCase("Public")]
        [TestCase("Private")]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void UserDefinedTypeMembers_ObjectField(string accessibility)
        {
            string inputCode =
$@"
Option Explicit

{accessibility} mThe|Class As Class1";

            string classContent =
$@"
Option Explicit

Private Sub Class_Initialize()
End Sub
";

            var presenterAction = Support.UserAcceptsDefaults();

            var codeString = inputCode.ToCodeString();
            var actualModuleCode = RefactoredCode(
                "Module1",
                codeString.CaretPosition.ToOneBased(),
                presenterAction,
                null,
                false,
                ("Class1", classContent, ComponentType.ClassModule),
                ("Module1", codeString.Code, ComponentType.StandardModule));

            var actualCode = actualModuleCode["Module1"];

            var defaults = new EncapsulationIdentifiers("mTheClass");

            StringAssert.Contains($"Private {defaults.Field} As Class1", actualCode);
            StringAssert.Contains($"Set {defaults.Field} = value", actualCode);
            StringAssert.Contains($"MTheClass = {defaults.Field}", actualCode);
            StringAssert.Contains($"Public Property Set {defaults.Property}", actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void UserDefinedType_MultipleFields_ToNewUDT()
        {
            string inputCode =
$@"
Public fo|o As Long
Public bar As String
Public foobar As Byte
";

            var userInput = new UserInputDataObject("foo", encapsulationFlag: true);
            userInput.AddAttributeSet("bar", encapsulationFlag: true);
            userInput.AddAttributeSet("foobar", encapsulationFlag: true);
            userInput.EncapsulateAsUDT = true;

            var presenterAction = Support.SetParameters(userInput);
            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            StringAssert.Contains("Private this As This_Type", actualCode);
            StringAssert.Contains("Private Type This_Type", actualCode);
            StringAssert.Contains("Foo As Long", actualCode);
            StringAssert.Contains("Bar As String", actualCode);
            StringAssert.Contains("Foobar As Byte", actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void UserDefinedType_MultipleFieldsWithUDT_ToNewUDT()
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

            var userInput = new UserInputDataObject("foo", encapsulationFlag: true);
            userInput.AddAttributeSet("myBar", encapsulationFlag: true);
            userInput.AddUDTMemberNameFlagPairs(("myBar", "First", true), ("myBar", "Second", true));
            userInput.EncapsulateAsUDT = true;

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

        protected override IRefactoring TestRefactoring(IRewritingManager rewritingManager, RubberduckParserState state, IRefactoringPresenterFactory factory, ISelectionService selectionService)
        {
            return Support.SupportTestRefactoring(rewritingManager, state, factory, selectionService);
        }
    }
}
