using System;
using NUnit.Framework;
using Moq;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.EncapsulateField;
using Rubberduck.VBEditor;
using RubberduckTests.Mocks;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.VBEditor.Utility;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Grammar;

namespace RubberduckTests.Refactoring.EncapsulateField
{

    [TestFixture]
    public class EncapsulateFieldTests : InteractiveRefactoringTestBase<IEncapsulateFieldPresenter, EncapsulateFieldModel>
    {
        private EncapsulateFieldTestSupport Support { get; } = new EncapsulateFieldTestSupport();

        [TestCase("fizz", true, "baz", true, "buzz", true)]
        [TestCase("fizz", false, "baz", true, "buzz", true)]
        [TestCase("fizz", false, "baz", false, "buzz", true)]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulateMultipleFields(
            string var1, bool var1Flag,
            string var2, bool var2Flag,
            string var3, bool var3Flag)
        {
            string inputCode =
$@"Public {var1} As Integer
Public {var2} As Integer
Public {var3} As Integer";

            var selection = new Selection(1, 1);

            var userInput = new UserInputDataObject(var1, $"{var1}Prop", var1Flag);
            userInput.AddAttributeSet(var2, $"{var2}Prop", var2Flag);
            userInput.AddAttributeSet(var3, $"{var3}Prop", var3Flag);

            var flags = new Dictionary<string, bool>()
            {
                [var1] = var1Flag,
                [var2] = var2Flag,
                [var3] = var3Flag
            };

            var presenterAction = Support.SetParameters(userInput);

            var actualCode = RefactoredCode(inputCode, selection, presenterAction);

            var notEncapsulated = flags.Keys.Where(k => !flags[k])
                   .Select(k => k);

            var encapsulated = flags.Keys.Where(k => flags[k])
                   .Select(k => k);

            foreach ( var variable in notEncapsulated)
            {
                StringAssert.Contains($"Public {variable} As Integer", actualCode);
            }

            foreach (var variable in encapsulated)
            {
                StringAssert.Contains($"Private {variable} As", actualCode);
                StringAssert.Contains($"{variable}Prop = {variable}", actualCode);
                StringAssert.Contains($"{variable} = value", actualCode);
                StringAssert.Contains($"Let {variable}Prop(ByVal value As", actualCode);
                StringAssert.Contains($"Property Get {variable}Prop()", actualCode);
            }
        }

        [TestCase("fizz", true, "baz", true, "buzz", true, "boink", true)]
        [TestCase("fizz", false, "baz", true, "buzz", true, "boink", true)]
        [TestCase("fizz", false, "baz", true, "buzz", true, "boink", false)]
        [TestCase("fizz", false, "baz", true, "buzz", false, "boink", false)]
        [TestCase("fizz", false, "baz", false, "buzz", true, "boink", true)]
        [TestCase("fizz", false, "baz", false, "buzz", false, "boink", true)]
        [TestCase("fizz", false, "baz", true, "buzz", false, "boink", true)]
        [TestCase("fizz", false, "baz", false, "buzz", true, "boink", true)]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulateMultipleFieldsInList(
            string var1, bool var1Flag,
            string var2, bool var2Flag,
            string var3, bool var3Flag,
            string var4, bool var4Flag)
        {
            string inputCode =
$@"Public {var1} As Integer, {var2} As Integer, {var3} As Integer, {var4} As Integer";

            var selection = new Selection(1, 9);

            var userInput = new UserInputDataObject(var1, $"{var1}Prop", var1Flag);
            userInput.AddAttributeSet(var2, $"{var2}Prop", var2Flag);
            userInput.AddAttributeSet(var3, $"{var3}Prop", var3Flag);
            userInput.AddAttributeSet(var4, $"{var4}Prop", var4Flag);

            var flags = new Dictionary<string, bool>()
            {
                [var1] = var1Flag,
                [var2] = var2Flag,
                [var3] = var3Flag,
                [var4] = var4Flag
            };

            var presenterAction = Support.SetParameters(userInput);

            var actualCode = RefactoredCode(inputCode, selection, presenterAction);

            var remainInList = flags.Keys.Where(k => !flags[k])
                   .Select(k => $"{k} As Integer");

            if (remainInList.Any())
            {
                var declarationList = $"Public {string.Join(", ", remainInList)}";
                StringAssert.Contains(declarationList, actualCode);
            }

            foreach (var key in flags.Keys)
            {
                if (flags[key])
                {
                    StringAssert.Contains($"Private {key} As", actualCode);
                    StringAssert.Contains($"{key}Prop = {key}", actualCode);
                    StringAssert.Contains($"{key} = value", actualCode);
                    StringAssert.Contains($"Let {key}Prop(ByVal value As", actualCode);
                    StringAssert.Contains($"Property Get {key}Prop()", actualCode);
                }
            }
        }

//        [Test]
//        [Category("Refactorings")]
//        [Category("Encapsulate Field")]
//        public void UserDefinedType_MultipleFieldsToNewUDT()
//        {
//            string inputCode =
//$@"
//Public fo|o As Long
//Public bar As String
//Public foobar As Byte
//";

//            var userInput = new UserInputDataObject("foo", encapsulationFlag: true);
//            userInput.AddAttributeSet("bar", encapsulationFlag: true);
//            userInput.AddAttributeSet("foobar", encapsulationFlag: true);
//            userInput.EncapsulateAsUDT = true;

//            var presenterAction = Support.SetParameters(userInput);
//            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
//            StringAssert.Contains("Private this As This_Type", actualCode);
//            StringAssert.Contains("Private Type This_Type", actualCode);
//            StringAssert.Contains("Foo As Long", actualCode);
//            StringAssert.Contains("Bar As String", actualCode);
//            StringAssert.Contains("Foobar As Byte", actualCode);
//        }

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
        public void UserDefinedType_UserAcceptsDefaults_ToUDTInstances(string accessibility)
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
            var actualCode = RefactoredCode(selectionModule, codeString.CaretPosition.ToOneBased(), presenterAdjustment, null, false, (selectionModule, codeString.Code, ComponentType.ClassModule), (otherModule, moduleCode, ComponentType.StandardModule) );
            StringAssert.Contains("Private this1 As TBar", actualCode[selectionModule]);
            StringAssert.Contains("this1 = value", actualCode[selectionModule]);
            StringAssert.Contains($"This = this1", actualCode[selectionModule]);
            StringAssert.DoesNotContain($"this1.First1 = value", actualCode[selectionModule]);
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

            var expectedThis = new EncapsulationIdentifiers("this") { Property = "MyType"};
            var expectedThat = new EncapsulationIdentifiers("that") { Property = "MyOtherType"};

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
        public void EncapsulatePublicField_InvalidDeclarationType_Throws()
        {
            const string inputCode =
                @"Public fizz As Integer";

            var presenterAction = Support.SetParametersForSingleTarget("fizz", "Name");
            var actualCode = RefactoredCode(inputCode, "TestModule1", DeclarationType.ProceduralModule, presenterAction, typeof(InvalidDeclarationTypeException));
            Assert.AreEqual(inputCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulatePublicField_InvalidIdentifierSelected_Throws()
        {
            const string inputCode =
                @"Public Function fiz|z() As Integer
End Function";

            var presenterAction = Support.SetParametersForSingleTarget("fizz", "Name");

            var codeString = inputCode.ToCodeString();
            var actualCode = RefactoredCode(codeString.Code, codeString.CaretPosition.ToOneBased(), presenterAction, typeof(NoDeclarationForSelectionException));
            Assert.AreEqual(codeString.Code, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulatePublicField_FieldIsOverMultipleLines()
        {
            const string inputCode =
                @"Public _
fi|zz _
As _
Integer";
            const string expectedCode =
                @"Private _
fizz _
As _
Integer

Public Property Get Name() As Integer
    Name = fizz
End Property

Public Property Let Name(ByVal value As Integer)
    fizz = value
End Property
";
            var presenterAction = Support.SetParametersForSingleTarget("fizz", "Name");
            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            Assert.AreEqual(expectedCode, actualCode);
        }


        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulatePublicField_ReadOnly()
        {
            const string inputCode =
                @"|Public fizz As Integer";

            const string expectedCode =
                @"Private fizz As Integer

Public Property Get Name() As Integer
    Name = fizz
End Property
";
            var presenterAction = Support.SetParametersForSingleTarget("fizz", "Name", isReadonly: true);
            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulatePublicField_NewPropertiesInsertedAboveExistingCode()
        {
            const string inputCode =
                @"|Public fizz As Integer

Sub Foo()
End Sub

Function Bar() As Integer
    Bar = 0
End Function";

            var presenterAction = Support.SetParametersForSingleTarget("fizz", "Name");

            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            Assert.Greater(actualCode.IndexOf("Sub Foo"), actualCode.LastIndexOf("End Property"));
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulatePublicField_OtherPropertiesInClass()
        {
            const string inputCode =
                @"|Public fizz As Integer

Property Get Foo() As Variant
    Foo = True
End Property

Property Let Foo(ByVal vall As Variant)
End Property

Property Set Foo(ByVal vall As Variant)
End Property";

            var presenterAction = Support.SetParametersForSingleTarget("fizz", "Name");

            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            Assert.Greater(actualCode.IndexOf("fizz = value"), actualCode.IndexOf("fizz As Integer"));
            Assert.Less(actualCode.IndexOf("fizz = value"), actualCode.IndexOf("Get Foo"));
        }

        [TestCase("|Public fizz As Integer\r\nPublic buzz As Boolean", "Private fizz As Integer\r\nPublic buzz As Boolean")]
        [TestCase("Public buzz As Boolean\r\n|Public fizz As Integer", "Public buzz As Boolean\r\nPrivate fizz As Integer")]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulatePublicField_OtherNonSelectedFieldsInClass(string inputFields, string expectedFields)
        {
            string inputCode = inputFields;

            string expectedCode =
$@"{expectedFields}

Public Property Get Name() As Integer
    Name = fizz
End Property

Public Property Let Name(ByVal value As Integer)
    fizz = value
End Property
";
            var presenterAction = Support.SetParametersForSingleTarget("fizz", "Name");
            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [TestCase(1, 10, "fizz", "Public buzz", "Private fizz As Variant", "Public fizz")]
        [TestCase(2, 2, "buzz", "Public fizz, _\r\nbazz", "Private buzz As Boolean", "")]
        [TestCase(3, 2, "bazz", "Public fizz, _\r\nbuzz", "Private bazz As Date", "Boolean, bazz As Date")]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulatePublicField_SelectedWithinDeclarationList(int rowSelection, int columnSelection, string fieldName, string contains1, string contains2, string doesNotContain)
        {
            string inputCode =
$@"Public fizz, _
buzz As Boolean, _
bazz As Date";

            var selection = new Selection(rowSelection, columnSelection);
            var presenterAction = Support.SetParametersForSingleTarget(fieldName, "Name");
            var actualCode = RefactoredCode(inputCode, selection, presenterAction);
            StringAssert.Contains(contains1, actualCode);
            StringAssert.Contains(contains1, actualCode);
            if (doesNotContain.Length > 0)
            {
                StringAssert.DoesNotContain(doesNotContain, actualCode);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulatePrivateField()
        {
            const string inputCode =
                @"|Private fizz As Integer";

            const string expectedCode =
                @"Private fizz As Integer

Public Property Get Name() As Integer
    Name = fizz
End Property

Public Property Let Name(ByVal value As Integer)
    fizz = value
End Property
";
            var presenterAction = Support.SetParametersForSingleTarget("fizz", "Name");
            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulatePrivateField_NameConflict()
        {
            const string inputCode =
                @"Private fizz As String
Private mName As String

Public Property Get Name() As String
    Name = mName
End Property

Public Property Let Name(ByVal value As String)
    mName = value
End Property
";
            var fieldName = "fizz";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _).Object;

            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe);
            using (state)
            {
                IEncapsulatedFieldDeclaration efd = null;
                var fields = new List<IEncapsulatedFieldDeclaration>();
                var validator = new EncapsulateFieldNamesValidator(state, () => fields);

                var match = state.DeclarationFinder.MatchName(fieldName).Single();
                efd = new EncapsulatedFieldDeclaration(match, validator);
                fields.Add(efd);
                efd.PropertyName = "Name";

                var hasConflict = !validator.HasValidEncapsulationAttributes(efd.EncapsulationAttributes, efd.QualifiedModuleName, (Declaration dec) => match.Equals(dec));
                Assert.IsTrue(hasConflict);
            }
        }

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

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulatePrivateField_Defaults()
        {
            const string inputCode =
                @"|Private fizz As Integer";

            const string expectedCode =
                @"Private fizz1 As Integer

Public Property Get Fizz() As Integer
    Fizz = fizz1
End Property

Public Property Let Fizz(ByVal value As Integer)
    fizz1 = value
End Property
";
            var presenterAction = Support.UserAcceptsDefaults();
            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulatePrivateField_DefaultsAsUDT()
        {
            const string inputCode =
                @"|Private fizz As Integer";

            var presenterAction = Support.UserAcceptsDefaults(asUDT: true);
            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            StringAssert.Contains("Fizz As Integer", actualCode);
            StringAssert.Contains("this As This_Type", actualCode);
            StringAssert.Contains("this.Fizz = value", actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulatePublicField_FieldHasReferences()
        {
            const string inputCode =
                @"|Public fizz As Integer

Sub Foo()
    fizz = 0
    Bar fizz
End Sub

Sub Bar(ByVal name As Integer)
End Sub";

            var presenterAction = Support.SetParametersForSingleTarget("fizz", "Name");

            var enapsulationIdentifiers = new EncapsulationIdentifiers("fizz") { Property = "Name" };

            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            StringAssert.AreEqualIgnoringCase(enapsulationIdentifiers.Field, "fizz");
            StringAssert.Contains($"Private {enapsulationIdentifiers.Field} As Integer", actualCode);
            StringAssert.Contains("Property Get Name", actualCode);
            StringAssert.Contains("Property Let Name", actualCode);
            StringAssert.Contains($"Name = {enapsulationIdentifiers.Field}", actualCode);
            StringAssert.Contains($"{enapsulationIdentifiers.Field} = value", actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void GivenReferencedPublicField_UpdatesReferenceToNewProperty()
        {
            const string codeClass1 =
                @"|Public fizz As Integer

Sub Foo()
    fizz = 1
End Sub";
            const string codeClass2 =
                @"Sub Foo()
    Dim c As Class1
    c.fizz = 0
    Bar c.fizz
End Sub

Sub Bar(ByVal v As Integer)
End Sub";

            var presenterAction = Support.SetParametersForSingleTarget("fizz", "Name", true);

            var class1CodeString = codeClass1.ToCodeString();
            var actualCode = RefactoredCode(
                "Class1",
                class1CodeString.CaretPosition.ToOneBased(), 
                presenterAction, 
                null, 
                false, 
                ("Class1", class1CodeString.Code, ComponentType.ClassModule),
                ("Class2", codeClass2, ComponentType.ClassModule));

            StringAssert.Contains("Name = 1", actualCode["Class1"]);
            StringAssert.Contains("c.Name = 0", actualCode["Class2"]);
            StringAssert.Contains("Bar c.Name", actualCode["Class2"]);
            StringAssert.DoesNotContain("fizz", actualCode["Class2"]);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulateField_PresenterIsNull()
        {
            const string inputCode =
                @"Private fizz As Variant";
            
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using(state)
            {
                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), Selection.Home);
                var selectionService = MockedSelectionService();
                var factory = new Mock<IRefactoringPresenterFactory>();
                factory.Setup(f => f.Create<IEncapsulateFieldPresenter, EncapsulateFieldModel>(It.IsAny<EncapsulateFieldModel>()))
                    .Returns(() => null); // resolves ambiguous method overload

                var refactoring = TestRefactoring(rewritingManager, state, factory.Object, selectionService);

                Assert.Throws<InvalidRefactoringPresenterException>(() => refactoring.Refactor(qualifiedSelection));

                var actualCode = component.CodeModule.Content();
                Assert.AreEqual(inputCode, actualCode);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulateField_ModelIsNull()
        {
            const string inputCode =
                @"|Private fizz As Variant";

            Func<EncapsulateFieldModel, EncapsulateFieldModel> presenterAction = model => null;

            var codeString = inputCode.ToCodeString();
            var actualCode = Support.RefactoredCode(codeString, presenterAction, typeof(InvalidRefactoringModelException));
            Assert.AreEqual(codeString.Code, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulatePublicField_OptionExplicit_NotMoved()
        {
            const string inputCode =
                @"Option Explicit

Public foo As String";

            const string expectedCode =
                @"Option Explicit

Private foo As String

Public Property Get Name() As String
    Name = foo
End Property

Public Property Let Name(ByVal value As String)
    foo = value
End Property
";
            var presenterAction = Support.SetParametersForSingleTarget("foo", "Name");
            var actualCode = RefactoredCode(inputCode, "foo", DeclarationType.Variable, presenterAction);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void Refactoring_Puts_Code_In_Correct_Place()
        {
            const string inputCode =
                @"Option Explicit

Public Fo|o As String";

            const string expectedCode =
                @"Option Explicit

Private Foo As String

Public Property Get bar() As String
    bar = Foo
End Property

Public Property Let bar(ByVal value As String)
    Foo = value
End Property
";
            var presenterAction = Support.SetParametersForSingleTarget("Foo", "bar");
            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [TestCase("Private", "mArray(5) As String", "mArray(5) As String")]
        [TestCase("Public", "mArray(5) As String", "mArray(5) As String")]
        [TestCase("Private", "mArray(5,2,3) As String", "mArray(5,2,3) As String")]
        [TestCase("Public", "mArray(5,2,3) As String", "mArray(5,2,3) As String")]
        [TestCase("Private", "mArray(1 to 10) As String", "mArray(1 to 10) As String")]
        [TestCase("Public", "mArray(1 to 10) As String", "mArray(1 to 10) As String")]
        [TestCase("Private", "mArray() As String", "mArray() As String")]
        [TestCase("Public", "mArray() As String", "mArray() As String")]
        [TestCase("Private", "mArray(5)", "mArray(5) As Variant")]
        [TestCase("Public", "mArray(5)", "mArray(5) As Variant")]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulateArray(string visibility, string arrayDeclaration, string expectedArrayDeclaration)
        {
            string inputCode =
                $@"Option Explicit

{visibility} {arrayDeclaration}";

            var selection = new Selection(3, 8, 3, 11);

            string expectedCode =
                $@"Option Explicit

Private {expectedArrayDeclaration}

Public Property Get MyArray() As Variant
    MyArray = mArray
End Property
";
            var userInput = new UserInputDataObject("mArray", "MyArray", true);

            var presenterAction = Support.SetParameters(userInput);
            var actualCode = RefactoredCode(inputCode, selection, presenterAction);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [TestCase("5")]
        [TestCase("5,2,3")]
        [TestCase("1 to 100")]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulateArray_DeclaredInList(string dimensions)
        {
            string inputCode =
                $@"Option Explicit

Public mArray({dimensions}) As String, anotherVar As Long, andOneMore As Variant";

            var selection = new Selection(3, 8, 3, 11);

            string expectedCode =
                $@"Option Explicit

Public anotherVar As Long, andOneMore As Variant
Private mArray({dimensions}) As String

Public Property Get MyArray() As Variant
    MyArray = mArray
End Property
";
            var presenterAction = Support.SetParametersForSingleTarget("mArray", "MyArray");
            var actualCode = RefactoredCode(inputCode, selection, presenterAction);
            StringAssert.Contains("Public anotherVar As Long, andOneMore As Variant", actualCode);
            StringAssert.Contains($"Private mArray({dimensions}) As String", actualCode);
            StringAssert.Contains("Get MyArray() As Variant", actualCode);
            StringAssert.Contains("MyArray = mArray", actualCode);
            StringAssert.DoesNotContain("Let MyArray", actualCode);
            StringAssert.DoesNotContain("Set MyArray", actualCode);
        }

        [TestCase("mArr|ay(5) As String, mNextVar As Long", "Private mArray(5) As String")]
        [TestCase("mNextVar As Long, mArr|ay(5) As String", "Private mArray(5) As String")]
        [TestCase("mArr|ay(5), mNextVar As Long", "Private mArray(5) As Variant")]
        [TestCase("mNextVar As Long, mAr|ray(5)", "Private mArray(5) As Variant")]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulateArray_newFieldNameForFieldInList(string declarationList, string expectedDeclaration)
        {
            string inputCode =
                $@"Option Explicit

Public {declarationList}";

            string expectedCode =
                $@"Option Explicit

Public mNextVar As Long
{expectedDeclaration}

Public Property Get MyArray() As Variant
    MyArray = mArray
End Property
";
            var presenterAction = Support.SetParametersForSingleTarget("mArray", "MyArray");
            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            Assert.AreEqual(expectedCode, actualCode);
        }

        #region setup

        //private TestEncapsulationAttributes UserModifiedEncapsulationAttributes(string field, string property = null, bool? isReadonly = null, bool encapsulateFlag = true)
        //{
        //    var testAttrs = new TestEncapsulationAttributes(field, encapsulateFlag, isReadonly ?? false);
        //    if (property != null)
        //    {
        //        testAttrs.PropertyName = property;
        //    }
        //    return testAttrs;
        //}

        //private Func<EncapsulateFieldModel, EncapsulateFieldModel> UserAcceptsDefaults(bool asUDT = false)
        //{
        //    return model => {model.EncapsulateWithUDT = asUDT; return model; };
        //}

        //private Func<EncapsulateFieldModel, EncapsulateFieldModel> Support.SetParametersForSingleTarget(string field, string property = null, bool? isReadonly = null, bool encapsulateFlag = true, bool asUDT = false)
        //{
        //    var clientAttrs = UserModifiedEncapsulationAttributes(field, property, isReadonly, encapsulateFlag);

        //    return Support.SetParameters(field, clientAttrs, asUDT);
        //}

        //public Func<EncapsulateFieldModel, EncapsulateFieldModel> Support.SetParameters(UserInputDataObject userModifications)
        //{
        //    return model =>
        //    {
        //        foreach (var testModifiedAttribute in userModifications.EncapsulateFieldAttributes)
        //        {
        //            var attrsInitializedByTheRefactoring = model[testModifiedAttribute.TargetFieldName].EncapsulationAttributes;

        //            attrsInitializedByTheRefactoring.PropertyName = testModifiedAttribute.PropertyName;
        //            attrsInitializedByTheRefactoring.EncapsulateFlag = testModifiedAttribute.EncapsulateFlag;

        //            var currentAttributes = model[testModifiedAttribute.TargetFieldName].EncapsulationAttributes;
        //            currentAttributes.PropertyName = attrsInitializedByTheRefactoring.PropertyName;
        //            currentAttributes.EncapsulateFlag = attrsInitializedByTheRefactoring.EncapsulateFlag;

        //            foreach ((string instanceVariable, string memberName, bool flag) in userModifications.UDTMemberNameFlagPairs)
        //            {
        //                model[$"{instanceVariable}.{memberName}"].EncapsulateFlag = flag;
        //            }
        //        }
        //        return model;
        //    };
        //}

        //private Func<EncapsulateFieldModel, EncapsulateFieldModel> Support.SetParameters(string originalField, TestEncapsulationAttributes attrs, bool asUDT = false)
        //{
        //    return model =>
        //    {
        //        var encapsulatedField = model[originalField];
        //        encapsulatedField.EncapsulationAttributes.PropertyName = attrs.PropertyName;
        //        encapsulatedField.EncapsulationAttributes.ReadOnly = attrs.IsReadOnly;
        //        encapsulatedField.EncapsulationAttributes.EncapsulateFlag = attrs.EncapsulateFlag;

        //        model.EncapsulateWithUDT = asUDT;
        //        return model;
        //    };
        //}

        //public class TestEncapsulationAttributes
        //{
        //    public TestEncapsulationAttributes(string fieldName, bool encapsulationFlag = true, bool isReadOnly = false)
        //    {
        //        _identifiers = new EncapsulationIdentifiers(fieldName);
        //        EncapsulateFlag = encapsulationFlag;
        //        IsReadOnly = isReadOnly;
        //    }

        //    private EncapsulationIdentifiers _identifiers;
        //    public string TargetFieldName => _identifiers.TargetFieldName;

        //    public string NewFieldName
        //    {
        //        get => _identifiers.Field;
        //        set => _identifiers.Field = value;
        //    }
        //    public string PropertyName
        //    {
        //        get => _identifiers.Property;
        //        set => _identifiers.Property = value;
        //    }
        //    public bool EncapsulateFlag { get; set; }
        //    public bool IsReadOnly { get; set; }
        //}

        //public class UserInputDataObject
        //{
        //    private List<TestEncapsulationAttributes> _userInput = new List<TestEncapsulationAttributes>();
        //    private List<(string, string, bool)> _udtNameFlagPairs = new List<(string, string, bool)>();

        //    public UserInputDataObject(string fieldName, string propertyName = null, bool encapsulationFlag = true, bool isReadOnly = false)
        //        => AddAttributeSet(fieldName, propertyName, encapsulationFlag, isReadOnly);

        //    public void AddAttributeSet(string fieldName, string propertyName = null, bool encapsulationFlag = true, bool isReadOnly = false)
        //    {
        //        var attrs = new TestEncapsulationAttributes(fieldName, encapsulationFlag, isReadOnly);
        //        attrs.PropertyName = propertyName ?? attrs.PropertyName;

        //        _userInput.Add(attrs);
        //    }

        //    public TestEncapsulationAttributes this[string fieldName] 
        //        => EncapsulateFieldAttributes.Where(efa => efa.TargetFieldName == fieldName).Single();


        //    public IEnumerable<TestEncapsulationAttributes> EncapsulateFieldAttributes => _userInput;

        //    public void AddUDTMemberNameFlagPairs(params (string, string, bool)[] nameFlagPairs)
        //        => _udtNameFlagPairs.AddRange(nameFlagPairs);

        //    public IEnumerable<(string, string, bool)> UDTMemberNameFlagPairs => _udtNameFlagPairs;
        //}

        //private string Support.RefactoredCode(CodeString codeString, Func<EncapsulateFieldModel, EncapsulateFieldModel> presenterAdjustment, Type expectedException = null, bool executeViaActiveSelection = false) 
        //    => Support.RefactoredCode(codeString.Code, codeString.CaretPosition.ToOneBased(), presenterAdjustment, expectedException, executeViaActiveSelection);

        public static IIndenter CreateIndenter(IVBE vbe = null)
        {
            return new Indenter(vbe, () => Settings.IndenterSettingsTests.GetMockIndenterSettings());
        }

        //protected override IRefactoring TestRefactoring(IRewritingManager rewritingManager, RubberduckParserState state, IRefactoringPresenterFactory factory, ISelectionService selectionService)
        //{
        //    var indenter = CreateIndenter(); //The refactoring only uses method independent of the VBE instance.
        //    var selectedDeclarationProvider = new SelectedDeclarationProvider(selectionService, state);
        //    return new EncapsulateFieldRefactoring(state, indenter, factory, rewritingManager, selectionService, selectedDeclarationProvider);
        //}
        protected override IRefactoring TestRefactoring(IRewritingManager rewritingManager, RubberduckParserState state, IRefactoringPresenterFactory factory, ISelectionService selectionService)
        {
            var indenter = CreateIndenter(); //The refactoring only uses method independent of the VBE instance.
            var selectedDeclarationProvider = new SelectedDeclarationProvider(selectionService, state);
            return new EncapsulateFieldRefactoring(state, indenter, factory, rewritingManager, selectionService, selectedDeclarationProvider);
        }

        #endregion
    }
}
