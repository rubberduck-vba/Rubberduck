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

namespace RubberduckTests.Refactoring
{
    [TestFixture]
    public class EncapsulateFieldTests : InteractiveRefactoringTestBase<IEncapsulateFieldPresenter, EncapsulateFieldModel>
    {
        [TestCase("fizz")]
        [TestCase("mFizz")]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulatePublicField_WithLet(string newFieldName)
        {
            //Input
            const string inputCode =
                @"Public fizz As Integer";
            var selection = new Selection(1, 1);

            //Expectation
            string expectedCode =
                $@"Private {newFieldName} As Integer

Public Property Get Name() As Integer
    Name = {newFieldName}
End Property

Public Property Let Name(ByVal value As Integer)
    {newFieldName} = value
End Property
";
            var presenterAction = SetParametersForSingleTarget("fizz", "Name", newFieldName: newFieldName);
            var actualCode = RefactoredCode(inputCode, selection, presenterAction);
            Assert.AreEqual(expectedCode, actualCode);
        }

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

            var userInput = new TestInputDataObject(var1, $"{var1}Prop", var1Flag, var1);
            userInput.AddAttributeSet(var2, $"{var2}Prop", var2Flag, var2);
            userInput.AddAttributeSet(var3, $"{var3}Prop", var3Flag, var3);

            var flags = new Dictionary<string, bool>()
            {
                [var1] = var1Flag,
                [var2] = var2Flag,
                [var3] = var3Flag
            };

            var presenterAction = SetParameters(userInput);

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

            var userInput = new TestInputDataObject(var1, $"{var1}Prop", var1Flag);
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

            var presenterAction = SetParameters(userInput);

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
                    StringAssert.Contains($"Private {key}1 As", actualCode);
                    StringAssert.Contains($"{key}Prop = {key}1", actualCode);
                    StringAssert.Contains($"{key}1 = value", actualCode);
                    StringAssert.Contains($"Let {key}Prop(ByVal value As", actualCode);
                    StringAssert.Contains($"Property Get {key}Prop()", actualCode);
                }
            }
        }

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

{accessibility} this As TBar";

            var selection = new Selection(7, 10); //Selects 'this' declaration

            var presenterAction = UserAcceptsDefaults();
            var actualCode = RefactoredCode(inputCode, selection, presenterAction);
            StringAssert.Contains("Private this1 As TBar", actualCode);
            StringAssert.Contains("this1 = value", actualCode);
            StringAssert.Contains($"This = this1", actualCode);
            StringAssert.DoesNotContain($"this1.First1 = value", actualCode);
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

{accessibility} this As TBar
{accessibility} that As TBar";

            var selection = new Selection(7, 10); //Selects 'this' declaration

            var userInput = new TestInputDataObject("this", "MyType", encapsulateThis);
            userInput.AddAttributeSet("that", "MyOtherType", encapsulateThat);
            userInput.AddUDTMemberNameFlagPairs(("this", "First", true));
            userInput.AddUDTMemberNameFlagPairs(("this", "Second", true));
            userInput.AddUDTMemberNameFlagPairs(("that", "First", true));
            userInput.AddUDTMemberNameFlagPairs(("that", "Second", true));

            var thisDefaults = new ClientEncapsulationAttributes("this");
            var thatDefaults = new ClientEncapsulationAttributes("that");

            var presenterAction = SetParameters(userInput);
            var actualCode = RefactoredCode(inputCode, selection, presenterAction);
            if (encapsulateThis)
            {
                StringAssert.Contains($"Private {thisDefaults.NewFieldName} As TBar", actualCode);
                StringAssert.Contains($"MyType = {thisDefaults.NewFieldName}", actualCode);
                StringAssert.Contains($"{thisDefaults.NewFieldName} = value", actualCode);
                StringAssert.Contains($"{thisDefaults.NewFieldName}.First = value", actualCode);
                StringAssert.Contains($"{thisDefaults.NewFieldName}.First = value", actualCode);
                StringAssert.Contains($"{thisDefaults.NewFieldName}.Second = value", actualCode);
            }
            else
            {
                StringAssert.Contains($"{accessibility} this As TBar", actualCode);
                StringAssert.Contains($"this.First = value", actualCode);
                StringAssert.Contains($"this.Second = value", actualCode);
            }

            if (encapsulateThat)
            {
                StringAssert.Contains($"Private {thatDefaults.NewFieldName} As TBar", actualCode);
                StringAssert.Contains($"MyOtherType = {thatDefaults.NewFieldName}", actualCode);
                StringAssert.Contains($"{thatDefaults.NewFieldName} = value", actualCode);
                StringAssert.Contains($"{thatDefaults.NewFieldName}.First = value", actualCode);
                StringAssert.Contains($"{thatDefaults.NewFieldName}.Second = value", actualCode);
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

Private this As TBar";

            var selection = new Selection(7, 10); //Selects 'this' declaration


            var userInput = new TestInputDataObject("this", "MyType", true, "this");
            userInput.AddUDTMemberNameFlagPairs(("this", memberToEncapsulate, true));

            var presenterAction = SetParameters(userInput);

            var actualCode = RefactoredCode(inputCode, selection, presenterAction);
            StringAssert.Contains($"this.{memberToEncapsulate} = value", actualCode);
            StringAssert.Contains($"{memberToEncapsulate} = this.{memberToEncapsulate}", actualCode);
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

{accessibility} this As TBar";

            var selection = new Selection(7, 10); //Selects 'this' declaration

            var userInput = new TestInputDataObject("this", "MyType", true, "this");
            userInput.AddUDTMemberNameFlagPairs(("this", "First", true), ("this", "Second", true));
            userInput.AddAttributeSet("mFoo", "Foo", true, "mFoo");
            userInput.AddAttributeSet("mBar", "Bar", true, "mBar");
            userInput.AddAttributeSet("mFizz", "Fizz", true, "mFizz");

            var presenterAction = SetParameters(userInput);

            var actualCode = RefactoredCode(inputCode, selection, presenterAction);

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
        public void UserDefinedTypeMembersAndFieldsWithTypeNameChange(string accessibility)
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

{accessibility} that As TBar";

            var selection = new Selection(7, 10); //Selects 'that' declaration

            var userInput = new TestInputDataObject("that", "MyType", true, "this");
            userInput.AddUDTMemberNameFlagPairs(("that", "First", true), ("that", "Second", true));
            userInput.AddAttributeSet("mFoo", "Foo", true,"mFoo");
            userInput.AddAttributeSet("mBar", "Bar", true, "mBar");
            userInput.AddAttributeSet("mFizz", "Fizz", true, "mFizz");

            var presenterAction = SetParameters(userInput);

            var actualCode = RefactoredCode(inputCode, selection, presenterAction);

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

{accessibility} this As TBar";

            string class1Code =
@"Option Explicit

Public Sub Foo()
End Sub
";

            var selection = new Selection(7, 10); //Selects 'this' declaration

            var userInput = new TestInputDataObject("this", "MyType", true, "this");

            userInput.AddUDTMemberNameFlagPairs(("this", "First", true), ("this", "Second", true));

            var presenterAction = SetParameters(userInput);

            var actualModuleCode = RefactoredCode(
                "Module1",
                selection,
                presenterAction,
                null,
                false,
                ("Class1", class1Code, ComponentType.ClassModule),
                ("Module1", inputCode, ComponentType.StandardModule));

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

{accessibility} this As TBar";

            var selection = new Selection(7, 10); //Selects 'this' declaration

            var userInput = new TestInputDataObject("this", "MyType", true, "this");
            userInput.AddUDTMemberNameFlagPairs(("this", "First", true), ("this", "Second", true));

            var presenterAction = SetParameters(userInput);

            var actualCode = RefactoredCode(inputCode, selection, presenterAction);
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

{accessibility} this As TBar";

            var selection = new Selection(8, 10); //Selects 'this' declaration

            var userInput = new TestInputDataObject("this", "MyType", false, "this");
            userInput.AddUDTMemberNameFlagPairs(("this", "First", true), ("this", "Second", true), ("this", "Third", true));

            var presenterAction = SetParameters(userInput);

            var actualCode = RefactoredCode(inputCode, selection, presenterAction);
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

{accessibility} this As TBar";

            var selection = new Selection(7, 10); //Selects 'this' declaration

            var userInput = new TestInputDataObject("this", "MyType", false, "this");
            userInput.AddUDTMemberNameFlagPairs(("this", "First", true), ("this", "Second", true));

            var presenterAction = SetParameters(userInput);

            var actualCode = RefactoredCode(inputCode, selection, presenterAction);
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

{accessibility} this As TBar";

            string typeDefinition =
$@"
Public Type TBar
    First As String
    Second As Long
End Type
";

            var selection = new Selection(4, 10); //Selects 'this' declaration

            var userInput = new TestInputDataObject("this", "MyType", true, "this");
            userInput.AddUDTMemberNameFlagPairs(("this", "First", true), ("this", "Second", true));

            var presenterAction = SetParameters(userInput);

            var actualModuleCode = RefactoredCode(
                "Class1",
                selection,
                presenterAction,
                null,
                false,
                ("Class1", inputCode, ComponentType.ClassModule),
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

{accessibility} mTheClass As Class1";

            string classContent =
$@"
Option Explicit

Private Sub Class_Initialize()
End Sub
";

            var selection = new Selection(4, 10); //Selects 'mTheClass' declaration

            var presenterAction = UserAcceptsDefaults();

            var actualModuleCode = RefactoredCode(
                "Module1",
                selection,
                presenterAction,
                null,
                false,
                ("Class1", classContent, ComponentType.ClassModule),
                ("Module1", inputCode, ComponentType.StandardModule));

            var actualCode = actualModuleCode["Module1"];

            var defaults = new ClientEncapsulationAttributes("mTheClass");

            StringAssert.Contains($"Private {defaults.NewFieldName} As Class1", actualCode);
            StringAssert.Contains($"Set {defaults.NewFieldName} = value", actualCode);
            StringAssert.Contains($"MTheClass = {defaults.NewFieldName}", actualCode);
            StringAssert.Contains($"Public Property Set {defaults.PropertyName}", actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulatePublicField_InvalidDeclarationType_Throws()
        {
            //Input
            const string inputCode =
                @"Public fizz As Integer";

            var presenterAction = SetParametersForSingleTarget("fizz","Name", newFieldName: "fizz");
            var actualCode = RefactoredCode(inputCode, "TestModule1", DeclarationType.ProceduralModule, presenterAction, typeof(InvalidDeclarationTypeException));
            Assert.AreEqual(inputCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulatePublicField_InvalidIdentifierSelected_Throws()
        {
            //Input
            const string inputCode =
                @"Public Function fizz() As Integer
End Function";
            var selection = new Selection(1, 19);

            var presenterAction = SetParametersForSingleTarget("fizz", "Name", newFieldName: "fizz");
            var actualCode = RefactoredCode(inputCode, selection, presenterAction, typeof(NoDeclarationForSelectionException));
            Assert.AreEqual(inputCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulatePublicField_FieldIsOverMultipleLines()
        {
            //Input
            const string inputCode =
                @"Public _
fizz _
As _
Integer";
            var selection = new Selection(1, 1);

            //Expectation
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
            var presenterAction = SetParametersForSingleTarget("fizz", "Name", newFieldName: "fizz");
            var actualCode = RefactoredCode(inputCode, selection, presenterAction);
            Assert.AreEqual(expectedCode, actualCode);
        }

        //TODO: It does not seem appropriate to generate a variant property and limit
        //it to either Set or Let.
//        [Test]
//        [Category("Refactorings")]
//        [Category("Encapsulate Field")]
//        public void EncapsulatePublicField_WithSetter()
//        {
//            //Input
//            const string inputCode =
//                @"Public fizz As Variant";
//            var selection = new Selection(1, 1);

//            //Expectation
//            const string expectedCode =
//                @"Private fizz As Variant

//Public Property Get Name() As Variant
//    Set Name = fizz
//End Property

//Public Property Set Name(ByVal value As Variant)
//    Set fizz = value
//End Property
//";
//            var presenterAction = SetParameters("Name", implementSet: true);
//            var actualCode = RefactoredCode(inputCode, selection, presenterAction);
//            Assert.AreEqual(expectedCode, actualCode);
//        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulatePublicField_WithOnlyGetter()
        {
            //Input
            const string inputCode =
                @"Public fizz As Integer";
            var selection = new Selection(1, 1);

            //Expectation
            const string expectedCode =
                @"Private fizz As Integer

Public Property Get Name() As Integer
    Name = fizz
End Property
";
            var presenterAction = SetParametersForSingleTarget("fizz","Name", isReadonly: true, newFieldName: "fizz");
            var actualCode = RefactoredCode(inputCode, selection, presenterAction);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulatePublicField_OtherMethodsInClass()
        {
            //Input
            const string inputCode =
                @"Public fizz As Integer

Sub Foo()
End Sub

Function Bar() As Integer
    Bar = 0
End Function";
            var selection = new Selection(1, 1);

            //var presenterAction = SetParameters("Name", implementLet: true);
            var presenterAction = SetParametersForSingleTarget("fizz", "Name", newFieldName: "fizz");
            var actualCode = RefactoredCode(inputCode, selection, presenterAction);
            Assert.Greater(actualCode.IndexOf("Sub Foo"), actualCode.LastIndexOf("End Property"));
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulatePublicField_OtherPropertiesInClass()
        {
            //Input
            const string inputCode =
                @"Public fizz As Integer

Property Get Foo() As Variant
    Foo = True
End Property

Property Let Foo(ByVal vall As Variant)
End Property

Property Set Foo(ByVal vall As Variant)
End Property";
            var selection = new Selection(1, 1);

            //var presenterAction = SetParameters("Name", implementLet: true);
            var presenterAction = SetParametersForSingleTarget("fizz", "Name", newFieldName: "fizz");
            var actualCode = RefactoredCode(inputCode, selection, presenterAction);
            Assert.Greater(actualCode.IndexOf("fizz = value"), actualCode.IndexOf("fizz As Integer"));
            Assert.Less(actualCode.IndexOf("fizz = value"), actualCode.IndexOf("Get Foo"));
        }

        [TestCase("Public fizz As Integer\r\nPublic buzz As Boolean", "Private fizz As Integer\r\nPublic buzz As Boolean", 1, 1)]
        [TestCase("Public buzz As Boolean\r\nPublic fizz As Integer", "Public buzz As Boolean\r\nPrivate fizz As Integer", 2, 1)]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulatePublicField_OtherFieldsInClass(string inputFields, string expectedFields, int rowSelection, int columnSelection)
        {
            string inputCode = inputFields;

            var selection = new Selection(rowSelection, columnSelection);

            string expectedCode =
$@"{expectedFields}

Public Property Get Name() As Integer
    Name = fizz
End Property

Public Property Let Name(ByVal value As Integer)
    fizz = value
End Property
";
            var presenterAction = SetParametersForSingleTarget("fizz", "Name", newFieldName: "fizz");
            var actualCode = RefactoredCode(inputCode, selection, presenterAction);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [TestCase(1, 10, "fizz", "Public buzz", "Private fizz As Variant", "Public fizz")]
        [TestCase(2, 2, "buzz", "Public fizz, _\r\nbazz", "Private buzz As Boolean", "")]
        [TestCase(3, 2, "bazz", "Public fizz, _\r\nbuzz", "Private bazz As Date", "Boolean, bazz As Date")]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulatePublicField_SelectedWithinDeclarationList(int rowSelection, int columnSelection, string fieldName, string contains1, string contains2, string doesNotContain)
        {
            //Input
            string inputCode =
$@"Public fizz, _
buzz As Boolean, _
bazz As Date";

            var selection = new Selection(rowSelection, columnSelection);
            var presenterAction = SetParametersForSingleTarget(fieldName, "Name", newFieldName: fieldName);
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
            //Input
            const string inputCode =
                @"Private fizz As Integer";
            var selection = new Selection(1, 1);

            //Expectation
            const string expectedCode =
                @"Private fizz As Integer

Public Property Get Name() As Integer
    Name = fizz
End Property

Public Property Let Name(ByVal value As Integer)
    fizz = value
End Property
";
            var presenterAction = SetParametersForSingleTarget("fizz", "Name", newFieldName: "fizz");
            var actualCode = RefactoredCode(inputCode, selection, presenterAction);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulatePrivateField_Defaults()
        {
            //Input
            const string inputCode =
                @"Private fizz As Integer";
            var selection = new Selection(1, 1);

            //Expectation
            const string expectedCode =
                @"Private fizz1 As Integer

Public Property Get Fizz() As Integer
    Fizz = fizz1
End Property

Public Property Let Fizz(ByVal value As Integer)
    fizz1 = value
End Property
";
            var presenterAction = UserAcceptsDefaults();
            var actualCode = RefactoredCode(inputCode, selection, presenterAction);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [TestCase("fizz")]
        [TestCase("mFizz")]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulatePublicField_FieldHasReferences(string newName)
        {
            //Input
            const string inputCode =
                @"Public fizz As Integer

Sub Foo()
    fizz = 0
    Bar fizz
End Sub

Sub Bar(ByVal name As Integer)
End Sub";
            var selection = new Selection(1, 1);

            var presenterAction = SetParametersForSingleTarget("fizz", "Name", newFieldName: newName);

            var actualCode = RefactoredCode(inputCode, selection, presenterAction);
            StringAssert.Contains($"Private {newName} As Integer", actualCode);
            StringAssert.Contains("Property Get Name", actualCode);
            StringAssert.Contains("Property Let Name", actualCode);
            StringAssert.Contains($"Name = {newName}", actualCode);
            StringAssert.Contains($"{newName} = value", actualCode);
            Assert.Greater(actualCode.IndexOf("Sub Foo"), actualCode.LastIndexOf("End Property"));
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void GivenReferencedPublicField_UpdatesReferenceToNewProperty()
        {
            //Input
            const string codeClass1 =
                @"Public fizz As Integer

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

            var selection = new Selection(1, 1);

            var presenterAction = SetParametersForSingleTarget("fizz", "Name", true);

            var actualCode = RefactoredCode(
                "Class1", 
                selection, 
                presenterAction, 
                null, 
                false, 
                ("Class1", codeClass1, ComponentType.ClassModule),
                ("Class2", codeClass2, ComponentType.ClassModule));

            StringAssert.Contains("Name = 1", actualCode["Class1"]);
            StringAssert.Contains("c.Name = 0", actualCode["Class2"]);
            StringAssert.Contains("Bar c.Name", actualCode["Class2"]);
            StringAssert.DoesNotContain("fizz", actualCode["Class2"]);
        }

//        [Test]
//        [Category("Refactorings")]
//        [Category("Encapsulate Field")]
//        public void EncapsulatePublicField_PassInTarget()
//        {
//            //Input
//            const string inputCode =
//                @"Private fizz As Integer";

//            //Expectation
//            const string expectedCode =
//                @"Private fizz As Integer

//Public Property Get Name() As Integer
//    Name = fizz
//End Property

//Public Property Let Name(ByVal value As Integer)
//    fizz = value
//End Property
//";
//            //var presenterAction = SetParameters("Name", implementLet: true);
//            var presenterAction = SetParametersForSingleTarget("fizz", "Name", newFieldName: "fizz");
//            var actualCode = RefactoredCode(inputCode, "fizz", DeclarationType.Variable, presenterAction);
//            Assert.AreEqual(expectedCode, actualCode);
//        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulateField_PresenterIsNull()
        {
            //Input
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
            //Input
            const string inputCode =
                @"Private fizz As Variant";
            var selection = new Selection(1, 1);

            Func<EncapsulateFieldModel, EncapsulateFieldModel> presenterAction = model => null;
            var actualCode = RefactoredCode(inputCode, selection, presenterAction, typeof(InvalidRefactoringModelException));
            Assert.AreEqual(inputCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulatePublicField_OptionExplicit_NotMoved()
        {
            //Input
            const string inputCode =
                @"Option Explicit

Public foo As String";

            //Expectation
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
            //var presenterAction = SetParameters("Name", implementLet: true);
            var presenterAction = SetParametersForSingleTarget("foo", "Name", newFieldName: "foo");
            var actualCode = RefactoredCode(inputCode, "foo", DeclarationType.Variable, presenterAction);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void Refactoring_Puts_Code_In_Correct_Place()
        {
            //Input
            const string inputCode =
                @"Option Explicit

Public Foo As String";

            var selection = new Selection(3, 8, 3, 11);

            //Output
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
            var presenterAction = SetParametersForSingleTarget("Foo", "bar", newFieldName: "Foo");
            var actualCode = RefactoredCode(inputCode, selection, presenterAction);
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
            var userInput = new TestInputDataObject("mArray", "MyArray", true, "mArray");
            //userInput.AddUDTMemberNameFlagPairs(("First", true), ("Second", true));

            //var presenterAction = SetParameters("MyArray");
            var presenterAction = SetParameters(userInput);
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
            var presenterAction = SetParametersForSingleTarget("mArray", "MyArray", newFieldName: "mArray");
            var actualCode = RefactoredCode(inputCode, selection, presenterAction);
            StringAssert.Contains("Public anotherVar As Long, andOneMore As Variant", actualCode);
            StringAssert.Contains($"Private mArray({dimensions}) As String", actualCode);
            StringAssert.Contains("Get MyArray() As Variant", actualCode);
            StringAssert.Contains("MyArray = mArray", actualCode);
            StringAssert.DoesNotContain("Let MyArray", actualCode);
            StringAssert.DoesNotContain("Set MyArray", actualCode);
        }

        [TestCase("Private")]
        [TestCase("Public")]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulateArray_newFieldName(string visibility)
        {
            string inputCode =
                $@"Option Explicit

{visibility} mArray(5) As String";

            var selection = new Selection(3, 8, 3, 11);

            string expectedCode =
                $@"Option Explicit

Private xArray(5) As String

Public Property Get MyArray() As Variant
    MyArray = xArray
End Property
";

            var presenterAction = SetParametersForSingleTarget("mArray", "MyArray", newFieldName: "xArray");

            var actualCode = RefactoredCode(inputCode, selection, presenterAction);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [TestCase("mArray(5) As String, mNextVar As Long", "xArray(5) As String", 8)]
        [TestCase("mNextVar As Long, mArray(5) As String", "xArray(5) As String", 27)]
        [TestCase("mArray(5), mNextVar As Long", "xArray(5) As Variant", 8)]
        [TestCase("mNextVar As Long, mArray(5)", "xArray(5) As Variant", 27)]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulateArray_newFieldNameForFieldInList(string declarationList, string newArrayDeclaration, int column)
        {
            string inputCode =
                $@"Option Explicit

Public {declarationList}";

            var selection = new Selection(3, column);

            string expectedCode =
                $@"Option Explicit

Public mNextVar As Long
Private {newArrayDeclaration}

Public Property Get MyArray() As Variant
    MyArray = xArray
End Property
";
            var presenterAction = SetParametersForSingleTarget("mArray", "MyArray", newFieldName: "xArray");
            var actualCode = RefactoredCode(inputCode, selection, presenterAction);
            Assert.AreEqual(expectedCode, actualCode);
        }


        #region setup

        private Func<EncapsulateFieldModel, EncapsulateFieldModel> UserAcceptsDefaults()
        {
            return model => { model.EncapsulateFlag = true; return model; };
        }

        private Func<EncapsulateFieldModel, EncapsulateFieldModel> SetParametersForSingleTarget(string field, string property = null, bool? isReadonly = null, bool encapsulateFlag = true, string newFieldName = null)
        {
            var clientAttrs = new ClientEncapsulationAttributes(field);
            clientAttrs.NewFieldName = newFieldName ?? clientAttrs.NewFieldName;
            clientAttrs.PropertyName = property ?? clientAttrs.PropertyName;
            clientAttrs.ReadOnly = isReadonly ?? clientAttrs.ReadOnly;
            clientAttrs.EncapsulateFlag = encapsulateFlag;

            return SetParameters(field, clientAttrs);
        }

        private Func<EncapsulateFieldModel, EncapsulateFieldModel> SetParameters(TestInputDataObject testInput)
        {
            return model =>
            {
                foreach (var testModifiedAttribute in testInput.EncapsulateFieldAttributes)
                {
                    var attrsInitializedByTheRefactoring = model[testModifiedAttribute.FieldName].EncapsulationAttributes as IClientEditableFieldEncapsulationAttributes;
                    var testSupport = model as ISupportEncapsulateFieldTests;

                    attrsInitializedByTheRefactoring.NewFieldName = testModifiedAttribute.NewFieldName;
                    attrsInitializedByTheRefactoring.PropertyName = testModifiedAttribute.PropertyName;
                    attrsInitializedByTheRefactoring.EncapsulateFlag = testModifiedAttribute.EncapsulateFlag;

                    testSupport.SetEncapsulationFieldAttributes(testModifiedAttribute.FieldName, attrsInitializedByTheRefactoring);

                    foreach ((string instanceVariable, string memberName, bool flag) in testInput.UDTMemberNameFlagPairs)
                    {
                        testSupport.SetMemberEncapsulationFlag($"{instanceVariable}.{memberName}", flag);
                    }
                }
                return model;
            };
        }

        private Func<EncapsulateFieldModel, EncapsulateFieldModel> SetParameters(string originalField, IClientEditableFieldEncapsulationAttributes clientEdits)
        {
            return model =>
            {
                model.ApplyAttributes(originalField, clientEdits);
                return model;
            };
        }

        private class UserModifiableFieldEncapsulationAttributes : IClientEditableFieldEncapsulationAttributes
        {
            public UserModifiableFieldEncapsulationAttributes(string fieldName, string propertyName, bool encapsulationFlag = true, string newFieldName = null)
            {
                FieldName = fieldName;
                NewFieldName = newFieldName ?? fieldName;
                PropertyName = propertyName;
                EncapsulateFlag = encapsulationFlag;
            }

            public string FieldName { get; set; }
            public string NewFieldName { get; set; }
            public string PropertyName { get; set; }
            public bool EncapsulateFlag { get; set; }
            public bool ReadOnly { get; set; }
        }

        private class TestInputDataObject
        {
            private List<IClientEditableFieldEncapsulationAttributes> _userInput = new List<IClientEditableFieldEncapsulationAttributes>();
            private List<(string, string, bool)> _udtNameFlagPairs = new List<(string, string, bool)>();

            public TestInputDataObject(string fieldName, string propertyName = null, bool encapsulationFlag = true, string newFieldName = null, bool isReadOnly = false)
                => AddAttributeSet(fieldName, propertyName, encapsulationFlag, newFieldName, isReadOnly);

            public void AddAttributeSet(string fieldName, string propertyName = null, bool encapsulationFlag = true, string newFieldName = null, bool isReadOnly = false)
            {
                var attrs = new ClientEncapsulationAttributes(fieldName)
                {
                    EncapsulateFlag = encapsulationFlag,
                    ReadOnly = isReadOnly,
                };
                attrs.PropertyName = propertyName ?? attrs.PropertyName;
                attrs.NewFieldName = newFieldName ?? attrs.NewFieldName;

                _userInput.Add(attrs as IClientEditableFieldEncapsulationAttributes);
            }

            public IEnumerable<IClientEditableFieldEncapsulationAttributes> EncapsulateFieldAttributes => _userInput;

            public void AddUDTMemberNameFlagPairs(params (string, string, bool)[] nameFlagPairs)
                => _udtNameFlagPairs.AddRange(nameFlagPairs);

            public IEnumerable<(string, string, bool)> UDTMemberNameFlagPairs => _udtNameFlagPairs;
        }

        private static IIndenter CreateIndenter(IVBE vbe = null)
        {
            return new Indenter(vbe, () => Settings.IndenterSettingsTests.GetMockIndenterSettings());
        }

        protected override IRefactoring TestRefactoring(IRewritingManager rewritingManager, RubberduckParserState state, IRefactoringPresenterFactory factory, ISelectionService selectionService)
        {
            var indenter = CreateIndenter(); //The refactoring only uses method independent of the VBE instance.
            var selectedDeclarationProvider = new SelectedDeclarationProvider(selectionService, state);
            return new EncapsulateFieldRefactoring(state, indenter, factory, rewritingManager, selectionService, selectedDeclarationProvider);
        }

        #endregion
    }
}
