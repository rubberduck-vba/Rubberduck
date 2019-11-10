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
            var presenterAction = SetParametersForSingleTarget("fizz", "Name", true, newFieldName);
            var actualCode = RefactoredCode(inputCode, selection, presenterAction);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulateMultipleFields()
        {
            //Input
            const string inputCode =
@"Public fizz As Integer
Public baz As String
Public buzz As Boolean";

            var selection = new Selection(1, 1);

            var userInput = new UserInputDataObject("fizz", "FizzProp", true);
            userInput.AddAttributeSet("baz", "BazProp", true);
            userInput.AddAttributeSet("buzz", "BuzzProp", true);

            var presenterAction = SetParameters(userInput);

            var actualCode = RefactoredCode(inputCode, selection, presenterAction);
            StringAssert.Contains("Private fizz As Integer", actualCode);
            StringAssert.Contains("Private baz As String", actualCode);
            StringAssert.Contains("Private buzz As Boolean", actualCode);
            StringAssert.Contains("Property Get FizzProp() As Integer", actualCode);
            StringAssert.Contains("FizzProp = fizz", actualCode);
            StringAssert.Contains("Let FizzProp(ByVal value As Integer)", actualCode);
            StringAssert.Contains("fizz = value", actualCode);
            StringAssert.Contains("Property Get BazProp() As String", actualCode);
            StringAssert.Contains("BazProp = baz", actualCode);
            StringAssert.Contains("Let BazProp(ByVal value As String)", actualCode);
            StringAssert.Contains("baz = value", actualCode);
            StringAssert.Contains("Property Get BuzzProp() As Boolean", actualCode);
            StringAssert.Contains("BuzzProp = buzz", actualCode);
            StringAssert.Contains("Let BuzzProp(ByVal value As Boolean)", actualCode);
            StringAssert.Contains("buzz = value", actualCode);
        }

        [TestCase("Public")]
        [TestCase("Private")]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulateUserDefinedType(string accessibility)
        {
            string inputCode =
$@"
Private Type TBar
    First As String
    Second As Long
End Type

{accessibility} this As TBar";

            var selection = new Selection(7, 10); //Selects 'this' declaration

            var presenterAction = SetParametersForSingleTarget("this", "MyType", true);
            var actualCode = RefactoredCode(inputCode, selection, presenterAction);
            StringAssert.Contains("Private this As TBar", actualCode);
            StringAssert.Contains("this = value", actualCode);
            StringAssert.Contains($"MyType = this", actualCode);
            StringAssert.DoesNotContain($"this.First = value", actualCode);
        }

        [TestCase("First")]
        [TestCase("Second")]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulateUserDefinedTypeMembers_Subset(string memberToEncapsulate)
        {
            string inputCode =
$@"
Private Type TBar
    First As String
    Second As Long
End Type

Private this As TBar";

            var selection = new Selection(7, 10); //Selects 'this' declaration


            var userInput = new UserInputDataObject("this", "MyType", true);
            userInput.AddUDTMemberNameFlagPairs((memberToEncapsulate, true));

            var presenterAction = SetParameters(userInput);

            var actualCode = RefactoredCode(inputCode, selection, presenterAction);
            StringAssert.Contains($"this.{memberToEncapsulate} = value", actualCode);
            StringAssert.Contains($"{memberToEncapsulate} = this.{memberToEncapsulate}", actualCode);
        }

        [TestCase("Public")]
        [TestCase("Private")]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulateUserDefinedTypeMember(string accessibility)
        {
            string inputCode =
$@"
Private Type TBar
    First As String
    Second As Long
End Type

{accessibility} this As TBar";

            var selection = new Selection(7, 10); //Selects 'this' declaration

            var userInput = new UserInputDataObject("this", "MyType", true);
            userInput.AddUDTMemberNameFlagPairs(("First", true), ("Second", true));

            var attributes = new FieldEncapsulationAttributes()
            {
                FieldName = "this",
                PropertyName = "MyType",
                AsTypeName = "TBar",
                ImplementLetSetterType = true,
                EncapsulateFlag = true
            };

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
        }

        [TestCase("Public")]
        [TestCase("Private")]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulateUserDefinedTypeMember_ContainsObjects(string accessibility)
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

            var userInput = new UserInputDataObject("this", "MyType", true);
            userInput.AddUDTMemberNameFlagPairs(("First", true), ("Second", true));

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
        public void EncapsulateUserDefinedTypeMember_ContainsVariant(string accessibility)
        {
            string inputCode =
$@"
Private Type TBar
    First As String
    Second As Variant
End Type

{accessibility} this As TBar";

            var selection = new Selection(7, 10); //Selects 'this' declaration

            var userInput = new UserInputDataObject("this", "MyType", true);
            userInput.AddUDTMemberNameFlagPairs(("First", true), ("Second", true));

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
        public void EncapsulateUserDefinedTypeMember_ContainsArrays(string accessibility)
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

            var userInput = new UserInputDataObject("this", "MyType", false);
            userInput.AddUDTMemberNameFlagPairs(("First", true), ("Second", true), ("Third", true));

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
        public void EncapsulateUserDefinedTypeMembersOnly(string accessibility, string expectedAccessibility)
        {
            string inputCode =
$@"
Private Type TBar
    First As String
    Second As Long
End Type

{accessibility} this As TBar";

            var selection = new Selection(7, 10); //Selects 'this' declaration

            var userInput = new UserInputDataObject("this", "MyType", false);
            userInput.AddUDTMemberNameFlagPairs(("First", true), ("Second", true));

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
        public void EncapsulateUserDefinedTypeMembers_ExternallyDefinedType(string accessibility)
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

            var userInput = new UserInputDataObject("this", "MyType", true);
            userInput.AddUDTMemberNameFlagPairs(("First", true), ("Second", true));

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
        public void EncapsulateUserDefinedTypeMembers_ObjectField(string accessibility)
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

            var presenterAction = SetParametersForSingleTarget("mTheClass", "TheClass", true);

            var actualModuleCode = RefactoredCode(
                "Module1",
                selection,
                presenterAction,
                null,
                false,
                ("Class1", classContent, ComponentType.ClassModule),
                ("Module1", inputCode, ComponentType.StandardModule));

            var actualCode = actualModuleCode["Module1"];

            StringAssert.Contains("Private mTheClass As Class1", actualCode);
            StringAssert.Contains("Set mTheClass = value", actualCode);
            StringAssert.Contains("TheClass = mTheClass", actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulatePublicField_InvalidDeclarationType_Throws()
        {
            //Input
            const string inputCode =
                @"Public fizz As Integer";

            var presenterAction = SetParameters("Name", implementLet: true);
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

            var presenterAction = SetParameters("Name", implementLet: true);
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
            var presenterAction = SetParameters("Name", implementLet: true);
            var actualCode = RefactoredCode(inputCode, selection, presenterAction);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulatePublicField_WithSetter()
        {
            //Input
            const string inputCode =
                @"Public fizz As Variant";
            var selection = new Selection(1, 1);

            //Expectation
            const string expectedCode =
                @"Private fizz As Variant

Public Property Get Name() As Variant
    Set Name = fizz
End Property

Public Property Set Name(ByVal value As Variant)
    Set fizz = value
End Property
";
            var presenterAction = SetParameters("Name", implementSet: true);
            var actualCode = RefactoredCode(inputCode, selection, presenterAction);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulatePublicField_WithOnlyGetter()
        {
            //Input
            const string inputCode =
                @"Public fizz As Variant";
            var selection = new Selection(1, 1);

            //Expectation
            const string expectedCode =
                @"Private fizz As Variant

Public Property Get Name() As Variant
    Name = fizz
End Property
";
            var presenterAction = SetParameters("Name");
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

            //Expectation
            const string expectedCode =
                @"Private fizz As Integer

Public Property Get Name() As Integer
    Name = fizz
End Property

Public Property Let Name(ByVal value As Integer)
    fizz = value
End Property


Sub Foo()
End Sub

Function Bar() As Integer
    Bar = 0
End Function";
            var presenterAction = SetParameters("Name", implementLet: true);
            var actualCode = RefactoredCode(inputCode, selection, presenterAction);
            Assert.AreEqual(expectedCode, actualCode);
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

            //Expectation
            const string expectedCode =
                @"Private fizz As Integer

Public Property Get Name() As Integer
    Name = fizz
End Property

Public Property Let Name(ByVal value As Integer)
    fizz = value
End Property


Property Get Foo() As Variant
    Foo = True
End Property

Property Let Foo(ByVal vall As Variant)
End Property

Property Set Foo(ByVal vall As Variant)
End Property";

            var presenterAction = SetParameters("Name", implementLet: true);
            var actualCode = RefactoredCode(inputCode, selection, presenterAction);
            Assert.AreEqual(expectedCode, actualCode);
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
            var presenterAction = SetParameters("Name", implementLet: true);
            var actualCode = RefactoredCode(inputCode, selection, presenterAction);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [TestCase(1, 10, "Public buzz", "Private fizz As Variant", "Public fizz")]
        [TestCase(2, 2, "Public fizz, _\r\nbazz", "Private buzz As Boolean", "")]
        [TestCase(3, 2, "Public fizz, _\r\nbuzz", "Private bazz As Date", "Boolean, bazz As Date")]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulatePublicField_SelectedWithinDeclarationList(int rowSelection, int columnSelection, string contains1, string contains2, string doesNotContain)
        {
            //Input
            string inputCode =
$@"Public fizz, _
buzz As Boolean, _
bazz As Date";

            var selection = new Selection(rowSelection, columnSelection);
            var presenterAction = SetParameters("Name", implementSet: true, implementLet: true);
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
            var presenterAction = SetParameters("Name", implementLet: true);
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

            //Expectation
            string expectedCode =
                $@"Private {newName} As Integer

Public Property Get Name() As Integer
    Name = {newName}
End Property

Public Property Let Name(ByVal value As Integer)
    {newName} = value
End Property


Sub Foo()
    Name = 0
    Bar Name
End Sub

Sub Bar(ByVal name As Integer)
End Sub";

            var presenterAction = SetParametersForSingleTarget("fizz", "Name", true, newName);

            var actualCode = RefactoredCode(inputCode, selection, presenterAction);
            Assert.AreEqual(expectedCode, actualCode);
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

            //Expectation
            const string expectedCode1 =
                @"Private fizz As Integer

Public Property Get Name() As Integer
    Name = fizz
End Property

Public Property Let Name(ByVal value As Integer)
    fizz = value
End Property


Sub Foo()
    Name = 1
End Sub";

            const string expectedCode2 =
                @"Sub Foo()
    Dim c As Class1
    c.Name = 0
    Bar c.Name
End Sub

Sub Bar(ByVal v As Integer)
End Sub";

            var presenterAction = SetParametersForSingleTarget("fizz", "Name", true);

            var actualCode = RefactoredCode(
                "Class1", 
                selection, 
                presenterAction, 
                null, 
                false, 
                ("Class1", codeClass1, ComponentType.ClassModule),
                ("Class2", codeClass2, ComponentType.ClassModule));
            Assert.AreEqual(expectedCode1, actualCode["Class1"]);
            Assert.AreEqual(expectedCode2, actualCode["Class2"]);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulatePublicField_PassInTarget()
        {
            //Input
            const string inputCode =
                @"Private fizz As Integer";

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
            var presenterAction = SetParameters("Name", implementLet: true);
            var actualCode = RefactoredCode(inputCode, "fizz", DeclarationType.Variable, presenterAction);
            Assert.AreEqual(expectedCode, actualCode);
        }

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
            var presenterAction = SetParameters("Name", implementLet: true);
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
            var presenterAction = SetParameters("bar", implementLet: true);
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
            var presenterAction = SetParameters("MyArray");
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
            var presenterAction = SetParameters("MyArray");
            var actualCode = RefactoredCode(inputCode, selection, presenterAction);
            Assert.AreEqual(expectedCode, actualCode);
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

            var presenterAction = SetParametersForSingleTarget("mArray", "MyArray", true, "xArray");

            var actualCode = RefactoredCode(inputCode, selection, presenterAction);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [TestCase("mArray(5) As String", "xArray(5) As String")]
        [TestCase("mArray(5)", "xArray(5) As Variant")]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulateArray_newFieldNameForFieldInList(string arrayDeclaration, string newArrayDeclaration)
        {
            string inputCode =
                $@"Option Explicit

Public {arrayDeclaration}, mNextVar As Long";

            var selection = new Selection(3, 8, 3, 11);

            string expectedCode =
                $@"Option Explicit

Public mNextVar As Long
Private {newArrayDeclaration}

Public Property Get MyArray() As Variant
    MyArray = xArray
End Property
";
            var presenterAction = SetParametersForSingleTarget("mArray", "MyArray", true, "xArray");
            var actualCode = RefactoredCode(inputCode, selection, presenterAction);
            Assert.AreEqual(expectedCode, actualCode);
        }


        #region setup

        private Func<EncapsulateFieldModel, EncapsulateFieldModel> SetParameters(
            string propertyName,
            bool implementSet = false,
            bool implementLet = false)
        {
            return model =>
            {
                model.PropertyName = propertyName;
                model.ImplementLetSetterType = implementLet;
                model.ImplementSetSetterType = implementSet;
                return model;
            };
        }

        private Func<EncapsulateFieldModel, EncapsulateFieldModel> SetParametersForSingleTarget(string field, string property, bool flag = true, string newFieldName =  null)
            => SetParameters(new UserInputDataObject(field, property, flag, newFieldName));


        private Func<EncapsulateFieldModel, EncapsulateFieldModel> SetParameters(UserInputDataObject testInput)
        {
            return model =>
            {
                foreach (var testModifiedAttribute in testInput.EncapsulateFieldAttributes)
                {
                    var attrsInitializedByTheRefactoring = model[testModifiedAttribute.FieldName].EncapsulationAttributes as IUserModifiableFieldEncapsulationAttributes;
                    var testSupport = model as ISupportEncapsulateFieldTests;

                    attrsInitializedByTheRefactoring.FieldName = testModifiedAttribute.FieldName;
                    attrsInitializedByTheRefactoring.NewFieldName = testModifiedAttribute.NewFieldName;
                    attrsInitializedByTheRefactoring.PropertyName = testModifiedAttribute.PropertyName;
                    attrsInitializedByTheRefactoring.EncapsulateFlag = testModifiedAttribute.EncapsulateFlag;

                    //If the Set/Let flags have been set by a test, they override the values applied by the refactoring
                    if (testModifiedAttribute.ImplLet.HasValue)
                    {
                        attrsInitializedByTheRefactoring.ImplementLetSetterType = testModifiedAttribute.ImplementLetSetterType;
                    }

                    if (testModifiedAttribute.ImplSet.HasValue)
                    {
                        attrsInitializedByTheRefactoring.ImplementSetSetterType = testModifiedAttribute.ImplementSetSetterType;
                    }

                    testSupport.SetEncapsulationFieldAttributes(attrsInitializedByTheRefactoring);

                    foreach ((string memberName, bool flag) in testInput.UDTMemberNameFlagPairs)
                    {
                        testSupport.SetMemberEncapsulationFlag(memberName, flag);
                    }
                }
                return model;
            };
        }

        private class UserModifiableFieldEncapsulationAttributes : IUserModifiableFieldEncapsulationAttributes
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
            public bool? ImplLet;
            public bool ImplementLetSetterType
            {
                get => ImplLet.HasValue ? ImplLet.Value : false;
                set => ImplLet = value;
            }

            public bool? ImplSet;
            public bool ImplementSetSetterType
            {
                get => ImplSet.HasValue? ImplSet.Value : false;
                set => ImplSet = value;
            }
    }

    private class UserInputDataObject
        {
            private List<UserModifiableFieldEncapsulationAttributes> _userInput = new List<UserModifiableFieldEncapsulationAttributes>();
            private List<(string, bool)> _udtNameFlagPairs = new List<(string, bool)>();

            public UserInputDataObject(string fieldName, string propertyName, bool encapsulationFlag = true, string newFieldName = null)
                => _userInput.Add(new UserModifiableFieldEncapsulationAttributes(fieldName, propertyName, encapsulationFlag, newFieldName));

            public void AddAttributeSet(string fieldName, string propertyName, bool encapsulationFlag = true, string newFieldName = null)
                => _userInput.Add(new UserModifiableFieldEncapsulationAttributes(fieldName, propertyName, encapsulationFlag, newFieldName));

            public IEnumerable<UserModifiableFieldEncapsulationAttributes> EncapsulateFieldAttributes => _userInput;

            public void AddUDTMemberNameFlagPairs(params (string, bool)[] nameFlagPairs)
                => _udtNameFlagPairs.AddRange(nameFlagPairs);

            public IEnumerable<(string, bool)> UDTMemberNameFlagPairs => _udtNameFlagPairs;
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
