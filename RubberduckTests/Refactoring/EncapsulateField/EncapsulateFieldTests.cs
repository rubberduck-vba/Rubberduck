using System;
using NUnit.Framework;
using Moq;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.EncapsulateField;
using Rubberduck.VBEditor;
using RubberduckTests.Mocks;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.VBEditor.Utility;
using System.Collections.Generic;
using System.Linq;

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

            var userInput = new UserInputDataObject()
                .AddUserInputSet(var1, $"{var1}Prop", var1Flag)
                .AddUserInputSet(var2, $"{var2}Prop", var2Flag)
                .AddUserInputSet(var3, $"{var3}Prop", var3Flag);

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
                StringAssert.Contains($"{variable} = {Support.RHSIdentifier}", actualCode);
                StringAssert.Contains($"Let {variable}Prop(ByVal {Support.RHSIdentifier} As", actualCode);
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

            var userInput = new UserInputDataObject()
                .AddUserInputSet(var1, $"{var1}Prop", var1Flag)
                .AddUserInputSet(var2, $"{var2}Prop", var2Flag)
                .AddUserInputSet(var3, $"{var3}Prop", var3Flag)
                .AddUserInputSet(var4, $"{var4}Prop", var4Flag);

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
            else
            {
                StringAssert.DoesNotContain($"Public {Environment.NewLine}", actualCode);
            }

            foreach (var key in flags.Keys)
            {
                if (flags[key])
                {
                    StringAssert.Contains($"Private {key} As", actualCode);
                    StringAssert.Contains($"{key}Prop = {key}", actualCode);
                    StringAssert.Contains($"{key} = {Support.RHSIdentifier}", actualCode);
                    StringAssert.Contains($"Let {key}Prop(ByVal {Support.RHSIdentifier} As", actualCode);
                    StringAssert.Contains($"Property Get {key}Prop()", actualCode);
                }
            }
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
            Assert.AreEqual(expectedCode.Trim(), actualCode);
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
            Assert.Greater(actualCode.IndexOf($"fizz = {Support.RHSIdentifier}"), actualCode.IndexOf("fizz As Integer"));
            Assert.Less(actualCode.IndexOf($"fizz = {Support.RHSIdentifier}"), actualCode.IndexOf("Get Foo"));
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

Public Property Let Name(ByVal {Support.RHSIdentifier} As Integer)
    fizz = {Support.RHSIdentifier}
End Property
";
            var presenterAction = Support.SetParametersForSingleTarget("fizz", "Name");
            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            Assert.AreEqual(expectedCode.Trim(), actualCode);
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

            var presenterAction = Support.SetParametersForSingleTarget("fizz", "Name");

            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            StringAssert.Contains("Public Property Get Name() As Integer", actualCode);
            StringAssert.Contains("Public Property Let Name(", actualCode);
            StringAssert.Contains($"(ByVal {Support.RHSIdentifier} As Integer)", actualCode);
            StringAssert.Contains("Name = fizz", actualCode);
            StringAssert.Contains($"fizz = {Support.RHSIdentifier}", actualCode);
            StringAssert.Contains("End Property", actualCode);
        }

        [TestCase("Public")]
        [TestCase("Private")]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulateEnumField_PublicEnum(string fieldAccessibility)
        {
            string inputCode =
                $@"

Public Enum NumberTypes 
 Whole = -1 
 Integral = 0 
 Rational = 1 
End Enum

{fieldAccessibility} numberT|ype As NumberTypes
";

            var presenterAction = Support.UserAcceptsDefaults();
            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            StringAssert.Contains("Private numberType_1 As NumberTypes", actualCode);
            StringAssert.Contains("Public Property Get NumberType() As NumberTypes", actualCode);
            StringAssert.Contains("NumberType = numberType_1", actualCode);
        }

        //5.3.1 The declared type of a function declaration may not be a private enum name.
        [TestCase("Public", null)]
        [TestCase("Private", null)]
        [TestCase("Public", ", foo As Long, bar As String")]
        [TestCase("Private", ", foo As Long, bar As String")]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulateEnumField(string enumTypeAccessibility, string declarationList)
        {
            string inputCode =
                $@"

{enumTypeAccessibility} Enum NumberTypes 
 Whole = -1 
 Integral = 0 
 Rational = 1 
End Enum

Private numberT|ype As NumberTypes{declarationList ?? string.Empty}
";

            var presenterAction = Support.UserAcceptsDefaults();
            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            var expectedPropertyType = enumTypeAccessibility == "Public" ? "NumberTypes" : "Long";
            StringAssert.Contains("Private numberType_1 As NumberTypes", actualCode);
            StringAssert.Contains($"Public Property Get NumberType() As {expectedPropertyType}", actualCode);
            StringAssert.Contains("NumberType = numberType_1", actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulatePrivateFieldInList()
        {
            const string inputCode =
                @"Private fi|zz As Integer, fuzz As Integer, fazz As Integer";

            var presenterAction = Support.SetParametersForSingleTarget("fizz");

            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            StringAssert.Contains("Private fizz_1 As Integer, fuzz As Integer,", actualCode);
            StringAssert.Contains("Public Property Get Fizz() As Integer", actualCode);
            StringAssert.Contains("Public Property Let Fizz(", actualCode);
            StringAssert.Contains($"(ByVal {Support.RHSIdentifier} As Integer)", actualCode);
            StringAssert.Contains("Fizz = fizz_1", actualCode);
            StringAssert.Contains($"fizz_1 = {Support.RHSIdentifier}", actualCode);
            StringAssert.Contains("End Property", actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulatePrivateField_Defaults()
        {
            const string inputCode =
                @"|Private fizz As Integer";

            var presenterAction = Support.UserAcceptsDefaults();

            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            StringAssert.Contains("Public Property Get Fizz() As Integer", actualCode);
            StringAssert.Contains("Public Property Let Fizz(", actualCode);
            StringAssert.Contains($"(ByVal {Support.RHSIdentifier} As Integer)", actualCode);
            StringAssert.Contains("Fizz = fizz_1", actualCode);
            StringAssert.Contains($"fizz_1 = {Support.RHSIdentifier}", actualCode);
            StringAssert.Contains("End Property", actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulatePrivateField_DefaultsAsUDT()
        {
            const string inputCode =
                @"|Private fizz As Integer";

            var presenterAction = Support.UserAcceptsDefaults(convertFieldToUDTMember: true);

            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            StringAssert.Contains("Fizz As Integer", actualCode);
            StringAssert.Contains($"this As {Support.StateUDTDefaultType}", actualCode);
            StringAssert.Contains($"this.Fizz = {Support.RHSIdentifier}", actualCode);
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

            var validator = EncapsulateFieldValidationsProvider.NameOnlyValidator(NameValidators.Default);
            var enapsulationIdentifiers = new EncapsulationIdentifiers("fizz", validator) { Property = "Name" };

            var actualCode = Support.RefactoredCode(inputCode.ToCodeString(), presenterAction);
            StringAssert.AreEqualIgnoringCase(enapsulationIdentifiers.TargetFieldName, "fizz");
            StringAssert.Contains($"Private {enapsulationIdentifiers.TargetFieldName} As Integer", actualCode);
            StringAssert.Contains("Property Get Name", actualCode);
            StringAssert.Contains("Property Let Name", actualCode);
            StringAssert.Contains($"Name = {enapsulationIdentifiers.TargetFieldName}", actualCode);
            StringAssert.Contains($"{enapsulationIdentifiers.TargetFieldName} = {Support.RHSIdentifier}", actualCode);
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
    Dim theClass As Class1
    theClass.fizz = 0
    Bar theClass.fizz
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
            StringAssert.Contains("theClass.Name = 0", actualCode["Class2"]);
            StringAssert.Contains("Bar theClass.Name", actualCode["Class2"]);
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

        [TestCase(false)]
        [TestCase(true)]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void StandardModuleSource_ExternalReferences(bool moduleResolve)
        {
            var sourceModuleName = "SourceModule";
            var referenceExpression = moduleResolve ? $"{sourceModuleName}." : string.Empty;
            var sourceModuleCode =
$@"

Public th|is As Long";

            var procedureModuleReferencingCode =
$@"Option Explicit

Private Const bar As Long = 7

Public Sub Bar()
    {referenceExpression}this = bar
End Sub

Public Sub Foo()
    With {sourceModuleName}
        .this = bar
    End With
End Sub
";

            string classModuleReferencingCode =
$@"Option Explicit

Private Const bar As Long = 7

Public Sub Bar()
    {referenceExpression}this = bar
End Sub

Public Sub Foo()
    With {sourceModuleName}
        .this = bar
    End With
End Sub
";

            var userInput = new UserInputDataObject()
                .UserSelectsField("this", "MyProperty");

            var presenterAction = Support.SetParameters(userInput);

            var sourceCodeString = sourceModuleCode.ToCodeString();
            var actualModuleCode = RefactoredCode(
                sourceModuleName,
                sourceCodeString.CaretPosition.ToOneBased(),
                presenterAction,
                null,
                false,
                ("StdModule", procedureModuleReferencingCode, ComponentType.StandardModule),
                ("ClassModule", classModuleReferencingCode, ComponentType.ClassModule),
                (sourceModuleName, sourceCodeString.Code, ComponentType.StandardModule));

            var referencingModuleCode = actualModuleCode["StdModule"];
            StringAssert.Contains($"{sourceModuleName}.MyProperty = ", referencingModuleCode);
            StringAssert.DoesNotContain($"{sourceModuleName}.{sourceModuleName}.MyProperty = ", referencingModuleCode);
            StringAssert.Contains($"  .MyProperty = bar", referencingModuleCode);

            var referencingClassCode = actualModuleCode["ClassModule"];
            StringAssert.Contains($"{sourceModuleName}.MyProperty = ", referencingClassCode);
            StringAssert.DoesNotContain($"{sourceModuleName}.{sourceModuleName}.MyProperty = ", referencingClassCode);
            StringAssert.Contains($"  .MyProperty = bar", referencingClassCode);
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
