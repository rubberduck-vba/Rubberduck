using NUnit.Framework;
using Rubberduck.Common;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings;
using Rubberduck.SmartIndenter;
using RubberduckTests.Mocks;
using RubberduckTests.Settings;
using System;
using System.Collections.Generic;
using System.Linq;

namespace RubberduckTests
{
    [TestFixture]
    public class CodeBuilderTests
    {
        private static string _rhsIdentifier = Rubberduck.Resources.Refactorings.Refactorings.CodeBuilder_DefaultPropertyRHSParam;
        private static string _defaultUDTIdentifier = "TestUDT";
        private static string _defaultProcIdentifier = "TestProcedure";
        private static string _defaultPropertyIdentifier = "TestProperty";

        [TestCase("fizz", DeclarationType.Variable, "Integer")]
        [TestCase("FirstValue", DeclarationType.UserDefinedTypeMember, "Long")]
        [TestCase("fazz", DeclarationType.Variable, "Long")]
        [TestCase("fuzz", DeclarationType.Variable, "ETestType2")]
        [Category(nameof(CodeBuilder))]
        public void PropertyBlockFromPrototype_PropertyGet(string targetIdentifier, DeclarationType declarationType, string typeName)
        {
            var testParams = new PropertyBlockFromPrototypeParams("Bazz", DeclarationType.PropertyGet);
            string inputCode =
$@"

Private Type TTestType
    FirstValue As Long
    SecondValue As Variant
End Type

Private Enum ETestType
    EFirstValue = 0
    ESecondValue
End Enum

Public Enum ETestType2
    EThirdValue = 0
    EFourthValue
End Enum

Private fizz As Integer

Private fazz As ETestType

Private fuzz As ETestType2
";
            var result = ParseAndTest<Declaration>(inputCode,
                targetIdentifier,
                declarationType,
                testParams,
                PropertyGetBlockFromPrototypeTest);

            StringAssert.Contains($"Property Get {testParams.Identifier}() As {typeName}", result);
        }

        [TestCase("fizz", DeclarationType.Variable, "Integer", Accessibility.Public)]
        [TestCase("FirstValue", DeclarationType.UserDefinedTypeMember, "Long", Accessibility.Public)]
        [TestCase("fazz", DeclarationType.Variable, "Long", Accessibility.Public)]
        [TestCase("fuzz", DeclarationType.Variable, "ETestType2", Accessibility.Private)]
        [Category(nameof(CodeBuilder))]
        public void PropertyBlockFromPrototype_PropertyGetAccessibility(string targetIdentifier, DeclarationType declarationType, string typeName, Accessibility accessibility)
        {
            var testParams = new PropertyBlockFromPrototypeParams("Bazz", DeclarationType.PropertyGet, accessibility);
            string inputCode =
$@"

Private Type TTestType
    FirstValue As Long
    SecondValue As Variant
End Type

Private Enum ETestType
    EFirstValue = 0
    ESecondValue
End Enum

Public Enum ETestType2
    EThirdValue = 0
    EFourthValue
End Enum

Private fizz As Integer

Private fazz As ETestType

Private fuzz As ETestType2
";
            var result = ParseAndTest<Declaration>(inputCode,
                targetIdentifier,
                declarationType,
                testParams,
                PropertyGetBlockFromPrototypeTest);

            StringAssert.Contains($"{accessibility} Property Get {testParams.Identifier}() As {typeName}", result);
        }

        [TestCase("fizz", DeclarationType.Variable, "Bazz = fizz")]
        [TestCase("FirstValue", DeclarationType.UserDefinedTypeMember, "Bazz = fozz.FirstValue")]
        [TestCase("fazz", DeclarationType.Variable, "Bazz = fazz")]
        [TestCase("fezz", DeclarationType.Variable, "Bazz = fezz")]
        [TestCase("fuzz", DeclarationType.Variable, "Bazz = fuzz")]
        [Category(nameof(CodeBuilder))]
        public void PropertyBlockFromPrototype_PropertyGetContent(string targetIdentifier, DeclarationType declarationType, string content)
        {
            var testParams = new PropertyBlockFromPrototypeParams("Bazz", DeclarationType.PropertyGet, content: content);
            string inputCode =
$@"

Private Type TTestType
    FirstValue As Long
    SecondValue As Variant
End Type

Private Enum ETestType
    EFirstValue = 0
    ESecondValue
End Enum

Public Enum ETestType2
    EThirdValue = 0
    EFourthValue
End Enum

Private fizz As Integer

Private fozz As TTestType

Private fazz As ETestType

Private fezz As ETestType2

Private fuzz As TTestType2
";
            var result = ParseAndTest<Declaration>(inputCode,
                targetIdentifier,
                declarationType,
                testParams,
                PropertyGetBlockFromPrototypeTest);

            StringAssert.Contains(content, result);
        }

        [Test]
        [Category(nameof(CodeBuilder))]
        public void PropertyBlockFromPrototype_PropertyGetChangeParamName()
        {
            var testParams = new PropertyBlockFromPrototypeParams("Bazz", DeclarationType.PropertyGet, paramIdentifier: "testParam");
            string inputCode =
$@"
Private fizz As Integer
";
            var result = ParseAndTest<Declaration>(inputCode,
                "fizz",
                DeclarationType.Variable,
                testParams,
                PropertyGetBlockFromPrototypeTest);

            StringAssert.Contains("Property Get Bazz() As Integer", result);
        }

        [TestCase("Private Const fizz As Integer = 5", DeclarationType.Constant, "Integer")]
        [TestCase("Private Type TTestType\r\nfizz As String\r\nEnd Type", DeclarationType.UserDefinedTypeMember, "String")]
        [Category(nameof(CodeBuilder))]
        public void PropertyBlockFromVariousPrototypeTypes_PropertyGet(string inputCode, DeclarationType declarationType, string expectedTypeName)
        {
            var testParams = new PropertyBlockFromPrototypeParams("Bazz", DeclarationType.PropertyGet);

            var result = ParseAndTest<Declaration>(inputCode,
                "fizz",
                declarationType,
                testParams,
                PropertyGetBlockFromPrototypeTest);

            StringAssert.Contains($"Property Get Bazz() As {expectedTypeName}", result);
        }

        [TestCase("Property Get", "Property", DeclarationType.PropertyGet, "Variant")]
        [TestCase("Function", "Function", DeclarationType.Function, "Variant")]
        [Category(nameof(CodeBuilder))]
        public void PropertyBlockFromFromFunctionPrototypes(string memberType, string memberEndStatement, DeclarationType declarationType, string typeName)
        {
            var targetIdentifier = "TestValue";
            var testParams = new PropertyBlockFromPrototypeParams("Bazz", DeclarationType.PropertyLet);
            var inputCode =
$@"

Private mTestValue As {typeName}

Public {memberType} TestValue() As {typeName}
    TestValue = mTestValue
End {memberEndStatement}
";

            var result = ParseAndTest<Declaration>(inputCode,
                targetIdentifier,
                declarationType,
                testParams,
                PropertyLetBlockFromPrototypeTest);

            StringAssert.Contains($"Property Let {testParams.Identifier}(ByVal RHS As {typeName})", result);
        }

        [TestCase("Public", "ByRef arg1 As Long", "String")]
        [TestCase("Private", "ByRef arg1 As Long", "String")]
        [TestCase("Public", "ByRef arg1 As Long, ByVal arg2 As Double", "String")]
        [TestCase("Public", "ByRef arg1 As Long", "Variant")]
        [Category(nameof(CodeBuilder))]
        public void PropertyLetFromFromPropertyGetWithParameters(string accessibilityToken,
            string paramList,
            string propertyType)
        {
            var procType = ProcedureTypeIdentifier(DeclarationType.PropertyGet);

            string inputCode =
$@"
{accessibilityToken} {procType.procType} {_defaultPropertyIdentifier}({paramList}) As {propertyType}
End {procType.endStmt}
";

            var prototype = GetPrototypeDeclaration<ModuleBodyElementDeclaration>(
                inputCode, _defaultPropertyIdentifier, DeclarationType.PropertyGet);

            var tryResult = CreateCodeBuilder().TryBuildPropertyLetCodeBlock(
                prototype, prototype.IdentifierName,
                out var result, prototype.Accessibility);

            var expected =
                $"{accessibilityToken} Property Let {_defaultPropertyIdentifier}({paramList}"
                    + $", ByVal RHS As {propertyType})";

            Assert.IsTrue(tryResult,
                "TryBuildPropertyLetCodeBlock(...) returned false");

            StringAssert.Contains(expected, result);
        }

        [TestCase("Public", "ByRef arg1 As Long", "Collection")]
        [TestCase("Private", "ByRef arg1 As Long", "Collection")]
        [TestCase("Public", "ByRef arg1 As Long, ByVal arg2 As Double", "Collection")]
        [TestCase("Public", "ByRef arg1 As Long", "Variant")]
        [Category(nameof(CodeBuilder))]
        public void PropertySetFromFromPropertyGetWithParameters(string accessibilityToken,
            string paramList,
            string propertyType)
        {
            var procType = ProcedureTypeIdentifier(DeclarationType.PropertyGet);

            string inputCode =
$@"
{accessibilityToken} {procType.procType} {_defaultPropertyIdentifier}({paramList}) As {propertyType}
End {procType.endStmt}
";

            var prototype = GetPrototypeDeclaration<ModuleBodyElementDeclaration>(
                inputCode, _defaultPropertyIdentifier, DeclarationType.PropertyGet);

            var tryResult = CreateCodeBuilder().TryBuildPropertySetCodeBlock(
                prototype, prototype.IdentifierName,
                out var result, prototype.Accessibility);

            var expected =
                $"{accessibilityToken} Property Set {_defaultPropertyIdentifier}({paramList}"
                    + $", ByVal RHS As {propertyType})";

            Assert.IsTrue(tryResult,
                "TryBuildPropertySetCodeBlock(...) returned false");

            StringAssert.Contains(expected, result);
        }

        [TestCase("Public", 
            "Public Property Get TestProperty(ByRef arg1 As Long) As String", 
            "ByRef arg1 As Long", "ByVal RHS As String")]
        [TestCase("Private", 
            "Private Property Get TestProperty(ByRef arg1 As Long) As String", 
            "ByRef arg1 As Long", "ByVal RHS As String")]
        [TestCase("Public", 
            "Public Property Get TestProperty(ByRef arg1 As Long) As Variant", 
            "ByRef arg1 As Long", "ByVal RHS As Variant")]
        [TestCase("Public", 
            "Public Property Get TestProperty(ByRef arg1 As Long) As Variant", 
            "ByRef arg1 As Long", "ByVal RHS")]
        [TestCase("Public", 
            "Public Property Get TestProperty(ByRef arg1 As Long, ByRef arg2 As String) As String", 
            "ByRef arg1 As Long", "ByRef arg2 As String", "ByVal RHS As String")]
        [Category(nameof(CodeBuilder))]
        public void PropertyGetFromFromPropertyLetWithParameters(string accessibilityToken, 
            string expected,
            params string[] propertLetParamsList)
        {
            var procType = ProcedureTypeIdentifier(DeclarationType.PropertyLet);

            string inputCode =
$@"
{accessibilityToken} {procType.procType} {_defaultPropertyIdentifier}({string.Join(", ", propertLetParamsList)})
End {procType.endStmt}
";

            var prototype = GetPrototypeDeclaration<ModuleBodyElementDeclaration>(
                inputCode, _defaultPropertyIdentifier, DeclarationType.PropertyLet);

            var tryResult = CreateCodeBuilder().TryBuildPropertyGetCodeBlock(
                prototype, prototype.IdentifierName,
                out var result, prototype.Accessibility);

            Assert.IsTrue(tryResult, "TryBuildPropertyGetCodeBlock(...) returned false");

            StringAssert.Contains(expected, result);
        }

        [TestCase("Public",
            "Public Property Get TestProperty(ByRef arg1 As Long) As Collection", 
            "ByRef arg1 As Long", "ByVal RHS As Collection")]
        [TestCase("Private",
            "Private Property Get TestProperty(ByRef arg1 As Long) As Collection",
            "ByRef arg1 As Long", "ByVal RHS As Collection")]
        [TestCase("Public",
            "Public Property Get TestProperty(ByRef arg1 As Long) As Variant",
            "ByRef arg1 As Long", "ByVal RHS As Variant")]
        [TestCase("Public",
            "Public Property Get TestProperty(ByRef arg1 As Long) As Variant",
            "ByRef arg1 As Long", "ByVal RHS")]
        [TestCase("Public",
            "Public Property Get TestProperty(ByRef arg1 As Long, ByRef arg2 As String) As Collection",
            "ByRef arg1 As Long", "ByRef arg2 As String", "ByVal RHS As Collection")]
        [Category(nameof(CodeBuilder))]
        public void PropertyGetFromFromPropertySetWithParameters(string accessibilityToken, 
            string expected,
            params string[] propertySetParamsList)
        {
            var procType = ProcedureTypeIdentifier(DeclarationType.PropertySet);

            string inputCode =
$@"
{accessibilityToken} {procType.procType} {_defaultPropertyIdentifier}({string.Join(", ", propertySetParamsList)})
End {procType.endStmt}
";

            var prototype = GetPrototypeDeclaration<ModuleBodyElementDeclaration>(
                inputCode, _defaultPropertyIdentifier, DeclarationType.PropertySet);

            var tryResult = CreateCodeBuilder().TryBuildPropertyGetCodeBlock(
                prototype, prototype.IdentifierName,
                out var result, prototype.Accessibility);

            Assert.IsTrue(tryResult, "TryBuildPropertyGetCodeBlock(...) returned false");

            StringAssert.Contains(expected, result);
        }

        //Creating a Set from a Set prototype typicaly needs a new name
        [TestCase(DeclarationType.PropertySet, "NewSetProperty")]
        [TestCase(DeclarationType.PropertyLet, null)]
        [Category(nameof(CodeBuilder))]
        public void PropertySetFromPropertyMutator(DeclarationType prototypeDeclarationType, string propertyName)
        {
            (string procType, string endStmt) = ProcedureTypeIdentifier(prototypeDeclarationType);

            string inputCode =
$@"
Public {procType} {_defaultPropertyIdentifier}(ByVal RHS As Variant)
End Property
";

            var prototype = GetPrototypeDeclaration<ModuleBodyElementDeclaration>(
                inputCode, _defaultPropertyIdentifier, prototypeDeclarationType);

            var tryResult = CreateCodeBuilder().TryBuildPropertySetCodeBlock(
                prototype, propertyName ?? _defaultPropertyIdentifier,
                out var result, prototype.Accessibility);

            Assert.IsTrue(tryResult,
                "TryBuildPropertySetCodeBlock(...) returned false");

            var expected =
                $"Public Property Set {propertyName ?? _defaultPropertyIdentifier}(ByVal RHS As Variant)";

            StringAssert.Contains(expected, result);
        }

        //Creating a Set from a Set prototype typicaly needs a new name
        [TestCase(DeclarationType.PropertySet, "NewSetProperty")]
        [TestCase(DeclarationType.PropertyLet, null)]
        [Category(nameof(CodeBuilder))]
        public void PropertySetFromParameterizedPropertyMutator(DeclarationType prototypeDeclarationType, string propertyName)
        {
            (string procType, string endStmt) = ProcedureTypeIdentifier(prototypeDeclarationType);

            string inputCode =
$@"
Public {procType} {_defaultPropertyIdentifier}(ByVal index1 As Long, ByVal index2 As Long, ByVal RHS As Variant)
End Property
";

            var prototype = GetPrototypeDeclaration<ModuleBodyElementDeclaration>(
                inputCode, _defaultPropertyIdentifier, prototypeDeclarationType);

            var tryResult = CreateCodeBuilder().TryBuildPropertySetCodeBlock(
                prototype, propertyName ?? _defaultPropertyIdentifier,
                out var result, prototype.Accessibility);

            Assert.IsTrue(tryResult,
                "TryBuildPropertySetCodeBlock(...) returned false");

            var expected =
                $"Public Property Set {propertyName ?? _defaultPropertyIdentifier}(ByVal index1 As Long, ByVal index2 As Long, ByVal RHS As Variant)";

            StringAssert.Contains(expected, result);
        }

        //Creating a Let from a Let prototype typicaly needs a new name
        [TestCase(DeclarationType.PropertyLet, "NewLetProperty")]
        [TestCase(DeclarationType.PropertySet, null)]
        [Category(nameof(CodeBuilder))]
        public void PropertyLetFromPropertyMutator(DeclarationType prototypeDeclarationType, string propertyName)
        {
            (string procType, string endStmt) = ProcedureTypeIdentifier(prototypeDeclarationType);

            string inputCode =
$@"
Public {procType} {_defaultPropertyIdentifier}(ByVal RHS As Variant)
End Property
";

            var prototype = GetPrototypeDeclaration<ModuleBodyElementDeclaration>(
                inputCode, _defaultPropertyIdentifier, prototypeDeclarationType);

            var tryResult = CreateCodeBuilder().TryBuildPropertyLetCodeBlock(
                prototype, propertyName ?? _defaultPropertyIdentifier,
                out var result, prototype.Accessibility);

            Assert.IsTrue(tryResult,
                "TryBuildPropertyLetCodeBlock(...) returned false");

            var expected =
                $"Public Property Let {propertyName ?? _defaultPropertyIdentifier}(ByVal RHS As Variant)";

            StringAssert.Contains(expected, result);
        }

        //Creating a Let from a Let prototype typicaly needs a new name
        [TestCase(DeclarationType.PropertyLet, "NewLetProperty")]
        [TestCase(DeclarationType.PropertySet, null)]
        [Category(nameof(CodeBuilder))]
        public void PropertyLetFromParameterizedPropertyMutator(DeclarationType prototypeDeclarationType, string propertyName)
        {
            (string procType, string endStmt) = ProcedureTypeIdentifier(prototypeDeclarationType);

            string inputCode =
$@"
Public {procType} {_defaultPropertyIdentifier}(ByVal index1 As Long, ByVal index2 As Long, ByVal RHS As Variant)
End Property
";

            var prototype = GetPrototypeDeclaration<ModuleBodyElementDeclaration>(
                inputCode, _defaultPropertyIdentifier, prototypeDeclarationType);

            var tryResult = CreateCodeBuilder().TryBuildPropertyLetCodeBlock(
                prototype, propertyName ?? _defaultPropertyIdentifier,
                out var result, prototype.Accessibility);

            Assert.IsTrue(tryResult,
                "TryBuildPropertyLetCodeBlock(...) returned false");

            var expected =
                $"Public Property Let {propertyName ?? _defaultPropertyIdentifier}(ByVal index1 As Long, ByVal index2 As Long, ByVal RHS As Variant)";

            StringAssert.Contains(expected, result);
        }

        [Test]
        [Category(nameof(CodeBuilder))]
        public void PropertyLetFromObjectPropertySet_ReturnsFalse()
        {
            (string procType, string endStmt) = ProcedureTypeIdentifier(DeclarationType.PropertySet);

            string inputCode =
$@"
Public {procType} {_defaultPropertyIdentifier}(ByVal RHS As Collection)
End Property
";

            var prototype = GetPrototypeDeclaration<ModuleBodyElementDeclaration>(
                inputCode, _defaultPropertyIdentifier, DeclarationType.PropertySet);

            var tryResult = CreateCodeBuilder().TryBuildPropertyLetCodeBlock(
                prototype, _defaultPropertyIdentifier,
                out var result, prototype.Accessibility);

            Assert.IsFalse(tryResult);
        }

        [Test]
        [Category(nameof(CodeBuilder))]
        public void PropertySetFromSimpleValueTypePropertyLet_ReturnsFalse()
        {
            (string procType, string endStmt) = ProcedureTypeIdentifier(DeclarationType.PropertyLet);

            string inputCode =
$@"
Public {procType} {_defaultPropertyIdentifier}(ByVal RHS As Long)
End Property
";

            var prototype = GetPrototypeDeclaration<ModuleBodyElementDeclaration>(
                inputCode, _defaultPropertyIdentifier, DeclarationType.PropertyLet);

            var tryResult = CreateCodeBuilder().TryBuildPropertySetCodeBlock(
                prototype, _defaultPropertyIdentifier,
                out var result, prototype.Accessibility);

            Assert.IsFalse(tryResult);
        }

        [TestCase("fizz", DeclarationType.Variable, "Integer")]
        [TestCase("FirstValue", DeclarationType.UserDefinedTypeMember, "Long")]
        [TestCase("fazz", DeclarationType.Variable, "Long")]
        [TestCase("fuzz", DeclarationType.Variable, "ETestType2")]
        [Category(nameof(CodeBuilder))]
        public void PropertyBlockFromPrototype_PropertyLet(string targetIdentifier, DeclarationType declarationType, string typeName)
        {
            var testParams = new PropertyBlockFromPrototypeParams("Bazz", DeclarationType.PropertyLet);
            string inputCode =
$@"

Private Type TTestType
    FirstValue As Long
    SecondValue As Variant
End Type

Private Enum ETestType
    EFirstValue = 0
    ESecondValue
End Enum

Public Enum ETestType2
    EThirdValue = 0
    EFourthValue
End Enum

Private fizz As Integer

Private fazz As ETestType

Private fuzz As ETestType2
";
            var result = ParseAndTest<Declaration>(inputCode,
                targetIdentifier,
                declarationType,
                testParams,
                PropertyLetBlockFromPrototypeTest);
            StringAssert.Contains($"Property Let {testParams.Identifier}(ByVal RHS As {typeName})", result);
        }

        [TestCase("fizz", DeclarationType.Variable, "Variant")]
        [TestCase("SecondValue", DeclarationType.UserDefinedTypeMember, "Variant")]
        [Category(nameof(CodeBuilder))]
        public void PropertyBlockFromPrototype_PropertySet(string targetIdentifier, DeclarationType declarationType, string typeName)
        {
            var testParams = new PropertyBlockFromPrototypeParams("Bazz", DeclarationType.PropertySet);
            string inputCode =
$@"

Private Type TTestType
    FirstValue As Long
    SecondValue As Variant
End Type

Private fizz As Variant

";
            var result = ParseAndTest<Declaration>(inputCode,
                targetIdentifier,
                declarationType,
                testParams,
                PropertySetBlockFromPrototypeTest);

            StringAssert.Contains($"Property Set {testParams.Identifier}(ByVal {_rhsIdentifier} As {typeName})", result);
        }

        [TestCase(DeclarationType.PropertyLet)]
        [TestCase(DeclarationType.PropertySet)]
        [TestCase(DeclarationType.Procedure)]
        [Category(nameof(CodeBuilder))]
        public void MemberBlockFromPrototype_AppliesByVal(DeclarationType declarationType)
        {
            var procType = ProcedureTypeIdentifier(declarationType);

            string inputCode =
$@"
Public {procType.procType} {_defaultProcIdentifier}(arg1 As Long, arg2 As String)
End {procType.endStmt}
";
            var result = ParseAndTest<ModuleBodyElementDeclaration>(inputCode,
                _defaultProcIdentifier,
                declarationType,
                MemberBlockFromPrototypeTest);

            var expected = declarationType.HasFlag(DeclarationType.Property)
                ? "(arg1 As Long, ByVal arg2 As String)"
                : "(arg1 As Long, arg2 As String)";

            StringAssert.Contains($"Public {procType.procType} {_defaultProcIdentifier}{expected}", result);
        }

        [TestCase(DeclarationType.PropertyLet)]
        [TestCase(DeclarationType.PropertySet)]
        [TestCase(DeclarationType.Procedure)]
        [Category(nameof(CodeBuilder))]
        public void ImprovedArgumentList_AppliesByVal(DeclarationType declarationType)
        {
            var procType = ProcedureTypeIdentifier(declarationType);

            string inputCode =
$@"
Public {procType.procType} {_defaultProcIdentifier}(arg1 As Long, arg2 As String)
End {procType.endStmt}
";
            var result = ParseAndTest<ModuleBodyElementDeclaration>(inputCode,
                _defaultProcIdentifier,
                declarationType,
                ImprovedArgumentListTest);

            var expected = declarationType.HasFlag(DeclarationType.Property)
                ? "arg1 As Long, ByVal arg2 As String"
                : "arg1 As Long, arg2 As String";

            StringAssert.AreEqualIgnoringCase(expected, result);
        }

        [TestCase(DeclarationType.PropertyGet)]
        [TestCase(DeclarationType.Function)]
        [Category(nameof(CodeBuilder))]
        public void ImprovedArgumentList_FunctionTypes(DeclarationType declarationType)
        {
            var procType = ProcedureTypeIdentifier(declarationType);

            string inputCode =
$@"
Public {procType.procType} {_defaultProcIdentifier}(arg1 As Long, arg2 As String) As Long
End {procType.endStmt}
";
            var result = ParseAndTest<ModuleBodyElementDeclaration>(inputCode,
                _defaultProcIdentifier,
                declarationType,
                ImprovedArgumentListTest);

            StringAssert.AreEqualIgnoringCase($"arg1 As Long, arg2 As String", result);
        }

        [Test]
        [Category(nameof(CodeBuilder))]
        public void UDT_CreateFromFields()
        {
            var inputCode =
@"
    Public field1 As Long
    Public field2 As String";

            var expected =
$@"Private Type {_defaultUDTIdentifier}
    Field1 As Long
    Field2 As String
End Type";
            var actual = CodeBuilderUDTResult(inputCode, DeclarationType.Variable, "field1", "field2");
            StringAssert.AreEqualIgnoringCase(expected, actual);
        }

        [Test]
        [Category(nameof(CodeBuilder))]
        public void UDT_ImplicitTypeMadeExplicit()
        {
            var inputCode = "Public field1";
            var actual = CodeBuilderUDTResult(inputCode, DeclarationType.Variable, "field1");
            StringAssert.Contains("Field1 As Variant", actual);
        }

        [TestCase("()", "Long")]
        [TestCase("(50)", "Long")]
        [TestCase("(1 To 10)", "Long")]
        [TestCase("()", "")]
        [TestCase("(50)", "")]
        [TestCase("(1 To 10)", "")]
        [Category(nameof(CodeBuilder))]
        public void UDT_FromArrayField(string dimensions, string type)
        {
            var field = "field1";

            var inputCode = string.IsNullOrEmpty(type)
                ? $"Public {field}{dimensions}"
                : $"Public {field}{dimensions} As {type}";

            var expectedType = string.IsNullOrEmpty(type)
                ? "Variant"
                : type;

            var actual = CodeBuilderUDTResult(inputCode, DeclarationType.Variable, field);
            StringAssert.Contains($"Field1{dimensions} As {expectedType}", actual);
        }

        [Test]
        [Category(nameof(CodeBuilder))]
        public void UDT_CreateFromConstants()
        {
            var inputCode =
@"
    Public Const field1 As Long = 5
    Public Const field2 As String = ""Yo""
";

            var expected =
$@"Private Type {_defaultUDTIdentifier}
    Field1 As Long
    Field2 As String
End Type";
            var actual = CodeBuilderUDTResult(inputCode, DeclarationType.Constant, "field1", "field2");
            StringAssert.AreEqualIgnoringCase(expected, actual);
        }

        [TestCase(DeclarationType.PropertyGet)]
        [TestCase(DeclarationType.Function)]
        [Category(nameof(CodeBuilder))]
        public void UDT_CreateFromFunctionPrototypes(DeclarationType declarationType)
        {
            var procStrings = ProcedureTypeIdentifier(declarationType);
            var inputCode =
$@"

Private mTestValue As Long
Private mTestValue2 As Variant

Public {procStrings.procType} TestValue() As Long
    TestValue = mTestValue
End {procStrings.endStmt}


Public {procStrings.procType} TestValue2() As Variant
    TestValue2 = mTestValue2
End {procStrings.endStmt}
";

            var expected =
$@"Private Type {_defaultUDTIdentifier}
    TestValue As Long
    TestValue2 As Variant
End Type";

            var actual = CodeBuilderUDTResult(inputCode, declarationType, "TestValue", "TestValue2");
            StringAssert.AreEqualIgnoringCase(expected, actual);
        }

        [TestCase(DeclarationType.PropertyLet)]
        [TestCase(DeclarationType.PropertySet)]
        [Category(nameof(CodeBuilder))]
        public void UDT_CreateFromPropertyLetPrototypes(DeclarationType declarationType)
        {
            var procStrings = ProcedureTypeIdentifier(declarationType);
            var inputCode =
$@"

Private mTestValue As Long
Private mTestValue2 As Variant

Public {procStrings.procType} TestValue(ByVal RHS As Long)
    mTestValue = RHS
End {procStrings.endStmt}


Public {procStrings.procType} TestValue2(ByVal RHS As Variant)
    mTestValue2 = RHS
End {procStrings.endStmt}
";

            var expected =
$@"Private Type {_defaultUDTIdentifier}
    TestValue As Long
    TestValue2 As Variant
End Type";

            var actual = CodeBuilderUDTResult(inputCode, declarationType, "TestValue", "TestValue2");
            StringAssert.AreEqualIgnoringCase(expected, actual);
        }

        [Test]
        [Category(nameof(CodeBuilder))]
        public void UDT_CreateFromUDTMemberPrototypes()
        {
            var inputCode =
$@"
Private Type ExistingType
    FirstValue As Long
    SecondValue As Byte
    ThirdValue As String
End Type
";

            var expected =
$@"Private Type {_defaultUDTIdentifier}
    FirstValue As Long
    ThirdValue As String
End Type";

            var actual = CodeBuilderUDTResult(inputCode, DeclarationType.UserDefinedTypeMember, "FirstValue", "ThirdValue");
            StringAssert.AreEqualIgnoringCase(expected, actual);
        }

        [TestCase(DeclarationType.Procedure)]
        [Category(nameof(CodeBuilder))]
        public void UDT_InvalidPrototypes_NoResult(DeclarationType declarationType)
        {
            var procStrings = ProcedureTypeIdentifier(declarationType);

            var inputCode =
$@"
Public {procStrings.procType} TestValue(arg As Long)
End {procStrings.endStmt}
";
            var actual = CodeBuilderUDTResult(inputCode, declarationType, "TestValue", "TestValue2");
            Assert.IsTrue(string.IsNullOrEmpty(actual));
        }

        [Test]
        [Category(nameof(CodeBuilder))]
        public void UDT_NullUDTIdentifierBuildUDT_NoResult()
        {
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule("Private test As Long", out _).Object;
            var state = MockParser.CreateAndParse(vbe);
            using (state)
            {
                var targets = state.DeclarationFinder.DeclarationsWithType(DeclarationType.Variable)
                    .Where(d => d.IdentifierName == "test")
                    .Select(d => (d, d.IdentifierName));

                var result = CreateCodeBuilder().TryBuildUserDefinedTypeDeclaration(null, targets, out var declaration);

                Assert.IsFalse(result);
                Assert.IsTrue(string.IsNullOrEmpty(declaration));
            }
        }

        [Test]
        [Category(nameof(CodeBuilder))]
        public void UDT_EmptyPrototypeList_NoResult()
        {
            var result = CreateCodeBuilder().TryBuildUserDefinedTypeDeclaration(_defaultUDTIdentifier, Enumerable.Empty<(Declaration, string)>(), out var declaration);
            Assert.IsFalse(result);
            Assert.IsTrue(string.IsNullOrEmpty(declaration));
        }

        [Test]
        [Category(nameof(CodeBuilder))]
        public void UDT_NullDeclarationInPrototypeList_NoResult()
        {
            var nullInList = new List<(Declaration, string)>() { (null, "Fizz") };
            var result = CreateCodeBuilder().TryBuildUserDefinedTypeDeclaration(_defaultUDTIdentifier, nullInList, out var declaration);
            Assert.IsFalse(result);
            Assert.IsTrue(string.IsNullOrEmpty(declaration));
        }

        [Test]
        [Category(nameof(CodeBuilder))]
        public void UDT_NullIdentifierInPrototypeList_NoResult()
        {
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule("Private test As Long", out _).Object;
            var state = MockParser.CreateAndParse(vbe);
            using (state)
            {
                string nullIdentifier = null;
                var targets = state.DeclarationFinder.DeclarationsWithType(DeclarationType.Variable)
                    .Where(d => d.IdentifierName == "test")
                    .Select(d => (d, nullIdentifier));

                var result = CreateCodeBuilder().TryBuildUserDefinedTypeDeclaration("TestType", targets, out var declaration);

                Assert.IsFalse(result);
                Assert.IsTrue(string.IsNullOrEmpty(declaration));
            }
        }

        [Test]
        [Category(nameof(CodeBuilder))]
        public void UDT_NullPrototype_NoResult()
        {
            var result = CreateCodeBuilder().TryBuildUDTMemberDeclaration(null, _defaultUDTIdentifier, out var declaration);
            Assert.IsFalse(result);
            Assert.IsTrue(string.IsNullOrEmpty(declaration));
        }

        [Test]
        [Category(nameof(CodeBuilder))]
        public void UDT_NullUDTIdentifierBuildUDTMember_NoResult()
        {
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule("Private test As Long", out _).Object;
            var state = MockParser.CreateAndParse(vbe);
            using (state)
            {
                var target = state.DeclarationFinder.DeclarationsWithType(DeclarationType.Variable)
                    .Single(d => d.IdentifierName == "test");

                var result = CreateCodeBuilder().TryBuildUDTMemberDeclaration(target, null, out var declaration);

                Assert.IsFalse(result);
                Assert.IsTrue(string.IsNullOrEmpty(declaration));
            }
        }

        private string CodeBuilderUDTResult(string inputCode, DeclarationType declarationType, params string[] prototypes)
        {
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _).Object;
            var state = MockParser.CreateAndParse(vbe);
            using (state)
            {
                var targets = state.DeclarationFinder.DeclarationsWithType(declarationType)
                    .Where(d => prototypes.Contains(d.IdentifierName))
                    .Select(prototype => (prototype, prototype.IdentifierName.CapitalizeFirstLetter()));

                return CreateCodeBuilder().TryBuildUserDefinedTypeDeclaration(_defaultUDTIdentifier, targets, out string declaration)
                    ? declaration
                    : string.Empty;
            }
        }

        private string ParseAndTest<T>(string inputCode, string targetIdentifier, 
            DeclarationType declarationType, Func<T, string> theTest) where T: Declaration

        {
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _).Object;
            var state = MockParser.CreateAndParse(vbe);
            using (state)
            {
                var target = state.DeclarationFinder.DeclarationsWithType(declarationType)
                    .Where(d => d.IdentifierName == targetIdentifier).OfType<T>()
                    .Single();
                return theTest(target);
            }
        }

        private string ParseAndTest<T>(string inputCode, string targetIdentifier, 
            DeclarationType declarationType, PropertyBlockFromPrototypeParams testParams, 
            Func<T, PropertyBlockFromPrototypeParams, string> theTest) where T: Declaration

        {
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _).Object;
            var state = MockParser.CreateAndParse(vbe);
            using (state)
            {
                var target = state.DeclarationFinder.DeclarationsWithType(declarationType)
                    .Where(d => d.IdentifierName == targetIdentifier).OfType<T>()
                    .Single();
                return theTest(target, testParams);
            }
        }

        private Declaration GetPrototypeDeclaration<T>(
            string inputCode, 
            string targetIdentifier, 
            DeclarationType declarationType) where T:Declaration
        {
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _).Object;
            var state = MockParser.CreateAndParse(vbe);
            using (state)
            {
                var target = state.DeclarationFinder.DeclarationsWithType(declarationType)
                    .Where(d => d.IdentifierName == targetIdentifier).OfType<T>()
                    .Single();
                return target;
            }
        }

        private static string PropertyGetBlockFromPrototypeTest<T>(T target, PropertyBlockFromPrototypeParams testParams) where T : Declaration
        {
            CreateCodeBuilder().TryBuildPropertyGetCodeBlock(target, testParams.Identifier, out string result, testParams.Accessibility, testParams.Content);
            return result;
        }

        private static string PropertyLetBlockFromPrototypeTest<T>(T target, PropertyBlockFromPrototypeParams testParams) where T : Declaration
        {
            CreateCodeBuilder().TryBuildPropertyLetCodeBlock(target, testParams.Identifier, out string result, testParams.Accessibility, testParams.Content, testParams.WriteParam);
            return result;
        }

        private static string PropertySetBlockFromPrototypeTest<T>(T target, PropertyBlockFromPrototypeParams testParams) where T : Declaration
        {
            CreateCodeBuilder().TryBuildPropertySetCodeBlock(target, testParams.Identifier, out string result, testParams.Accessibility, testParams.Content, testParams.WriteParam);
            return result;
        }

        private static string ImprovedArgumentListTest(ModuleBodyElementDeclaration mbed)
            => CreateCodeBuilder().ImprovedArgumentList(mbed);

        private static string MemberBlockFromPrototypeTest(ModuleBodyElementDeclaration mbed)
            => CreateCodeBuilder().BuildMemberBlockFromPrototype(mbed, string.Empty, Accessibility.Public, _defaultProcIdentifier);

        private static ICodeBuilder CreateCodeBuilder()
            => new CodeBuilder(new Indenter(null, CreateIndenterSettings));

        private static IndenterSettings CreateIndenterSettings()
        {
            var s = IndenterSettingsTests.GetMockIndenterSettings();
            s.VerticallySpaceProcedures = true;
            s.LinesBetweenProcedures = 1;
            return s;
        }

        private (string procType, string endStmt) ProcedureTypeIdentifier(DeclarationType declarationType)
        {
            switch (declarationType)
            {
                case DeclarationType.Function:
                    return ("Function", "Function");
                case DeclarationType.Procedure:
                    return ("Sub", "Sub");
                case DeclarationType.PropertyGet:
                    return ("Property Get", "Property");
                case DeclarationType.PropertyLet:
                    return ("Property Let", "Property");
                case DeclarationType.PropertySet:
                    return ("Property Set", "Property");
                default:
                    throw new ArgumentOutOfRangeException();
            }
        }

        private struct PropertyBlockFromPrototypeParams
        {
            public PropertyBlockFromPrototypeParams(string identifier, DeclarationType propertyType, Accessibility accessibility = Accessibility.Public, string content = null, string paramIdentifier = null)
            {
                Identifier = identifier;
                DeclarationType = propertyType;
                Accessibility = accessibility;
                Content = content;
                WriteParam = paramIdentifier;
            }
            public DeclarationType DeclarationType { get; }
            public string Identifier { get; }
            public Accessibility Accessibility {get; }
            public string Content { get; }
            public string WriteParam { get; }
        }
    }
}
