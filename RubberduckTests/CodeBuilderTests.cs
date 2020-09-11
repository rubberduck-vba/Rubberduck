using NUnit.Framework;
using Rubberduck.Common;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings;
using RubberduckTests.Mocks;
using System;
using System.Collections.Generic;
using System.Linq;

namespace RubberduckTests
{
    [TestFixture]
    public class CodeBuilderTests
    {
        private static string _rhsIdentifier = Rubberduck.Resources.Refactorings.Refactorings.CodeBuilder_DefaultPropertyRHSParam;

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

        [TestCase("fizz", DeclarationType.Variable, "Integer", "Public")]
        [TestCase("FirstValue", DeclarationType.UserDefinedTypeMember, "Long", "Public")]
        [TestCase("fazz", DeclarationType.Variable, "Long", "Public")]
        [TestCase("fuzz", DeclarationType.Variable, "ETestType2", "Private")]
        [Category(nameof(CodeBuilder))]
        public void PropertyBlockFromPrototype_PropertyGetAccessibility(string targetIdentifier, DeclarationType declarationType, string typeName, string accessibility)
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

        [TestCase("fizz", DeclarationType.Variable, "Integer", "Bazz = fizz")]
        [TestCase("FirstValue", DeclarationType.UserDefinedTypeMember, "Long", "Bazz = fozz.FirstValue")]
        [TestCase("fazz", DeclarationType.Variable, "Long", "Bazz = fazz")]
        [TestCase("fuzz", DeclarationType.Variable, "TTestType2", "Bazz = fuzz")]
        [Category(nameof(CodeBuilder))]
        public void PropertyBlockFromPrototype_PropertyGetContent(string targetIdentifier, DeclarationType declarationType, string typeName, string content)
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

Private fuzz As TTestType2
";
            var result = ParseAndTest<Declaration>(inputCode,
                targetIdentifier,
                declarationType,
                testParams,
                PropertyGetBlockFromPrototypeTest);

            StringAssert.Contains(content, result);
        }

        [TestCase("fizz", DeclarationType.Variable, "Integer", "Bazz = fizz")]
        [Category(nameof(CodeBuilder))]
        public void PropertyBlockFromPrototype_PropertyGetChangeParamName(string targetIdentifier, DeclarationType declarationType, string typeName, string content)
        {
            var testParams = new PropertyBlockFromPrototypeParams("Bazz", DeclarationType.PropertyGet, paramIdentifier: "testParam");
            string inputCode =
$@"
Private fizz As Integer
";
            var result = ParseAndTest<Declaration>(inputCode,
                targetIdentifier,
                declarationType,
                testParams,
                PropertyGetBlockFromPrototypeTest);

            StringAssert.Contains("Property Get Bazz() As Integer", result);
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
            var procedureIdentifier = "TestProcedure";
            var procType = ProcedureTypeIdentifier(declarationType);

            string inputCode =
$@"
Public {procType.procType} {procedureIdentifier}(arg1 As Long, arg2 As String)
End {procType.endStmt}
";
            var result = ParseAndTest<ModuleBodyElementDeclaration>(inputCode,
                procedureIdentifier,
                declarationType,
                new MemberBlockFromPrototypeTestParams(),
                MemberBlockFromPrototypeTest);

            var expected = declarationType.HasFlag(DeclarationType.Property)
                ? "(arg1 As Long, ByVal arg2 As String)"
                : "(arg1 As Long, arg2 As String)";

            StringAssert.Contains($"{procType.procType} {procedureIdentifier}{expected}", result);
        }

        [TestCase(DeclarationType.PropertyLet)]
        [TestCase(DeclarationType.PropertySet)]
        [TestCase(DeclarationType.Procedure)]
        [Category(nameof(CodeBuilder))]
        public void ImprovedArgumentList_AppliesByVal(DeclarationType declarationType)
        {
            var procedureIdentifier = "TestProperty";
            var procType = ProcedureTypeIdentifier(declarationType);

            string inputCode =
$@"
Public {procType.procType} {procedureIdentifier}(arg1 As Long, arg2 As String)
End {procType.endStmt}
";
            var result = ParseAndTest<ModuleBodyElementDeclaration>(inputCode,
                procedureIdentifier,
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
            var procedureIdentifier = "TestProperty";
            var procType = ProcedureTypeIdentifier(declarationType);

            string inputCode =
$@"
Public {procType.procType} {procedureIdentifier}(arg1 As Long, arg2 As String) As Long
End {procType.endStmt}
";
            var result = ParseAndTest<ModuleBodyElementDeclaration>(inputCode,
                procedureIdentifier,
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
@"Private Type TestUDT
    Field1 As Long
    Field2 As String
End Type";
            var actual = CodeBuilderUDTResult(inputCode, "field1", "field2");
            StringAssert.AreEqualIgnoringCase(expected, actual);
        }

        [TestCase("Duplicate", "Duplicate")]
        [TestCase("Duplicate", "DuplicATE")]
        [Category(nameof(CodeBuilder))]
        public void UDT_DuplicateMemberIdentifiers_Throws(string field1Identifier, string field2Identifier)
        {
            var inputCode = "Public field1 : Public field2";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _).Object;
            var state = MockParser.CreateAndParse(vbe);
            using (state)
            {
                var targets = state.DeclarationFinder.UserDeclarations(DeclarationType.Variable)
                    .OfType<VariableDeclaration>()
                    .Where(d => new string[] { "field1", "field2" }.Contains(d.IdentifierName));

                var inputPairs = new List<(VariableDeclaration, string)>();
                inputPairs.Add((targets.First(), field1Identifier));
                inputPairs.Add((targets.Last(), field2Identifier));
                ICodeBuilder codeBuilder = new CodeBuilder();
                Assert.Throws<ArgumentException>(() => { codeBuilder.BuildUserDefinedTypeDeclaration("TestUDT", inputPairs); });
            }

        }

        [Test]
        [Category(nameof(CodeBuilder))]
        public void UDT_EmptyMembersList_Throws()
        {
            Assert.Throws<ArgumentOutOfRangeException>(() => { new CodeBuilder().BuildUserDefinedTypeDeclaration("TestUDT", Enumerable.Empty<(VariableDeclaration, string)>()); });
        }

        [Test]
        [Category(nameof(CodeBuilder))]
        public void UDT_ImplicitTypeMadeExplicit()
        {
            var inputCode = "Public field1";
            var actual = CodeBuilderUDTResult(inputCode, "field1");
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

            var actual = CodeBuilderUDTResult(inputCode, field);
            StringAssert.Contains($"Field1{dimensions} As {expectedType}", actual);
        }

        private string CodeBuilderUDTResult(string inputCode, params string[] fields)
        {
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _).Object;
            var state = MockParser.CreateAndParse(vbe);
            using (state)
            {
                var targets = state.DeclarationFinder.DeclarationsWithType(DeclarationType.Variable)
                    .Where(d => fields.Contains(d.IdentifierName))
                    .Select(field => (field as VariableDeclaration, field.IdentifierName.CapitalizeFirstLetter()));
                return new CodeBuilder().BuildUserDefinedTypeDeclaration("TestUDT", targets);
            }
        }

        private string ParseAndTest<T>(string inputCode, string targetIdentifier, DeclarationType declarationType, Func<T, string> theTest)
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

        private string ParseAndTest<T>(string inputCode, string targetIdentifier, DeclarationType declarationType, MemberBlockFromPrototypeTestParams testParams, Func<T, MemberBlockFromPrototypeTestParams, string> theTest)
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

        private string ParseAndTest<T>(string inputCode, string targetIdentifier, DeclarationType declarationType, PropertyBlockFromPrototypeParams testParams, Func<T, PropertyBlockFromPrototypeParams, string> theTest)
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

        private static string PropertyGetBlockFromPrototypeTest<T>(T target, PropertyBlockFromPrototypeParams testParams) where T : Declaration
        {
            new CodeBuilder().TryBuildPropertyGetCodeBlock(target, testParams.Identifier, out string result, testParams.Accessibility, testParams.Content);
            return result;
        }

        private static string PropertyLetBlockFromPrototypeTest<T>(T target, PropertyBlockFromPrototypeParams testParams) where T : Declaration
        {
            new CodeBuilder().TryBuildPropertyLetCodeBlock(target, testParams.Identifier, out string result, testParams.Accessibility, testParams.Content, testParams.WriteParam);
            return result;
        }

        private static string PropertySetBlockFromPrototypeTest<T>(T target, PropertyBlockFromPrototypeParams testParams) where T : Declaration
        {
            new CodeBuilder().TryBuildPropertySetCodeBlock(target, testParams.Identifier, out string result, testParams.Accessibility, testParams.Content, testParams.WriteParam);
            return result;
        }

        private static string ImprovedArgumentListTest(ModuleBodyElementDeclaration mbed)
            => new CodeBuilder().ImprovedArgumentList(mbed);

        private static string MemberBlockFromPrototypeTest(ModuleBodyElementDeclaration mbed, MemberBlockFromPrototypeTestParams testParams)
            => new CodeBuilder().BuildMemberBlockFromPrototype(mbed, testParams.Accessibility, testParams.Content, testParams.NewIdentifier);

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
            public PropertyBlockFromPrototypeParams(string identifier, DeclarationType propertyType, string accessibility = null, string content = null, string paramIdentifier = null)
            {
                Identifier = identifier;
                DeclarationType = propertyType;
                Accessibility = accessibility;
                Content = content;
                WriteParam = paramIdentifier;
            }
            public DeclarationType DeclarationType { get; }
            public string Identifier { get; }
            public string Accessibility {get; }
            public string Content { get; }
            public string WriteParam { get; }
        }

        private struct MemberBlockFromPrototypeTestParams
        {
            public MemberBlockFromPrototypeTestParams(ModuleBodyElementDeclaration mbed, string accessibility = null, string content = null, string newIdentifier = null)
            {
                Accessibility = accessibility;
                Content = content;
                NewIdentifier = newIdentifier;
            }

            public string Accessibility { get; }
            public string Content { get; }
            public string NewIdentifier { get; }
        }
    }
}
