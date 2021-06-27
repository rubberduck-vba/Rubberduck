using System;
using System.Collections.Generic;
using System.Linq;
using Castle.Windsor;
using NUnit.Framework;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.AddInterfaceImplementations;
using Rubberduck.Refactorings.DeleteDeclarations;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.Refactorings.ImplementInterface;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;
using RubberduckTests.Settings;

namespace RubberduckTests.Refactoring.DeleteDeclarations
{
    
    [TestFixture]
    public class DeleteDeclarationsRefactoringActionTests
    {
        private static string threeConsecutiveNewLines = $"{Environment.NewLine}{Environment.NewLine}{Environment.NewLine}";

        private readonly DeleteDeclarationsTestSupport _support = new DeleteDeclarationsTestSupport();

        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void RemoveMultipleDeclarationTypes()
        {
            const string inputCode =
@"
Sub Foo(ByVal arg As Long)

Label1:    Dim var0 As Long: var0 = arg

    Dim var2 As Variant
End Sub

Public Property Get Test() As Long
    Test = 6
End Property

Public Sub DoNothing()
End Sub
";

            var expected =
@"
Sub Foo(ByVal arg As Long)

           Dim var0 As Long: var0 = arg

    Dim var2 As Variant
End Sub

Public Sub DoNothing()
End Sub
";

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "Test", "Label1"));
            StringAssert.Contains(expected, actualCode);
            StringAssert.AreEqualIgnoringCase(expected, actualCode);
        }

        [TestCase("var0", "Label1")]
        [TestCase("Label1", "var0")]
        [Category("Refactorings")]
        [Category("DeleteDeclarationWithLineLabel")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void RemoveMultipleDeclarationTypesLabelAndVariableAnyOrder(params string[] targets)
        {
            const string inputCode =
@"
Sub Foo(ByVal arg As Long)

Label1:    Dim var0 As Long: var0 = arg

    Dim var2 As Variant
End Sub
";

            var expected =
@"
Sub Foo(ByVal arg As Long)

var0 = arg

    Dim var2 As Variant
End Sub
";

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, targets));
            StringAssert.Contains(expected, actualCode);
            StringAssert.AreEqualIgnoringCase(expected, actualCode);
        }

        [TestCase("var0", "Label1", "var1")]
        [TestCase("Label1", "var1", "var0")]
        [Category("Refactorings")]
        [Category("DeleteDeclarationWithLineLabel")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void RemoveMultipleDeclarationTypesLabelAndVariableListAnyOrder(params string[] targets)
        {
            const string inputCode =
@"
Sub Foo(ByVal arg As Long)

Label1:    Dim var0 As Long, var1 As String

    var0 = arg

    Dim var2 As Variant
End Sub
";

            var expected =
@"
Sub Foo(ByVal arg As Long)

    var0 = arg

    Dim var2 As Variant
End Sub
";

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, targets));
            StringAssert.Contains(expected, actualCode);
            StringAssert.AreEqualIgnoringCase(expected, actualCode);
        }

        [TestCase("Test", "localVar")]
        [TestCase("localVar", "Test")]
        [TestCase("Test", "Label1")]
        [TestCase("Label1", "Test")]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void RemoveMemberContainingOtherLocalTargetsAnyOrder(params string[] targets)
        {
            var inputCode =
@"
Private mTestVal As Long

Public Function Test() As Long
    Dim localVar As String
    localVar = ""asdf""
Label1:
    Test = mTestVal
End Function

Public Sub DoNothing()
End Sub
";

            var expected =
@"
Private mTestVal As Long

Public Sub DoNothing()
End Sub
";

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, targets));
            StringAssert.Contains(expected, actualCode);
            StringAssert.AreEqualIgnoringCase(expected, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void InvalidDeclarationType_throws()
        {
            var inputCode =
$@"
Option Explicit

Public Sub TestSub(arg1 As Long)
End Sub
";
            Assert.Throws<InvalidDeclarationTypeException>(() => GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "arg1")));
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void DeclarationsDeleteModel_AddDeclarationsToDelete()
        {
            var inputCode =
$@"
Option Explicit

Private mVar0 As Long

Private mVar1 As Long

Private mVar2 As String

Private mVar31 As Integer

Private mVar32 As Boolean

Public Sub TestSub(arg1 As Long)
End Sub

Public Function Bizz() As Long
End Function
";
            var vbe = MockVbeBuilder.BuildFromModules((MockVbeBuilder.TestModuleName, inputCode, ComponentType.StandardModule)).Object;
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe);
            using (state)
            {
                var dVar0 = state.DeclarationFinder.MatchName("mVar0").First();

                var model = new DeleteDeclarationsModel(dVar0);

                var dVar1 = state.DeclarationFinder.MatchName("mVar1").First();
                var dVar2 = state.DeclarationFinder.MatchName("mVar2").First();

                model.AddDeclarationsToDelete(dVar1, dVar2);

                var multiple = state.DeclarationFinder.UserDeclarations(DeclarationType.Variable).Where(d => d.IdentifierName.StartsWith("mVar3"));
                model.AddRangeOfDeclarationsToDelete(multiple);

                var funcBizz = state.DeclarationFinder.MatchName("Bizz").First();

                model.AddDeclarationsToDelete(funcBizz);

                var deleteDeclarationsResolver = new DeleteDeclarationsTestsResolver(state, rewritingManager);
                var deleteDeclarationRefactoringAction = deleteDeclarationsResolver.Resolve<DeleteDeclarationsRefactoringAction>();

                var session = rewritingManager.CheckOutCodePaneSession();
                deleteDeclarationRefactoringAction.Refactor(model, session);

                session.TryRewrite();

                var results = vbe.ActiveVBProject.VBComponents
                    .ToDictionary(component => component.Name, component => component.CodeModule.Content());

                var actualCode = results[MockVbeBuilder.TestModuleName];

                StringAssert.Contains("Public Sub TestSub(arg1 As Long)", actualCode);
                StringAssert.DoesNotContain("mVar", actualCode);
                StringAssert.DoesNotContain("Bizz", actualCode);
            }
        }

        //Every Enum must declare at least one member...Removing all members is equivalent to removing the entire Enum declaration
        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void RemoveAllEnumMembers()
        {
            var inputCode =
@"
Option Explicit

Public mVar1 As Long

Private Enum TestEnum
    FirstValue
    SecondValue
    ThirdValue
End Enum

Public Sub Test1()
End Sub
";
            var expected =
@"
Option Explicit

Public mVar1 As Long

Public Sub Test1()
End Sub
";

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "FirstValue", "SecondValue", "ThirdValue"));
            StringAssert.Contains(expected, actualCode);
            StringAssert.AreEqualIgnoringCase(expected, actualCode);
        }

        //Every UserDefinedType must declare at least one member...Removing all members is equivalent to removing the entire UDT declaration
        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void RemoveAllUdtMembers()
        {
            var inputCode =
@"
Option Explicit

Public mVar1 As Long

Private Type TestType
    FirstValue As Long
    SecondValue As String
    ThirdValue As Boolean
End Type

Public Sub Test1()
End Sub
";
            var expected =
@"
Option Explicit

Public mVar1 As Long

Public Sub Test1()
End Sub
";

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "FirstValue", "SecondValue", "ThirdValue"));
            StringAssert.Contains(expected, actualCode);
            StringAssert.AreEqualIgnoringCase(expected, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void RemoveAllUdtMembersAndAllEnumMembers()
        {
            var inputCode =
@"
Option Explicit

Public mVar1 As Long

Private Enum TestEnum
    EFirstValue
    ESecondValue
    EThirdValue
End Enum

Private Type TestType
    TFirstValue As Long
    TSecondValue As String
    TThirdValue As Boolean
End Type

Public Sub Test1()
End Sub
";
            var expected =
@"
Option Explicit

Public mVar1 As Long

Public Sub Test1()
End Sub
";

            var actualCode = GetRetainedCodeBlock(inputCode, 
                state => _support.TestTargets(state, "EFirstValue", "ESecondValue", "EThirdValue", "TFirstValue", "TSecondValue", "TThirdValue"));
            StringAssert.Contains(expected, actualCode);
            StringAssert.AreEqualIgnoringCase(expected, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void RemoveMemberAndDeclarationUDT()
        {
            var inputCode =
@"
Option Explicit

Public mVar1 As Long

Private Type TestType
    FirstValue As Long
    SecondValue As String
    ThirdValue As Boolean
End Type

Public Sub Test1()
End Sub
";
            var expected =
@"
Option Explicit

Public mVar1 As Long

Public Sub Test1()
End Sub
";

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "ThirdValue", "TestType"));
            StringAssert.Contains(expected, actualCode);
            StringAssert.AreEqualIgnoringCase(expected, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void RemoveMemberAndDeclarationEnum()
        {
            var inputCode =
@"
Option Explicit

Public mVar1 As Long

Private Enum TestEnum
    FirstValue
    SecondValue
    ThirdValue
End Enum

Public Sub Test1()
End Sub
";
            var expected =
@"
Option Explicit

Public mVar1 As Long

Public Sub Test1()
End Sub
";

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "ThirdValue", "TestEnum"));
            StringAssert.Contains(expected, actualCode);
            StringAssert.AreEqualIgnoringCase(expected, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void RemovesAllSupportedTypes()
        {
            var inputCode =
@"
Option Explicit

Private Type TTestType1
    FirstValue1 As Long
    SecondValue1 As String
End Type

Private Type TTestType2
    FirstValue2 As Long
End Type

Private Enum ETestEnum
    First
    Second
End Enum

Private Enum ETestEnum2
    First2
End Enum

Public Const PI As Double = 3.14

Public BreakEncapsulation1 As Long, _
    BreakEncapsulation2 As String, BreakEncapsulation3 As Boolean

Public Sub TestSub1()
    Dim localVar As String
End Sub

Public Sub SecondTestSub()
    Const localConst As Long = 5 
End Sub

Public Sub ThirdTestSub()
End Sub

Public Function TestFunction1()

ErrorFound:
End Function

Public Function SecondTestFunction()
End Function

Public Property Get TestProperty1() As Variant
End Property
Public Property Let TestProperty1(ByVal RHS As Variant)
End Property
Public Property Set TestProperty1(ByVal RHS As Variant)
End Property

Public Property Get TestProperty2() As Variant
End Property

";

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargetsUsingDeclarationType(state,
                ("FirstValue1", DeclarationType.UserDefinedTypeMember),
                ("FirstValue2", DeclarationType.UserDefinedTypeMember),
                ("First", DeclarationType.EnumerationMember),
                ("First2", DeclarationType.EnumerationMember),
                ("PI", DeclarationType.Constant),
                ("BreakEncapsulation2", DeclarationType.Variable),
                ("localConst", DeclarationType.Constant),
                ("localVar", DeclarationType.Variable),
                ("ThirdTestSub", DeclarationType.Procedure),
                ("SecondTestFunction", DeclarationType.Function),
                ("ErrorFound", DeclarationType.LineLabel),
                ("TestProperty1", DeclarationType.PropertyGet),
                ("TestProperty1", DeclarationType.PropertyLet),
                ("TestProperty1", DeclarationType.PropertySet)));

            StringAssert.DoesNotContain("FirstValue1", actualCode);
            StringAssert.DoesNotContain("FirstValue2", actualCode);
            StringAssert.DoesNotContain("TTestType2", actualCode);
            StringAssert.DoesNotContain("First", actualCode);
            StringAssert.DoesNotContain("First2", actualCode);
            StringAssert.DoesNotContain("PI", actualCode);
            StringAssert.DoesNotContain("BreakEncapsulation2", actualCode);
            StringAssert.DoesNotContain("localConst", actualCode);
            StringAssert.DoesNotContain("localVar", actualCode);
            StringAssert.DoesNotContain("ThirdTestSub", actualCode);
            StringAssert.DoesNotContain("SecondTestFunction", actualCode);
            StringAssert.DoesNotContain("TestProperty1", actualCode);
            StringAssert.DoesNotContain("ErrorFound", actualCode);

            StringAssert.Contains("Private Type TTestType1", actualCode);
            StringAssert.Contains("SecondValue1", actualCode);
            StringAssert.Contains("TestFunction1", actualCode);
            StringAssert.Contains("TestProperty2", actualCode);
        }

        [TestCase("1:BreakEncapsulation3", "2:BreakEncapsulation1")]
        [TestCase("1:FirstValue", "1:ThirdValue", "2:FirstValue", "2:ThirdValue")]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void RespectsModuleBoundariesForDuplicateIdentifiers(params string[] targetIDs)
        {
            var component1 = MockVbeBuilder.TestModuleName;
            var inputCode1 =
@"
Option Explicit

Private Type TTestType
    FirstValue As Long
    SecondValue As String
    ThirdValue As Boolean
End Type

Private Enum ETestEnum
    First
    Second
    Third
    AfterThird
End Enum

Public Const PI As Double = 3.14

Public BreakEncapsulation1 As Long, _
    BreakEncapsulation2 As String, BreakEncapsulation3 As Boolean

Public Const TEN As Double = 10#

Public Sub FirstTestSub()
    Dim localVar As String
End Sub

Public Sub SecondTestSub()
End Sub

Public Sub ThirdTestSub()
End Sub
";

            var component2 = "AnotherTestModule";
            var inputCode2 =
@"
Option Explicit

Private Type TAnotherTestModule
    FirstValue As Long
    SecondValue As String
    ThirdValue As Boolean
    FourthValue As Currency
End Type

Private Enum ETestEnum
    First
    Second
    Third
End Enum

Public Const PI As Double = 3.14

Public BreakEncapsulation1 As Long, _
    BreakEncapsulation2 As String, BreakEncapsulation3 As Boolean

Public Const TEN As Double = 10#

Public Sub FirstTestSub()
    Dim localVar As String
End Sub

Public Sub SecondTestSub()
End Sub

Public Sub ThirdTestSub()
End Sub
";

            var vbe = MockVbeBuilder.BuildFromModules((component1, inputCode1, ComponentType.StandardModule), (component2, inputCode2, ComponentType.StandardModule)).Object;
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe);
            using (state)
            {
                var module1QMN = state.DeclarationFinder.AllModules.Single(m => m.ComponentName == component1);
                var module2QMN = state.DeclarationFinder.AllModules.Single(m => m.ComponentName == component2);

                var module1Declarations = state.DeclarationFinder.Members(module1QMN).Where(m => !(m is ModuleDeclaration));
                var module2Declarations = state.DeclarationFinder.Members(module2QMN).Where(m => !(m is ModuleDeclaration));

                var module1MemberMap = module1Declarations.ToDictionary(m => m.IdentifierName);
                var module2MemberMap = module2Declarations.ToDictionary(m => m.IdentifierName);

                var module1TargetIDs = new List<string>();

                var module2TargetIDs = new List<string>();

                foreach (var id in targetIDs)
                {
                    if (id.StartsWith("1:"))
                    {
                        module1TargetIDs.Add(id.Substring(2));
                        continue;
                    }
                    module2TargetIDs.Add(id.Substring(2));
                }

                var module1Targets = module1TargetIDs.Select(id => module1MemberMap[id]);

                var module2Targets = module2TargetIDs.Select(id => module2MemberMap[id]);

                var session = RefactorAllElementTypes(state, module1Targets.Concat(module2Targets), rewritingManager);

                session.TryRewrite();

                var resultsMap = vbe.ActiveVBProject.VBComponents
                    .ToDictionary(component => component.Name, component => component.CodeModule.Content());

                var module1Results = resultsMap[component1];
                var module2Results = resultsMap[component2];

                foreach (var removed in module1TargetIDs)
                {
                    StringAssert.DoesNotContain(removed, module1Results);
                }

                var module1RetainedIDs = module1MemberMap.Keys.Where(k => !module1TargetIDs.Contains(k));
                foreach (var retained in module1RetainedIDs)
                {
                    StringAssert.Contains(retained, module1Results);
                }

                foreach (var removed in module2TargetIDs)
                {
                    StringAssert.DoesNotContain(removed, module2Results);
                }

                var module2RetainedIDs = module2MemberMap.Keys.Where(k => !module2TargetIDs.Contains(k));
                foreach (var retained in module2RetainedIDs)
                {
                    StringAssert.Contains(retained, module2Results);
                }
            }
        }

        [TestCase("Dim X As Long", true, true)]
        [TestCase("Dim X As Long", true, false)]
        [TestCase("Dim X As Long", false, true)]
        [TestCase("Dim X As Long", false, false)]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void RespectsDeleteDeclarationsOnlyFlag(string declaration, bool injectTODO, bool deleteOnly)
        {
            var inputCode =
$@"
Option Explicit

Public mVar1 As Long

Public Sub DoSomethingElse(arg As Long)
    'There is already a comment
    '@Ignore UseMeaningfulName
    'And then another
    {declaration}

    Dim usedVar As Long
    arg = usedVar
    
End Sub
";

            void modelFlags(IDeleteDeclarationsModel model)
            {
                model.InsertValidationTODOForRetainedComments = injectTODO;
                model.DeleteDeclarationsOnly = deleteOnly;
            }

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "X"), modelFlags);

            var injectedContent = injectTODO
                ? DeleteDeclarationsTestSupport.TodoContent
                : string.Empty;

            if (deleteOnly) 
            {
                injectedContent = string.Empty;
                StringAssert.DoesNotContain(DeleteDeclarationsTestSupport.TodoContent, actualCode);
            }

            StringAssert.Contains($"{injectedContent}There is already a comment", actualCode);
            StringAssert.Contains($"{injectedContent}And then another", actualCode);
        }

        [TestCase("Dim X As Long", true)]
        [TestCase("Dim X As Long", false)]
        [TestCase("Const X As Long = 9", true)]
        [TestCase("Const X As Long = 9", false)]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void RespectsDeleteAnnotationWithInsertTODOFlag(string declaration, bool deleteAnnotation)
        {
            var inputCode =
$@"
Option Explicit

Public mVar1 As Long

Public Sub DoSomethingElse(arg As Long)
    'There is already a comment
    '@Ignore UseMeaningfulName
    'And then another
    {declaration}

    Dim usedVar As Long
    arg = usedVar
    
End Sub
";
            void modelFlags(IDeleteDeclarationsModel model)
            {
                model.InsertValidationTODOForRetainedComments = false;
                model.DeleteAnnotations = deleteAnnotation;
            }

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "X"), modelFlags);

            if (deleteAnnotation)
            {
                StringAssert.DoesNotContain("'@Ignore UseMeaningfulName", actualCode);
            }
            else
            {
                StringAssert.Contains("'@Ignore UseMeaningfulName", actualCode);
            }
        }

        [TestCase("Dim X As Long", true, false)]
        [TestCase("Dim X As Long", true, true)]
        [TestCase("Dim X As Long", false, false)]
        [TestCase("Dim X As Long", false, true)]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void RespectsDeleteDeclarationsOnlyWithDeleteLogicalLineFlag(string declaration, bool deleteLogicalLineComments, bool deleteOnly)
        {
            var inputCode =
$@"
Option Explicit

Public mVar1 As Long

Public Sub DoSomethingElse(arg As Long)
    'There is already a comment
    'And then another
    {declaration} 'This is a declaration logical line comment

    Dim usedVar As Long
    arg = usedVar
    
End Sub
";
            void modelFlags(IDeleteDeclarationsModel model)
            {
                model.InsertValidationTODOForRetainedComments = false;
                model.DeleteDeclarationLogicalLineComments = deleteLogicalLineComments;
                model.DeleteDeclarationsOnly = deleteOnly;
            }

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "X"), modelFlags);

            if (deleteLogicalLineComments)
            {
                if (deleteOnly)
                {
                    StringAssert.Contains("'This is a declaration logical line comment", actualCode);
                }
                else
                {
                    StringAssert.DoesNotContain("'This is a declaration logical line comment", actualCode);
                }
            }
            else
            {
                StringAssert.Contains("'This is a declaration logical line comment", actualCode);
                StringAssert.DoesNotContain("'And then another 'This is a declaration logical line comment", actualCode);
            }
        }

        [TestCase(true, true)]
        [TestCase(true, false)]
        [TestCase(false, true)]
        [TestCase(false, false)]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void RespectsDeleteAnnotationWithInsertTODO_ModuleVariable(bool injectTODO, bool deleteOnly)
        {
            var inputCode =
@"
Option Explicit

'@Ignore ""VariableNotUsed""
'A comment following an Annotation
Public mVar1 As Long

Public Sub Test1()
End Sub
";

            void modelFlags(IDeleteDeclarationsModel model)
            {
                model.InsertValidationTODOForRetainedComments = injectTODO;
                model.DeleteDeclarationsOnly = deleteOnly;
            }

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "mVar1"), modelFlags);
            StringAssert.DoesNotContain("mVar1", actualCode);
            StringAssert.Contains("Test1", actualCode);
            
            var injectedContent = injectTODO
                ? DeleteDeclarationsTestSupport.TodoContent
                : string.Empty;

            if (deleteOnly)
            {
                injectedContent = string.Empty;
                StringAssert.DoesNotContain(DeleteDeclarationsTestSupport.TodoContent, actualCode);
            }

            StringAssert.Contains($"{injectedContent}A comment following an Annotation", actualCode);
        }

        private string GetRetainedCodeBlock(string moduleCode, Func<RubberduckParserState, IEnumerable<Declaration>> targetListBuilder, Action<IDeleteDeclarationsModel> modelFlagsAction = null)
        {
            var refactoredCode = _support.TestRefactoring(
                targetListBuilder,
                RefactorAllElementTypes,
                modelFlagsAction ?? _support.DefaultModelFlagAction,
                (MockVbeBuilder.TestModuleName, moduleCode, ComponentType.StandardModule));

            return refactoredCode[MockVbeBuilder.TestModuleName];
        }

        private static IExecutableRewriteSession RefactorAllElementTypes(RubberduckParserState state, IEnumerable<Declaration> targets, IRewritingManager rewritingManager, Action<IDeleteDeclarationsModel> modelFlagsAction = null)
        {
            var model = new DeleteDeclarationsModel(targets);
            if (modelFlagsAction != null)
            {
                modelFlagsAction(model);
            }

            var session = rewritingManager.CheckOutCodePaneSession();

            var deleteDeclarationRefactoringAction = new DeleteDeclarationsTestsResolver(state, rewritingManager)
                .Resolve<DeleteDeclarationsRefactoringAction>();
            
            deleteDeclarationRefactoringAction.Refactor(model, session);

            return session;
        }
    }
}
