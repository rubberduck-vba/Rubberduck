using System;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.AddInterfaceImplementations;
using Rubberduck.Refactorings.DeleteDeclarations;
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

        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void LocalConstant()
        {
            var inputCode =
@"
Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            var actualCodeLines = GetRetainedLines(inputCode, state => TestModel(state, "const1"));
            Assert.IsFalse(actualCodeLines.Contains("Const const1 As Integer = 9"));
            Assert.AreEqual(2, actualCodeLines.Count());
        }

        [TestCase("Public")]
        [TestCase("Private")]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void ModuleConstant(string visibility)
        {
            var inputCode =
$@"
Option Explicit

{visibility} Const const1 As Integer = 9
";

            var actualCodeLines = GetRetainedLines(inputCode, state => TestModel(state, "const1"));
            Assert.IsFalse(actualCodeLines.Contains($"{visibility} Const const1 As Integer = 9"));
            Assert.AreEqual(1, actualCodeLines.Count());
        }

        [TestCase("Const const1 As Integer = 9, const3", "const2")]
        [TestCase("Const const3 As Integer = 7", "const1", "const2")]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void LocalConstantsDeclarationLists(string expected, params string[] toDelete)
        {
            var inputCode =
@"
Public Sub Foo()
    Const const1 As Integer = 9, const2 As Integer = 8, const3 As Integer = 7
End Sub";

            var actualCode = GetRetainedCodeBlock(inputCode, state => TestModel(state, toDelete));
            StringAssert.Contains(expected, actualCode);
        }

        [TestCase("Const const3 As Integer = 7", "const1", "const2")]
        [TestCase("Const const1 As Integer = 9", "const2", "const3")]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void LocalConstantsDeclarationListsLineExtensions(string expectedCode, params string[] toDelete)
        {
            var inputCode =
@"
Public Sub Foo()
    Const const1 As Integer = 9, const2 As Integer = 8, _
            const3 As Integer = 7
End Sub";

            var actualCode = GetRetainedCodeBlock(inputCode, state => TestModel(state, toDelete));
            StringAssert.Contains(expectedCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void LocalConstantsDeclarationListsLineExtensionsDeleteOneOfTwoInSingleLine()
        {
            var inputCode =
@"
Public Sub Foo()
    Const const1 As Integer = 9, const2 As Integer = 8, _
            const3 As Integer = 7
End Sub";

            var expectedCode =
@"
Public Sub Foo()
    Const const1 As Integer = 9, _
            const3 As Integer = 7
End Sub";

            var actualCode = GetRetainedCodeBlock(inputCode, state => TestModel(state, "const2"));
            StringAssert.Contains(expectedCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void LocalConstantsDeclarationListsDeleteAll()
        {
            var inputCode =
@"
Public Sub Foo()
    Const const1 As Integer = 9, const2 As Integer = 8, const3 As Integer = 7
End Sub";

            var actualCodeLines = GetRetainedLines(inputCode, state => TestModel(state, "const1", "const2", "const3"));
            Assert.IsTrue(2 == actualCodeLines.Count(), $"Unexpected line count: {Environment.NewLine} {string.Join(Environment.NewLine, actualCodeLines)}");
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void ModuleVariablesMultiple()
        {
            var inputCode =
@"
Option Explicit

Public notUsed1 As Long
Public notUsed2 As Long
Public notUsed3 As Long
";

            var actualCodeLines = GetRetainedLines(inputCode, state => TestModel(state, "notUsed1"));
            Assert.IsTrue(4 == actualCodeLines.Count(), $"Unexpected line count: {Environment.NewLine} {string.Join(Environment.NewLine, actualCodeLines)}");
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void ModuleVariablesMultiplNewLinesPriorToNextVariableDeclaration()
        {
            var inputCode =
@"
Option Explicit

Public notUsed1 As Long



'Comment after many blank lines
Public notUsed2 As Long
Public notUsed3 As Long
";

            var actualCodeLines = GetRetainedLines(inputCode, state => TestModel(state, "notUsed1"));
            Assert.IsTrue(8 == actualCodeLines.Count(), $"Unexpected line count: {Environment.NewLine} {string.Join(Environment.NewLine, actualCodeLines)}");
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void DeclarationListLineContinuationWithCommentRemoveAll_RemovesComment()
        {
            var inputCode =
@"
Option Explicit

Public notUsed1 As Long, _
    notUsed2 As Long, _
        notUsed3 As Long _
            'These fields are not used


'This field is used
Public used As String
";

            var actualCode = GetRetainedCodeBlock(inputCode, state => TestModel(state, "notUsed1", "notUsed2", "notUsed3"));
            StringAssert.Contains("This field is used", actualCode);
            StringAssert.Contains("Public used As", actualCode);
            StringAssert.DoesNotContain("notUsed1", actualCode);
            StringAssert.DoesNotContain("notUsed2", actualCode);
            StringAssert.DoesNotContain("notUsed3", actualCode);
            StringAssert.DoesNotContain("These fields", actualCode);
        }

        [TestCase("notUsed1","notUsed2", "notUsed3")]
        [TestCase("notUsed2", "notUsed1", "notUsed3")]
        [TestCase("notUsed3", "notUsed1", "notUsed2")]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void DeclarationListLineContinuationWithCommentRemoveSingle_RetainsComment(string toRemove, params string[] retained)
        {
            var inputCode =
@"
Option Explicit

Public notUsed1 As Long, _
    notUsed2 As Long, _
        notUsed3 As Long _
            'One of these fields are not used


'This field is used
Public used As String
";

            var actualCode = GetRetainedCodeBlock(inputCode, state => TestModel(state, toRemove));
            StringAssert.Contains("This field is used", actualCode);
            StringAssert.Contains("Public used As", actualCode);
            StringAssert.Contains("One of these fields are not used", actualCode);
            foreach (var field in retained)
            {
                StringAssert.Contains(field, actualCode);
            }
            StringAssert.DoesNotContain(toRemove, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void ModuleVariablesListAllRemoved()
        {
            var inputCode =
@"
Option Explicit

Private used As Long
Public notUsed1 As Long, notUsed2 As Long, notUsed3 As Long
";

            var actualCodeLines = GetRetainedLines(inputCode, state => TestModel(state, "notUsed1", "notUsed2", "notUsed3"));
            Assert.IsTrue(3 == actualCodeLines.Count(), $"Unexpected line count: {Environment.NewLine} {string.Join(Environment.NewLine, actualCodeLines)}");
        }

        [TestCase("Dim var1 As String: Dim var2 As String 'Comment on var2")]
        [TestCase("Dim var1 As String, var2 As String 'Comment on var2")]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void MultipleDeclarationsLineWithTrailingComment_RetainsComments(string declarationList)
        {
            var inputCode =
$@"
Option Explicit

Sub Foo()
    {declarationList}
    ' More Comments
End Sub";

            var actualCode = GetRetainedCodeBlock(inputCode, state => TestModel(state, "var1"));
            StringAssert.Contains("var2 As String 'Comment ", actualCode);
            StringAssert.Contains("More Comments", actualCode);
            StringAssert.DoesNotContain("Dim var1", actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void MultiLineWithTrailingComment_RetainsComments()
        {
            var inputCode =
$@"
Option Explicit

Sub Foo()
    Dim var1 As String, _
        var2 As String _ 
            'Comment on var2
    ' More Comments
End Sub";

            var actualCode = GetRetainedCodeBlock(inputCode, state => TestModel(state, "var1"));
            StringAssert.Contains("var2 As String", actualCode);
            StringAssert.Contains("'Comment ", actualCode);
            StringAssert.Contains("More Comments", actualCode);
            StringAssert.DoesNotContain("Dim var1", actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void MultiLineWithTrailingCommentRemoveAll_RemovesComment()
        {
            var inputCode =
$@"
Option Explicit

Sub Foo()
    Dim var1 As String, _
        var2 As String _
            'Comment on var2
    ' More Comments
End Sub";

            var actualCode = GetRetainedCodeBlock(inputCode, state => TestModel(state, "var1", "var2"));
            StringAssert.Contains("More Comments", actualCode);
            StringAssert.DoesNotContain("var2", actualCode);
            StringAssert.DoesNotContain("'Comment ", actualCode);
            StringAssert.DoesNotContain("var1", actualCode);
        }

        [TestCase("Dim bar As Boolean", "1", "bar")]
        [TestCase("Dim bar As Boolean, bazz As String", "1   Dim bazz As String", "bar")]
        [TestCase("Dim bar As Boolean, bazz As String, bizz As String", "1   Dim bar As Boolean", "bizz", "bazz")]
        [Category("Refactorings")]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void VariableDeletionWithPreceedingLineNumber(string expression, string expected, params string[] targets)
        {
            var inputCode =
$@"
Private Sub Foo()
1   {expression}
2   Dim bat As Integer
End Sub
";

            var actualCode = GetRetainedCodeBlock(inputCode, state => TestModel(state, targets));
            StringAssert.Contains(expected, actualCode);
            StringAssert.Contains("2   Dim bat As Integer", actualCode);
            foreach (var deletedIdentifier in targets)
            {
                StringAssert.DoesNotContain(deletedIdentifier, actualCode);
            }
        }

        [TestCase("Dim bar As Boolean", "Label1:", "bar")]
        [TestCase("Dim bar As Boolean, bazz As String", "Label1:   Dim bazz As String", "bar")]
        [TestCase("Dim bar As Boolean, bazz As String, bizz As String", "Label1:   Dim bar As Boolean", "bizz", "bazz")]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void VariableDeletionWithPreceedingLineLabel(string expression, string expected, params string[] targets)
        {
            var inputCode =
$@"
Private Sub Foo()
Label1:   {expression}
   Dim bat As Integer
End Sub
";

            var actualCode = GetRetainedCodeBlock(inputCode, state => TestModel(state, targets));
            StringAssert.Contains(expected, actualCode);
            StringAssert.Contains("   Dim bat As Integer", actualCode);
            foreach (var deletedIdentifier in targets)
            {
                StringAssert.DoesNotContain(deletedIdentifier, actualCode);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void RemoveLabelWithFollowingExpression()
        {
            const string inputCode =
@"
Sub Foo(ByVal arg As Long)
    Dim var0 As Long
Label1:    var0 = arg

    Dim var2 As Variant
End Sub";

            var expected =
@"
Sub Foo(ByVal arg As Long)
    Dim var0 As Long
    var0 = arg

    Dim var2 As Variant
End Sub";

            var actualCode = GetRetainedCodeBlock(inputCode, state => TestModel(state, "Label1"));
            StringAssert.Contains(expected, actualCode);
            StringAssert.DoesNotContain("Label1:", actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void RemoveLabelAndFollowingDeclaration()
        {
            const string inputCode =
@"
Sub Foo(ByVal arg As Long)

Label1:    Dim var0 As Long: var0 = arg

    Dim var2 As Variant
End Sub";

            var expected =
@"
Sub Foo(ByVal arg As Long)

Label1:    var0 = arg

    Dim var2 As Variant
End Sub";

            var actualCode = GetRetainedCodeBlock(inputCode, state => TestModel(state, "var0"));
            var actLines = actualCode.Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
            var expLines = expected.Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
            StringAssert.Contains(expected, actualCode);
        }

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

            var actualCode = GetRetainedCodeBlock(inputCode, state => TestModel(state, "Test", "Label1"));
            var actLines = actualCode.Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
            var expLines = expected.Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
            StringAssert.Contains(expected, actualCode);
            StringAssert.DoesNotContain("Label1:", actualCode);
        }

        [TestCase("var0", "Label1")]
        [TestCase("Label1", "var0")]
        [Category("Refactorings")]
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

            var actualCode = GetRetainedCodeBlock(inputCode, state => TestModel(state, targets));
            var actLines = actualCode.Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
            var expLines = expected.Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
            StringAssert.Contains(expected, actualCode);
            StringAssert.DoesNotContain("Label1:", actualCode);
        }

        [TestCase("Test", "localVar")]
        [TestCase("localVar", "Test")]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void RemoveFunctionAndLocalVariableAnyOrder(params string[] targets)
        {
            var inputCode =
@"
Private mTestVal As Long

Public Function Test() As Long
    Dim localVar As String
    localVar = ""asdf""
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

            var actualCode = GetRetainedCodeBlock(inputCode, state => TestModel(state, targets));
            var actLines = actualCode.Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
            var expLines = expected.Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
            StringAssert.Contains(expected, actualCode);
            StringAssert.DoesNotContain("Label1:", actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void RemoveFieldDeclarations()
        {
            var inputCode =
@"
Option Explicit

Public mVar1 As Long

Public mVar2 As Long

Private mVar3 As String, mVar4 As Long, mVar5 As String

Private Type TestType
    FirstValue As Long
    SecondValue As Long
End Type

Public Sub Test()
End Sub";

            var actualCode = GetRetainedCodeBlock(inputCode, state => TestModel(state, "mVar1", "mVar2", "mVar3", "mVar4", "mVar5"));
            StringAssert.DoesNotContain(threeConsecutiveNewLines, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void RemoveFieldDeclarationsUsesColonStmtDelimiter()
        {
            var inputCode =
@"
Option Explicit

Public mVar1 As Long: Public mVar2 As Long

Private mVar3 As String, mVar4 As Long, mVar5 As String

Private Type TestType
    FirstValue As Long
    SecondValue As Long
End Type

Public Sub Test()
End Sub";

            var actualCode = GetRetainedCodeBlock(inputCode, state => TestModel(state, "mVar1", "mVar3", "mVar4", "mVar5"));
            StringAssert.DoesNotContain(threeConsecutiveNewLines, actualCode);
            StringAssert.DoesNotContain(":", actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void RemoveFieldDeclarations_LineContinuations()
        {
            var inputCode =
@"
Option Explicit

Public mVar1 As Long

Public mVar2 As Long

Private mVar3 As String _
        , mVar4 As Long _
                , mVar5 As String

Private Type TestType
    FirstValue As Long
    SecondValue As Long
End Type

Public Sub Test()
End Sub";

            var actualCode = GetRetainedCodeBlock(inputCode, state => TestModel(state, "mVar1", "mVar2", "mVar3", "mVar5"));
            StringAssert.DoesNotContain(threeConsecutiveNewLines, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void RemoveFieldDeclarations_RemovesAllBlankLines()
        {
            var inputCode =
@"
Option Explicit

Public mVar1 As Long





Public mVar2 As Long
Private mVar3 As String, mVar4 As Long, mVar5 As String





Private Type TestType
    FirstValue As Long
    SecondValue As Long
End Type

Public Sub Test()
End Sub";

            var expected =
@"
Option Explicit

Private Type TestType
    FirstValue As Long
    SecondValue As Long
End Type

Public Sub Test()
End Sub";

            var actualCode = GetRetainedCodeBlock(inputCode, state => TestModel(state, "mVar1", "mVar2", "mVar3", "mVar4", "mVar5"));
            StringAssert.AreEqualIgnoringCase(expected, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void RemoveMemberDeclarations()
        {
            var inputCode =
@"
Option Explicit

Public mVar1 As Long

Private Type TestType
    FirstValue As Long
    SecondValue As Long
End Type

Public Sub Test1()
End Sub

Public Sub Test2()
End Sub


Public Function Test3() As String
End Function

Public Sub Test4()
End Sub
";

            var actualCode = GetRetainedCodeBlock(inputCode, state => TestModel(state, "Test2", "Test3"));
            StringAssert.Contains("Test1", actualCode);
            StringAssert.Contains("Test4", actualCode);
            StringAssert.DoesNotContain(threeConsecutiveNewLines, actualCode);
            StringAssert.DoesNotContain("Test2", actualCode);
            StringAssert.DoesNotContain("Test3", actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void RemoveUDTDeclaration()
        {
            var inputCode =
@"
Option Explicit

Public mVar1 As Long

Private Type TestType
    FirstValue As Long
    SecondValue As Long
End Type

Public Sub Test1()
End Sub

Public Sub Test2()
End Sub
";

            var actualCode = GetRetainedCodeBlock(inputCode, state => TestModel(state, "TestType"));
            StringAssert.Contains("mVar1", actualCode);
            StringAssert.Contains("Test1", actualCode);
            StringAssert.Contains("Test2", actualCode);
            StringAssert.DoesNotContain("TestType", actualCode);
            StringAssert.DoesNotContain("FirstValue", actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void RemoveUDTMemberDeclaration()
        {
            var inputCode =
@"
Option Explicit

Public mVar1 As Long

Private Type TestType
    FirstValue As Long
    SecondValue As Long
End Type

Public Sub Test1()
End Sub

Public Sub Test2()
End Sub
";
            var modifiedDeclaration =
@"
Private Type TestType
    SecondValue As Long
End Type
";

            var actualCode = GetRetainedCodeBlock(inputCode, state => TestModel(state, "FirstValue"));
            StringAssert.Contains("mVar1", actualCode);
            StringAssert.Contains("Test1", actualCode);
            StringAssert.Contains("Test2", actualCode);
            StringAssert.Contains(modifiedDeclaration, actualCode);
            StringAssert.DoesNotContain("FirstValue", actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void RemoveEnumDeclaration()
        {
            var inputCode =
@"
Option Explicit

Public mVar1 As Long

Private Enum TestEnum
    FirstValue
    SecondValue
End Enum

Public Sub Test1()
End Sub

Public Sub Test2()
End Sub
";

            var actualCode = GetRetainedCodeBlock(inputCode, state => TestModel(state, "TestEnum"));
            StringAssert.Contains("mVar1", actualCode);
            StringAssert.Contains("Test1", actualCode);
            StringAssert.Contains("Test2", actualCode);
            StringAssert.DoesNotContain("TestEnum", actualCode);
            StringAssert.DoesNotContain("FirstValue", actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void RemoveEnumMemberDeclaration()
        {
            var inputCode =
@"
Option Explicit

Public mVar1 As Long

Private Enum TestEnum
    FirstValue
    SecondValue
End Enum

Public Sub Test1()
End Sub

Public Sub Test2()
End Sub
";
            var modifiedDeclaration =
@"
Private Enum TestEnum
    SecondValue
End Enum
";

            var actualCode = GetRetainedCodeBlock(inputCode, state => TestModel(state, "FirstValue"));
            StringAssert.Contains("mVar1", actualCode);
            StringAssert.Contains("Test1", actualCode);
            StringAssert.Contains("Test2", actualCode);
            StringAssert.Contains(modifiedDeclaration, actualCode);
            StringAssert.DoesNotContain("FirstValue", actualCode);
        }

        private IDictionary<string, string> RefactoredCode(Func<RubberduckParserState, DeleteDeclarationsModel> modelBuilder, params (string componentName, string content, ComponentType componentType)[] modules)
        {
            var vbe = MockVbeBuilder.BuildFromModules(modules).Object;
            return RefactoredCode(vbe, modelBuilder);
        }

        private IDictionary<string, string> RefactoredCode(IVBE vbe, Func<RubberduckParserState, DeleteDeclarationsModel> modelBuilder)
        {
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe);
            using (state)
            {
                var action = new DeleteDeclarationsRefactoringAction(state, rewritingManager, CreateIndenter());
                var model = modelBuilder(state);

                var session = rewritingManager.CheckOutCodePaneSession();
                action.Refactor(model, session);

                session.TryRewrite();

                return vbe.ActiveVBProject.VBComponents
                    .ToDictionary(component => component.Name, component => component.CodeModule.Content());
            }
        }

        private List<string> GetRetainedLines(string moduleCode, Func<RubberduckParserState, DeleteDeclarationsModel> modelBuilder) 
            => GetRetainedCodeBlock(moduleCode, modelBuilder)
                .Trim()
                .Split(new string[] { Environment.NewLine }, StringSplitOptions.None)
                .ToList();

        private string GetRetainedCodeBlock(string moduleCode, Func<RubberduckParserState, DeleteDeclarationsModel> modelBuilder)
        {
            var refactoredCode = RefactoredCode(
                modelBuilder,
                (MockVbeBuilder.TestModuleName, moduleCode, ComponentType.StandardModule));

            return refactoredCode[MockVbeBuilder.TestModuleName];
        }

        private static DeleteDeclarationsModel TestModel(RubberduckParserState state, params string[] identifiers)
        {
            var finder = state.DeclarationFinder;
            var targets = new List<Declaration>();
            foreach (var tgt in identifiers)
            {
                targets.Add(finder.MatchName(tgt).Single());
            }
            return new DeleteDeclarationsModel(targets);
        }

        private static IIndenter CreateIndenter()
            => new Indenter(null, CreateIndenterSettings);

        private static IndenterSettings CreateIndenterSettings()
        {
            var s = IndenterSettingsTests.GetMockIndenterSettings();
            s.VerticallySpaceProcedures = true;
            s.LinesBetweenProcedures = 1;
            return s;
        }
    }
}
