using Antlr4.Runtime;
using NUnit.Framework;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
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
using System;
using System.Collections.Generic;
using System.Linq;

namespace RubberduckTests.Refactoring.DeleteDeclarations
{
    [TestFixture]
    public class DeleteDeclarationsLabelsTests
    {
        private readonly DeleteDeclarationsTestSupport _support = new DeleteDeclarationsTestSupport();

        [TestCase("var0 = arg")]
        [TestCase("Dim var1 As Long: var1 = arg")]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void RemoveLabelWithFollowingExpression(string expression)
        {
            var inputCode =
$@"
Sub Foo(ByVal arg As Long)
    Dim var0 As Long
Label1:    {expression}

End Sub";

            var expected =
$@"
Sub Foo(ByVal arg As Long)
    Dim var0 As Long
           {expression}

End Sub";

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "Label1"));
            StringAssert.Contains(expected, actualCode);
            StringAssert.DoesNotContain("Label1:", actualCode);
        }

        [TestCase("Label1:    'Comment on Label1 line\r\n\r\n", "")]
        [TestCase("Label1:    Dim var0 As Long    'Comment on Label1 line\r\n\r\n", "           Dim var0 As Long    'Comment on Label1 line\r\n\r\n")]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void LabelWithSameLineContent(string testExpression, string expectedRewrite)
        {
            var inputCode =
$@"
Sub Foo(ByVal arg As Long)

{testExpression}End Sub";

            var expected =
$@"
Sub Foo(ByVal arg As Long)

{expectedRewrite}End Sub";

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "Label1"));
            StringAssert.AreEqualIgnoringCase(expected, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void LabelWithPrecedingAnnotation()
        {
            const string inputCode =
@"
Sub Foo(ByVal arg As Long)

    Dim var2 As Variant

'@Ignore UseMeaningfulName
Label1:    Dim X As Long: X = arg

End Sub
";

            var expected =
@"
Sub Foo(ByVal arg As Long)

    Dim var2 As Variant

'@Ignore UseMeaningfulName
           Dim X As Long: X = arg

End Sub
";

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "Label1"));
            StringAssert.Contains(expected, actualCode);
            StringAssert.AreEqualIgnoringCase(expected, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void LabelWithinForNext()
        {
            const string inputCode =
@"
Sub Foo(ByVal arg As Long, ByRef sum As Long)
    sum = 0
On Error GoTo Label1

    Dim idx As Long
    For idx = 0 To arg
Label1:
        sum = sum + idx
    Next idx

End Sub
";

            var expected =
@"
    Dim idx As Long
    For idx = 0 To arg
        sum = sum + idx
    Next idx
";

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "Label1"));
            StringAssert.Contains(expected, actualCode);
        }

        [TestCase("Do While idx < arg", "Loop")]
        [TestCase("While idx < arg", "Wend")]
        [TestCase("Do Until idx = arg", "Loop")]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void LabelWithinDoLoop(string loopStart, string loopTerminator)
        {
            var inputCode =
$@"
Sub Foo(ByVal arg As Long, ByRef sum As Long)
    sum = 0
On Error GoTo Label1

    Dim idx As Long
    idx = 0
    {loopStart}
Label1:
        sum = sum + idx
        idx = idx + 1
    {loopTerminator}

End Sub
";

            var expected =
$@"
    Dim idx As Long
    idx = 0
    {loopStart}
        sum = sum + idx
        idx = idx + 1
    {loopTerminator}
";

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "Label1"));
            StringAssert.Contains(expected, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void LabelWithinForEach()
        {
            const string inputCode =
@"
Sub Foo(ByVal indexes As Collection)

On Error GoTo Label1

    Dim element As Variant
    For Each element in indexes
Label1:
        element = element + 1
    Next

End Sub
";

            var expected =
@"
    Dim element As Variant
    For Each element in indexes
        element = element + 1
    Next
";

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "Label1"));
            StringAssert.Contains(expected, actualCode);
        }

        [TestCase(1)]
        [TestCase(2)]
        [TestCase(3)]
        [TestCase(4)]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void LabelInsideWithStatement(long testCase)
        {
            var withStatementVersions = new Dictionary<long, string>()
            {
                [1] =
@"
    With tType
TestLabel1:
        .FirstVal = arg
        .SecondVal = CStr(arg)
        .ThirdVal = CDbl(arg)
    End With
",
                [2] =
@"
    With tType
        .FirstVal = arg
TestLabel1:
        .SecondVal = CStr(arg)
        .ThirdVal = CDbl(arg)
    End With
",
                [3] =
@"
    With tType
        .FirstVal = arg
        .SecondVal = CStr(arg)
TestLabel1:
        .ThirdVal = CDbl(arg)
    End With
",
                [4] =
@"
    With tType
        .FirstVal = arg
        .SecondVal = CStr(arg)
        .ThirdVal = CDbl(arg)
TestLabel1:
    End With"
        };

        var inputCode =
$@"

Private Type TestType
    FirstVal As Long
    SecondVal As String
    ThirdVal As Double
End Type

Private tType As TestType

Sub Foo(ByVal arg As Long)
    If arg < 0 Then
        arg = arg * 10
        GoTo TestLabel1
    End If

    {withStatementVersions[testCase]}

End Sub
";

            var expected =
@"
    With tType
        .FirstVal = arg
        .SecondVal = CStr(arg)
        .ThirdVal = CDbl(arg)
    End With
";

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "TestLabel1"));
            StringAssert.Contains(expected, actualCode);
        }

        [TestCase(1)]
        [TestCase(2)]
        [TestCase(3)]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void LabelWithinSelectCase(long testCase)
        {
            var selectCaseVersions = new Dictionary<long, string>()
            {
                [1] =
@"
    Select Case arg
        Case Is >= 80
Label1:
            result = ""very good""
        Case Is >= 70
            result = ""good""
        Case Is >= 60
            result = ""sufficient""
        Case Else
            result = ""insufficient""
    End Select
",
                [2] =
@"
    Select Case arg
        Case Is >= 80
            result = ""very good""
Label1:
        Case Is >= 70
            result = ""good""
        Case Is >= 60
            result = ""sufficient""
        Case Else
            result = ""insufficient""
    End Select
",
                [3] =
@"
    Select Case arg
        Case Is >= 80
            result = ""very good""
        Case Is >= 70
            result = ""good""
        Case Is >= 60
            result = ""sufficient""
        Case Else
Label1:
            result = ""insufficient""
    End Select
"
            };

            var inputCode =
$@"
Private Function Test(arg As Long) As String

    Dim result As String
    GoTo Label1
    
{selectCaseVersions[testCase]}

    Test = result
End Function
";

            var expected =
@"
    Select Case arg
        Case Is >= 80
            result = ""very good""
        Case Is >= 70
            result = ""good""
        Case Is >= 60
            result = ""sufficient""
        Case Else
            result = ""insufficient""
    End Select
";

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "Label1"));
            StringAssert.Contains(expected, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void LabelWithinSelectCaseOnExpressionLine()
        {
            var inputCode =
$@"
Private Function Test(arg As Long) As String

    Dim result As String
    GoTo Label1
    
    Select Case arg
        Case Is >= 80
Label1:     result = ""very good""
        Case Is >= 70
            result = ""good""
        Case Is >= 60
            result = ""sufficient""
        Case Else
            result = ""insufficient""
    End Select

    Test = result
End Function
";

            var expected =
@"
    Select Case arg
        Case Is >= 80
            result = ""very good""
        Case Is >= 70
            result = ""good""
        Case Is >= 60
            result = ""sufficient""
        Case Else
            result = ""insufficient""
    End Select
";

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "Label1"));
            StringAssert.Contains(expected, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void LabelGroups()
        {
            const string inputCode =
@"
Sub Foo(ByVal arg As Long)
Label1: 'Group1

Label2:

Label3:     'Group2

4     'Group2

5     'Group2

6

7

Label4: 'Group3

End Sub
";

            var expected =
@"
Sub Foo(ByVal arg As Long)
Label2:

6

7

End Sub
";

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "Label1", "Label3", "4", "5", "Label4"));
            StringAssert.Contains(expected, actualCode);
            StringAssert.AreEqualIgnoringCase(expected, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void UsesCorrectScopingContextWithStmt()
        {
            const string inputCode =
@"
Private Type TestType
    FirstVal As Long
    SecondVal As String
    ThirdVal As Double
End Type

Private tType As TestType

Sub Foo(ByVal arg As Long)

    Dim types As Collection
    Set types = new Collection
    Dim idx As Long

    For idx = 0 To arg
        Dim localType As tType
        With localType
Label1: 
            .FirstVal = idx
            types.Add localType
        End With
    Next idx

End Sub
";

            var expected =
@"
    For idx = 0 To arg
        Dim localType As tType
        With localType
            .FirstVal = idx
            types.Add localType
        End With
    Next idx
";

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "Label1"));
            StringAssert.Contains(expected, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void UsesCorrectScopingContextForNext()
        {
            const string inputCode =
@"
Private Type TestType
    FirstVal As Long
    SecondVal As String
    ThirdVal As Double
End Type

Private tType As TestType

Sub Foo(ByVal arg As Long)

    Dim types As Collection
    Set types = new Collection
    Dim idx As Long

    For idx = 0 To arg
Label1:
        Dim localType As tType
        With localType
            .FirstVal = idx
            types.Add localType
        End With
    Next idx

End Sub
";

            var expected =
@"
    For idx = 0 To arg
        Dim localType As tType
        With localType
            .FirstVal = idx
            types.Add localType
        End With
    Next idx
";

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "Label1"));
            StringAssert.Contains(expected, actualCode);
        }

        private string GetRetainedCodeBlock(string moduleCode, Func<RubberduckParserState, IEnumerable<Declaration>> targetListBuilder, Action<IDeleteDeclarationsModel> modelFlagAction = null)
        {
            var refactoredCode = _support.TestRefactoring(
                targetListBuilder,
                RefactorProcedureScopeElements,
                modelFlagAction ?? _support.DefaultModelFlagAction,
                (MockVbeBuilder.TestModuleName, moduleCode, ComponentType.StandardModule));

            return refactoredCode[MockVbeBuilder.TestModuleName];
        }

        private static IExecutableRewriteSession RefactorProcedureScopeElements(RubberduckParserState state, IEnumerable<Declaration> targets, IRewritingManager rewritingManager, Action<IDeleteDeclarationsModel> modelFlagAction)
        {
            var model = new DeleteProcedureScopeElementsModel(targets);
            modelFlagAction(model);

            var session = rewritingManager.CheckOutCodePaneSession();

            var refactoringAction = new DeleteDeclarationsTestsResolver(state, rewritingManager)
                .Resolve<DeleteProcedureScopeElementsRefactoringAction>();

            refactoringAction.Refactor(model, session);

            return session;
        }
    }
}
