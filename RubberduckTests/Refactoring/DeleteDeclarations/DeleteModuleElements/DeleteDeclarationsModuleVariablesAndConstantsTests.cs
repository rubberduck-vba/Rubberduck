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
using System.Text.RegularExpressions;

namespace RubberduckTests.Refactoring.DeleteDeclarations
{
    [TestFixture]

    public class DeleteDeclarationsModuleVariablesAndConstantsTests : ModuleSectionElementsTestsBase
    {
        [TestCase("Option Explicit\r\n\r\n")]
        [TestCase("")]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void RemoveAllFieldDeclarations(string optionExplicit)
        {
            var inputCode =
$@"
{optionExplicit}Public mVar1 As Long

Public mVar2 As Long

Private mVar3 As String, mVar4 As Long, mVar5 As String

Public Sub Test()
End Sub";

            var expected =
$@"
{optionExplicit}Public Sub Test()
End Sub";
            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "mVar1", "mVar2", "mVar3", "mVar4", "mVar5"));
            StringAssert.AreEqualIgnoringCase(expected, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void RemoveAllConstantDeclarations()
        {
            var inputCode =
@"
Option Explicit

Public Const mVar1 As Long = 100

Public Const mVar2 As Long = 200

Private Const mVar3 As String = 300, mVar4 As Long = 400, mVar5 As String = ""Test5""

Public Sub Test()
End Sub";

            var expected =
@"
Option Explicit

Public Sub Test()
End Sub";
            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "mVar1", "mVar2", "mVar3", "mVar4", "mVar5"));
            StringAssert.AreEqualIgnoringCase(expected, actualCode);
        }

        [TestCase("Public", "target As Integer")]
        [TestCase("Private", "target As Integer")]
        [TestCase("Public", "Const target As Integer = 9")]
        [TestCase("Private", "Const target As Integer = 9")]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void PublicPrivateAccessibilities(string visibility, string declaration)
        {
            var inputCode =
$@"
Option Explicit

{visibility} {declaration}
";

            var expected =
$@"
Option Explicit
";

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "target"));
            StringAssert.AreEqualIgnoringCase(expected, actualCode);
        }

        [TestCase("Public notUsed1 As Long")]
        [TestCase("Public Const notUsed1 As Long = 100")]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void MultiplNewLinesPriorToNextDeclarationRetained(string declaration)
        {
            var inputCode =
$@"
Option Explicit

{declaration}



'Comment after many blank lines

     'Another indented Comment
Public notUsed2 As Long
Public notUsed3 As Long
";
            var expected =
@"
Option Explicit

'Comment after many blank lines

     'Another indented Comment
Public notUsed2 As Long
Public notUsed3 As Long
";

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "notUsed1"));
            StringAssert.AreEqualIgnoringCase(expected, actualCode);
        }

        [TestCase("Public  mVar1 As Long: ", "Public mVar2 As Long")]
        [TestCase("Public Const mVar1 As Long = 100: ", "Public Const mVar2 As Long = 200")]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void RemoveFieldDeclarationsUsesColonStmtDelimiter(string declaration1, string declaration2)
        {
            var inputCode =
$@"
Option Explicit

{declaration1}{declaration2}

Private mVar3 As String, mVar4 As Long, mVar5 As String

Public Sub Test()
End Sub";

            var expected =
$@"
Option Explicit

{declaration2}

Public Sub Test()
End Sub";
            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "mVar1", "mVar3", "mVar4", "mVar5"));
            StringAssert.AreEqualIgnoringCase(expected, actualCode);
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

Public Sub Test()
End Sub";

            var expected =
@"
Option Explicit

Private mVar4 As Long

Public Sub Test()
End Sub";

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "mVar1", "mVar2", "mVar3", "mVar5"));
            StringAssert.AreEqualIgnoringCase(expected, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void RemoveConstDeclarations_LineContinuations()
        {
            var inputCode =
@"
Option Explicit

Public Const mVar1 As Long = 100

Public Const mVar2 As Long = 200

Private Const mVar3 As Long = 300 _
        , mVar4 As Long = 400 _
                , mVar5 As Long = 500

Public Sub Test()
End Sub";

            var expected =
@"
Option Explicit

Private Const mVar4 As Long = 400

Public Sub Test()
End Sub";

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "mVar1", "mVar2", "mVar3", "mVar5"));
            StringAssert.AreEqualIgnoringCase(expected, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void InlineFieldDeclarationListWithLineContinuation_RemovesSameLineComment()
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

            var expected =
@"
Option Explicit

'This field is used
Public used As String
";
            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "notUsed1", "notUsed2", "notUsed3"));
            StringAssert.AreEqualIgnoringCase(expected, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void InlineConstantDeclarationListWithLineContinuation_RemovesSameLineComment()
        {
            var inputCode =
@"
Option Explicit

Public Const notUsed1 As Long = 100, _
    notUsed2 As Long = 200, _
        notUsed3 As Long = 300 _
            'These constants are not used


'This field is used
Public used As String
";

            var expected =
@"
Option Explicit

'This field is used
Public used As String
";
            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "notUsed1", "notUsed2", "notUsed3"));
            StringAssert.Contains(expected, actualCode);
            StringAssert.AreEqualIgnoringCase(expected, actualCode);
        }

        [TestCase("notUsed1", "notUsed2", "notUsed3")]
        [TestCase("notUsed2", "notUsed1", "notUsed3")]
        [TestCase("notUsed3", "notUsed1", "notUsed2")]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void FieldDeclarationListLineContinuationWithCommentRemoveSingle_RetainsComment(string toRemove, params string[] retained)
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

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, toRemove));
            StringAssert.Contains("This field is used", actualCode);
            StringAssert.Contains("Public used As", actualCode);
            StringAssert.Contains("One of these fields are not used", actualCode);
            foreach (var field in retained)
            {
                StringAssert.Contains(field, actualCode);
            }
            StringAssert.DoesNotContain(toRemove, actualCode);
        }

        [TestCase("notUsed1", "notUsed2", "notUsed3")]
        [TestCase("notUsed2", "notUsed1", "notUsed3")]
        [TestCase("notUsed3", "notUsed1", "notUsed2")]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void ConstantDeclarationListLineContinuationWithCommentRemoveSingle_RetainsComment(string toRemove, params string[] retained)
        {
            var inputCode =
@"
Option Explicit

Public Const notUsed1 As Long = 100, _
    notUsed2 As Long = 200, _
        notUsed3 As Long = 300 _
            'One of these fields are not used


'This field is used
Public used As String
";

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, toRemove));
            StringAssert.Contains("This field is used", actualCode);
            StringAssert.Contains("Public used As", actualCode);
            StringAssert.Contains("One of these fields are not used", actualCode);
            foreach (var field in retained)
            {
                StringAssert.Contains(field, actualCode);
            }
            StringAssert.DoesNotContain(toRemove, actualCode);
        }

        [TestCase("Public notUsed1 As Long, notUsed2 As Long, notUsed3 As Long")]
        [TestCase("Public Const notUsed1 As Long = 100, notUsed2 As Long = 200, notUsed3 As Long = 300")]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void DeclarationListAllRemoved(string declarationList)
        {
            var inputCode =
$@"
Option Explicit

Private used As Long
{declarationList}
";

            var expected =
$@"
Option Explicit

Private used As Long
";

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "notUsed1", "notUsed2", "notUsed3"));
            StringAssert.AreEqualIgnoringCase(expected, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void SingleDeclarationOfManyWithAnnotationRetainsOtherComments()
        {
            var inputCode =
$@"
Option Explicit 'Use this to force explicit declarations

'This is a single deletion test with comments and annotations to deal with


'This comment above the deleted Annotation needs to remain
'@VariableDescription(""This is the delete target of the Test"")
'This comment below the deleted Annotation also needs to remain
Private target As Long

'@VariableDescription(""This is NOT the delete target of the Test"")
Private somethingElse As Double
";
            var expected =
$@"
Option Explicit 'Use this to force explicit declarations

'This is a single deletion test with comments and annotations to deal with


'This comment above the deleted Annotation needs to remain
'This comment below the deleted Annotation also needs to remain

'@VariableDescription(""This is NOT the delete target of the Test"")
Private somethingElse As Double
";

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "target"));
            StringAssert.AreEqualIgnoringCase(expected, actualCode);
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

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "mVar1", "mVar2", "mVar3", "mVar4", "mVar5"));
            StringAssert.AreEqualIgnoringCase(expected, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void RemoveConstantDeclarations_RemovesAllBlankLines()
        {
            var inputCode =
@"
Option Explicit

Public Const mVar1 As Long = 100





Public Const mVar2 As Long = 200
Private Const mVar3 As String = ""Test3"", mVar4 As Long = 400, mVar5 As String = ""Test5""





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

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "mVar1", "mVar2", "mVar3", "mVar4", "mVar5"));
            StringAssert.AreEqualIgnoringCase(expected, actualCode);
        }

        [TestCase("Public mVar1 As Long")]
        [TestCase("Public Const mVar1 As Long = 100")]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void RemovesLogicalLineCommentwithLineContinuation(string declaration)
        {
            var inputCode =
$@"
Option Explicit

    Private retained As Long

    'Comment above mVar1
{declaration} 'This is a comment for mVar1 _
        'and so is this

        'Comment below mVar1
Public Sub Test()
End Sub";

            var expected =
@"
Option Explicit

    Private retained As Long

    'Comment above mVar1
        'Comment below mVar1
Public Sub Test()
End Sub";
            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "mVar1"));
            StringAssert.AreEqualIgnoringCase(expected, actualCode);
        }

        [TestCase("Public mVar1 As Long")]
        [TestCase("Public Const mVar1 As Long = 100")]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void RemovesLogicalLineCommentwithLineContinuationInjectTODO(string declaration)
        {
            var inputCode =
$@"
Option Explicit

    Private retained As Long

    'Comment above mVar1
{declaration} 'This is a comment for mVar1 _
        'and so is this

        'Comment below mVar1
Public Sub Test()
End Sub";

            var expected =
$@"
Option Explicit

    Private retained As Long

    {DeleteDeclarationsTestSupport.TodoContent}Comment above mVar1
        {DeleteDeclarationsTestSupport.TodoContent}Comment below mVar1
Public Sub Test()
End Sub";
            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "mVar1"), (m) => m.InsertValidationTODOForRetainedComments = true);
            StringAssert.Contains(expected, actualCode);
            StringAssert.AreEqualIgnoringCase(expected, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void VariableDescriptionAnnotation()
        {
            var inputCode =
$@"
Option Explicit

'@VariableDescription(""Exposes a read / write value."")
Public SomeValue As Long

Public Sub DoSomething()
End Sub
";

            var expected =
$@"
Option Explicit

Public Sub DoSomething()
End Sub
";
            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "SomeValue"));
            StringAssert.AreEqualIgnoringCase(expected, actualCode);
        }

        [TestCase("Public X As Long", "VariableNotUsed")]
        [TestCase("Public Const X As Long = 9", "ConstantNotUsed")]
        [Category("Refactorings")]
        [Category("DeleteDeclarationWithLineLabel")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void AnnotationWithSubsequentComment_RemovesAnnotations(string declaration, string ignoreNotUsed)
        {
            var inputCode =
$@"
Option Explicit

'There is already a comment above the deleted Annotation
'@Ignore {ignoreNotUsed}, UseMeaningfulName
'And then another below the deleted Annotation
{declaration}

'The Sub has a comment
Public Sub DoSomething()
End Sub
";

            var expectedCode =
$@"
Option Explicit

'There is already a comment above the deleted Annotation
'And then another below the deleted Annotation

'The Sub has a comment
Public Sub DoSomething()
End Sub
";

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "X"));
            StringAssert.AreEqualIgnoringCase(expectedCode, actualCode);
        }
        [TestCase("    ", "Public X As Long", "VariableNotUsed")]
        [TestCase("    ", "Public Const X As Long = 9", "ConstantNotUsed")]
        [TestCase("        ", "Public X As Long", "VariableNotUsed")]
        [TestCase("        ", "Public Const X As Long = 9", "ConstantNotUsed")]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void MultipleAnnotationLists_RetainsNextStatementIndentation(string nextStatementIndentation, string declaration, string ignoreNotAssigned)
        {
            var inputCode =
$@"
Option Explicit

Public AnotherVar As String

    '@Ignore {ignoreNotAssigned}
    '@Ignore UseMeaningfulName
    {declaration}

{nextStatementIndentation}Private usedVar As Long

Public Sub DoSomething()
End Sub
";

            var expectedCode =
$@"
Option Explicit

Public AnotherVar As String

{nextStatementIndentation}Private usedVar As Long

Public Sub DoSomething()
End Sub
";

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "X"));
            StringAssert.AreEqualIgnoringCase(expectedCode, actualCode);
        }

        [TestCase(true)]
        [TestCase(false)]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void RespectsInjectTODOCommentFlag(bool injectTODO)
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

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "mVar1"), (m) => m.InsertValidationTODOForRetainedComments = injectTODO);
            StringAssert.DoesNotContain("mVar1", actualCode);
            StringAssert.Contains("Test1", actualCode);
            var injectedContent = injectTODO
                ? DeleteDeclarationsTestSupport.TodoContent
                : string.Empty;
            StringAssert.Contains($"{injectedContent}A comment following an Annotation", actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void MultipleDeleteGroupsCheck_DeclarationSection()
        {
            var inputCode =
@"
    Private firstLong As Long 'Group1

    Private mainCollection As Collection

    Private firstStr As String 'Group2
    Private secondStr As String 'Group2
    
    Private thirdStr As String 'Group2

    Private firstVar As Variant

    Private i As Long

    Private firstBool As Boolean

    Private secondBool As Boolean 'Group3
";

            var expected =
@"
    Private mainCollection As Collection

    Private firstVar As Variant

    Private i As Long

    Private firstBool As Boolean
";

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "firstLong", "firstStr", "secondStr", "thirdStr", "secondBool"));
            StringAssert.Contains(expected, actualCode);
            StringAssert.AreEqualIgnoringCase(expected, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void MultipleDeleteGroupsCheck_DeclarationSectionEveryOtherDeclaration()
        {
            var inputCode =
@"
    Private firstLong As Long 'Group1

    Private mainCollection As Collection

    Private firstStr As String 'Group2
    Private secondStr As String 'Group2
    
    Private thirdStr As String 'Group2

    Private firstVar As Variant
";

            var expected =
@"
    Private mainCollection As Collection

    Private firstVar As Variant
";

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "firstLong", "firstStr", "secondStr", "thirdStr"));
            StringAssert.Contains(expected, actualCode);
            StringAssert.AreEqualIgnoringCase(expected, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void MultipleDeleteGroupsCheck_DeclarationSectionEveryOtherDeclarationLastWithMultiple()
        {
            var inputCode =
@"
    Private firstLong As Long 'Group1

    Private mainCollection As Collection

    Private firstStr As String 'Group2

    Private firstVar As Variant
";

            var expected =
@"
    Private mainCollection As Collection

    Private firstVar As Variant
";

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "firstLong", "firstStr"));
            StringAssert.Contains(expected, actualCode);
            StringAssert.AreEqualIgnoringCase(expected, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void FindsDeclarationLogicalLineContextListDeclarationMultiLineComment()
        {
            var inputCode =
$@"
Option Explicit

    Private retained As Long

    'Comment above mVar1
Private mVar1 As Long 'This is a comment for mVar1 _
        'and so is this

        'Comment below mVar1
Public Sub Test()
End Sub";

            void thisTest(IDeclarationDeletionTarget sut)
            {
                var commentContext = sut.GetDeclarationLogicalLineCommentContext();
                var content = commentContext?.GetText() ?? string.Empty;
                StringAssert.Contains("'This is a comment for mVar1", content);
                StringAssert.Contains("'and so is this", content);
            }

            _support.SetupAndInvokeIDeclarationDeletionTargetTest(inputCode, "mVar1", thisTest);
        }
    }
}

