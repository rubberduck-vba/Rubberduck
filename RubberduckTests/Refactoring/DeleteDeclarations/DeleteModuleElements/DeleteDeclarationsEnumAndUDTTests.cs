using NUnit.Framework;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.DeleteDeclarations;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RubberduckTests.Refactoring.DeleteDeclarations
{
    [TestFixture]
    public class DeleteDeclarationsEnumAndUDTTests : ModuleSectionElementsTestsBase
    {
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
";

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "TestType"));
            StringAssert.Contains("mVar1", actualCode);
            StringAssert.Contains("Test1", actualCode);
            StringAssert.DoesNotContain("TestType", actualCode);
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
";

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "TestEnum"));
            StringAssert.Contains("mVar1", actualCode);
            StringAssert.Contains("Test1", actualCode);
            StringAssert.DoesNotContain("TestEnum", actualCode);
            StringAssert.DoesNotContain("FirstValue", actualCode);
        }
    }
}
