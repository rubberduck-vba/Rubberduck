using Antlr4.Runtime;
using Castle.Windsor;
using NUnit.Framework;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.AddInterfaceImplementations;
using Rubberduck.Refactorings.DeleteDeclarations;
using Rubberduck.Refactorings.DeleteDeclarations.Abstract;
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
    public class DeleteDeclarationTargetBaseTests
    {
        private readonly DeleteDeclarationsTestSupport _support = new DeleteDeclarationsTestSupport();

        [TestCase("\r\n")]
        [TestCase("\r\n\r\n")]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void FindsSeparationToNextDeclaration(string separation)
        {
            var inputCode =
$@"
Option Explicit

Public Sub DoSomething(arg As Long)
End Sub{separation}Public Sub DoSomethingElse(arg As Long)
    Dim X As Long
End Sub
";

            void thisTest(IDeclarationDeletionTarget sut)
            {
                var concrete = sut as DeclarationDeletionTargetBase;
                StringAssert.AreEqualIgnoringCase(separation, concrete.EOSSeparation);
            }

            _support.SetupAndInvokeIDeclarationDeletionTargetTest(inputCode, "DoSomething", thisTest);
        }

        [TestCase("    ")]
        [TestCase("")]
        [TestCase("        ")]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void FindsIndentationToNextContext(string expectedIndentation)
        {
            var inputCode =
$@"
Option Explicit

Public Sub DoSomethingElse(arg As Long)

Dim Target As String
{expectedIndentation}Dim X As Long

End Sub
";
            void thisTest(IDeclarationDeletionTarget sut)
            {
                var concrete = sut as DeclarationDeletionTargetBase;
                StringAssert.Contains(expectedIndentation, concrete.EOSIndentation);
            }

            _support.SetupAndInvokeIDeclarationDeletionTargetTest(inputCode, "Target", thisTest);
        }
    }
}
