using Castle.Windsor;
using NUnit.Framework;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command;
using Rubberduck.UI.Command.Refactorings;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.Utility;
using RubberduckTests.Refactoring.EncapsulateField;

namespace RubberduckTests.Commands.RefactorCommands
{
    public class EncapsulateFieldCommandTests : RefactorCodePaneCommandTestBase
    {
        [Test]
        [Category("Commands")]
        [Category("Encapsulate Field")]
        public void EncapsulateField_CanExecute_LocalVariable()
        {
            const string input =
                @"Sub Foo()
    Dim d As Boolean
End Sub";
            var selection = new Selection(2, 9, 2, 9);
            Assert.IsFalse(CanExecute(input, selection));
        }

        [Test]
        [Category("Commands")]
        [Category("Encapsulate Field")]
        public void EncapsulateField_CanExecute_Proc()
        {
            const string input =
                @"Dim d As Boolean
Sub Foo()
End Sub";
            var selection = new Selection(2, 7, 2, 7);
            Assert.IsFalse(CanExecute(input, selection));
        }

        [Test]
        [Category("Commands")]
        [Category("Encapsulate Field")]
        public void EncapsulateField_CanExecute_Field()
        {
            const string input =
                @"Dim d As Boolean
Sub Foo()
End Sub";
            var selection = new Selection(1, 5, 1, 5);
            Assert.IsTrue(CanExecute(input, selection));
        }

        protected override CommandBase TestCommand(IVBE vbe, RubberduckParserState state, IRewritingManager rewritingManager, ISelectionService selectionService)
        {
            var resolver = new EncapsulateFieldTestsResolver(state, rewritingManager, selectionService);
            resolver.Install(new WindsorContainer(), null);
            return resolver.Resolve<RefactorEncapsulateFieldCommand>();
        }

        protected override IVBE SetupAllowingExecution()
        {
            const string input =
                @"Dim d As Boolean
Sub Foo()
End Sub";
            var selection = new Selection(1, 5, 1, 5);
            return TestVbe(input, selection);
        }
    }
}
