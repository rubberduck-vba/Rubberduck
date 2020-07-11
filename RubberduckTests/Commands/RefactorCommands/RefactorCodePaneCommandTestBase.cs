using NUnit.Framework;

namespace RubberduckTests.Commands.RefactorCommands
{
    [TestFixture]
    public abstract class RefactorCodePaneCommandTestBase : RefactorCommandTestBase
    {
        [Category("Commands")]
        [Test]
        public void RefactoringCommand_CanExecute_ValidInput()
        {
            var vbe = SetupAllowingExecution();

            Assert.IsTrue(CanExecute(vbe), GetType().FullName);
        }

        [Category("Commands")]
        [Test]
        public void RefactoringCommand_CanExecute_NullActiveCodePane()
        {
            var vbe = SetupAllowingExecution();
            vbe.ActiveCodePane = null;

            Assert.IsFalse(CanExecute(vbe), GetType().FullName);
        }
    }
}