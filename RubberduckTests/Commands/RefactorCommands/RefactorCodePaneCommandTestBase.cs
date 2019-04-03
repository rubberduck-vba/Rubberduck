using NUnit.Framework;

namespace RubberduckTests.Commands.RefactorCommands
{
    [TestFixture]
    public abstract class RefactorCodePaneCommandTestBase : RefactorCommandTestBase
    {
        [Category("Commands")]
        [Test]
        public void EncapsulateField_CanExecute_NullActiveCodePane()
        {
            var vbe = SetupAllowingExecution();
            vbe.ActiveCodePane = null;

            Assert.IsFalse(CanExecute(vbe));
        }
    }
}