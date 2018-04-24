using NUnit.Framework;
using Rubberduck.VBEditor.ComManagement;

namespace RubberduckTests.VBEditor
{
    [TestFixture()]
    public class ComSafeManagerTests
    {
        [Test]
        [Category("COM")]
        public void ComSafeReturnedOnSecondIvokationOfGetCurrentComSafeIsTheSame()
        {
            ComSafeManager.DisposeAndResetComSafe(); //Resetting to get a claen start.

            var comSafe1 = ComSafeManager.GetCurrentComSafe();
            var comSafe2 = ComSafeManager.GetCurrentComSafe();

            Assert.AreSame(comSafe1, comSafe2);
        }

        [Test]
        [Category("COM")]
        public void AfterCallingResetComSafeGetCurrentComSafeReturnsDifferentSafe()
        {
            ComSafeManager.DisposeAndResetComSafe(); //Resetting to get a claen start.

            var comSafe1 = ComSafeManager.GetCurrentComSafe();
            ComSafeManager.DisposeAndResetComSafe();
            var comSafe2 = ComSafeManager.GetCurrentComSafe();

            Assert.AreNotSame(comSafe1, comSafe2);
        }
    }
}