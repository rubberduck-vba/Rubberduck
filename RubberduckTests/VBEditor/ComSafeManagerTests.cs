using NUnit.Framework;
using Rubberduck.VBEditor.ComManagement;

namespace RubberduckTests.VBEditor
{
    [TestFixture()]
    public class ComSafeManagerTests
    {
        [Test()]
        public void ComSafeReturnedOnSecondIvokationOfGetCurrentComSafeIsTheSame()
        {
            ComSafeManager.ResetComSafe(); //Resetting to get a claen start.

            var comSafe1 = ComSafeManager.GetCurrentComSafe();
            var comSafe2 = ComSafeManager.GetCurrentComSafe();

            Assert.AreSame(comSafe1, comSafe2);
        }

        [Test()]
        public void AfterCallingResetComSafeGetCurrentComSafeReturnsDifferentSafe()
        {
            ComSafeManager.ResetComSafe(); //Resetting to get a claen start.

            var comSafe1 = ComSafeManager.GetCurrentComSafe();
            ComSafeManager.ResetComSafe();
            var comSafe2 = ComSafeManager.GetCurrentComSafe();

            Assert.AreNotSame(comSafe1, comSafe2);
        }
    }
}