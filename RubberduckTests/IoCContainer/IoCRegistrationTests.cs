using Castle.Windsor;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Settings;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.Root;
using RubberduckTests.Mocks;

namespace RubberduckTests.IoCContainer
{
    [TestClass]
    public class IoCRegistrationTests
    {
        [TestMethod]
        public void RegistrationOfRubberduckIoCContainerWithSC_NoException()
        {
            var vbeBuilder = new MockVbeBuilder();
            var ide = vbeBuilder.Build().Object;
            var addin = new Mock<IAddIn>().Object;
            var initialSettings = new GeneralSettings {SourceControlEnabled = true};

            using (var container =
                new WindsorContainer().Install(new RubberduckIoCInstaller(ide, addin, initialSettings)))
            {
            }

            //This test does not need an assert because it tests that no exception has been thrown.
        }

        [TestMethod]
        public void RegistrationOfRubberduckIoCContainerWithoutSC_NoException()
        {
            var vbeBuilder = new MockVbeBuilder();
            var ide = vbeBuilder.Build().Object;
            var addin = new Mock<IAddIn>().Object;
            var initialSettings = new GeneralSettings {SourceControlEnabled = false};

            using (var container =
                new WindsorContainer().Install(new RubberduckIoCInstaller(ide, addin, initialSettings)))
            {
            }

            //This test does not need an assert because it tests that no exception has been thrown.
        }
    }
}
