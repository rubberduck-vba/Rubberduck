using Castle.Windsor;
using NUnit.Framework;
using Moq;
using Rubberduck.Settings;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.Root;
using RubberduckTests.Mocks;

namespace RubberduckTests.IoCContainer
{
    [TestFixture]
    public class IoCRegistrationTests
    {
        [Test]
        public void RegistrationOfRubberduckIoCContainerWithSC_NoException()
        {
            var vbeBuilder = new MockVbeBuilder();
            var ide = vbeBuilder.Build().Object;
            var addin = new Mock<IAddIn>().Object;
            var initialSettings = new GeneralSettings {IsSourceControlEnabled = true};

            using (var container =
                new WindsorContainer().Install(new RubberduckIoCInstaller(ide, addin, initialSettings)))
            {
            }

            //This test does not need an assert because it tests that no exception has been thrown.
        }

        [Test]
        public void RegistrationOfRubberduckIoCContainerWithoutSC_NoException()
        {
            var vbeBuilder = new MockVbeBuilder();
            var ide = vbeBuilder.Build().Object;
            var addin = new Mock<IAddIn>().Object;
            var initialSettings = new GeneralSettings {IsSourceControlEnabled = false};

            using (var container =
                new WindsorContainer().Install(new RubberduckIoCInstaller(ide, addin, initialSettings)))
            {
            }

            //This test does not need an assert because it tests that no exception has been thrown.
        }
    }
}
