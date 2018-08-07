using System.Collections.Generic;
using Castle.Windsor;
using NUnit.Framework;
using Rubberduck.Settings;
using Rubberduck.Root;
using RubberduckTests.Mocks;

namespace RubberduckTests.IoCContainer
{
    [TestFixture]
    public class IoCRegistrationTests
    {
        [Test]
        [Category("IoC_Registration")]
        public void RegistrationOfRubberduckIoCContainerWithSC_NoException()
        {
            var vbeBuilder = new MockVbeBuilder();
            var addInBuilder = new MockAddInBuilder();
            var ide = vbeBuilder.Build().Object;            
            var addin = addInBuilder.Build().Object;
            var initialSettings = new GeneralSettings
            {
                EnableExperimentalFeatures = new List<ExperimentalFeatures>
                {
                    new ExperimentalFeatures()
                }
            };

            using (var container =
                new WindsorContainer().Install(new RubberduckIoCInstaller(ide, addin, initialSettings)))
            {
            }

            //This test does not need an assert because it tests that no exception has been thrown.
        }

        [Test]
        [Category("IoC_Registration")]
        public void RegistrationOfRubberduckIoCContainerWithoutSC_NoException()
        {
            var vbeBuilder = new MockVbeBuilder();
            var addInBuilder = new MockAddInBuilder();
            var ide = vbeBuilder.Build().Object;
            var addin = addInBuilder.Build().Object;
            var initialSettings = new GeneralSettings {EnableExperimentalFeatures = new List<ExperimentalFeatures>()};

            using (var container =
                new WindsorContainer().Install(new RubberduckIoCInstaller(ide, addin, initialSettings)))
            {
            }

            //This test does not need an assert because it tests that no exception has been thrown.
        }
    }
}
