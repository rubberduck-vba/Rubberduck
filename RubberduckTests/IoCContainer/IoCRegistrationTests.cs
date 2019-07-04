using System.Collections.Generic;
using Castle.Windsor;
using Moq;
using NUnit.Framework;
using Rubberduck.Settings;
using Rubberduck.Root;
using Rubberduck.Runtime;
using Rubberduck.VBEditor.SourceCodeHandling;
using Rubberduck.VBEditor.VbeRuntime;
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
            var ideMock = vbeBuilder.Build();
            var sourceFileHandler = new Mock<ITempSourceFileHandler>().Object;
            ideMock.Setup(m => m.TempSourceFileHandler).Returns(() => sourceFileHandler);
            var ide = ideMock.Object;
            var addInBuilder = new MockAddInBuilder();
            var addin = addInBuilder.Build().Object;
            var vbeNativeApi = new Mock<IVbeNativeApi>();
            var beepInterceptor = new Mock<IBeepInterceptor>();
            var initialSettings = new GeneralSettings
            {
                EnableExperimentalFeatures = new List<ExperimentalFeature>
                {
                    new ExperimentalFeature()
                }
            };

            using (var container =
                new WindsorContainer().Install(new RubberduckIoCInstaller(ide, addin, initialSettings, vbeNativeApi.Object, beepInterceptor.Object)))
            {
            }

            //This test does not need an assert because it tests that no exception has been thrown.
        }

        [Test]
        [Category("IoC_Registration")]
        public void RegistrationOfRubberduckIoCContainerWithoutSC_NoException()
        {
            var vbeBuilder = new MockVbeBuilder();
            var ideMock = vbeBuilder.Build();
            var sourceFileHandler = new Mock<ITempSourceFileHandler>().Object;
            ideMock.Setup(m => m.TempSourceFileHandler).Returns(() => sourceFileHandler);
            var ide = ideMock.Object;
            var addInBuilder = new MockAddInBuilder();
            var addin = addInBuilder.Build().Object;
            var vbeNativeApi = new Mock<IVbeNativeApi>();
            var beepInterceptor = new Mock<IBeepInterceptor>();

            var initialSettings = new GeneralSettings {EnableExperimentalFeatures = new List<ExperimentalFeature>()};

            using (var container =
                new WindsorContainer().Install(new RubberduckIoCInstaller(ide, addin, initialSettings, vbeNativeApi.Object, beepInterceptor.Object)))
            {
            }

            //This test does not need an assert because it tests that no exception has been thrown.
        }
    }
}
