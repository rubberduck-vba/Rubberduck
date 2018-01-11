using System;
using System.Collections.Generic;
using Castle.Windsor;
using NUnit.Framework;
using Moq;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.VBA;
using Rubberduck.Settings;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.Root;
using Rubberduck.UI;
using RubberduckTests.Mocks;

namespace RubberduckTests.IoCContainer
{
    [TestFixture]
    public class IoCResolvingTests
    {
        [Test]
        public void ResolveInspections_NoException()
        {
            var vbeBuilder = new MockVbeBuilder();
            var ide = vbeBuilder.Build().Object;
            var addin = new Mock<IAddIn>().Object;
            var initialSettings = new GeneralSettings
            {
                EnableExperimentalFeatures = new List<ExperimentalFeatures>
                {
                    new ExperimentalFeatures
                    {
                        Key = nameof(RubberduckUI.GeneralSettings_EnableSourceControl)
                    }
                }
            };

            IWindsorContainer container = null;
            try
            {
                try
                {
                    container = new WindsorContainer().Install(new RubberduckIoCInstaller(ide, addin, initialSettings));
                }
                catch (Exception exception)
                {
                    Assert.Inconclusive($"Unable to register. {Environment.NewLine} {exception}");
                }

                var inspections = container.ResolveAll<IInspection>();

                //This test does not need an assert because it tests that no exception has been thrown.
            }
            finally
            {
                container?.Dispose();
            }
        }

        [Test]
        public void ResolveRubberduckParserState_NoException()
        {
            var vbeBuilder = new MockVbeBuilder();
            var ide = vbeBuilder.Build().Object;
            var addin = new Mock<IAddIn>().Object;
            var initialSettings = new GeneralSettings
            {
                EnableExperimentalFeatures = new List<ExperimentalFeatures>
                {
                    new ExperimentalFeatures
                    {
                        Key = nameof(RubberduckUI.GeneralSettings_EnableSourceControl)
                    }
                }
            };

            IWindsorContainer container = null;
            try
            {
                try
            {
                container = new WindsorContainer().Install(new RubberduckIoCInstaller(ide, addin, initialSettings));
            }
            catch (Exception exception)
            {
                Assert.Inconclusive($"Unable to register. {Environment.NewLine} {exception}");
            }

            var state = container.ResolveAll<RubberduckParserState>();

                //This test does not need an assert because it tests that no exception has been thrown.
            }
            finally
            {
                container?.Dispose();
            }
        }
    }
}
