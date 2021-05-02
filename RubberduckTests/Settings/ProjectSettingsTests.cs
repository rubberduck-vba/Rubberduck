using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NUnit.Framework;
using Rubberduck.Settings;
using Moq;

namespace RubberduckTests.Settings
{
    [TestFixture]
    class ProjectSettingsTests
    {
        public static ProjectSettings GetMockProjectSettings()
        {
            var mockSettings = new Mock<ProjectSettings>();

            return mockSettings.Object;
        }

        private Configuration GetDefaultConfig()
        {
            var userSettings = new UserSettings(null, null, null, null, null, null, null, null, new ProjectSettings());
            return new Configuration(userSettings);
        }

        [Category("Settings")]
        [Test]
        public void Foo()
        {
            var bar = new ProjectSettings();
            //bar.
        }
    }
}
