using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Settings;

namespace RubberduckTests.Settings
{
    [TestClass]
    public class HotkeySettingsTests
    {
        [TestCategory("Settings")]
        [TestMethod]
        public void DefaultsSetInCtor()
        {
            var expected  = new HotkeySetting[0];

            // TODO: Use costructor with parameter
            var settings = new HotkeySettings();
            var actual = settings.Settings;

            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Ignore]
        [TestCategory("Settings")]
        [TestMethod]
        public void InvalidSettingNameWontAdd()
        {
            var settings = new HotkeySettings();
            var expected = settings.Settings;

            settings.Settings = new[] { new HotkeySetting { Name = "Foobar", IsEnabled = false, Key1 = "CTRL-C" } };

            var actual = settings.Settings;

            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Ignore]
        [TestCategory("Settings")]
        [TestMethod]
        public void InvalidSettingKeyWontAdd()
        {
            var settings = new HotkeySettings();
            var expected = settings.Settings;

            settings.Settings = new[] { new HotkeySetting { Name = "ParseAll", IsEnabled = false, Key1 = "Foobar" } };

            var actual = settings.Settings;

            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Ignore]
        [TestCategory("Settings")]
        [TestMethod]
        public void DuplicateKeysAreDeactivated()
        {
            var duplicate1 = new HotkeySetting { Name = "ParseAll", IsEnabled = true, Key1 = "X" };
            var duplicate2 = new HotkeySetting { Name = "FindSymbol", IsEnabled = true, Key1 = "X" };

            // ReSharper disable once UnusedVariable
            var settings = new HotkeySettings
            {
                Settings = new[] { duplicate1, duplicate2 }
            };

            Assert.IsFalse(duplicate1.IsEnabled == duplicate2.IsEnabled);
        }

        [Ignore]
        [TestCategory("Settings")]
        [TestMethod]
        public void DuplicateNamesAreIgnored()
        {
            var expected = new HotkeySetting { Name = "ParseAll", IsEnabled = true, Key1 = "X" };
            var duplicate = new HotkeySetting { Name = "ParseAll", IsEnabled = true, Key1 = "Y" };

            var settings = new HotkeySettings
            {
                Settings = new[] { expected, duplicate }
            };

            Assert.IsFalse(settings.Settings.Contains(duplicate));
        }
    }
}
