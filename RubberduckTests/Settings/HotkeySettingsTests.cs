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
            var expected = new []
            {
                new HotkeySetting {CommandTypeName = "FooCommand", Key1 = "F"},
                new HotkeySetting {CommandTypeName = "BarCommand", Key1 = "B"}
            };

            var settings = new HotkeySettings(new []
            {
                new HotkeySetting {CommandTypeName = "FooCommand", Key1 = "F"},
                new HotkeySetting {CommandTypeName = "BarCommand", Key1 = "B"}
            });

            var actual = settings.Settings;

            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [TestCategory("Settings")]
        [TestMethod]
        public void InvalidSettingNameWontAdd()
        {
            var settings = new HotkeySettings(new[]
            {
                new HotkeySetting {CommandTypeName = "FooCommand", Key1 = "F"}
            });

            var expected = settings.Settings;

            settings.Settings = new[]
                {new HotkeySetting {CommandTypeName = "BarCommand", IsEnabled = false, Key1 = "CTRL-C"}};

            var actual = settings.Settings;

            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [TestCategory("Settings")]
        [TestMethod]
        public void InvalidSettingKeyWontAdd()
        {
            var settings = new HotkeySettings(new[]
            {
                new HotkeySetting {CommandTypeName = "FooCommand", Key1 = "F"}
            });

            var expected = settings.Settings;

            settings.Settings = new[]
                {new HotkeySetting {CommandTypeName = "FooCommand", IsEnabled = false, Key1 = "Foobar"}};

            var actual = settings.Settings;

            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [TestCategory("Settings")]
        [TestMethod]
        public void DuplicateKeysAreDeactivated()
        {
            var duplicate1 = new HotkeySetting {CommandTypeName = "FooCommand", IsEnabled = true, Key1 = "X"};
            var duplicate2 = new HotkeySetting {CommandTypeName = "BarCommand", IsEnabled = true, Key1 = "X"};

            // ReSharper disable once UnusedVariable
            var settings = new HotkeySettings
            {
                Settings = new[] { duplicate1, duplicate2 }
            };

            Assert.IsFalse(duplicate1.IsEnabled == duplicate2.IsEnabled);
        }

        [TestCategory("Settings")]
        [TestMethod]
        public void DuplicateNamesAreIgnored()
        {
            var expected = new HotkeySetting {CommandTypeName = "FooCommand", IsEnabled = true, Key1 = "X"};
            var duplicate = new HotkeySetting {CommandTypeName = "FooCommand", IsEnabled = true, Key1 = "Y"};

            var settings = new HotkeySettings
            {
                Settings = new[] { expected, duplicate }
            };

            Assert.IsFalse(settings.Settings.Contains(duplicate));
        }
    }
}
