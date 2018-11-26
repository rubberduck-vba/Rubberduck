using System.Linq;
using NUnit.Framework;
using Rubberduck.Settings;

namespace RubberduckTests.Settings
{
    [TestFixture]
    public class HotkeySettingsTests
    {
        [Category("Settings")]
        [Test]
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

        [Category("Settings")]
        [Test]
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

        [Category("Settings")]
        [Test]
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

        [Category("Settings")]
        [Test]
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

        [Category("Settings")]
        [Test]
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

        [Category("Settings")]
        [Test]
        public void HotkeyModifier_None_IsNotValid()
        {
            var setting = new HotkeySetting { CommandTypeName = "Foo", Key1 = "I" };

            Assert.IsFalse(setting.IsValid);
        }

        [Category("Settings")]
        [Test]
        public void HotkeyModifier_Shift_IsNotValid()
        {
            var setting = new HotkeySetting { CommandTypeName = "Foo", HasShiftModifier = true, Key1 = "I" };

            Assert.IsFalse(setting.IsValid);
        }

        [Category("Settings")]
        [Test]
        public void HotkeyModifier_ShiftAlt_IsValid()
        {
            var setting = new HotkeySetting { CommandTypeName = "Foo", HasShiftModifier = true, HasAltModifier = true, Key1 = "I" };

            Assert.IsTrue(setting.IsValid);
        }

        [Category("Settings")]
        [Test]
        public void HotkeyModifier_ShiftCtrl_IsValid()
        {
            var setting = new HotkeySetting { CommandTypeName = "Foo", HasShiftModifier = true, HasCtrlModifier = true, Key1 = "I" };

            Assert.IsTrue(setting.IsValid);
        }

        [Category("Settings")]
        [Test]
        public void HotkeyModifier_ShiftAltCtrl_IsValid()
        {
            var setting = new HotkeySetting { CommandTypeName = "Foo", HasShiftModifier = true, HasAltModifier = true, HasCtrlModifier = true, Key1 = "I" };

            Assert.IsTrue(setting.IsValid);
        }

        [Category("Settings")]
        [Test]
        public void HotkeyModifier_Alt_IsValid()
        {
            var setting = new HotkeySetting { CommandTypeName = "Foo", HasAltModifier = true, Key1 = "I" };

            Assert.IsTrue(setting.IsValid);
        }

        [Category("Settings")]
        [Test]
        public void HotkeyModifier_Ctrl_IsValid()
        {
            var setting = new HotkeySetting { CommandTypeName = "Foo", HasCtrlModifier = true, Key1 = "I" };

            Assert.IsTrue(setting.IsValid);
        }

        [Category("Settings")]
        [Test]
        public void HotkeyModifier_AltCtrl_IsValid()
        {
            var setting = new HotkeySetting { CommandTypeName = "Foo", HasAltModifier = true, HasCtrlModifier = true, Key1 = "I" };

            Assert.IsTrue(setting.IsValid);
        }
    }
}
