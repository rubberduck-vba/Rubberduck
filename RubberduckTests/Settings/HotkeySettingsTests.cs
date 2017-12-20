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
            var expected = new[] 
            {
                new HotkeySetting{Name=RubberduckHotkey.ParseAll.ToString(), IsEnabled=true, HasCtrlModifier = true, Key1="`" },
                new HotkeySetting{Name=RubberduckHotkey.IndentProcedure.ToString(), IsEnabled=true, HasCtrlModifier = true, Key1="P" },
                new HotkeySetting{Name=RubberduckHotkey.IndentModule.ToString(), IsEnabled=true, HasCtrlModifier = true, Key1="M" },
                new HotkeySetting{Name=RubberduckHotkey.CodeExplorer.ToString(), IsEnabled=true, HasCtrlModifier = true, Key1="R" },
                new HotkeySetting{Name=RubberduckHotkey.FindSymbol.ToString(), IsEnabled=true, HasCtrlModifier = true, Key1="T" },
                new HotkeySetting{Name=RubberduckHotkey.InspectionResults.ToString(), IsEnabled=true, HasCtrlModifier = true, HasShiftModifier = true, Key1="I" },
                new HotkeySetting{Name=RubberduckHotkey.TestExplorer.ToString(), IsEnabled=true, HasCtrlModifier = true, HasShiftModifier = true, Key1="T" },
                new HotkeySetting{Name=RubberduckHotkey.RefactorMoveCloserToUsage.ToString(), IsEnabled=true, HasCtrlModifier = true, HasShiftModifier = true, Key1="C" },
                new HotkeySetting{Name=RubberduckHotkey.RefactorRename.ToString(), IsEnabled=true, HasCtrlModifier = true, HasShiftModifier = true, Key1="R" },
                new HotkeySetting{Name=RubberduckHotkey.RefactorExtractMethod.ToString(), IsEnabled=true, HasCtrlModifier = true, HasShiftModifier = true, Key1="M" },
                new HotkeySetting{Name=RubberduckHotkey.SourceControl.ToString(), IsEnabled=true, HasCtrlModifier = true, HasShiftModifier = true, Key1="D6" },
                new HotkeySetting{Name=RubberduckHotkey.RefactorEncapsulateField.ToString(), IsEnabled=true, HasCtrlModifier = true, HasShiftModifier = true, Key1="F" },
                new HotkeySetting{Name=RubberduckHotkey.ExportActiveProject.ToString(), IsEnabled = true, HasCtrlModifier = true, HasShiftModifier = true, Key1="E" }
            };

            var settings = new HotkeySettings();
            var actual = settings.Settings;

            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [TestCategory("Settings")]
        [TestMethod]
        public void InvalidSettingNameWontAdd()
        {
            var settings = new HotkeySettings();
            var expected = settings.Settings;

            settings.Settings = new[] {new HotkeySetting {Name = "Foobar", IsEnabled = false, Key1 = "CTRL-C"}};

            var actual = settings.Settings;

            Assert.IsTrue(expected.SequenceEqual(actual));
        }

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

        [TestCategory("Settings")]
        [TestMethod]
        public void DuplicateKeysAreDeactivated()
        {
            var duplicate1 = new HotkeySetting {Name = "ParseAll", IsEnabled = true, HasAltModifier = true, Key1 = "X"};
            var duplicate2 = new HotkeySetting {Name = "FindSymbol", IsEnabled = true, HasAltModifier = true, Key1 = "X"};

            // ReSharper disable once UnusedVariable
            var settings = new HotkeySettings
            {
                Settings = new[] {duplicate1, duplicate2}
            };

            Assert.IsFalse(duplicate1.IsEnabled == duplicate2.IsEnabled);
        }

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

        [TestMethod]
        [TestCategory("Settings)")]
        public void HotkeyModifier_None_IsNotValid()
        {
            var setting = new HotkeySetting { Name = "Foo", Key1 = "I" };

            Assert.IsFalse(setting.IsValid);
        }

        [TestMethod]
        [TestCategory("Settings)")]
        public void HotkeyModifier_Shift_IsNotValid()
        {
            var setting = new HotkeySetting { Name = "Foo", HasShiftModifier = true, Key1 = "I" };

            Assert.IsFalse(setting.IsValid);
        }

        [TestMethod]
        [TestCategory("Settings)")]
        public void HotkeyModifier_ShiftAlt_IsValid()
        {
            var setting = new HotkeySetting { Name = "Foo", HasShiftModifier = true, HasAltModifier = true, Key1 = "I" };

            Assert.IsTrue(setting.IsValid);
        }

        [TestMethod]
        [TestCategory("Settings)")]
        public void HotkeyModifier_ShiftCtrl_IsValid()
        {
            var setting = new HotkeySetting { Name = "Foo", HasShiftModifier = true, HasCtrlModifier = true, Key1 = "I" };

            Assert.IsTrue(setting.IsValid);
        }

        [TestMethod]
        [TestCategory("Settings)")]
        public void HotkeyModifier_ShiftAltCtrl_IsValid()
        {
            var setting = new HotkeySetting { Name = "Foo", HasShiftModifier = true, HasAltModifier = true, HasCtrlModifier = true, Key1 = "I" };

            Assert.IsTrue(setting.IsValid);
        }

        [TestMethod]
        [TestCategory("Settings)")]
        public void HotkeyModifier_Alt_IsValid()
        {
            var setting = new HotkeySetting { Name = "Foo", HasAltModifier = true, Key1 = "I" };

            Assert.IsTrue(setting.IsValid);
        }

        [TestMethod]
        [TestCategory("Settings)")]
        public void HotkeyModifier_Ctrl_IsValid()
        {
            var setting = new HotkeySetting { Name = "Foo", HasCtrlModifier = true, Key1 = "I" };

            Assert.IsTrue(setting.IsValid);
        }

        [TestMethod]
        [TestCategory("Settings)")]
        public void HotkeyModifier_AltCtrl_IsValid()
        {
            var setting = new HotkeySetting { Name = "Foo", HasAltModifier = true, HasCtrlModifier = true, Key1 = "I" };

            Assert.IsTrue(setting.IsValid);
        }
    }
}
