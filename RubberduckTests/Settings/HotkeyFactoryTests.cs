using System;
using NUnit.Framework;
using Moq;
using Rubberduck.Common.Hotkeys;
using Rubberduck.Settings;
using Rubberduck.UI.Command;

namespace RubberduckTests.Settings
{
    [TestFixture]
    public class HotkeyFactoryTests
    {
        [Test]
        [Category("Hotkeys")]
        public void CreatingHotkeyReturnsNullWhenNoSettingProvided()
        {
            var factory = new HotkeyFactory(null);

            var hotkey = factory.Create(null, IntPtr.Zero);

            Assert.IsNull(hotkey);
        }

        [Test]
        [Category("Hotkeys")]
        public void CreatingHotkeyReturnsNullWhenNoMatchingCommandExists()
        {
            var mockCommand = new Mock<CommandBase>(null).Object;
            var factory = new HotkeyFactory(new[] {mockCommand});
            var setting = new HotkeySetting { CommandTypeName = "Foo" };

            var hotkey = factory.Create(setting, IntPtr.Zero);

            Assert.IsNull(hotkey);
        }

        [Test]
        [Category("Hotkeys")]
        public void CreatingHotkeyReturnsCorrectResult()
        {
            var mockCommand = new Mock<CommandBase>(null).Object;
            var factory = new HotkeyFactory(new[] {mockCommand});
            var setting = new HotkeySetting
            {
                CommandTypeName = mockCommand.GetType().Name,
                Key1 = "X",
                HasCtrlModifier = true
            };

            var hotkey = factory.Create(setting, IntPtr.Zero);

            Assert.Multiple(() => {
                Assert.AreEqual(mockCommand, hotkey.Command);
                Assert.AreEqual(setting.ToString(), hotkey.Key);
            });
        }
    }
}
