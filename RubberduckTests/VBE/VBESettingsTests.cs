using Microsoft.Win32;
using Moq;
using NUnit.Framework;
using Rubberduck.VBEditor.Utility;
using Rubberduck.VBEditor.VbeRuntime;
using Rubberduck.VBEditor.VbeRuntime.Settings;
using RubberduckTests.Mocks;

namespace RubberduckTests.VBE
{
    [TestFixture]
    public class VbeSettingsTests
    {
        private const string Vbe7SettingPath = @"HKEY_CURRENT_USER\Software\Microsoft\VBA\7.0\Common";
        private const string Vbe6SettingPath = @"HKEY_CURRENT_USER\Software\Microsoft\VBA\6.0\Common";
        private const int DWordFalseValue = 0;
        private const int DWordTrueValue = 1;

        private Mock<IRegistryWrapper> GetRegistryMock()
        {
            return new Mock<IRegistryWrapper>();
        }

        [Category("VBE")]
        [Test]
        public void DllVersion_MustBe6()
        {
            var vbe = new MockVbeBuilder().Build();
            var registry = GetRegistryMock();

            vbe.SetupGet(s => s.Version).Returns("6.00");
            var settings = new VbeSettings(vbe.Object, registry.Object);

            Assert.AreEqual(DllVersion.Vbe6, settings.Version);
        }

        [Category("VBE")]
        [Test]
        public void DllVersion_MustBe7()
        {
            var vbe = new MockVbeBuilder().Build();
            var registry = GetRegistryMock();

            vbe.SetupGet(s => s.Version).Returns("7.00");
            var settings = new VbeSettings(vbe.Object, registry.Object);

            Assert.AreEqual(DllVersion.Vbe7, settings.Version);
        }
        
        [Category("VBE")]
        [Test]
        public void DllVersion_IsBogus()
        {
            var vbe = new MockVbeBuilder().Build();
            var registry = GetRegistryMock();

            vbe.SetupGet(s => s.Version).Returns("foo");
            var settings = new VbeSettings(vbe.Object, registry.Object);

            Assert.AreEqual(DllVersion.Unknown, settings.Version);
        }

        [Category("VBE")]
        [Test]
        public void DllVersion_IsNull()
        {
            var vbe = new MockVbeBuilder().Build();
            var registry = GetRegistryMock();

            vbe.SetupGet(s => s.Version).Returns((string)null);
            var settings = new VbeSettings(vbe.Object, registry.Object);

            Assert.IsTrue(settings.Version == DllVersion.Unknown);
        }

        [Category("VBE")]
        [Test]
        public void CompileOnDemand_Write_IsTrue()
        {
            var vbe = new MockVbeBuilder().Build();
            var registry = GetRegistryMock();

            vbe.SetupGet(s => s.Version).Returns("7.00");
            registry.Setup(s => s.SetValue(Vbe7SettingPath, "CompileOnDemand", true, RegistryValueKind.DWord));
            registry.Setup(s => s.GetValue(Vbe7SettingPath, "CompileOnDemand", DWordFalseValue)).Returns(DWordTrueValue);

            var settings = new VbeSettings(vbe.Object, registry.Object);

            settings.CompileOnDemand = true;
            Assert.IsTrue(settings.CompileOnDemand);
        }

        [Category("VBE")]
        [Test]
        public void BackGroundCompile_Write_IsFalse()
        {
            var vbe = new MockVbeBuilder().Build();
            var registry = GetRegistryMock();

            vbe.SetupGet(s => s.Version).Returns("7.00");
            registry.Setup(s => s.SetValue(Vbe7SettingPath, "BackGroundCompile", false, RegistryValueKind.DWord));
            registry.Setup(s => s.GetValue(Vbe7SettingPath, "BackGroundCompile", DWordFalseValue)).Returns(DWordFalseValue);

            var settings = new VbeSettings(vbe.Object, registry.Object);
            
            settings.BackGroundCompile = false;
            Assert.IsTrue(settings.BackGroundCompile == false);
        }
    }
}
