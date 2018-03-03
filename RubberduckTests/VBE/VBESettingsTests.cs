using NUnit.Framework;
using Rubberduck.VBERuntime;
using RubberduckTests.Mocks;

namespace RubberduckTests.VBE
{
    [TestFixture]
    public class VBESettingsTests
    {
        [Category("VBE")]
        [Test]
        public void DllVersion_MustBe6()
        {
            var vbe = new MockVbeBuilder().Build();
            vbe.SetupGet(s => s.Version).Returns("6.00");
            var settings = new VBESettings(vbe.Object);

            Assert.IsTrue(settings.Version == VBESettings.DllVersion.Vbe6);
        }

        [Category("VBE")]
        [Test]
        public void DllVersion_MustBe7()
        {
            var vbe = new MockVbeBuilder().Build();
            vbe.SetupGet(s => s.Version).Returns("7.00");
            var settings = new VBESettings(vbe.Object);

            Assert.IsTrue(settings.Version == VBESettings.DllVersion.Vbe7);
        }
        
        [Category("VBE")]
        [Test]
        public void DllVersion_IsBogus()
        {
            var vbe = new MockVbeBuilder().Build();
            vbe.SetupGet(s => s.Version).Returns("foo");
            var settings = new VBESettings(vbe.Object);

            Assert.IsTrue(settings.Version == VBESettings.DllVersion.Unknown);
        }

        [Category("VBE")]
        [Test]
        public void DllVersion_IsNull()
        {
            var vbe = new MockVbeBuilder().Build();
            vbe.SetupGet(s => s.Version).Returns((string)null);
            var settings = new VBESettings(vbe.Object);

            Assert.IsTrue(settings.Version == VBESettings.DllVersion.Unknown);
        }

        [Category("VBE")]
        [Test]
        public void CompileOnDemand_WriteRead()
        {
            var vbe = new MockVbeBuilder().Build();
            vbe.SetupGet(s => s.Version).Returns("7.00");
            var settings = new VBESettings(vbe.Object);

            settings.CompileOnDemand = true;
            Assert.IsTrue(settings.CompileOnDemand);

            settings.CompileOnDemand = false;
            Assert.IsTrue(settings.CompileOnDemand == false);
        }

        [Category("VBE")]
        [Test]
        public void BackGroundCompile_WriteRead()
        {
            var vbe = new MockVbeBuilder().Build();
            vbe.SetupGet(s => s.Version).Returns("7.00");
            var settings = new VBESettings(vbe.Object);

            settings.BackGroundCompile = true;
            Assert.IsTrue(settings.BackGroundCompile);

            settings.BackGroundCompile = false;
            Assert.IsTrue(settings.BackGroundCompile == false);
        }
    }
}
