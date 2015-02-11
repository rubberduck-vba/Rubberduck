using Microsoft.Vbe.Interop;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.VBA;

namespace RubberduckTests
{
    [TestClass]
    public class ComponentTypeExtensionTests
    {
        [TestMethod]
        public void ClassReturnsCls()
        {
            var type = vbext_ComponentType.vbext_ct_ClassModule;

            Assert.AreEqual(".cls", type.FileExtension());
        }

        [TestMethod]
        public void FormReturnsFrm()
        {
            var type = vbext_ComponentType.vbext_ct_MSForm;
            Assert.AreEqual(".frm", type.FileExtension());
        }

        [TestMethod]
        public void StandardReturnsBas()
        {
            var type = vbext_ComponentType.vbext_ct_StdModule;
            Assert.AreEqual(".bas", type.FileExtension());
        }
    }
}
