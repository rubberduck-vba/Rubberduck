using Microsoft.Vbe.Interop;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.VBEditor.Extensions;

namespace RubberduckTests
{
    [TestClass]
    public class ComponentTypeExtensionTests
    {
        [TestMethod, Timeout(1000)]
        public void ClassReturnsCls()
        {
            var type = vbext_ComponentType.vbext_ct_ClassModule;

            Assert.AreEqual(".cls", type.FileExtension());
        }

        [TestMethod, Timeout(1000)]
        public void FormReturnsFrm()
        {
            var type = vbext_ComponentType.vbext_ct_MSForm;
            Assert.AreEqual(".frm", type.FileExtension());
        }

        [TestMethod, Timeout(1000)]
        public void StandardReturnsBas()
        {
            var type = vbext_ComponentType.vbext_ct_StdModule;
            Assert.AreEqual(".bas", type.FileExtension());
        }
    }
}
