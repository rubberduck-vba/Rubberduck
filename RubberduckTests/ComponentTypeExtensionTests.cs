using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.Extensions;

namespace RubberduckTests
{
    [TestClass]
    public class ComponentTypeExtensionTests
    {
        [TestMethod]
        public void ClassReturnsCls()
        {
            var type = ComponentType.ClassModule;

            Assert.AreEqual(".cls", type.FileExtension());
        }

        [TestMethod]
        public void FormReturnsFrm()
        {
            var type = ComponentType.UserForm;
            Assert.AreEqual(".frm", type.FileExtension());
        }

        [TestMethod]
        public void StandardReturnsBas()
        {
            var type = ComponentType.StandardModule;
            Assert.AreEqual(".bas", type.FileExtension());
        }
    }
}
