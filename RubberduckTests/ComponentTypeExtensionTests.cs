using System;
using UnitTesting = Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Reflection;
using Microsoft.Vbe.Interop;

namespace RubberduckTests
{
    [UnitTesting.TestClass]
    public class ComponentTypeExtensionTests
    {
        [UnitTesting.TestMethod]
        public void ClassReturnsCls()
        {
            var type = vbext_ComponentType.vbext_ct_ClassModule;

            UnitTesting.Assert.AreEqual(".cls", type.FileExtension());
        }

        [UnitTesting.TestMethod]
        public void FormReturnsFrm()
        {
            var type = vbext_ComponentType.vbext_ct_MSForm;
            UnitTesting.Assert.AreEqual(".frm", type.FileExtension());
        }

        [UnitTesting.TestMethod]
        public void StandardReturnsBas()
        {
            var type = vbext_ComponentType.vbext_ct_StdModule;
            UnitTesting.Assert.AreEqual(".bas", type.FileExtension());
        }
    }
}
