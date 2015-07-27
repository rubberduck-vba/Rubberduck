using Microsoft.Vbe.Interop;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Navigations.RegexSearchReplace;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;
using RubberduckTests.Mocks;

namespace RubberduckTests.RegexSearchReplace
{
    [TestClass]
    public class RegexSearchReplaceTests : VbeTestBase
    {
        [TestMethod]
        public void RegexSearchReplace_PrivateFunction()
        {
            const string inputCode = @"
Private Sub Foo()
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode)
                .Build().Object;

            var codePaneFactory = new RubberduckCodePaneFactory();
            var parseResult = new RubberduckParser(codePaneFactory).Parse(project);

            var regexSearchReplace = new Rubberduck.Navigations.RegexSearchReplace.RegexSearchReplace(parseResult);
            var results = regexSearchReplace.Search("Foo");

            //assert
            Assert.AreEqual(1, results.Count);
        }
    }
}
