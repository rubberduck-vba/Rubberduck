using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;

namespace RubberduckTests.Grammar
{
    //https://github.com/rubberduck-vba/Rubberduck/issues/2164
    [TestClass]
    public class ParserRuleContextExtensionTests
    {
        [TestMethod]
        [TestCategory("Inspections")]
        public void Evil_Code_Selection_Not_Evil()
        {
            const string inputCode =
@" _
 _
 Function _
 _
 Foo _
 _
 ( _
 _
 fizz _
 _
 ) As Boolean

 End _
 _
 Function";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);

            var state = MockParser.CreateAndParse(vbe.Object);

            var declaration = state.AllDeclarations.Single(d => d.IdentifierName.Equals("Foo"));

            var actual = ((VBAParser.FunctionStmtContext)declaration.Context).GetProcedureSelection();
            var expected = new Selection(3, 2, 11, 14);

            Assert.AreEqual(actual, expected);
        }
    }
}
