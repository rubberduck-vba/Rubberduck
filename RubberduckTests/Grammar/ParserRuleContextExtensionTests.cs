using System.Linq;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Application;
using Rubberduck.VBEditor.Events;
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

            //Arrange
            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var declaration = parser.State.AllDeclarations.Single(d => d.IdentifierName.Equals("Foo"));

            var actual = ((VBAParser.FunctionStmtContext)declaration.Context).GetProcedureSelection();
            var expected = new Selection(3, 2, 11, 14);

            Assert.AreEqual(actual, expected);
        }
    }
}
