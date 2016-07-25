using System;
using System.Threading;
using Microsoft.Vbe.Interop;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Parsing;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command;
using Rubberduck.UnitTesting;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.VBEHost;
using RubberduckTests.Mocks;

namespace RubberduckTests.Commands
{
    [TestClass]
    public class AddTestMethodCommandTests
    {
        [TestMethod]
        public void AddsTest()
        {
            var input =
                @"Option Explicit
Option Private Module


'@TestModule
Private Assert As Object
{0}";

            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(string.Format(input, string.Empty), out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(new Mock<ISinks>().Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error)
            {
                Assert.Inconclusive("Parser Error");
            }

            var newTestMethodCommand = new Mock<NewTestMethodCommand>(vbe.Object, parser.State);
            var addTestMethodCommand = new AddTestMethodCommand(vbe.Object, parser.State, newTestMethodCommand.Object);

            addTestMethodCommand.Execute(null);

            Assert.AreEqual(
                string.Format(input,
                    NewTestMethodCommand.TestMethodTemplate.Replace(NewTestMethodCommand.NamePlaceholder, "TestMethod1")) +
                Environment.NewLine, vbe.Object.ActiveCodePane.CodeModule.Lines());
        }

        [TestMethod]
        public void AddsExpectedErrorTest()
        {
            var input =
@"Option Explicit
Option Private Module


'@TestModule
Private Assert As Object
{0}";

            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(string.Format(input, string.Empty), out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(new Mock<ISinks>().Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var newTestMethodCommand = new Mock<NewTestMethodCommand>(vbe.Object, parser.State);
            var addTestMethodCommand = new AddTestMethodExpectedErrorCommand(vbe.Object, parser.State, newTestMethodCommand.Object);

            addTestMethodCommand.Execute(null);

            Assert.AreEqual(
                string.Format(input,
                    NewTestMethodCommand.TestMethodExpectedErrorTemplate.Replace(NewTestMethodCommand.NamePlaceholder,
                        "TestMethod1")) + Environment.NewLine, vbe.Object.ActiveCodePane.CodeModule.Lines());
        }
    }
}