using System;
using Moq;
using Rubberduck.Common;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.UnitTesting;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace RubberduckTests.UnitTesting
{
    internal class MockedTestExplorer : IDisposable
    {
        public MockedTestExplorer(MockedTestExplorerModel model)
        {
            Vbe = model.Engine.Vbe.Object;
            State = model.Engine.ParserState;
            Model = model.Model;
            ViewModel = new TestExplorerViewModel(null, Model, ClipboardWriter.Object, null, null, null, null);
        }

        public RubberduckParserState State { get; set; }

        public IVBE Vbe { get; }

        public TestExplorerViewModel ViewModel { get; set; }

        public TestExplorerModel Model { get; set; }

        public Mock<IClipboardWriter> ClipboardWriter { get; } = new Mock<IClipboardWriter>();

        public void Dispose()
        {
            Model?.Dispose();
            ViewModel?.Dispose();
        }
    }
}
