using System.Collections.Generic;
using System.Collections.ObjectModel;
using Moq;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace RubberduckTests.Mocks
{
    public class MockAddInBuilder
    {
        private readonly Mock<IAddIn> _addIn;

        public MockAddInBuilder()
        {
            _addIn = CreateAddInMock();
        }

        private Mock<IAddIn> CreateAddInMock()
        {
            var addIn = new Mock<IAddIn>();

            addIn.Setup(a => a.CommandBarLocations).Returns(new ReadOnlyDictionary<CommandBarSite, CommandBarLocation>(new Dictionary<CommandBarSite, CommandBarLocation>
            {
                {CommandBarSite.MenuBar, new CommandBarLocation(1, 1)},
                {CommandBarSite.CodePaneContextMenu, new CommandBarLocation(2, 2)},
                {CommandBarSite.ProjectExplorerContextMenu, new CommandBarLocation(3, 3)},
                {CommandBarSite.FormDesignerContextMenu, new CommandBarLocation(4, 4)},
                {CommandBarSite.FormDesignerControlContextMenu, new CommandBarLocation(5, 5)}
            }));

            return addIn;
        }

        public Mock<IAddIn> Build()
        {
            return _addIn;
        }
    }
}
