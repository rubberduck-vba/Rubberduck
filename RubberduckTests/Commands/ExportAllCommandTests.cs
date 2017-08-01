using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Navigation.Folders;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.SmartIndenter;
using Rubberduck.UI;
using Rubberduck.UI.Command;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;
using System.Collections.Generic;
using System.Threading;
using System.Windows.Forms;

namespace RubberduckTests.Commands
{


    [TestClass]
    public class ExportAllTests
    {
        private Mock<IFolderBrowserFactory> _folderBrowserFactory;
        private Mock<IFolderBrowser> _folderBrowser;

        [TestInitialize]
        public void Initalize()
        {
            _folderBrowser = new Mock<IFolderBrowser>();
            _folderBrowserFactory = new Mock<IFolderBrowserFactory>();
            _folderBrowserFactory.Setup(f => f.CreateFolderBrowser(It.IsAny<string>())).Returns(_folderBrowser.Object);
            _folderBrowserFactory.Setup(f => f.CreateFolderBrowser(It.IsAny<string>(), false)).Returns(_folderBrowser.Object);
        }    

        [TestCategory("Commands")]
        [TestMethod]
        public void ExportAllModule()
        {
            var builder = new MockVbeBuilder();
            var projectMock = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, "")
                .AddComponent("ClassModule1", ComponentType.ClassModule, "")
                .AddComponent("Document1", ComponentType.Document, "")
                .AddComponent("UserForm1", ComponentType.UserForm, "");

            var project = projectMock.Build();
            var vbe = builder.AddProject(project).Build();
            //var component = projectMock.MockComponents;
            _folderBrowser.Object.SelectedPath = @"C:\Users\Rubberduck\Desktop\ExportAll\";
            _folderBrowser.Setup(o => o.ShowDialog()).Returns(DialogResult.OK);

            var state = new RubberduckParserState(vbe.Object, new DeclarationFinderFactory());
            //var commands = new List<CommandBase>
            //{
            //    new ExportAllCommand(vbe.Object, folderBrowser.Object)
            //};

            var ExportAllCommand = new ExportAllCommand(vbe.Object, _folderBrowser.Object);
            //var vm = new CodeExplorerViewModel(new FolderHelper(state), state, commands);

            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            //vm.SelectedItem = vm.Projects.First().Items.First().Items.First();
            ExportAllCommand.Execute(project.Object);

            project.Verify(c => c.ExportSourceFiles(@"C:\Users\Rubberduck\Desktop\ExportAll\"), Times.Once);
        }

        //[TestCategory("Commands")]
        //[TestMethod]
        //public void ExportAllModule_Cancel()
        //{
        //    var builder = new MockVbeBuilder();
        //    var projectMock = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
        //        .AddComponent("Module1", ComponentType.StandardModule, "")
        //        .AddComponent("UserForm1", ComponentType.UserForm, "");

        //    var project = projectMock.Build();
        //    var vbe = builder.AddProject(project).Build();
        //    //var component = projectMock.MockComponents.First();

        //    var folderBrowser = new Mock<IFolderBrowserFactory>();
        //    //folderBrowser.Setup(o => o.SelectedPath).Returns("C:\\Users\\Rubberduck\\Desktop\\ExportAll\\");
        //    //folderBrowser.Setup(o => o.ShowDialog()).Returns(DialogResult.Cancel);

        //    var state = new RubberduckParserState(vbe.Object, new DeclarationFinderFactory());
        //    var commands = new List<CommandBase>
        //    {
        //        new ExportAllCommand(folderBrowser.Object)
        //    };

        //    var vm = new CodeExplorerViewModel(new FolderHelper(state), state, commands);

        //    var parser = MockParser.Create(vbe.Object, state);
        //    parser.Parse(new CancellationTokenSource());
        //    if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

        //    vm.SelectedItem = vm.Projects.First().Items.First().Items.First();
        //    vm.ExportCommand.Execute(vm.SelectedItem);

        //    project.Verify(c => c.ExportSourceFiles("C:\\Users\\Rubberduck\\Desktop\\ExportAll\\"), Times.Never);
        //}
    }
}