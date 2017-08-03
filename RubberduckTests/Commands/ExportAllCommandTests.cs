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
        private const string _path = @"C:\Users\Rubberduck\Desktop\ExportAll";

        [TestCategory("Commands")]
        [TestMethod]
        public void ExportAllCommand_SingleProject_ExpectExecution()
        {
            var builder = new MockVbeBuilder();

            var projectMock = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, "")
                .AddComponent("ClassModule1", ComponentType.ClassModule, "")
                .AddComponent("Document1", ComponentType.Document, "")
                .AddComponent("UserForm1", ComponentType.UserForm, "");

            var project = projectMock.Build();
            project.Setup(m => m.SaveAs(@"C:\Project.xlsm")).Returns(project.Object);

            project.Object.SaveAs(@"C:\Users\Rubberduck\Documents\Subfolder\Project.xlsm");
            var result = project.Object.IsSaved;
           
            var vbe = builder.AddProject(project).Build();
            vbe.Setup(m => m.ActiveVBProject.SaveAs(@"C:\Users\Rubberduck\Documents\Subfolder\Project.xlsm"));
            //project.Object.FileName = @"C:\Users\Rubberduck\Documents\Subfolder\Project.xlsm";
            var mockFolderBrowser = new Mock<IFolderBrowser>();
            mockFolderBrowser.Setup(m => m.SelectedPath).Returns(_path);
            mockFolderBrowser.Setup(m => m.ShowDialog()).Returns(DialogResult.OK);

            var mockFolderBrowserFactory = new Mock<IFolderBrowserFactory>();
            mockFolderBrowserFactory.Setup(m => m.CreateFolderBrowser(It.IsAny<string>(), true, @"C:\Users\Rubberduck\Documents\Subfolder\Project.xlsm")).Returns(mockFolderBrowser.Object);
            project.Setup(m => m.ExportSourceFiles(_path));

            var ExportAllCommand = new ExportAllCommand(vbe.Object, mockFolderBrowserFactory.Object);

            ExportAllCommand.Execute(project.Object);

            project.Verify(m => m.ExportSourceFiles(_path), Times.Once);
        }

        [TestCategory("Commands")]
        [TestMethod]
        public void ExportAllCommand_SingleProject_BrowserCanceled_ExpectNoExecution()
        {
            var builder = new MockVbeBuilder();

            var projectMock = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, "")
                .AddComponent("ClassModule1", ComponentType.ClassModule, "")
                .AddComponent("Document1", ComponentType.Document, "")
                .AddComponent("UserForm1", ComponentType.UserForm, "");

            var project = projectMock.Build();
            var vbe = builder.AddProject(project).Build();

            var mockFolderBrowserFactory = new Mock<IFolderBrowserFactory>();
            var mockFolderBrowser = new Mock<IFolderBrowser>();

            mockFolderBrowser.Setup(m => m.SelectedPath).Returns(_path);
            mockFolderBrowser.Setup(m => m.ShowDialog()).Returns(DialogResult.Cancel);
            mockFolderBrowserFactory.Setup(m => m.CreateFolderBrowser(It.IsAny<string>(), true, It.IsAny<string>())).Returns(mockFolderBrowser.Object);
            project.Setup(m => m.ExportSourceFiles(_path));

            var ExportAllCommand = new ExportAllCommand(vbe.Object, mockFolderBrowserFactory.Object);

            ExportAllCommand.Execute(project.Object);

            project.Verify(m => m.ExportSourceFiles(_path), Times.Never);
        }

        [TestCategory("Commands")]
        [TestMethod]
        public void ExportAllCommand_MultipleProjects_ExpectExecution()
        {
            var builder = new MockVbeBuilder();

            var projectMock1 = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, "")
                .AddComponent("ClassModule1", ComponentType.ClassModule, "")
                .AddComponent("Document1", ComponentType.Document, "")
                .AddComponent("UserForm1", ComponentType.UserForm, "");

            var projectMock2 = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, "")
                .AddComponent("ClassModule1", ComponentType.ClassModule, "")
                .AddComponent("Document1", ComponentType.Document, "")
                .AddComponent("UserForm1", ComponentType.UserForm, "");

            var project1 = projectMock1.Build();
            var project2 = projectMock2.Build();
            var vbe = builder.AddProject(project1).Build();
            vbe = builder.AddProject(project2).Build();

            var mockFolderBrowserFactory = new Mock<IFolderBrowserFactory>();
            var mockFolderBrowser = new Mock<IFolderBrowser>();

            mockFolderBrowser.Setup(m => m.SelectedPath).Returns(_path);
            mockFolderBrowser.Setup(m => m.ShowDialog()).Returns(DialogResult.OK);
            mockFolderBrowserFactory.Setup(m => m.CreateFolderBrowser(It.IsAny<string>(), true, It.IsAny<string>())).Returns(mockFolderBrowser.Object);

            // Can't seem to activate project1 in the mock VBE, but the second project will be active
            project2.Setup(m => m.ExportSourceFiles(_path));

            var ExportAllCommand = new ExportAllCommand(vbe.Object, mockFolderBrowserFactory.Object);

            ExportAllCommand.Execute(project2.Object);

            project1.Verify(m => m.ExportSourceFiles(_path), Times.Never);
            project2.Verify(m => m.ExportSourceFiles(_path), Times.Once);
        }
    }
}