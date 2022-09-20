using NUnit.Framework;
using Moq;
using Rubberduck.UI;
using Rubberduck.UI.Command.ComCommands;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;
using System.Linq;
using System.Windows.Forms;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using System.Collections.Generic;
using System;
using Path = System.IO.Path;
using Directory = System.IO.Directory;

namespace RubberduckTests.Commands
{
    [TestFixture]
    public class ExportAllTests
    {
        private const string _path = @"C:\Users\Rubberduck\Desktop\ExportAll";
        private const string _projectPath = @"C:\Users\Rubberduck\Documents\Subfolder";
        private const string _projectFullPath = @"C:\Users\Rubberduck\Documents\Subfolder\Project.xlsm";
        private const string _projectFullPath2 = @"C:\Users\Rubberduck\Documents\Subfolder\Project2.xlsm";

        [Category("Commands")]
        [Category(nameof(ExportAllCommand))]
        [Test]
        public void ExportAllCommand_CanExecute_PassedNull_ExpectTrue()
        {
            var builder = new MockVbeBuilder();
            var projectMock = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected);
            var project = projectMock.Build();
            var vbe = builder.AddProject(project).Build();

            vbe.SetupGet(m => m.ActiveVBProject.VBComponents.Count).Returns(1);

            var mockFolderBrowserFactory = new Mock<IFileSystemBrowserFactory>();
            var exportAllCommand = ArrangeExportAllCommand(vbe, mockFolderBrowserFactory);

            Assert.IsTrue(exportAllCommand.CanExecute(null));
        }

        [Category("Commands")]
        [Category(nameof(ExportAllCommand))]
        [Test]
        public void ExportAllCommand_CanExecute_PassedNull_NoComponents_ExpectFalse()
        {
            var builder = new MockVbeBuilder();
            var projectMock = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected);
            var project = projectMock.Build();
            var vbe = builder.AddProject(project).Build();

            vbe.SetupGet(m => m.ActiveVBProject.VBComponents.Count).Returns(0);

            var mockFolderBrowserFactory = new Mock<IFileSystemBrowserFactory>();
            var exportAllCommand = ArrangeExportAllCommand(vbe, mockFolderBrowserFactory);

            Assert.IsFalse(exportAllCommand.CanExecute(null));
        }

        [Category("Commands")]
        [Category(nameof(ExportAllCommand))]
        [Test]
        public void ExportAllCommand_CanExecute_PassedIVBProject_ExpectTrue()
        {
            var builder = new MockVbeBuilder();
            var projectMock = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected);
            var project = projectMock.Build();

            var vbe = builder.AddProject(project).Build();

            project.SetupGet(m => m.VBComponents.Count).Returns(1);

            var mockFolderBrowserFactory = new Mock<IFileSystemBrowserFactory>();
            var exportAllCommand = ArrangeExportAllCommand(vbe, mockFolderBrowserFactory);

            Assert.IsTrue(exportAllCommand.CanExecute(vbe.Object.VBProjects.First()));
        }

        [Category("Commands")]
        [Category(nameof(ExportAllCommand))]
        [Test]
        public void ExportAllCommand_CanExecute_PassedIVBProject_NoComponents_ExpectFalse()
        {
            var builder = new MockVbeBuilder();
            var projectMock = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected);
            var project = projectMock.Build();
            var vbe = builder.AddProject(project).Build();

            vbe.SetupGet(m => m.ActiveVBProject.VBComponents.Count).Returns(0);

            var mockFolderBrowserFactory = new Mock<IFileSystemBrowserFactory>();
            var exportAllCommand = ArrangeExportAllCommand(vbe, mockFolderBrowserFactory);

            Assert.IsFalse(exportAllCommand.CanExecute(vbe.Object));
        }


        [Category("Commands")]
        [Category(nameof(ExportAllCommand))]
        [Test]
        public void ExportAllCommand_Execute_PassedNull_SingleProject_ExpectExecution()
        {
            var project = CreateTestProjectMocks("TestProject1").Values.First();

            project.SetupGet(m => m.IsSaved).Returns(true);
            project.SetupGet(m => m.FileName).Returns(_projectFullPath);

            var builder = new MockVbeBuilder();
            var vbe = builder.AddProject(project).Build();

            project.Setup(m => m.ExportSourceFiles(_path));

            var exportAllCommand 
                = ArrangeExportAllCommand(vbe, CreateMockFolderBrowserFactory(_projectPath, _path, DialogResult.OK));

            exportAllCommand.Execute(null);

            project.Verify(m => m.ExportSourceFiles(_path), Times.Once);
        }

        [Category("Commands")]
        [Category(nameof(ExportAllCommand))]
        [Test]
        public void ExportAllCommand_Execute_PassedIVBProject_SingleProject_ExpectExecution()
        {
            var project = CreateTestProjectMocks("TestProject1").Values.First();
            project.SetupGet(m => m.IsSaved).Returns(true);
            project.SetupGet(m => m.FileName).Returns(_projectFullPath);

            var builder = new MockVbeBuilder();
            var vbe = builder.AddProject(project).Build();

            project.Setup(m => m.ExportSourceFiles(_path));

            var exportAllCommand 
                = ArrangeExportAllCommand(vbe, CreateMockFolderBrowserFactory(_projectPath, _path, DialogResult.OK));

            exportAllCommand.Execute(project.Object);

            project.Verify(m => m.ExportSourceFiles(_path), Times.Once);
        }


        [Category("Commands")]
        [Category(nameof(ExportAllCommand))]
        [Test]
        public void ExportAllCommand_Execute_PassedNull_MultipleProjects_ExpectExecution()
        {
            var projects = CreateTestProjectMocks("TestProject1", "TestProject2");
            var project1 = projects["TestProject1"];
            var project2 = projects["TestProject2"];
            project1.SetupGet(m => m.IsSaved).Returns(true);
            project1.SetupGet(m => m.FileName).Returns(_projectFullPath);
            project2.SetupGet(m => m.IsSaved).Returns(true);
            project2.SetupGet(m => m.FileName).Returns(_projectFullPath2);

            var builder = new MockVbeBuilder();
            var vbe = builder
                .AddProject(project1)
                .AddProject(project2)
                .Build();

            project2.Setup(m => m.ExportSourceFiles(_path));

            var exportAllCommand
                = ArrangeExportAllCommand(vbe, CreateMockFolderBrowserFactory(_projectPath, _path, DialogResult.OK));

            exportAllCommand.Execute(null);

            // project2 added last, will be active project
            project1.Verify(m => m.ExportSourceFiles(_path), Times.Never);
            project2.Verify(m => m.ExportSourceFiles(_path), Times.Once);
        }

        [Category("Commands")]
        [Category(nameof(ExportAllCommand))]
        [Test]
        public void ExportAllCommand_Execute_PassedIVBProject_MultipleProjects_ExpectExecution()
        {
            var projects = CreateTestProjectMocks("TestProject1", "TestProject2");
            var project1 = projects["TestProject1"];
            var project2 = projects["TestProject2"];
            project1.SetupGet(m => m.IsSaved).Returns(true);
            project1.SetupGet(m => m.FileName).Returns(_projectFullPath);
            project2.SetupGet(m => m.IsSaved).Returns(true);
            project2.SetupGet(m => m.FileName).Returns(_projectFullPath2);

            var builder = new MockVbeBuilder();
            var vbe = builder
                .AddProject(project1)
                .AddProject(project2)
                .Build();

            project1.Setup(m => m.ExportSourceFiles(_path));

            var exportAllCommand
                = ArrangeExportAllCommand(vbe, CreateMockFolderBrowserFactory(_projectPath, _path, DialogResult.OK));

            exportAllCommand.Execute(project1.Object);

            project1.Verify(m => m.ExportSourceFiles(_path), Times.Once);
            project2.Verify(c => c.ExportSourceFiles(_path), Times.Never);
        }

        [Category("Commands")]
        [Category(nameof(ExportAllCommand))]
        [Test]
        public void ExportAllCommand_Execute_PassedNull_SingleProject_BrowserCanceled_ExpectNoExecution()
        {
            var project = CreateTestProjectMocks("TestProject1").Values.First();
            project.SetupGet(m => m.IsSaved).Returns(true);
            project.SetupGet(m => m.FileName).Returns(_projectFullPath);

            var builder = new MockVbeBuilder();
            var vbe = builder.AddProject(project).Build();

            project.Setup(m => m.ExportSourceFiles(_path));

            var exportAllCommand
                = ArrangeExportAllCommand(vbe, CreateMockFolderBrowserFactory(_projectPath, _path, DialogResult.Cancel));

            exportAllCommand.Execute(null);

            project.Verify(m => m.ExportSourceFiles(_path), Times.Never);
        }

        [Category("Commands")]
        [Category(nameof(ExportAllCommand))]
        [Test]
        public void ExportAllCommand_Execute_PassedIVBProject_SingleProject_BrowserCanceled_ExpectNoExecution()
        {
            var project = CreateTestProjectMocks("TestProject1").Values.First();
            project.SetupGet(m => m.IsSaved).Returns(true);
            project.SetupGet(m => m.FileName).Returns(_projectFullPath);

            var builder = new MockVbeBuilder();
            var vbe = builder.AddProject(project).Build();

            project.Setup(m => m.ExportSourceFiles(_path));

            var exportAllCommand
                = ArrangeExportAllCommand(vbe, CreateMockFolderBrowserFactory(_projectPath, _path, DialogResult.Cancel));

            exportAllCommand.Execute(project.Object);

            project.Verify(m => m.ExportSourceFiles(_path), Times.Never);
        }


        [Category("Commands")]
        [Category(nameof(ExportAllCommand))]
        [Test]
        public void ExportAllCommand_Execute_PassedNull_MultipleProjects_BrowserCanceled_ExpectNoExecution()
        {
            var projects = CreateTestProjectMocks("TestProject1", "TestProject2");
            var project1 = projects["TestProject1"];
            var project2 = projects["TestProject2"];
            project1.SetupGet(m => m.IsSaved).Returns(true);
            project1.SetupGet(m => m.FileName).Returns(_projectFullPath);
            project2.SetupGet(m => m.IsSaved).Returns(true);
            project2.SetupGet(m => m.FileName).Returns(_projectFullPath2);

            var builder = new MockVbeBuilder();
            var vbe = builder
                .AddProject(project1)
                .AddProject(project2)
                .Build();

            project2.Setup(m => m.ExportSourceFiles(_path));

            var exportAllCommand
                = ArrangeExportAllCommand(vbe, CreateMockFolderBrowserFactory(_projectPath, _path, DialogResult.Cancel));

            exportAllCommand.Execute(null);

            // project2 added last, will be active project
            project1.Verify(m => m.ExportSourceFiles(_path), Times.Never);
            project2.Verify(m => m.ExportSourceFiles(_path), Times.Never);
        }

        [Category("Commands")]
        [Category(nameof(ExportAllCommand))]
        [Test]
        public void ExportAllCommand_Execute_PassedIVBProject_MultipleProjects_BrowserCanceled_ExpectNoExecution()
        {
            var projects = CreateTestProjectMocks("TestProject1", "TestProject2");
            var project1 = projects["TestProject1"];
            var project2 = projects["TestProject2"];
            project1.SetupGet(m => m.IsSaved).Returns(true);
            project1.SetupGet(m => m.FileName).Returns(_projectFullPath);
            project2.SetupGet(m => m.IsSaved).Returns(true);
            project2.SetupGet(m => m.FileName).Returns(_projectFullPath2);

            var builder = new MockVbeBuilder();
            var vbe = builder
                .AddProject(project1)
                .AddProject(project2)
                .Build();

            project1.Setup(m => m.ExportSourceFiles(_path));

            var exportAllCommand
                = ArrangeExportAllCommand(vbe, CreateMockFolderBrowserFactory(_projectPath, _path, DialogResult.Cancel));

            exportAllCommand.Execute(project1.Object);

            project1.Verify(m => m.ExportSourceFiles(_path), Times.Never);
            project2.Verify(m => m.ExportSourceFiles(_path), Times.Never);
        }

        [Category("Commands")]
        [Category(nameof(ExportAllCommand))]
        [TestCase(_projectFullPath, _projectPath)]
        [TestCase("   ", "")]
        [TestCase(null, "")]
        public void ExportAllCommand_GetDefaultExportFolder_HandleVariousProjectFileNamePropertyValues(
            string fileNamePropertyValue, 
            string expectedParentFolderpath)
        {
            var project = CreateTestProjectMocks("TestProject1").Values.First();

            project.SetupGet(m => m.FileName).Returns(fileNamePropertyValue);

            var builder = new MockVbeBuilder();
            var vbe = builder
                .AddProject(project)
                .Build();

            var exportAllCommandStub 
                = ArrangeFakeExportAllCommand(vbe, CreateMockFolderBrowserFactory(_projectPath, _path, DialogResult.Cancel));

            exportAllCommandStub.SetupFolderExists(_projectPath, !string.IsNullOrWhiteSpace(fileNamePropertyValue));

            var actualExportPath = exportAllCommandStub.GetDefaultExportFolder(project.Object.FileName);

            Assert.AreEqual(expectedParentFolderpath, actualExportPath);
        }

        [Category("Commands")]
        [Category(nameof(ExportAllCommand))]
        [Test]
        public void ExportAllCommand_LastFolderpathRetained()
        {
            var project = CreateTestProjectMocks("TestProject1").Values.First();

            project.SetupGet(m => m.FileName).Returns(_projectFullPath);

            var builder = new MockVbeBuilder();
            var vbe = builder
                .AddProject(project)
                .Build();

            var selectedPath = _projectPath + "\\ExportPath";
            var exportAllCommandStub
                = ArrangeFakeExportAllCommand(vbe, CreateMockFolderBrowserFactory(_projectPath, selectedPath, DialogResult.OK));

            exportAllCommandStub.SetupFolderExists(selectedPath, true);

            exportAllCommandStub.Execute(project);

            var actualInitialExportPath = exportAllCommandStub.GetInitialFolderBrowserPath(project.Object);

            Assert.AreEqual(selectedPath, actualInitialExportPath);
        }

        [Category("Commands")]
        [Category(nameof(ExportAllCommand))]
        [Test]
        public void ExportAllCommand_HandlesLastFolderpathDeleted_ReturnsWorkbookFolder()
        {
            var project = CreateTestProjectMocks("TestProject1").Values.First();

            project.SetupGet(m => m.FileName).Returns(_projectFullPath);

            var builder = new MockVbeBuilder();
            var vbe = builder
                .AddProject(project)
                .Build();

            var selectedPath = _projectPath + "\\ExportPath";
            var exportAllCommandStub
                = ArrangeFakeExportAllCommand(vbe, CreateMockFolderBrowserFactory(_projectPath, selectedPath, DialogResult.OK));

            exportAllCommandStub.SetupFolderExists(selectedPath, true);

            exportAllCommandStub.Execute(project);

            //User deletes the folder containing the last export 
            exportAllCommandStub.SetupFolderExists(selectedPath, false);

            //Initial path provided to the folder browser is now the folder containing the workbook
            var actual = exportAllCommandStub.GetInitialFolderBrowserPath(project.Object);

            Assert.AreEqual(_projectPath, actual);
        }

        [Category("Commands")]
        [Category(nameof(ExportAllCommand))]
        [Test] //User exports 3 projects.  Tests that all project folderpaths are cached
        public void ExportAllCommand_ExportMultipleProjectsInSequence()
        {
            //Arrange
            var projects = CreateTestProjectMocks("TestProject1", "TestProject2", "TestProject3");

            var project1 = projects["TestProject1"];
            var project2 = projects["TestProject2"];
            var project3 = projects["TestProject3"];

            project1.SetupGet(m => m.FileName).Returns(@"C:\Users\Rubberduck\Documents\Subfolder\Project1.xlsm");
            project2.SetupGet(m => m.FileName).Returns(@"C:\Users\Rubberduck\Documents\Subfolder\Project2.xlsm");
            project3.SetupGet(m => m.FileName).Returns(@"C:\Users\Rubberduck\Documents\Subfolder\Project3.xlsm");

            var builder = new MockVbeBuilder();
            var vbe = builder
                .AddProject(project1)
                .AddProject(project2)
                .AddProject(project3)
                .Build();

            var selected1 = Path.GetDirectoryName(project1.Object.FileName) + "\\Export1";
            var selected2 = Path.GetDirectoryName(project2.Object.FileName) + "\\Export2";
            var selected3 = Path.GetDirectoryName(project3.Object.FileName) + "\\Export3";

            var mockFolderBrowserFactory = new Mock<IFileSystemBrowserFactory>();
            mockFolderBrowserFactory
                .SetupSequence<IFolderBrowser>(m => m.CreateFolderBrowser(It.IsAny<string>(), true, It.IsAny<string>()))
                .Returns(CreateMockFolderBrowser(selected1).Object)
                .Returns(CreateMockFolderBrowser(selected2).Object)
                .Returns(CreateMockFolderBrowser(selected3).Object);

            var exportAllCommandStub = ArrangeFakeExportAllCommand(vbe, mockFolderBrowserFactory);

            exportAllCommandStub.SetupFolderExists(selected1, true);
            exportAllCommandStub.SetupFolderExists(selected2, true);
            exportAllCommandStub.SetupFolderExists(selected3, true);

            //Act

            //Get the initial folder which should be the workbook folder
            var initial1BeforeExportAll = exportAllCommandStub.GetInitialFolderBrowserPath(project1.Object);
            var initial2BeforeExportAll = exportAllCommandStub.GetInitialFolderBrowserPath(project2.Object);
            var initial3BeforeExportAll = exportAllCommandStub.GetInitialFolderBrowserPath(project3.Object);

            exportAllCommandStub.Execute(project1.Object);
            exportAllCommandStub.Execute(project2.Object);
            exportAllCommandStub.Execute(project3.Object);

            var initial1AfterExportAll = exportAllCommandStub.GetInitialFolderBrowserPath(project1.Object);
            var initial2AfterExportAll = exportAllCommandStub.GetInitialFolderBrowserPath(project2.Object);
            var initial3AfterExportAll = exportAllCommandStub.GetInitialFolderBrowserPath(project3.Object);

            //Assert

            Assert.AreEqual(_projectPath, initial1BeforeExportAll, "initial1BeforeExportAll failed");
            Assert.AreEqual(_projectPath, initial2BeforeExportAll, "initial2BeforeExportAll failed");
            Assert.AreEqual(_projectPath, initial3BeforeExportAll, "initial3BeforeExportAll failed");

            Assert.AreEqual(selected1, initial1AfterExportAll, "initial1AfterExportAll failed");
            Assert.AreEqual(selected2, initial2AfterExportAll, "initial2AfterExportAll failed");
            Assert.AreEqual(selected3, initial3AfterExportAll, "initial3AfterExportAll failed");
        }

        private Dictionary<string, Mock<IVBProject>> CreateTestProjectMocks(params string[] projectNames)
        {
            var results = new Dictionary<string, Mock<IVBProject>>();

            var builder = new MockVbeBuilder();

            for (int i = 0; i < projectNames.Length; i++)
            {
                var projectMock = builder.ProjectBuilder(projectNames[i], ProjectProtection.Unprotected)
                    .AddComponent("Module1", ComponentType.StandardModule, string.Empty)
                    .AddComponent("ClassModule1", ComponentType.ClassModule, string.Empty)
                    .AddComponent("Document1", ComponentType.Document, string.Empty)
                    .AddComponent("UserForm1", ComponentType.UserForm, string.Empty);

                results.Add(projectNames[i], projectMock.Build());
            }

            return results;
        }

        private static Mock<IFileSystemBrowserFactory> CreateMockFolderBrowserFactory(string projectPath, 
            string returnPath, DialogResult dialogResult)
        {
            var mockFolderBrowser = CreateMockFolderBrowser(returnPath, dialogResult);

            var mockFolderBrowserFactory = new Mock<IFileSystemBrowserFactory>();
            mockFolderBrowserFactory
                .Setup(m => m.CreateFolderBrowser(It.IsAny<string>(), true, projectPath))
                .Returns(mockFolderBrowser.Object);

            return mockFolderBrowserFactory;
        }

        private static Mock<IFolderBrowser> CreateMockFolderBrowser(string selectedFolder, DialogResult dialogResult = DialogResult.OK)
        {
            var mockFolderBrowser = new Mock<IFolderBrowser>();
            mockFolderBrowser.Setup(m => m.SelectedPath).Returns(selectedFolder);
            mockFolderBrowser.Setup(m => m.ShowDialog()).Returns(dialogResult);
            return mockFolderBrowser;
        }


        private static ExportAllCommand ArrangeExportAllCommand(
            Mock<IVBE> vbe,
            Mock<IFileSystemBrowserFactory> mockFolderBrowserFactory,
            IProjectsProvider projectsProvider = null)
        {
            return ArrangeExportAllCommand(vbe, mockFolderBrowserFactory, 
                MockVbeEvents.CreateMockVbeEvents(vbe), projectsProvider);
        }

        private static ExportAllCommand ArrangeExportAllCommand(
            Mock<IVBE> vbe,
            Mock<IFileSystemBrowserFactory> mockFolderBrowserFactory,
            Mock<IVbeEvents> vbeEvents,
            IProjectsProvider projectsProvider)
        {
            return new ExportAllCommand(vbe.Object, mockFolderBrowserFactory.Object, 
                vbeEvents.Object, projectsProvider, new ProjectToExportFolderMap());
        }

        private static ExportAllCommandFake ArrangeFakeExportAllCommand(
            Mock<IVBE> vbe,
            Mock<IFileSystemBrowserFactory> mockFolderBrowserFactory,
            IProjectsProvider projectsProvider = null)
        {
            return new ExportAllCommandFake(vbe.Object, mockFolderBrowserFactory.Object,
                MockVbeEvents.CreateMockVbeEvents(vbe).Object, projectsProvider);
        }

        /// <summary>
        /// ExportAllCommandFake inherits ExportAllCommand in order to access protected functions for testing
        /// </summary>
        class ExportAllCommandFake : ExportAllCommand
        {
            private static Dictionary<string, bool> _projectExportFolderExists;

            public ExportAllCommandFake(IVBE vbe, IFileSystemBrowserFactory browserFactory, 
                IVbeEvents vbeEvents, IProjectsProvider projectsProvider)
                   : base(vbe, browserFactory, vbeEvents, projectsProvider, new ProjectToExportFolderMap())
            {
                if (_projectExportFolderExists is null)
                {
                    _projectExportFolderExists = new Dictionary<string, bool>();
                }
            }

            public new string GetDefaultExportFolder(string projectFileName)
            {
                return base.GetDefaultExportFolder(projectFileName);
            }

            public void SetupFolderExists(string exportFolderpath, bool folderExists)
            {
                if (!_projectExportFolderExists.ContainsKey(exportFolderpath))
                {
                    _projectExportFolderExists.Add(exportFolderpath, folderExists);
                    return;
                }
                _projectExportFolderExists[exportFolderpath] = folderExists;
            }

            public new string GetInitialFolderBrowserPath(IVBProject project)
            {
                return base.GetInitialFolderBrowserPath(project);
            }

            protected override bool FolderExists(string path)
            {
                if (_projectExportFolderExists.ContainsKey(path))
                {
                    return _projectExportFolderExists[path];
                }

                return false;
            }
        }
    }
}
