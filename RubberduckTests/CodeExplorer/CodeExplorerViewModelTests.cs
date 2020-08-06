using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using NUnit.Framework;
using Moq;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.Interaction;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.UIContext;
using Rubberduck.Parsing.VBA;
using Rubberduck.Settings;
using Rubberduck.SettingsProvider;
using Rubberduck.UI.Command.ComCommands;
using RubberduckTests.Mocks;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.Utility;

namespace RubberduckTests.CodeExplorer
{
    [TestFixture]
    public class CodeExplorerViewModelTests
    {
        [Category("Code Explorer")]
        [Test]
        public void AddStdModule()
        {
            using (var explorer = new MockedCodeExplorer(ProjectType.HostProject).SelectFirstModule())
            {
                explorer.ExecuteAddStdModuleCommand();
                explorer.VbComponents.Verify(c => c.Add(ComponentType.StandardModule), Times.Once);
            }
        }

        [Category("Code Explorer")]
        [Test]
        [TestCase(ProjectType.StandardExe, ExpectedResult = true)]
        [TestCase(ProjectType.ActiveXExe, ExpectedResult = true)]
        [TestCase(ProjectType.ActiveXDll, ExpectedResult = true)]
        [TestCase(ProjectType.ActiveXControl, ExpectedResult = true)]
        [TestCase(ProjectType.HostProject, ExpectedResult = true)]
        [TestCase(ProjectType.StandAlone, ExpectedResult = true)]
        public bool AddStdModule_CanExecuteBasedOnProjectType(ProjectType projectType)
        {
            using (var explorer = new MockedCodeExplorer(projectType).ImplementAddStdModuleCommand().SelectFirstModule())
            {
                return explorer.ViewModel.AddStdModuleCommand.CanExecute(explorer.ViewModel.SelectedItem);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void AddClassModule()
        {
            using (var explorer = new MockedCodeExplorer(ProjectType.HostProject).SelectFirstModule())
            {
                explorer.ExecuteAddClassModuleCommand();
                explorer.VbComponents.Verify(c => c.Add(ComponentType.ClassModule), Times.Once);
            }
        }

        [Category("Code Explorer")]
        [Test]
        [TestCase(ProjectType.StandardExe, ExpectedResult = true)]
        [TestCase(ProjectType.ActiveXExe, ExpectedResult = true)]
        [TestCase(ProjectType.ActiveXDll, ExpectedResult = true)]
        [TestCase(ProjectType.ActiveXControl, ExpectedResult = true)]
        [TestCase(ProjectType.HostProject, ExpectedResult = true)]
        [TestCase(ProjectType.StandAlone, ExpectedResult = true)]
        public bool AddClassModule_CanExecuteBasedOnProjectType(ProjectType projectType)
        {
            using (var explorer = new MockedCodeExplorer(projectType).ImplementAddClassModuleCommand().SelectFirstModule())
            {
                return explorer.ViewModel.AddClassModuleCommand.CanExecute(explorer.ViewModel.SelectedItem);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void AddUserForm()
        {
            using (var explorer = new MockedCodeExplorer(ProjectType.HostProject).SelectFirstModule())
            {
                explorer.ExecuteAddUserFormCommand();
                explorer.VbComponents.Verify(c => c.Add(ComponentType.UserForm), Times.Once);
            }
        }

        [Category("Code Explorer")]
        [Test]
        [TestCase(ProjectType.StandardExe, ExpectedResult = false)]
        [TestCase(ProjectType.ActiveXExe, ExpectedResult = false)]
        [TestCase(ProjectType.ActiveXDll, ExpectedResult = false)]
        [TestCase(ProjectType.ActiveXControl, ExpectedResult = false)]
        [TestCase(ProjectType.HostProject, ExpectedResult = true)]
        [TestCase(ProjectType.StandAlone, ExpectedResult = true)]
        public bool AddUserForm_CanExecuteBasedOnProjectType(ProjectType projectType)
        {
            using (var explorer = new MockedCodeExplorer(projectType).ImplementAddUserFormCommand().SelectFirstModule())
            {
                return explorer.ViewModel.AddUserFormCommand.CanExecute(explorer.ViewModel.SelectedItem);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void AddVbForm()
        {
            using (var explorer = new MockedCodeExplorer(ProjectType.StandardExe).SelectFirstModule())
            {
                explorer.ExecuteAddVbFormCommand();
                explorer.VbComponents.Verify(c => c.Add(ComponentType.VBForm), Times.Once);
            }
        }

        [Category("Code Explorer")]
        [Test]
        [TestCase(ProjectType.StandardExe, ExpectedResult = true)]
        [TestCase(ProjectType.ActiveXExe, ExpectedResult = true)]
        [TestCase(ProjectType.ActiveXDll, ExpectedResult = true)]
        [TestCase(ProjectType.ActiveXControl, ExpectedResult = true)]
        [TestCase(ProjectType.HostProject, ExpectedResult = false)]
        [TestCase(ProjectType.StandAlone, ExpectedResult = false)]
        public bool AddVBForm_CanExecuteBasedOnProjectType(ProjectType projectType)
        {
            using (var explorer = new MockedCodeExplorer(projectType).ImplementAddVbFormCommand().SelectFirstModule())
            {
                return explorer.ViewModel.AddVBFormCommand.CanExecute(explorer.ViewModel.SelectedItem);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void AddMdiForm()
        {
            using (var explorer = new MockedCodeExplorer(ProjectType.StandardExe).SelectFirstModule())
            {
                explorer.ExecuteAddMdiFormCommand();
                explorer.VbComponents.Verify(c => c.Add(ComponentType.MDIForm), Times.Once);
            }
        }

        [Category("Code Explorer")]
        [Test]
        [TestCase(ProjectType.StandardExe, ExpectedResult = true)]
        [TestCase(ProjectType.ActiveXExe, ExpectedResult = true)]
        [TestCase(ProjectType.ActiveXDll, ExpectedResult = false)]
        [TestCase(ProjectType.ActiveXControl, ExpectedResult = false)]
        [TestCase(ProjectType.HostProject, ExpectedResult = false)]
        [TestCase(ProjectType.StandAlone, ExpectedResult = false)]
        public bool AddMDIForm_CanExecuteBasedOnProjectType(ProjectType projectType)
        {
            using (var explorer = new MockedCodeExplorer(projectType).ImplementAddMdiFormCommand().SelectFirstModule())
            {
                return explorer.ViewModel.AddMDIFormCommand.CanExecute(explorer.ViewModel.SelectedItem);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void AddMDIForm_CannotExecuteIfProjectAlreadyHasMDIForm()
        {
            using (var explorer = new MockedCodeExplorer(ProjectType.HostProject, ComponentType.MDIForm).ImplementAddMdiFormCommand().SelectFirstModule())
            {
                Assert.IsFalse(explorer.ViewModel.AddMDIFormCommand.CanExecute(explorer.ViewModel.SelectedItem));
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void AddUserControlForm()
        {
            using (var explorer = new MockedCodeExplorer(ProjectType.StandardExe).SelectFirstModule())
            {
                explorer.ExecuteAddUserControlCommand();
                explorer.VbComponents.Verify(c => c.Add(ComponentType.UserControl), Times.Once);
            }
        }

        [Category("Code Explorer")]
        [Test]
        [TestCase(ProjectType.StandardExe, ExpectedResult = true)]
        [TestCase(ProjectType.ActiveXExe, ExpectedResult = true)]
        [TestCase(ProjectType.ActiveXDll, ExpectedResult = true)]
        [TestCase(ProjectType.ActiveXControl, ExpectedResult = true)]
        [TestCase(ProjectType.HostProject, ExpectedResult = false)]
        [TestCase(ProjectType.StandAlone, ExpectedResult = false)]
        public bool AddUserControl_CanExecuteBasedOnProjectType(ProjectType projectType)
        {
            using (var explorer = new MockedCodeExplorer(projectType).ImplementAddUserControlCommand().SelectFirstModule())
            {
                return explorer.ViewModel.AddUserControlCommand.CanExecute(explorer.ViewModel.SelectedItem);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void AddPropertyPage()
        {
            using (var explorer = new MockedCodeExplorer(ProjectType.StandardExe).SelectFirstModule())
            {
                explorer.ExecuteAddPropertyPageCommand();
                explorer.VbComponents.Verify(c => c.Add(ComponentType.PropPage), Times.Once);
            }
        }

        [Category("Code Explorer")]
        [Test]
        [TestCase(ProjectType.StandardExe, ExpectedResult = true)]
        [TestCase(ProjectType.ActiveXExe, ExpectedResult = true)]
        [TestCase(ProjectType.ActiveXDll, ExpectedResult = true)]
        [TestCase(ProjectType.ActiveXControl, ExpectedResult = true)]
        [TestCase(ProjectType.HostProject, ExpectedResult = false)]
        [TestCase(ProjectType.StandAlone, ExpectedResult = false)]
        public bool AddPropertyPage_CanExecuteBasedOnProjectType(ProjectType projectType)
        {
            using (var explorer = new MockedCodeExplorer(projectType).ImplementAddPropertyPageCommand().SelectFirstModule())
            {
                return explorer.ViewModel.AddPropertyPageCommand.CanExecute(explorer.ViewModel.SelectedItem);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void AddUserDocument()
        {
            using (var explorer = new MockedCodeExplorer(ProjectType.ActiveXExe).SelectFirstModule())
            {
                explorer.ExecuteAddUserDocumentCommand();
                explorer.VbComponents.Verify(c => c.Add(ComponentType.DocObject), Times.Once);
            }
        }

        [Category("Code Explorer")]
        [Test]
        [TestCase(ProjectType.StandardExe, ExpectedResult = false)]
        [TestCase(ProjectType.ActiveXExe, ExpectedResult = true)]
        [TestCase(ProjectType.ActiveXDll, ExpectedResult = true)]
        [TestCase(ProjectType.ActiveXControl, ExpectedResult = false)]
        [TestCase(ProjectType.HostProject, ExpectedResult = false)]
        [TestCase(ProjectType.StandAlone, ExpectedResult = false)]
        public bool AddUserDocument_CanExecuteBasedOnProjectType(ProjectType projectType)
        {
            using (var explorer = new MockedCodeExplorer(projectType).ImplementAddUserDocumentCommand().SelectFirstModule())
            {
                return explorer.ViewModel.AddUserDocumentCommand.CanExecute(explorer.ViewModel.SelectedItem);
            }
        }

        [Category("Code Explorer")]
        [Test]
        [TestCase(ComponentType.ActiveXDesigner, ExpectedResult = true)]
        [TestCase(ComponentType.ClassModule, ExpectedResult = true)]
        [TestCase(ComponentType.ComComponent, ExpectedResult = true)]
        [TestCase(ComponentType.DocObject, ExpectedResult = true)]
        [TestCase(ComponentType.Document, ExpectedResult = true)]
        [TestCase(ComponentType.MDIForm, ExpectedResult = true)]
        [TestCase(ComponentType.PropPage, ExpectedResult = true)]
        [TestCase(ComponentType.RelatedDocument, ExpectedResult = false, Ignore = "Project doesn't contain selectable modules")]
        [TestCase(ComponentType.ResFile, ExpectedResult = false, Ignore = "Project doesn't contain selectable modules")]
        [TestCase(ComponentType.StandardModule, ExpectedResult = false)]
        [TestCase(ComponentType.Undefined, ExpectedResult = true)]
        [TestCase(ComponentType.UserControl, ExpectedResult = true)]
        [TestCase(ComponentType.UserForm, ExpectedResult = true)]
        [TestCase(ComponentType.VBForm, ExpectedResult = true)]
        public bool RefactorExtractInterface_CanExecuteBasedOnComponentType(ComponentType componentType)
        {
            using (var explorer = new MockedCodeExplorer(ProjectType.HostProject, componentType, @"Public Sub Foo():  MsgBox """":End Sub ")
                .ImplementExtractInterfaceCommand().SelectFirstModule())
            {
                return explorer.ViewModel.CodeExplorerExtractInterfaceCommand.CanExecute(explorer.ViewModel.SelectedItem);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void AddTestModule()
        {
            using (var explorer = new MockedCodeExplorer(ProjectType.HostProject).SelectFirstModule())
            {
                explorer.ExecuteAddTestModuleCommand();
                explorer.VbComponents.Verify(c => c.Add(ComponentType.StandardModule), Times.Once);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void AddTestModuleWithStubs()
        {
            using (var explorer = new MockedCodeExplorer(ProjectType.HostProject).SelectFirstModule())
            {
                explorer.ExecuteAddTestModuleWithStubsCommand();
                explorer.VbComponents.Verify(c => c.Add(ComponentType.StandardModule), Times.Once);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void AddTestModuleWithStubs_DisabledWhenParameterIsProject()
        {
            using (var explorer = new MockedCodeExplorer(ProjectType.HostProject).ImplementAddTestModuleWithStubsCommand().SelectFirstProject())
            {
                Assert.IsFalse(explorer.ViewModel.AddTestModuleWithStubsCommand.CanExecute(explorer.ViewModel.SelectedItem));
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void AddTestModuleWithStubs_DisabledWhenParameterIsFolder()
        {
            using (var explorer = new MockedCodeExplorer(ProjectType.HostProject).ImplementAddTestModuleWithStubsCommand().SelectFirstCustomFolder())
            {
                Assert.IsFalse(explorer.ViewModel.AddTestModuleWithStubsCommand.CanExecute(explorer.ViewModel.SelectedItem));
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void AddTestModuleWithStubs_DisabledWhenParameterIsModuleMember()
        {
            using (var explorer = new MockedCodeExplorer(ProjectType.HostProject, ComponentType.StandardModule, @"Private Sub Foo(): End Sub")
                .ImplementAddTestModuleWithStubsCommand()
                .SelectFirstMember())
            {
                Assert.IsFalse(explorer.ViewModel.AddTestModuleWithStubsCommand.CanExecute(explorer.ViewModel.SelectedItem));
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void ImportModule()
        {
            const string path = @"C:\Users\Rubberduck\Desktop\StdModule1.bas";

            using (var explorer = new MockedCodeExplorer(ProjectType.HostProject)
                .ConfigureOpenDialog(new[] { path }, DialogResult.OK)
                .SelectFirstProject())
            {
                explorer.ExecuteImportCommand();
                explorer.VbComponents.Verify(c => c.Import(path), Times.Once);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void ImportMultipleModules()
        {
            const string path1 = @"C:\Users\Rubberduck\Desktop\StdModule1.bas";
            const string path2 = @"C:\Users\Rubberduck\Desktop\ClsModule1.bas";

            using (var explorer = new MockedCodeExplorer(ProjectType.HostProject)
                .ConfigureOpenDialog(new[] { path1, path2 }, DialogResult.OK)
                .SelectFirstProject())
            {
                explorer.ExecuteImportCommand();
                explorer.VbComponents.Verify(c => c.Import(path1), Times.Once);
                explorer.VbComponents.Verify(c => c.Import(path2), Times.Once);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void ImportCommand_ModuleThere_DoesNotRemoveMatchingUserFormWithoutBinaryAndUsesSpecialImportMethod()
        {
            const string path = @"C:\Users\Rubberduck\Desktop\StdModule1.frm";
            const string binaryName = "myBinary.frx";
            var binaryPath = Path.Combine(Path.GetDirectoryName(path), binaryName);

            using (var explorer = new MockedCodeExplorer(
                    ProjectType.HostProject,
                    ("TestModule", ComponentType.UserForm, string.Empty),
                    ("OtherTestModule", ComponentType.StandardModule, string.Empty))
                .ConfigureOpenDialog(new[] { path }, DialogResult.OK)
                .SelectFirstProject())
            {
                var mockExtractor = new Mock<IRequiredBinaryFilesFromFileNameExtractor>();
                mockExtractor
                    .SetupGet(m => m.SupportedComponentTypes)
                    .Returns(new List<ComponentType> { ComponentType.UserForm });
                mockExtractor
                    .Setup(m => m.RequiredBinaryFiles(path, ComponentType.UserForm))
                    .Returns(new List<string> { binaryName });

                var mockFileExistenceChecker = new Mock<IFileExistenceChecker>();
                mockFileExistenceChecker.Setup(m => m.FileExists(binaryPath)).Returns(false);

                explorer.ExecuteImportCommand(
                    filename => filename == path ? "TestModule" : "YetAnotherModule",
                    null,
                    new List<IRequiredBinaryFilesFromFileNameExtractor> { mockExtractor.Object },
                    mockFileExistenceChecker);
                explorer.VbComponents.Verify(c => c.Remove(It.IsAny<IVBComponent>()), Times.Never);
                explorer.VbComponents.Verify(m => m.ImportSourceFile(path), Times.Once);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void ImportCommand_ModuleThere_ImportsUserFormWithBinary()
        {
            const string path = @"C:\Users\Rubberduck\Desktop\StdModule1.frm";
            const string binaryName = "myBinary.frx";
            var binaryPath = Path.Combine(Path.GetDirectoryName(path), binaryName);

            using (var explorer = new MockedCodeExplorer(
                    ProjectType.HostProject,
                    ("TestModule", ComponentType.UserForm, string.Empty),
                    ("OtherTestModule", ComponentType.StandardModule, string.Empty))
                .ConfigureOpenDialog(new[] { path }, DialogResult.OK)
                .SelectFirstProject())
            {
                var mockExtractor = new Mock<IRequiredBinaryFilesFromFileNameExtractor>();
                mockExtractor
                    .SetupGet(m => m.SupportedComponentTypes)
                    .Returns(new List<ComponentType> { ComponentType.UserForm });
                mockExtractor
                    .Setup(m => m.RequiredBinaryFiles(path, ComponentType.UserForm))
                    .Returns(new List<string> { binaryName });

                var mockFileExistenceChecker = new Mock<IFileExistenceChecker>();
                mockFileExistenceChecker.Setup(m => m.FileExists(binaryPath)).Returns(true);

                explorer.ExecuteImportCommand(
                    filename => filename == path ? "TestModule" : "YetAnotherModule",
                    null,
                    new List<IRequiredBinaryFilesFromFileNameExtractor> { mockExtractor.Object },
                    mockFileExistenceChecker);

                var modulesNames = explorer
                    .VbComponents
                    .Object
                    .Select(component => component.Name)
                    .ToList();

                explorer.VbComponents.Verify(c => c.Remove(It.IsAny<IVBComponent>()), Times.Never);
                //This depends on the setup of Import on the VBComponents mock, which determines the component name from the filename.
                Assert.IsTrue(modulesNames.Contains("StdModule1"));
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void ImportCommand_MultipleImports_RepeatedModuleName_Aborts()
        {
            const string path1 = @"C:\Users\Rubberduck\Desktop\StdModule1.bas";
            const string path2 = @"C:\Users\Rubberduck\Desktop\Class1.cls";
            const string path3 = @"C:\Users\Rubberduck\Desktop\StdModule2.bas";
            const string path4 = @"C:\Users\Rubberduck\Desktop\Class2.cls";

            using (var explorer = new MockedCodeExplorer(
                ProjectType.HostProject,
                ("TestModule", ComponentType.StandardModule, string.Empty),
                ("OtherTestModule", ComponentType.StandardModule, string.Empty),
                ("TestClass", ComponentType.ClassModule, string.Empty))
                .ConfigureOpenDialog(new[] { path1, path2, path3, path4 }, DialogResult.OK)
                .SelectFirstProject())
            {
                explorer.ExecuteImportCommand(filename =>
                {
                    switch (filename)
                    {
                        case path1:
                            return "TestModule";
                        case path2:
                            return "TestClass";
                        case path3:
                            return "TestModule";
                        case path4:
                            return "NewClass";
                        default:
                            return "YetAnotherModule";
                    }
                });

                var modulesNames = explorer
                    .VbComponents
                    .Object
                    .Select(component => component.Name)
                    .ToList();

                explorer.VbComponents.Verify(c => c.Remove(It.IsAny<IVBComponent>()), Times.Never);
                explorer.VbComponents.Verify(c => c.Import(It.IsAny<string>()), Times.Never);

                Assert.IsTrue(modulesNames.Contains("OtherTestModule"));
                Assert.IsTrue(modulesNames.Contains("TestModule"));
                Assert.IsTrue(modulesNames.Contains("TestClass"));
                Assert.AreEqual(3, modulesNames.Count);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void ImportCommand_NonMatchingComponentTypeForFormWithoutBinary_Aborts()
        {
            const string path1 = @"C:\Users\Rubberduck\Desktop\StdModule1.cls";
            const string path2 = @"C:\Users\Rubberduck\Desktop\Form1.frm";
            const string binaryName = "myBinary.frx";
            var binaryPath = Path.Combine(Path.GetDirectoryName(path2), binaryName);
            const string path3 = @"C:\Users\Rubberduck\Desktop\StdModule2.bas";
            const string path4 = @"C:\Users\Rubberduck\Desktop\Class2.cls";

            using (var explorer = new MockedCodeExplorer(
                ProjectType.HostProject,
                ("TestModule", ComponentType.StandardModule, string.Empty),
                ("OtherTestModule", ComponentType.StandardModule, string.Empty),
                ("TestClass", ComponentType.ClassModule, string.Empty))
                .ConfigureOpenDialog(new[] { path1, path2, path3, path4 }, DialogResult.OK)
                .SelectFirstProject())
            {
                var mockExtractor = new Mock<IRequiredBinaryFilesFromFileNameExtractor>();
                mockExtractor
                    .SetupGet(m => m.SupportedComponentTypes)
                    .Returns(new List<ComponentType> { ComponentType.UserForm });
                mockExtractor
                    .Setup(m => m.RequiredBinaryFiles(path2, ComponentType.UserForm))
                    .Returns(new List<string> { binaryName });

                var mockFileExistenceChecker = new Mock<IFileExistenceChecker>();
                mockFileExistenceChecker.Setup(m => m.FileExists(binaryPath)).Returns(false);

                explorer.ExecuteImportCommand(filename =>
                    {
                        switch (filename)
                        {
                            case path1:
                                return "TestModule";
                            case path2:
                                return "TestClass";
                            case path3:
                                return "NewModule";
                            case path4:
                                return "NewClass";
                            default:
                                return "YetAnotherModule";
                        }
                    },
                    null,
                    new List<IRequiredBinaryFilesFromFileNameExtractor> { mockExtractor.Object },
                    mockFileExistenceChecker);

                var modulesNames = explorer
                    .VbComponents
                    .Object
                    .Select(component => component.Name)
                    .ToList();

                explorer.VbComponents.Verify(c => c.Remove(It.IsAny<IVBComponent>()), Times.Never);
                explorer.VbComponents.Verify(c => c.Import(It.IsAny<string>()), Times.Never);

                Assert.IsTrue(modulesNames.Contains("OtherTestModule"));
                Assert.IsTrue(modulesNames.Contains("TestModule"));
                Assert.IsTrue(modulesNames.Contains("TestClass"));
                Assert.AreEqual(3, modulesNames.Count);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void ImportCommand_NonMatchingComponentTypeForDocument_Aborts()
        {
            const string path1 = @"C:\Users\Rubberduck\Desktop\StdModule1.cls";
            const string path2 = @"C:\Users\Rubberduck\Desktop\Document.doccls";
            const string path3 = @"C:\Users\Rubberduck\Desktop\StdModule2.bas";
            const string path4 = @"C:\Users\Rubberduck\Desktop\Class1.cls";

            using (var explorer = new MockedCodeExplorer(
                ProjectType.HostProject,
                ("TestModule", ComponentType.StandardModule, string.Empty),
                ("OtherTestModule", ComponentType.StandardModule, string.Empty),
                ("TestClass", ComponentType.ClassModule, string.Empty))
                .ConfigureOpenDialog(new[] { path1, path2, path3, path4 }, DialogResult.OK)
                .SelectFirstProject())
            {
                explorer.ExecuteImportCommand(filename =>
                {
                    switch (filename)
                    {
                        case path1:
                            return "TestModule";
                        case path2:
                            return "TestClass";
                        case path3:
                            return "NewModule";
                        case path4:
                            return "NewDocument";
                        default:
                            return "YetAnotherModule";
                    }
                });

                var modulesNames = explorer
                    .VbComponents
                    .Object
                    .Select(component => component.Name)
                    .ToList();

                explorer.VbComponents.Verify(c => c.Remove(It.IsAny<IVBComponent>()), Times.Never);
                explorer.VbComponents.Verify(c => c.Import(It.IsAny<string>()), Times.Never);

                Assert.IsTrue(modulesNames.Contains("OtherTestModule"));
                Assert.IsTrue(modulesNames.Contains("TestModule"));
                Assert.IsTrue(modulesNames.Contains("TestClass"));
                Assert.AreEqual(3, modulesNames.Count);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void ImportCommand_UserFormWithoutExistingFormOrBinary_Aborts()
        {
            const string path1 = @"C:\Users\Rubberduck\Desktop\StdModule1.cls";
            const string path2 = @"C:\Users\Rubberduck\Desktop\Class1.cls";
            const string path3 = @"C:\Users\Rubberduck\Desktop\StdModule2.bas";
            const string path4 = @"C:\Users\Rubberduck\Desktop\Class2.frm";
            const string binaryName = "myBinary.frx";
            var binaryPath = Path.Combine(Path.GetDirectoryName(path4), binaryName);

            using (var explorer = new MockedCodeExplorer(
                ProjectType.HostProject,
                ("TestModule", ComponentType.StandardModule, string.Empty),
                ("OtherTestModule", ComponentType.StandardModule, string.Empty),
                ("TestClass", ComponentType.ClassModule, string.Empty))
                .ConfigureOpenDialog(new[] { path1, path2, path3, path4 }, DialogResult.OK)
                .SelectFirstProject())
            {
                var mockExtractor = new Mock<IRequiredBinaryFilesFromFileNameExtractor>();
                mockExtractor
                    .SetupGet(m => m.SupportedComponentTypes)
                    .Returns(new List<ComponentType> { ComponentType.UserForm });
                mockExtractor
                    .Setup(m => m.RequiredBinaryFiles(path4, ComponentType.UserForm))
                    .Returns(new List<string> { binaryName });

                var mockFileExistenceChecker = new Mock<IFileExistenceChecker>();
                mockFileExistenceChecker.Setup(m => m.FileExists(binaryPath)).Returns(false);

                explorer.ExecuteImportCommand(filename =>
                    {
                        switch (filename)
                        {
                            case path1:
                                return "TestModule";
                            case path2:
                                return "TestClass";
                            case path3:
                                return "NewModule";
                            case path4:
                                return "NewForm";
                            default:
                                return "YetAnotherModule";
                        }
                    },
                    null,
                    new List<IRequiredBinaryFilesFromFileNameExtractor> { mockExtractor.Object },
                    mockFileExistenceChecker);

                var modulesNames = explorer
                    .VbComponents
                    .Object
                    .Select(component => component.Name)
                    .ToList();

                explorer.VbComponents.Verify(c => c.Remove(It.IsAny<IVBComponent>()), Times.Never);
                explorer.VbComponents.Verify(c => c.Import(It.IsAny<string>()), Times.Never);

                Assert.IsTrue(modulesNames.Contains("OtherTestModule"));
                Assert.IsTrue(modulesNames.Contains("TestModule"));
                Assert.IsTrue(modulesNames.Contains("TestClass"));
                Assert.AreEqual(3, modulesNames.Count);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void ImportCommand_DocumentWithoutExistingModule_Aborts()
        {
            const string path1 = @"C:\Users\Rubberduck\Desktop\StdModule1.cls";
            const string path2 = @"C:\Users\Rubberduck\Desktop\Class1.cls";
            const string path3 = @"C:\Users\Rubberduck\Desktop\StdModule2.bas";
            const string path4 = @"C:\Users\Rubberduck\Desktop\Class2.doccls";

            using (var explorer = new MockedCodeExplorer(
                ProjectType.HostProject,
                ("TestModule", ComponentType.StandardModule, string.Empty),
                ("OtherTestModule", ComponentType.StandardModule, string.Empty),
                ("TestClass", ComponentType.ClassModule, string.Empty))
                .ConfigureOpenDialog(new[] { path1, path2, path3, path4 }, DialogResult.OK)
                .SelectFirstProject())
            {
                explorer.ExecuteImportCommand(filename =>
                {
                    switch (filename)
                    {
                        case path1:
                            return "TestModule";
                        case path2:
                            return "TestClass";
                        case path3:
                            return "NewModule";
                        case path4:
                            return "NewDocument";
                        default:
                            return "YetAnotherModule";
                    }
                });

                var modulesNames = explorer
                    .VbComponents
                    .Object
                    .Select(component => component.Name)
                    .ToList();

                explorer.VbComponents.Verify(c => c.Remove(It.IsAny<IVBComponent>()), Times.Never);
                explorer.VbComponents.Verify(c => c.Import(It.IsAny<string>()), Times.Never);

                Assert.IsTrue(modulesNames.Contains("OtherTestModule"));
                Assert.IsTrue(modulesNames.Contains("TestModule"));
                Assert.IsTrue(modulesNames.Contains("TestClass"));
                Assert.AreEqual(3, modulesNames.Count);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void ImportModule_Cancel()
        {
            const string path = @"C:\Users\Rubberduck\Desktop\StdModule1.bas";

            using (var explorer = new MockedCodeExplorer(ProjectType.HostProject)
                .ConfigureSaveDialog(path, DialogResult.Cancel)
                .SelectFirstModule())
            {
                explorer.ExecuteImportCommand();
                explorer.VbComponents.Verify(c => c.Import(path), Times.Never);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void UpdateFromFile_ModuleNotThere_Imports()
        {
            const string path = @"C:\Users\Rubberduck\Desktop\StdModule1.bas";

            using (var explorer = new MockedCodeExplorer(
                ProjectType.HostProject,
                ("TestModule", ComponentType.StandardModule, string.Empty))
                .ConfigureOpenDialog(new[] { path }, DialogResult.OK)
                .SelectFirstProject())
            {
                explorer.ExecuteUpdateFromFileCommand(filename => "SomeOtherModule");
                explorer.VbComponents.Verify(c => c.Import(path), Times.Once);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void UpdateFromFile_ModuleThere_Imports()
        {
            const string path = @"C:\Users\Rubberduck\Desktop\StdModule1.bas";

            using (var explorer = new MockedCodeExplorer(
                ProjectType.HostProject,
                ("TestModule", ComponentType.StandardModule, string.Empty))
                .ConfigureOpenDialog(new[] { path }, DialogResult.OK)
                .SelectFirstProject())
            {
                explorer.ExecuteUpdateFromFileCommand(filename => "TestModule");
                explorer.VbComponents.Verify(c => c.Import(path), Times.Once);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void UpdateFromFile_ModuleNotThere_DoesNotRemove()
        {
            const string path = @"C:\Users\Rubberduck\Desktop\StdModule1.bas";

            using (var explorer = new MockedCodeExplorer(
                ProjectType.HostProject,
                ("TestModule", ComponentType.StandardModule, string.Empty))
                .ConfigureOpenDialog(new[] { path }, DialogResult.OK)
                .SelectFirstProject())
            {
                explorer.ExecuteUpdateFromFileCommand(filename => "SomeOtherModule");
                explorer.VbComponents.Verify(c => c.Remove(It.IsAny<IVBComponent>()), Times.Never);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void UpdateFromFile_ModuleThere_RemovesMatchingComponent()
        {
            const string path = @"C:\Users\Rubberduck\Desktop\StdModule1.bas";

            using (var explorer = new MockedCodeExplorer(
                ProjectType.HostProject,
                ("TestModule", ComponentType.StandardModule, string.Empty),
                ("OtherTestModule", ComponentType.StandardModule, string.Empty))
                .ConfigureOpenDialog(new[] { path }, DialogResult.OK)
                .SelectFirstProject())
            {
                explorer.ExecuteUpdateFromFileCommand(filename => filename == path ? "TestModule" : "YetAnotherModule");

                var modulesNames = explorer
                    .VbComponents
                    .Object
                    .Select(component => component.Name)
                    .ToList();

                explorer.VbComponents.Verify(c => c.Remove(It.IsAny<IVBComponent>()), Times.Once);

                Assert.IsTrue(modulesNames.Contains("OtherTestModule"));
                //This depends on the setup of Import on the VBComponents mock, which determines the component name from the filename.
                Assert.IsTrue(modulesNames.Contains("StdModule1"));
                Assert.IsFalse(modulesNames.Contains("TestModule"));
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void UpdateFromFile_ModuleThere_DoesNotRemoveMatchingDocument()
        {
            const string path = @"C:\Users\Rubberduck\Desktop\StdModule1.doccls";

            using (var explorer = new MockedCodeExplorer(
                    ProjectType.HostProject,
                    ("TestModule", ComponentType.Document, string.Empty),
                    ("OtherTestModule", ComponentType.StandardModule, string.Empty))
                .ConfigureOpenDialog(new[] { path }, DialogResult.OK)
                .SelectFirstProject())
            {
                explorer.ExecuteUpdateFromFileCommand(filename => filename == path ? "TestModule" : "YetAnotherModule");
                explorer.VbComponents.Verify(c => c.Remove(It.IsAny<IVBComponent>()), Times.Never);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void UpdateFromFile_ModuleThere_DoesNotRemoveMatchingUserFormWithoutBinaryAndUsesSpecialImportMethod()
        {
            const string path = @"C:\Users\Rubberduck\Desktop\StdModule1.frm";
            const string binaryName = "myBinary.frx";
            var binaryPath = Path.Combine(Path.GetDirectoryName(path), binaryName);
                
            using (var explorer = new MockedCodeExplorer(
                    ProjectType.HostProject,
                    ("TestModule", ComponentType.UserForm, string.Empty),
                    ("OtherTestModule", ComponentType.StandardModule, string.Empty))
                .ConfigureOpenDialog(new[] { path }, DialogResult.OK)
                .SelectFirstProject())
            {
                var mockExtractor = new Mock<IRequiredBinaryFilesFromFileNameExtractor>();
                mockExtractor
                    .SetupGet(m => m.SupportedComponentTypes)
                    .Returns(new List<ComponentType> {ComponentType.UserForm});
                mockExtractor
                    .Setup(m => m.RequiredBinaryFiles(path, ComponentType.UserForm))
                    .Returns(new List<string>{ binaryName });

                var mockFileExistenceChecker = new Mock<IFileExistenceChecker>();
                mockFileExistenceChecker.Setup(m => m.FileExists(binaryPath)).Returns(false);

                explorer.ExecuteUpdateFromFileCommand(
                    filename => filename == path ? "TestModule" : "YetAnotherModule",
                    null,
                    new List<IRequiredBinaryFilesFromFileNameExtractor>{mockExtractor.Object},
                    mockFileExistenceChecker);
                explorer.VbComponents.Verify(c => c.Remove(It.IsAny<IVBComponent>()), Times.Never);
                explorer.VbComponents.Verify(m => m.ImportSourceFile(path), Times.Once);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void UpdateFromFile_ModuleThere_RemovesMatchingUserFormWithBinary()
        {
            const string path = @"C:\Users\Rubberduck\Desktop\StdModule1.frm";
            const string binaryName = "myBinary.frx";
            var binaryPath = Path.Combine(Path.GetDirectoryName(path), binaryName);

            using (var explorer = new MockedCodeExplorer(
                    ProjectType.HostProject,
                    ("TestModule", ComponentType.UserForm, string.Empty),
                    ("OtherTestModule", ComponentType.StandardModule, string.Empty))
                .ConfigureOpenDialog(new[] { path }, DialogResult.OK)
                .SelectFirstProject())
            {
                var mockExtractor = new Mock<IRequiredBinaryFilesFromFileNameExtractor>();
                mockExtractor
                    .SetupGet(m => m.SupportedComponentTypes)
                    .Returns(new List<ComponentType> { ComponentType.UserForm });
                mockExtractor
                    .Setup(m => m.RequiredBinaryFiles(path, ComponentType.UserForm))
                    .Returns(new List<string> { binaryName });

                var mockFileExistenceChecker = new Mock<IFileExistenceChecker>();
                mockFileExistenceChecker.Setup(m => m.FileExists(binaryPath)).Returns(true);

                explorer.ExecuteUpdateFromFileCommand(
                    filename => filename == path ? "TestModule" : "YetAnotherModule",
                    null,
                    new List<IRequiredBinaryFilesFromFileNameExtractor> { mockExtractor.Object },
                    mockFileExistenceChecker);

                var modulesNames = explorer
                    .VbComponents
                    .Object
                    .Select(component => component.Name)
                    .ToList();

                explorer.VbComponents.Verify(c => c.Remove(It.IsAny<IVBComponent>()), Times.Once);
                //This depends on the setup of Import on the VBComponents mock, which determines the component name from the filename.
                Assert.IsTrue(modulesNames.Contains("StdModule1"));
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void UpdateFromFile_MultipleImports_DifferentNames()
        {
            const string path1 = @"C:\Users\Rubberduck\Desktop\StdModule1.bas";
            const string path2 = @"C:\Users\Rubberduck\Desktop\Class1.cls";
            const string path3 = @"C:\Users\Rubberduck\Desktop\StdModule2.bas";
            const string path4 = @"C:\Users\Rubberduck\Desktop\Class2.cls";

            using (var explorer = new MockedCodeExplorer(
                ProjectType.HostProject,
                ("TestModule", ComponentType.StandardModule, string.Empty),
                ("OtherTestModule", ComponentType.StandardModule, string.Empty),
                ("TestClass", ComponentType.ClassModule, string.Empty))
                .ConfigureOpenDialog(new[] { path1, path2, path3, path4 }, DialogResult.OK)
                .SelectFirstProject())
            {
                explorer.ExecuteUpdateFromFileCommand(filename =>
                {
                    switch (filename)
                    {
                        case path1:
                            return "TestModule";
                        case path2:
                            return "TestClass";
                        case path3:
                            return "NewModule";
                        case path4:
                            return "NewClass";
                        default:
                            return "YetAnotherModule";
                    }
                });

                var modulesNames = explorer
                    .VbComponents
                    .Object
                    .Select(component => component.Name)
                    .ToList();

                explorer.VbComponents.Verify(c => c.Remove(It.IsAny<IVBComponent>()), Times.Exactly(2));
                explorer.VbComponents.Verify(c => c.Import(path1), Times.Once);
                explorer.VbComponents.Verify(c => c.Import(path2), Times.Once);
                explorer.VbComponents.Verify(c => c.Import(path3), Times.Once);
                explorer.VbComponents.Verify(c => c.Import(path4), Times.Once);

                Assert.IsTrue(modulesNames.Contains("OtherTestModule"));
                //This depends on the setup of Import on the VBComponents mock, which determines the component name from the filename.
                Assert.IsTrue(modulesNames.Contains("StdModule1"));
                Assert.IsTrue(modulesNames.Contains("Class1"));
                Assert.IsTrue(modulesNames.Contains("StdModule2"));
                Assert.IsTrue(modulesNames.Contains("Class2"));
                Assert.IsFalse(modulesNames.Contains("TestModule"));
                Assert.IsFalse(modulesNames.Contains("TestClass"));
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void UpdateFromFile_MultipleImports_RepeatedModuleName_Aborts()
        {
            const string path1 = @"C:\Users\Rubberduck\Desktop\StdModule1.bas";
            const string path2 = @"C:\Users\Rubberduck\Desktop\Class1.cls";
            const string path3 = @"C:\Users\Rubberduck\Desktop\StdModule2.bas";
            const string path4 = @"C:\Users\Rubberduck\Desktop\Class2.cls";

            using (var explorer = new MockedCodeExplorer(
                ProjectType.HostProject,
                ("TestModule", ComponentType.StandardModule, string.Empty),
                ("OtherTestModule", ComponentType.StandardModule, string.Empty),
                ("TestClass", ComponentType.ClassModule, string.Empty))
                .ConfigureOpenDialog(new[] { path1, path2, path3, path4 }, DialogResult.OK)
                .SelectFirstProject())
            {
                explorer.ExecuteUpdateFromFileCommand(filename =>
                {
                    switch (filename)
                    {
                        case path1:
                            return "TestModule";
                        case path2:
                            return "TestClass";
                        case path3:
                            return "TestModule";
                        case path4:
                            return "NewClass";
                        default:
                            return "YetAnotherModule";
                    }
                });

                var modulesNames = explorer
                    .VbComponents
                    .Object
                    .Select(component => component.Name)
                    .ToList();

                explorer.VbComponents.Verify(c => c.Remove(It.IsAny<IVBComponent>()), Times.Never);
                explorer.VbComponents.Verify(c => c.Import(It.IsAny <string>()), Times.Never);

                Assert.IsTrue(modulesNames.Contains("OtherTestModule"));
                Assert.IsTrue(modulesNames.Contains("TestModule"));
                Assert.IsTrue(modulesNames.Contains("TestClass"));
                Assert.AreEqual(3, modulesNames.Count);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void UpdateFromFile_NonMatchingComponentType_Aborts()
        {
            const string path1 = @"C:\Users\Rubberduck\Desktop\StdModule1.cls";
            const string path2 = @"C:\Users\Rubberduck\Desktop\Class1.cls";
            const string path3 = @"C:\Users\Rubberduck\Desktop\StdModule2.bas";
            const string path4 = @"C:\Users\Rubberduck\Desktop\Class2.cls";

            using (var explorer = new MockedCodeExplorer(
                ProjectType.HostProject,
                ("TestModule", ComponentType.StandardModule, string.Empty),
                ("OtherTestModule", ComponentType.StandardModule, string.Empty),
                ("TestClass", ComponentType.ClassModule, string.Empty))
                .ConfigureOpenDialog(new[] { path1, path2, path3, path4 }, DialogResult.OK)
                .SelectFirstProject())
            {
                explorer.ExecuteUpdateFromFileCommand(filename =>
                {
                    switch (filename)
                    {
                        case path1:
                            return "TestModule";
                        case path2:
                            return "TestClass";
                        case path3:
                            return "NewModule";
                        case path4:
                            return "NewClass";
                        default:
                            return "YetAnotherModule";
                    }
                });

                var modulesNames = explorer
                    .VbComponents
                    .Object
                    .Select(component => component.Name)
                    .ToList();

                explorer.VbComponents.Verify(c => c.Remove(It.IsAny<IVBComponent>()), Times.Never);
                explorer.VbComponents.Verify(c => c.Import(It.IsAny<string>()), Times.Never);

                Assert.IsTrue(modulesNames.Contains("OtherTestModule"));
                Assert.IsTrue(modulesNames.Contains("TestModule"));
                Assert.IsTrue(modulesNames.Contains("TestClass"));
                Assert.AreEqual(3, modulesNames.Count);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void UpdateFromFile_UserFormWithoutExistingFormAndBinary_Aborts()
        {
            const string path1 = @"C:\Users\Rubberduck\Desktop\StdModule1.cls";
            const string path2 = @"C:\Users\Rubberduck\Desktop\Class1.cls";
            const string path3 = @"C:\Users\Rubberduck\Desktop\StdModule2.bas";
            const string path4 = @"C:\Users\Rubberduck\Desktop\Class2.frm";
            const string binaryName = "myBinary.frx";
            var binaryPath = Path.Combine(Path.GetDirectoryName(path4), binaryName);

            using (var explorer = new MockedCodeExplorer(
                ProjectType.HostProject,
                ("TestModule", ComponentType.StandardModule, string.Empty),
                ("OtherTestModule", ComponentType.StandardModule, string.Empty),
                ("TestClass", ComponentType.ClassModule, string.Empty))
                .ConfigureOpenDialog(new[] { path1, path2, path3, path4 }, DialogResult.OK)
                .SelectFirstProject())
            {
                var mockExtractor = new Mock<IRequiredBinaryFilesFromFileNameExtractor>();
                mockExtractor
                    .SetupGet(m => m.SupportedComponentTypes)
                    .Returns(new List<ComponentType> { ComponentType.UserForm });
                mockExtractor
                    .Setup(m => m.RequiredBinaryFiles(path4, ComponentType.UserForm))
                    .Returns(new List<string> { binaryName });

                var mockFileExistenceChecker = new Mock<IFileExistenceChecker>();
                mockFileExistenceChecker.Setup(m => m.FileExists(binaryPath)).Returns(false);

                explorer.ExecuteUpdateFromFileCommand(filename =>
                    {
                        switch (filename)
                        {
                            case path1:
                                return "TestModule";
                            case path2:
                                return "TestClass";
                            case path3:
                                return "NewModule";
                            case path4:
                                return "NewForm";
                            default:
                                return "YetAnotherModule";
                        }
                    },
                    null,
                    new List<IRequiredBinaryFilesFromFileNameExtractor> { mockExtractor.Object },
                    mockFileExistenceChecker);

                var modulesNames = explorer
                    .VbComponents
                    .Object
                    .Select(component => component.Name)
                    .ToList();

                explorer.VbComponents.Verify(c => c.Remove(It.IsAny<IVBComponent>()), Times.Never);
                explorer.VbComponents.Verify(c => c.Import(It.IsAny<string>()), Times.Never);

                Assert.IsTrue(modulesNames.Contains("OtherTestModule"));
                Assert.IsTrue(modulesNames.Contains("TestModule"));
                Assert.IsTrue(modulesNames.Contains("TestClass"));
                Assert.AreEqual(3, modulesNames.Count);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void UpdateFromFile_ModuleNotThere_ImportsUserFormWithBinary()
        {
            const string path = @"C:\Users\Rubberduck\Desktop\StdModule1.frm";
            const string binaryName = "myBinary.frx";
            var binaryPath = Path.Combine(Path.GetDirectoryName(path), binaryName);

            using (var explorer = new MockedCodeExplorer(
                    ProjectType.HostProject,
                    ("TestModule", ComponentType.UserForm, string.Empty),
                    ("OtherTestModule", ComponentType.StandardModule, string.Empty))
                .ConfigureOpenDialog(new[] { path }, DialogResult.OK)
                .SelectFirstProject())
            {
                var mockExtractor = new Mock<IRequiredBinaryFilesFromFileNameExtractor>();
                mockExtractor
                    .SetupGet(m => m.SupportedComponentTypes)
                    .Returns(new List<ComponentType> { ComponentType.UserForm });
                mockExtractor
                    .Setup(m => m.RequiredBinaryFiles(path, ComponentType.UserForm))
                    .Returns(new List<string> { binaryName });

                var mockFileExistenceChecker = new Mock<IFileExistenceChecker>();
                mockFileExistenceChecker.Setup(m => m.FileExists(binaryPath)).Returns(true);

                explorer.ExecuteUpdateFromFileCommand(
                    filename => filename == path ? "NewTestForm" : "YetAnotherModule",
                    null,
                    new List<IRequiredBinaryFilesFromFileNameExtractor> { mockExtractor.Object },
                    mockFileExistenceChecker);

                var modulesNames = explorer
                    .VbComponents
                    .Object
                    .Select(component => component.Name)
                    .ToList();

                explorer.VbComponents.Verify(c => c.Remove(It.IsAny<IVBComponent>()), Times.Never);
                //This depends on the setup of Import on the VBComponents mock, which determines the component name from the filename.
                Assert.IsTrue(modulesNames.Contains("StdModule1"));
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void UpdateFromFile_DocumentWithoutExistingModule_Aborts()
        {
            const string path1 = @"C:\Users\Rubberduck\Desktop\StdModule1.cls";
            const string path2 = @"C:\Users\Rubberduck\Desktop\Class1.cls";
            const string path3 = @"C:\Users\Rubberduck\Desktop\StdModule2.bas";
            const string path4 = @"C:\Users\Rubberduck\Desktop\Class2.doccls";

            using (var explorer = new MockedCodeExplorer(
                ProjectType.HostProject,
                ("TestModule", ComponentType.StandardModule, string.Empty),
                ("OtherTestModule", ComponentType.StandardModule, string.Empty),
                ("TestClass", ComponentType.ClassModule, string.Empty))
                .ConfigureOpenDialog(new[] { path1, path2, path3, path4 }, DialogResult.OK)
                .SelectFirstProject())
            {
                explorer.ExecuteUpdateFromFileCommand(filename =>
                {
                    switch (filename)
                    {
                        case path1:
                            return "TestModule";
                        case path2:
                            return "TestClass";
                        case path3:
                            return "NewModule";
                        case path4:
                            return "NewDocument";
                        default:
                            return "YetAnotherModule";
                    }
                });

                var modulesNames = explorer
                    .VbComponents
                    .Object
                    .Select(component => component.Name)
                    .ToList();

                explorer.VbComponents.Verify(c => c.Remove(It.IsAny<IVBComponent>()), Times.Never);
                explorer.VbComponents.Verify(c => c.Import(It.IsAny<string>()), Times.Never);

                Assert.IsTrue(modulesNames.Contains("OtherTestModule"));
                Assert.IsTrue(modulesNames.Contains("TestModule"));
                Assert.IsTrue(modulesNames.Contains("TestClass"));
                Assert.AreEqual(3, modulesNames.Count);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void UpdateFromFile_Cancel()
        {
            const string path = @"C:\Users\Rubberduck\Desktop\StdModule1.bas";

            using (var explorer = new MockedCodeExplorer(
                ProjectType.HostProject,
                ("TestModule", ComponentType.StandardModule, string.Empty))
                .ConfigureOpenDialog(new[] { path }, DialogResult.Cancel)
                .SelectFirstProject())
            {
                explorer.ExecuteUpdateFromFileCommand(filename => "TestModule");
                explorer.VbComponents.Verify(c => c.Import(It.IsAny<string>()), Times.Never);
                explorer.VbComponents.Verify(c => c.Remove(It.IsAny<IVBComponent>()), Times.Never);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void ReplaceProjectContentsFromFiles_Imports()
        {
            const string path = @"C:\Users\Rubberduck\Desktop\StdModule1.bas";

            using (var explorer = new MockedCodeExplorer(
                ProjectType.HostProject,
                ("TestModule", ComponentType.StandardModule, string.Empty))
                .ConfigureOpenDialog(new[] { path }, DialogResult.OK)
                .SelectFirstProject())
            {
                explorer.ExecuteReplaceProjectContentsFromFilesCommand();
                explorer.VbComponents.Verify(c => c.Import(path), Times.Once);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void ReplaceProjectContentsFromFiles_ImportsMultiple()
        {
            const string path1 = @"C:\Users\Rubberduck\Desktop\StdModule1.bas";
            const string path2 = @"C:\Users\Rubberduck\Desktop\Class1.cls";

            using (var explorer = new MockedCodeExplorer(
                ProjectType.HostProject,
                ("TestModule", ComponentType.StandardModule, string.Empty))
                .ConfigureOpenDialog(new[] { path1, path2 }, DialogResult.OK)
                .SelectFirstProject())
            {
                explorer.ExecuteReplaceProjectContentsFromFilesCommand();
                explorer.VbComponents.Verify(c => c.Import(path1), Times.Once);
                explorer.VbComponents.Verify(c => c.Import(path2), Times.Once);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void ReplaceProjectContentsFromFiles_RemovesReimportableComponents()
        {
            const string path = @"C:\Users\Rubberduck\Desktop\StdModule1.bas";

            using (var explorer = new MockedCodeExplorer(
                ProjectType.HostProject,
                ("TestModule", ComponentType.StandardModule, string.Empty),
                ("TestClass", ComponentType.ClassModule, string.Empty),
                ("TestUserForm", ComponentType.UserForm, string.Empty))
                .ConfigureOpenDialog(new[] { path }, DialogResult.OK)
                .SelectFirstProject())
            {
                explorer.ExecuteReplaceProjectContentsFromFilesCommand();

                var modulesNames = explorer
                    .VbComponents
                    .Object
                    .Select(component => component.Name)
                    .ToList();

                explorer.VbComponents.Verify(c => c.Remove(It.IsAny<IVBComponent>()), Times.Exactly(3));

                Assert.IsFalse(modulesNames.Contains("TestModule"));
                Assert.IsFalse(modulesNames.Contains("TestClass"));
                Assert.IsFalse(modulesNames.Contains("TestUserForm"));

                //This depends on the setup of Import on the VBComponents mock, which determines the component name from the filename.
                Assert.IsTrue(modulesNames.Contains("StdModule1"));
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void ReplaceProjectContentsFromFiles_DoesNotRemoveNonReimportableComponents()
        {
            const string path = @"C:\Users\Rubberduck\Desktop\StdModule1.bas";

            using (var explorer = new MockedCodeExplorer(
                ProjectType.HostProject,
                ("TestModule", ComponentType.StandardModule, string.Empty),
                ("TestDesigner", ComponentType.ActiveXDesigner, string.Empty))
                .ConfigureOpenDialog(new[] { path }, DialogResult.OK)
                .SelectFirstProject())
            {
                explorer.ExecuteReplaceProjectContentsFromFilesCommand();

                var modulesNames = explorer
                    .VbComponents
                    .Object
                    .Select(component => component.Name)
                    .ToList();

                explorer.VbComponents.Verify(c => c.Remove(It.IsAny<IVBComponent>()), Times.Once);

                Assert.IsTrue(modulesNames.Contains("TestDesigner"));
                Assert.IsFalse(modulesNames.Contains("TestModule"));

                //This depends on the setup of Import on the VBComponents mock, which determines the component name from the filename.
                Assert.IsTrue(modulesNames.Contains("StdModule1"));
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void ReplaceProjectContentsFromFiles_DoesNotRemoveDocuments()
        {
            const string path = @"C:\Users\Rubberduck\Desktop\StdModule1.bas";

            using (var explorer = new MockedCodeExplorer(
                    ProjectType.HostProject,
                    ("TestModule", ComponentType.StandardModule, string.Empty),
                    ("TestDocument", ComponentType.Document, string.Empty))
                .ConfigureOpenDialog(new[] { path }, DialogResult.OK)
                .SelectFirstProject())
            {
                explorer.ExecuteReplaceProjectContentsFromFilesCommand();

                var modulesNames = explorer
                    .VbComponents
                    .Object
                    .Select(component => component.Name)
                    .ToList();

                explorer.VbComponents.Verify(c => c.Remove(It.IsAny<IVBComponent>()), Times.Once);

                Assert.IsTrue(modulesNames.Contains("TestDocument"));
                Assert.IsFalse(modulesNames.Contains("TestModule"));

                //This depends on the setup of Import on the VBComponents mock, which determines the component name from the filename.
                Assert.IsTrue(modulesNames.Contains("StdModule1"));
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void ReplaceProjectContentsFromFiles_DocumentWithoutExistingModule_Aborts()
        {
            const string path1 = @"C:\Users\Rubberduck\Desktop\StdModule1.cls";
            const string path2 = @"C:\Users\Rubberduck\Desktop\Class1.cls";
            const string path3 = @"C:\Users\Rubberduck\Desktop\StdModule2.bas";
            const string path4 = @"C:\Users\Rubberduck\Desktop\Class2.doccls";

            using (var explorer = new MockedCodeExplorer(
                ProjectType.HostProject,
                ("TestModule", ComponentType.StandardModule, string.Empty),
                ("OtherTestModule", ComponentType.StandardModule, string.Empty),
                ("TestClass", ComponentType.ClassModule, string.Empty))
                .ConfigureOpenDialog(new[] { path1, path2, path3, path4 }, DialogResult.OK)
                .SelectFirstProject())
            {
                explorer.ExecuteReplaceProjectContentsFromFilesCommand();

                var modulesNames = explorer
                    .VbComponents
                    .Object
                    .Select(component => component.Name)
                    .ToList();

                explorer.VbComponents.Verify(c => c.Remove(It.IsAny<IVBComponent>()), Times.Never);
                explorer.VbComponents.Verify(c => c.Import(It.IsAny<string>()), Times.Never);

                Assert.IsTrue(modulesNames.Contains("OtherTestModule"));
                Assert.IsTrue(modulesNames.Contains("TestModule"));
                Assert.IsTrue(modulesNames.Contains("TestClass"));
                Assert.AreEqual(3, modulesNames.Count);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void ReplaceProjectContentsFromFiles_RemovesMatchingUserFormWithBinary()
        {
            const string path = @"C:\Users\Rubberduck\Desktop\StdModule1.frm";
            const string binaryName = "myBinary.frx";
            var binaryPath = Path.Combine(Path.GetDirectoryName(path), binaryName);

            using (var explorer = new MockedCodeExplorer(
                    ProjectType.HostProject,
                    ("TestModule", ComponentType.UserForm, string.Empty),
                    ("OtherTestModule", ComponentType.StandardModule, string.Empty))
                .ConfigureOpenDialog(new[] { path }, DialogResult.OK)
                .SelectFirstProject())
            {
                var mockExtractor = new Mock<IRequiredBinaryFilesFromFileNameExtractor>();
                mockExtractor
                    .SetupGet(m => m.SupportedComponentTypes)
                    .Returns(new List<ComponentType> { ComponentType.UserForm });
                mockExtractor
                    .Setup(m => m.RequiredBinaryFiles(path, ComponentType.UserForm))
                    .Returns(new List<string> { binaryName });

                var mockFileExistenceChecker = new Mock<IFileExistenceChecker>();
                mockFileExistenceChecker.Setup(m => m.FileExists(binaryPath)).Returns(true);

                explorer.ExecuteReplaceProjectContentsFromFilesCommand(
                    null,
                    null,
                    new List<IRequiredBinaryFilesFromFileNameExtractor> { mockExtractor.Object },
                    mockFileExistenceChecker);

                var modulesNames = explorer
                    .VbComponents
                    .Object
                    .Select(component => component.Name)
                    .ToList();

                explorer.VbComponents.Verify(c => c.Remove(It.IsAny<IVBComponent>()), Times.Exactly(2));
                //This depends on the setup of Import on the VBComponents mock, which determines the component name from the filename.
                Assert.IsTrue(modulesNames.Contains("StdModule1"));
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void ReplaceFromFiles_ModuleThere_DoesNotRemoveMatchingUserFormWithoutBinaryAndUsesSpecialImportMethod()
        {
            const string path = @"C:\Users\Rubberduck\Desktop\StdModule1.frm";
            const string binaryName = "myBinary.frx";
            var binaryPath = Path.Combine(Path.GetDirectoryName(path), binaryName);

            using (var explorer = new MockedCodeExplorer(
                    ProjectType.HostProject,
                    ("TestModule", ComponentType.UserForm, string.Empty),
                    ("OtherTestModule", ComponentType.StandardModule, string.Empty))
                .ConfigureOpenDialog(new[] { path }, DialogResult.OK)
                .SelectFirstProject())
            {
                var mockExtractor = new Mock<IRequiredBinaryFilesFromFileNameExtractor>();
                mockExtractor
                    .SetupGet(m => m.SupportedComponentTypes)
                    .Returns(new List<ComponentType> { ComponentType.UserForm });
                mockExtractor
                    .Setup(m => m.RequiredBinaryFiles(path, ComponentType.UserForm))
                    .Returns(new List<string> { binaryName });

                var mockFileExistenceChecker = new Mock<IFileExistenceChecker>();
                mockFileExistenceChecker.Setup(m => m.FileExists(binaryPath)).Returns(false);

                explorer.ExecuteReplaceProjectContentsFromFilesCommand(
                    filename => filename == path ? "TestModule" : "YetAnotherModule",
                    null,
                    new List<IRequiredBinaryFilesFromFileNameExtractor> { mockExtractor.Object },
                    mockFileExistenceChecker);
                explorer.VbComponents.Verify(c => c.Remove(It.IsAny<IVBComponent>()), Times.Once);
                explorer.VbComponents.Verify(m => m.ImportSourceFile(path), Times.Once);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void ReplaceFromFiles_ModuleThere_ImportsUserFormWithBinary()
        {
            const string path = @"C:\Users\Rubberduck\Desktop\StdModule1.frm";
            const string binaryName = "myBinary.frx";
            var binaryPath = Path.Combine(Path.GetDirectoryName(path), binaryName);

            using (var explorer = new MockedCodeExplorer(
                    ProjectType.HostProject,
                    ("TestModule", ComponentType.UserForm, string.Empty),
                    ("OtherTestModule", ComponentType.StandardModule, string.Empty))
                .ConfigureOpenDialog(new[] { path }, DialogResult.OK)
                .SelectFirstProject())
            {
                var mockExtractor = new Mock<IRequiredBinaryFilesFromFileNameExtractor>();
                mockExtractor
                    .SetupGet(m => m.SupportedComponentTypes)
                    .Returns(new List<ComponentType> { ComponentType.UserForm });
                mockExtractor
                    .Setup(m => m.RequiredBinaryFiles(path, ComponentType.UserForm))
                    .Returns(new List<string> { binaryName });

                var mockFileExistenceChecker = new Mock<IFileExistenceChecker>();
                mockFileExistenceChecker.Setup(m => m.FileExists(binaryPath)).Returns(true);

                explorer.ExecuteReplaceProjectContentsFromFilesCommand(
                    filename => filename == path ? "TestModule" : "YetAnotherModule",
                    null,
                    new List<IRequiredBinaryFilesFromFileNameExtractor> { mockExtractor.Object },
                    mockFileExistenceChecker);

                var modulesNames = explorer
                    .VbComponents
                    .Object
                    .Select(component => component.Name)
                    .ToList();

                explorer.VbComponents.Verify(c => c.Remove(It.IsAny<IVBComponent>()), Times.Exactly(2));
                //This depends on the setup of Import on the VBComponents mock, which determines the component name from the filename.
                Assert.IsTrue(modulesNames.Contains("StdModule1"));
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void ReplaceFromFiles_MultipleImports_RepeatedModuleName_Aborts()
        {
            const string path1 = @"C:\Users\Rubberduck\Desktop\StdModule1.bas";
            const string path2 = @"C:\Users\Rubberduck\Desktop\Class1.cls";
            const string path3 = @"C:\Users\Rubberduck\Desktop\StdModule2.bas";
            const string path4 = @"C:\Users\Rubberduck\Desktop\Class2.cls";

            using (var explorer = new MockedCodeExplorer(
                ProjectType.HostProject,
                ("TestModule", ComponentType.StandardModule, string.Empty),
                ("OtherTestModule", ComponentType.StandardModule, string.Empty),
                ("TestClass", ComponentType.ClassModule, string.Empty))
                .ConfigureOpenDialog(new[] { path1, path2, path3, path4 }, DialogResult.OK)
                .SelectFirstProject())
            {
                explorer.ExecuteReplaceProjectContentsFromFilesCommand(filename =>
                {
                    switch (filename)
                    {
                        case path1:
                            return "TestModule";
                        case path2:
                            return "TestClass";
                        case path3:
                            return "TestModule";
                        case path4:
                            return "NewClass";
                        default:
                            return "YetAnotherModule";
                    }
                });

                var modulesNames = explorer
                    .VbComponents
                    .Object
                    .Select(component => component.Name)
                    .ToList();

                explorer.VbComponents.Verify(c => c.Remove(It.IsAny<IVBComponent>()), Times.Never);
                explorer.VbComponents.Verify(c => c.Import(It.IsAny<string>()), Times.Never);

                Assert.IsTrue(modulesNames.Contains("OtherTestModule"));
                Assert.IsTrue(modulesNames.Contains("TestModule"));
                Assert.IsTrue(modulesNames.Contains("TestClass"));
                Assert.AreEqual(3, modulesNames.Count);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void ReplaceFromFiles_NonMatchingComponentTypeForFormWithoutBinary_Aborts()
        {
            const string path1 = @"C:\Users\Rubberduck\Desktop\StdModule1.cls";
            const string path2 = @"C:\Users\Rubberduck\Desktop\Form1.frm";
            const string binaryName = "myBinary.frx";
            var binaryPath = Path.Combine(Path.GetDirectoryName(path2), binaryName);
            const string path3 = @"C:\Users\Rubberduck\Desktop\StdModule2.bas";
            const string path4 = @"C:\Users\Rubberduck\Desktop\Class2.cls";

            using (var explorer = new MockedCodeExplorer(
                ProjectType.HostProject,
                ("TestModule", ComponentType.StandardModule, string.Empty),
                ("OtherTestModule", ComponentType.StandardModule, string.Empty),
                ("TestClass", ComponentType.ClassModule, string.Empty))
                .ConfigureOpenDialog(new[] { path1, path2, path3, path4 }, DialogResult.OK)
                .SelectFirstProject())
            {
                var mockExtractor = new Mock<IRequiredBinaryFilesFromFileNameExtractor>();
                mockExtractor
                    .SetupGet(m => m.SupportedComponentTypes)
                    .Returns(new List<ComponentType> { ComponentType.UserForm });
                mockExtractor
                    .Setup(m => m.RequiredBinaryFiles(path2, ComponentType.UserForm))
                    .Returns(new List<string> { binaryName });

                var mockFileExistenceChecker = new Mock<IFileExistenceChecker>();
                mockFileExistenceChecker.Setup(m => m.FileExists(binaryPath)).Returns(false);

                explorer.ExecuteReplaceProjectContentsFromFilesCommand(filename =>
                {
                    switch (filename)
                    {
                        case path1:
                            return "TestModule";
                        case path2:
                            return "TestClass";
                        case path3:
                            return "NewModule";
                        case path4:
                            return "NewClass";
                        default:
                            return "YetAnotherModule";
                    }
                },
                    null,
                    new List<IRequiredBinaryFilesFromFileNameExtractor> { mockExtractor.Object },
                    mockFileExistenceChecker);

                var modulesNames = explorer
                    .VbComponents
                    .Object
                    .Select(component => component.Name)
                    .ToList();

                explorer.VbComponents.Verify(c => c.Remove(It.IsAny<IVBComponent>()), Times.Never);
                explorer.VbComponents.Verify(c => c.Import(It.IsAny<string>()), Times.Never);

                Assert.IsTrue(modulesNames.Contains("OtherTestModule"));
                Assert.IsTrue(modulesNames.Contains("TestModule"));
                Assert.IsTrue(modulesNames.Contains("TestClass"));
                Assert.AreEqual(3, modulesNames.Count);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void ReplaceFromFiles_NonMatchingComponentTypeForDocument_Aborts()
        {
            const string path1 = @"C:\Users\Rubberduck\Desktop\StdModule1.cls";
            const string path2 = @"C:\Users\Rubberduck\Desktop\Document.doccls";
            const string path3 = @"C:\Users\Rubberduck\Desktop\StdModule2.bas";
            const string path4 = @"C:\Users\Rubberduck\Desktop\Class1.cls";

            using (var explorer = new MockedCodeExplorer(
                ProjectType.HostProject,
                ("TestModule", ComponentType.StandardModule, string.Empty),
                ("OtherTestModule", ComponentType.StandardModule, string.Empty),
                ("TestClass", ComponentType.ClassModule, string.Empty))
                .ConfigureOpenDialog(new[] { path1, path2, path3, path4 }, DialogResult.OK)
                .SelectFirstProject())
            {
                explorer.ExecuteReplaceProjectContentsFromFilesCommand(filename =>
                {
                    switch (filename)
                    {
                        case path1:
                            return "TestModule";
                        case path2:
                            return "TestClass";
                        case path3:
                            return "NewModule";
                        case path4:
                            return "NewDocument";
                        default:
                            return "YetAnotherModule";
                    }
                });

                var modulesNames = explorer
                    .VbComponents
                    .Object
                    .Select(component => component.Name)
                    .ToList();

                explorer.VbComponents.Verify(c => c.Remove(It.IsAny<IVBComponent>()), Times.Never);
                explorer.VbComponents.Verify(c => c.Import(It.IsAny<string>()), Times.Never);

                Assert.IsTrue(modulesNames.Contains("OtherTestModule"));
                Assert.IsTrue(modulesNames.Contains("TestModule"));
                Assert.IsTrue(modulesNames.Contains("TestClass"));
                Assert.AreEqual(3, modulesNames.Count);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void ReplaceFromFiles_UserFormWithoutExistingFormOrBinary_Aborts()
        {
            const string path1 = @"C:\Users\Rubberduck\Desktop\StdModule1.cls";
            const string path2 = @"C:\Users\Rubberduck\Desktop\Class1.cls";
            const string path3 = @"C:\Users\Rubberduck\Desktop\StdModule2.bas";
            const string path4 = @"C:\Users\Rubberduck\Desktop\Class2.frm";
            const string binaryName = "myBinary.frx";
            var binaryPath = Path.Combine(Path.GetDirectoryName(path4), binaryName);

            using (var explorer = new MockedCodeExplorer(
                ProjectType.HostProject,
                ("TestModule", ComponentType.StandardModule, string.Empty),
                ("OtherTestModule", ComponentType.StandardModule, string.Empty),
                ("TestClass", ComponentType.ClassModule, string.Empty))
                .ConfigureOpenDialog(new[] { path1, path2, path3, path4 }, DialogResult.OK)
                .SelectFirstProject())
            {
                var mockExtractor = new Mock<IRequiredBinaryFilesFromFileNameExtractor>();
                mockExtractor
                    .SetupGet(m => m.SupportedComponentTypes)
                    .Returns(new List<ComponentType> { ComponentType.UserForm });
                mockExtractor
                    .Setup(m => m.RequiredBinaryFiles(path4, ComponentType.UserForm))
                    .Returns(new List<string> { binaryName });

                var mockFileExistenceChecker = new Mock<IFileExistenceChecker>();
                mockFileExistenceChecker.Setup(m => m.FileExists(binaryPath)).Returns(false);

                explorer.ExecuteReplaceProjectContentsFromFilesCommand(filename =>
                {
                    switch (filename)
                    {
                        case path1:
                            return "TestModule";
                        case path2:
                            return "TestClass";
                        case path3:
                            return "NewModule";
                        case path4:
                            return "NewForm";
                        default:
                            return "YetAnotherModule";
                    }
                },
                    null,
                    new List<IRequiredBinaryFilesFromFileNameExtractor> { mockExtractor.Object },
                    mockFileExistenceChecker);

                var modulesNames = explorer
                    .VbComponents
                    .Object
                    .Select(component => component.Name)
                    .ToList();

                explorer.VbComponents.Verify(c => c.Remove(It.IsAny<IVBComponent>()), Times.Never);
                explorer.VbComponents.Verify(c => c.Import(It.IsAny<string>()), Times.Never);

                Assert.IsTrue(modulesNames.Contains("OtherTestModule"));
                Assert.IsTrue(modulesNames.Contains("TestModule"));
                Assert.IsTrue(modulesNames.Contains("TestClass"));
                Assert.AreEqual(3, modulesNames.Count);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void ReplaceFromFiles_DocumentWithoutExistingModule_Aborts()
        {
            const string path1 = @"C:\Users\Rubberduck\Desktop\StdModule1.cls";
            const string path2 = @"C:\Users\Rubberduck\Desktop\Class1.cls";
            const string path3 = @"C:\Users\Rubberduck\Desktop\StdModule2.bas";
            const string path4 = @"C:\Users\Rubberduck\Desktop\Class2.doccls";

            using (var explorer = new MockedCodeExplorer(
                ProjectType.HostProject,
                ("TestModule", ComponentType.StandardModule, string.Empty),
                ("OtherTestModule", ComponentType.StandardModule, string.Empty),
                ("TestClass", ComponentType.ClassModule, string.Empty))
                .ConfigureOpenDialog(new[] { path1, path2, path3, path4 }, DialogResult.OK)
                .SelectFirstProject())
            {
                explorer.ExecuteReplaceProjectContentsFromFilesCommand(filename =>
                {
                    switch (filename)
                    {
                        case path1:
                            return "TestModule";
                        case path2:
                            return "TestClass";
                        case path3:
                            return "NewModule";
                        case path4:
                            return "NewDocument";
                        default:
                            return "YetAnotherModule";
                    }
                });

                var modulesNames = explorer
                    .VbComponents
                    .Object
                    .Select(component => component.Name)
                    .ToList();

                explorer.VbComponents.Verify(c => c.Remove(It.IsAny<IVBComponent>()), Times.Never);
                explorer.VbComponents.Verify(c => c.Import(It.IsAny<string>()), Times.Never);

                Assert.IsTrue(modulesNames.Contains("OtherTestModule"));
                Assert.IsTrue(modulesNames.Contains("TestModule"));
                Assert.IsTrue(modulesNames.Contains("TestClass"));
                Assert.AreEqual(3, modulesNames.Count);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void ReplaceProjectContentsFromFiles_Cancel()
        {
            const string path = @"C:\Users\Rubberduck\Desktop\StdModule1.bas";

            using (var explorer = new MockedCodeExplorer(
                ProjectType.HostProject,
                ("TestModule", ComponentType.StandardModule, string.Empty))
                .ConfigureOpenDialog(new[] { path }, DialogResult.Cancel)
                .SelectFirstProject())
            {
                explorer.ExecuteReplaceProjectContentsFromFilesCommand();
                explorer.VbComponents.Verify(c => c.Import(It.IsAny<string>()), Times.Never);
                explorer.VbComponents.Verify(c => c.Remove(It.IsAny<IVBComponent>()), Times.Never);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void ExportModule_ExpectExecution()
        {
            const string folder = @"C:\Users\Rubberduck\Desktop";
            const string filename = "StdModule1.bas";
            var path = Path.Combine(folder, filename);

            using (var explorer = new MockedCodeExplorer(ProjectType.HostProject)
                .ConfigureSaveDialog(path, DialogResult.OK)
                .SelectFirstModule())
            {
                explorer.VbComponent.Setup(c => c.ExportAsSourceFile(folder, It.IsAny<bool>(), It.IsAny<bool>()));
                explorer.ExecuteExportCommand();
                explorer.VbComponent.Verify(c => c.ExportAsSourceFile(folder, false, true), Times.Once);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void ExportModule_CancelPressed_ExpectNoExecution()
        {
            const string path = @"C:\Users\Rubberduck\Desktop\StdModule1.bas";

            using (var explorer = new MockedCodeExplorer(ProjectType.HostProject)
                .ConfigureSaveDialog(path, DialogResult.Cancel)
                .SelectFirstModule())
            {
                explorer.VbComponent.Setup(c => c.Export(path));
                explorer.VbComponent.Setup(c => c.ExportAsSourceFile(path, It.IsAny<bool>(), It.IsAny<bool>()));
                explorer.ExecuteExportCommand();
                explorer.VbComponent.Verify(c => c.Export(path), Times.Never);
                explorer.VbComponent.Verify(c => c.ExportAsSourceFile(path, It.IsAny<bool>(), It.IsAny<bool>()), Times.Never);
            }
        }

        [Category("Commands")]
        [Test]
        public void ExportProject_TestCanExecute_ExpectTrue()
        {
            const string selected = @"C:\Users\Rubberduck\Desktop\ExportAll";

            using (var explorer = new MockedCodeExplorer(ProjectType.HostProject)
                .ImplementExportAllCommand()
                .ConfigureFolderBrowser(selected, DialogResult.OK)
                .SelectFirstProject())
            {
                Assert.IsTrue(explorer.ViewModel.ExportAllCommand.CanExecute(explorer.ViewModel.SelectedItem));
            }
        }

        [Category("Commands")]
        [Test]
        public void ExportProject_TestExecute_OKPressed_ExpectExecution()
        {
            const string selected = @"C:\Users\Rubberduck\Desktop\ExportAll";
            const string result = @"C:\Users\Rubberduck\Documents\Subfolder\Project.xlsm";

            var modules = new[]
            {
                ComponentType.StandardModule, ComponentType.ClassModule, ComponentType.Document, ComponentType.UserControl
            };

            using (var explorer = new MockedCodeExplorer(ProjectType.HostProject, modules)
                .ConfigureFolderBrowser(selected, DialogResult.OK)
                .SelectFirstProject())
            {
                explorer.VbProject.SetupGet(m => m.IsSaved).Returns(true);
                explorer.VbProject.SetupGet(m => m.FileName).Returns(result);
                explorer.ExecuteExportAllCommand();
                explorer.VbProject.Verify(m => m.ExportSourceFiles(selected), Times.Once);
            }
        }

        [Category("Commands")]
        [Test]
        public void ExportProject_TestExecute_CancelPressed_ExpectExecution()
        {
            const string selected = @"C:\Users\Rubberduck\Desktop\ExportAll";
            const string result = @"C:\Users\Rubberduck\Documents\Subfolder\Project.xlsm";

            var modules = new[]
            {
                ComponentType.StandardModule, ComponentType.ClassModule, ComponentType.Document, ComponentType.UserControl
            };

            using (var explorer = new MockedCodeExplorer(ProjectType.HostProject, modules)
                .ConfigureFolderBrowser(selected, DialogResult.Cancel)
                .SelectFirstProject())
            {
                explorer.VbProject.SetupGet(m => m.IsSaved).Returns(true);
                explorer.VbProject.SetupGet(m => m.FileName).Returns(result);
                explorer.ExecuteExportAllCommand();
                explorer.VbProject.Verify(m => m.ExportSourceFiles(selected), Times.Never);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void OpenDesigner()
        {
            using (var explorer = new MockedCodeExplorer(ProjectType.HostProject, new[] { ComponentType.UserForm })
                .SelectFirstModule())
            {
                explorer.ExecuteOpenDesignerCommand();
                explorer.VbComponent.Verify(c => c.DesignerWindow(), Times.Once);
                Assert.IsTrue(explorer.VbComponent.Object.DesignerWindow().IsVisible);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void RemoveCommand_RemovesModuleWhenPromptOk()
        {
            const string path = @"C:\Users\Rubberduck\Desktop\StdModule1.bas";

            using (var explorer = new MockedCodeExplorer(ProjectType.HostProject)
                .ConfigureMessageBox(ConfirmationOutcome.Yes)
                .ConfigureSaveDialog(path, DialogResult.OK)
                .SelectFirstModule())
            {
                var removing = explorer.ViewModel.SelectedItem;
                var component = explorer.VbComponent.Object;

                explorer.ViewModel.RemoveCommand.Execute(removing);
                explorer.ProjectsRepository.Verify(c => c.RemoveComponent(component.QualifiedModuleName), Times.Once);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void RemoveCommand_CancelsWhenFilePromptCancels()
        {
            const string path = @"C:\Users\Rubberduck\Desktop\StdModule1.bas";

            using (var explorer = new MockedCodeExplorer(ProjectType.HostProject)
                .ConfigureSaveDialog(path, DialogResult.Cancel)
                .SelectFirstModule())
            {
                explorer.MessageBox.Setup(m => m.ConfirmYesNo(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<bool>())).Returns(true);

                var removing = explorer.ViewModel.SelectedItem;
                var component = explorer.VbComponent.Object;

                explorer.ViewModel.RemoveCommand.Execute(removing);
                explorer.VbComponents.Verify(c => c.Remove(component), Times.Never);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void RemoveCommand_GivenMsgBoxNo_RemovesModuleNoExport()
        {
            using (var explorer = new MockedCodeExplorer(ProjectType.HostProject, new[] { ComponentType.UserForm })
                .ConfigureMessageBox(ConfirmationOutcome.No)
                .SelectFirstModule())
            {

                var removing = explorer.ViewModel.SelectedItem;
                var component = explorer.VbComponent.Object;


                explorer.ViewModel.RemoveCommand.Execute(removing);
                explorer.ProjectsRepository.Verify(c => c.RemoveComponent(component.QualifiedModuleName), Times.Once);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void RemoveModule_Cancel()
        {
            using (var explorer = new MockedCodeExplorer(ProjectType.HostProject, new[] { ComponentType.UserForm })
                .ConfigureMessageBox(ConfirmationOutcome.Cancel)
                .SelectFirstModule())
            {
                var removing = explorer.ViewModel.SelectedItem;
                var component = explorer.VbComponent.Object;

                explorer.ViewModel.RemoveCommand.Execute(removing);
                explorer.VbComponents.Verify(c => c.Remove(component), Times.Never);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void IndentModule()
        {
            const string inputCode =
@"Sub Foo()
Dim d As Boolean
d = True
End Sub";

            const string expectedCode =
@"Sub Foo()
    Dim d As Boolean
    d = True
End Sub
";

            using (var explorer = new MockedCodeExplorer(inputCode)
                .SelectFirstModule())
            {
                explorer.ExecuteIndenterCommand();
                var actualCode = explorer.VbComponent.Object.CodeModule.Content();
                Assert.AreEqual(expectedCode, actualCode);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void IndentModule_DisabledWithNoIndentAnnotation()
        {
            const string inputCode =
@"'@NoIndent

Sub Foo()
Dim d As Boolean
d = True
End Sub";

            using (var explorer = new MockedCodeExplorer(inputCode)
                .ImplementIndenterCommand()
                .SelectFirstModule())
            {
                Assert.IsFalse(explorer.ViewModel.IndenterCommand.CanExecute(explorer.ViewModel.SelectedItem));
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void IndentProject()
        {
            const string inputCode =
@"Sub Foo()
Dim d As Boolean
d = True
End Sub";

            const string expectedCode =
@"Sub Foo()
    Dim d As Boolean
    d = True
End Sub
";

            using (var explorer = new MockedCodeExplorer(ProjectType.HostProject, new[] { ComponentType.StandardModule, ComponentType.ClassModule }, new[] { inputCode, inputCode })
                .SelectFirstProject())
            {
                var module1 = explorer.VbComponents.Object[0].CodeModule;
                var module2 = explorer.VbComponents.Object[1].CodeModule;

                explorer.ExecuteIndenterCommand();

                Assert.AreEqual(expectedCode, module1.Content());
                Assert.AreEqual(expectedCode, module2.Content());
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void IndentProject_IndentsModulesWithoutNoIndentAnnotation()
        {
            const string inputCode1 =
@"Sub Foo()
Dim d As Boolean
d = True
End Sub";

            const string inputCode2 =
@"'@NoIndent

Sub Foo()
Dim d As Boolean
d = True
End Sub";

            const string expectedCode =
@"Sub Foo()
    Dim d As Boolean
    d = True
End Sub
";

            using (var explorer = new MockedCodeExplorer(ProjectType.HostProject, new[] { ComponentType.StandardModule, ComponentType.ClassModule }, new[] { inputCode1, inputCode2 })
                .SelectFirstProject())
            {
                var module1 = explorer.VbComponents.Object[0].CodeModule;
                var module2 = explorer.VbComponents.Object[1].CodeModule;

                explorer.ExecuteIndenterCommand();

                Assert.AreEqual(expectedCode, module1.Content());
                Assert.AreEqual(inputCode2, module2.Content());
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void IndentProject_DisabledWhenAllModulesHaveNoIndentAnnotation()
        {
            const string inputCode =
@"'@NoIndent

Sub Foo()
Dim d As Boolean
d = True
End Sub";

            using (var explorer = new MockedCodeExplorer(ProjectType.HostProject, new[] { ComponentType.StandardModule, ComponentType.ClassModule }, new[] { inputCode, inputCode })
                .ImplementIndenterCommand()
                .SelectFirstProject())
            {
                Assert.IsFalse(explorer.ViewModel.IndenterCommand.CanExecute(explorer.ViewModel.SelectedItem));
            }
        }
               
        [Category("Code Explorer")]
        [Test]
        public void IndentFolder()
        {
            const string inputCode =
@"'@Folder ""folder""

Sub Foo()
Dim d As Boolean
d = True
End Sub";

            const string expectedCode =
@"'@Folder ""folder""

Sub Foo()
    Dim d As Boolean
    d = True
End Sub
";

            using (var explorer = new MockedCodeExplorer(ProjectType.HostProject, new[] { ComponentType.StandardModule, ComponentType.ClassModule }, new[] { inputCode, inputCode })
                .SelectFirstCustomFolder())
            {
                var module1 = explorer.VbComponents.Object[0].CodeModule;
                var module2 = explorer.VbComponents.Object[1].CodeModule;

                explorer.ExecuteIndenterCommand();

                Assert.AreEqual(expectedCode, module1.Content());
                Assert.AreEqual(expectedCode, module2.Content());
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void IndentFolder_IndentsModulesWithoutNoIndentAnnotation()
        {
            const string inputCode1 =
@"'@Folder ""folder""

Sub Foo()
Dim d As Boolean
d = True
End Sub";

            const string inputCode2 =
@"'@NoIndent
'@Folder ""folder""

Sub Foo()
Dim d As Boolean
d = True
End Sub";

            const string expectedCode =
@"'@Folder ""folder""

Sub Foo()
    Dim d As Boolean
    d = True
End Sub
";

            using (var explorer = new MockedCodeExplorer(ProjectType.HostProject, new[] { ComponentType.StandardModule, ComponentType.ClassModule }, new[] { inputCode1, inputCode2 })
                .SelectFirstCustomFolder())
            {
                var module1 = explorer.VbComponents.Object[0].CodeModule;
                var module2 = explorer.VbComponents.Object[1].CodeModule;

                explorer.ExecuteIndenterCommand();

                Assert.AreEqual(expectedCode, module1.Content());
                Assert.AreEqual(inputCode2, module2.Content());
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void IndentFolder_DisabledWhenAllModulesHaveNoIndentAnnotation()
        {
            const string inputCode =
@"'@NoIndent
'@Folder ""folder""

Sub Foo()
Dim d As Boolean
d = True
End Sub";

            using (var explorer = new MockedCodeExplorer(ProjectType.HostProject, new[] { ComponentType.StandardModule, ComponentType.ClassModule }, new[] { inputCode, inputCode })
                .ImplementIndenterCommand()
                .SelectFirstCustomFolder())
            {
                Assert.IsFalse(explorer.ViewModel.IndenterCommand.CanExecute(explorer.ViewModel.SelectedItem));
            }
        }

        private IEnumerable<bool> GetNodeExpandedStates(ICodeExplorerNode root)
        {
            yield return root.IsExpanded;
            foreach (var node in root.Children)
            {
                foreach (var state in GetNodeExpandedStates(node))
                {
                    yield return state;
                }
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void ExpandAllNodes()
        {
            const string inputCode =
@"Sub Foo()
End Sub";

            using (var explorer = new MockedCodeExplorer(inputCode)
                .SelectFirstProject())
            {
                var node = explorer.ViewModel.SelectedItem;
                explorer.ViewModel.ExpandAllSubnodesCommand.Execute(node);
                Assert.IsTrue(GetNodeExpandedStates(node).All(state => state));
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void ExpandAllNodes_StartingWithSubNode()
        {
            const string foo = @"'@Folder ""Foo""";
            const string bar = @"'@Folder ""Bar""";

            using (var explorer = new MockedCodeExplorer(ProjectType.HostProject, new[] { ComponentType.StandardModule, ComponentType.ClassModule }, new[] { foo, bar })
                .SelectFirstCustomFolder())
            {
                var expanded = explorer.ViewModel.SelectedItem;
                var collapsed = explorer.ViewModel.Projects.Single().Children.Last();

                expanded.IsExpanded = true;
                collapsed.IsExpanded = false;

                explorer.ViewModel.ExpandAllSubnodesCommand.Execute(expanded);

                Assert.IsTrue(GetNodeExpandedStates(expanded).All(state => state));
                Assert.IsFalse(GetNodeExpandedStates(collapsed).All(state => state));
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void CollapseAllNodes()
        {
            const string inputCode =
@"Sub Foo()
End Sub";

            using (var explorer = new MockedCodeExplorer(inputCode)
                .SelectFirstProject())
            {
                var node = explorer.ViewModel.SelectedItem;
                explorer.ViewModel.ExpandAllSubnodesCommand.Execute(node);
                explorer.ViewModel.CollapseAllSubnodesCommand.Execute(node);

                Assert.IsFalse(GetNodeExpandedStates(node).All(state => state));
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void CollapseAllNodes_StartingWithSubNode()
        {
            const string foo = @"'@Folder ""Foo""";
            const string bar = @"'@Folder ""Bar""";

            using (var explorer = new MockedCodeExplorer(ProjectType.HostProject, new[] { ComponentType.StandardModule, ComponentType.ClassModule }, new[] { foo, bar })
                .SelectFirstProject())
            {
                explorer.ViewModel.ExpandAllSubnodesCommand.Execute(explorer.ViewModel.SelectedItem);
                var expanded = explorer.ViewModel.Projects.Single().Children.Last();

                explorer.SelectFirstCustomFolder();
                var collapsed = explorer.ViewModel.SelectedItem;
                explorer.ViewModel.CollapseAllSubnodesCommand.Execute(collapsed);

                Assert.IsTrue(GetNodeExpandedStates(expanded).All(state => state));
                Assert.IsFalse(GetNodeExpandedStates(collapsed).All(state => state));
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void UnparsedSetToTrue_NoProjects()
        {
            var builder = new MockVbeBuilder();
            var vbe = builder.Build();
            var vbeEvents = MockVbeEvents.CreateMockVbeEvents(vbe);
            var parser = MockParser.Create(vbe.Object, null, MockVbeEvents.CreateMockVbeEvents(vbe));
            var state = parser.State;
            var dispatcher = new Mock<IUiDispatcher>();
            var generalSettingsProvider = new Mock<IConfigurationService<GeneralSettings>>();
            var generalSettings = new GeneralSettings();
            generalSettingsProvider.Setup(m => m.Read()).Returns(generalSettings);

            dispatcher.Setup(m => m.Invoke(It.IsAny<Action>())).Callback((Action argument) => argument.Invoke());
            dispatcher.Setup(m => m.StartTask(It.IsAny<Action>(), It.IsAny<TaskCreationOptions>())).Returns((Action argument, TaskCreationOptions options) => Task.Factory.StartNew(argument.Invoke, options));

            var viewModel = new CodeExplorerViewModel(state, null, generalSettingsProvider.Object, null, dispatcher.Object, vbe.Object, null,
                new CodeExplorerSyncProvider(vbe.Object, state, vbeEvents.Object), new List<IAnnotation>());

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error)
            {
                Assert.Inconclusive("Parser Error");
            }

            Assert.IsTrue(viewModel.Unparsed);
        }
    }
}
