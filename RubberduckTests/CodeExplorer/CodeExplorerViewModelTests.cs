using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using NUnit.Framework;
using Moq;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.Interaction;
using Rubberduck.Parsing.UIContext;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.CodeExplorer.Commands;
using Rubberduck.UI.Command;
using RubberduckTests.Mocks;
using MessageBox = System.Windows.MessageBox;

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
            using (var explorer = new MockedCodeExplorer(ProjectType.HostProject).SelectFirstModule())
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
            using (var explorer = new MockedCodeExplorer(ProjectType.HostProject).SelectFirstModule())
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
            using (var explorer = new MockedCodeExplorer(ProjectType.HostProject).SelectFirstModule())
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
            using (var explorer = new MockedCodeExplorer(ProjectType.HostProject).SelectFirstModule())
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
            using (var explorer = new MockedCodeExplorer(ProjectType.HostProject).SelectFirstModule())
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
        public void ImportModule_Cancel()
        {
            const string path = @"C:\Users\Rubberduck\Desktop\StdModule1.bas";

            using (var explorer = new MockedCodeExplorer(ProjectType.HostProject)
                .ConfigureOpenDialog(new[] { path }, DialogResult.Cancel)
                .SelectFirstProject())
            {
                explorer.ExecuteImportCommand();
                explorer.VbComponents.Verify(c => c.Import(path), Times.Never);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void ExportModule_ExpectExecution()
        {
            const string path = @"C:\Users\Rubberduck\Desktop\StdModule1.bas";

            using (var explorer = new MockedCodeExplorer(ProjectType.HostProject)
                .ConfigureSaveDialog(path, DialogResult.OK)
                .SelectFirstModule())
            {
                explorer.ExecuteExportCommand();
                explorer.VbComponent.Verify(c => c.Export(path), Times.Once);
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
                explorer.ExecuteExportCommand();
                explorer.VbComponent.Verify(c => c.Export(path), Times.Never);
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
                explorer.VbComponents.Verify(c => c.Remove(component), Times.Once);
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
                explorer.VbComponents.Verify(c => c.Remove(component), Times.Once);
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
                Assert.AreEqual(expectedCode, explorer.VbComponent.Object.CodeModule.Content());
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
            var parser = MockParser.Create(vbe.Object, null, MockVbeEvents.CreateMockVbeEvents(vbe));
            var state = parser.State;
            var dispatcher = new Mock<IUiDispatcher>();

            dispatcher.Setup(m => m.Invoke(It.IsAny<Action>())).Callback((Action argument) => argument.Invoke());

            var viewModel = new CodeExplorerViewModel(state, null, null, null, dispatcher.Object, vbe.Object, null, new CodeExplorerSyncProvider(vbe.Object, state));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error)
            {
                Assert.Inconclusive("Parser Error");
            }

            Assert.IsTrue(viewModel.Unparsed);
        }
    }
}
