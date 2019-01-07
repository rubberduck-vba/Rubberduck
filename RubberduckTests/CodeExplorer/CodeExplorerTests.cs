using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using NUnit.Framework;
using Moq;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Navigation.Folders;
using Rubberduck.Parsing.VBA;
using Rubberduck.Settings;
using Rubberduck.SmartIndenter;
using Rubberduck.UI;
using Rubberduck.UI.CodeExplorer.Commands;
using Rubberduck.UI.Command;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;
using Rubberduck.Parsing.UIContext;
using Rubberduck.SettingsProvider;
using Rubberduck.Interaction;
using Rubberduck.UI.UnitTesting.Commands;
using Rubberduck.UnitTesting;

namespace RubberduckTests.CodeExplorer
{
    [TestFixture]
    public class CodeExplorerTests
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

            var modules = new []
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

            using (var explorer = new MockedCodeExplorer(ProjectType.HostProject, new [] { ComponentType.StandardModule, ComponentType.ClassModule }, new [] { inputCode, inputCode })
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

        private IEnumerable<bool> GetNodeExpandedStates(CodeExplorerItemViewModel root)
        {
            yield return root.IsExpanded;
            foreach (var node in root.Items)
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
                var collapsed = explorer.ViewModel.Projects.Single().Items.Last();

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
                var expanded = explorer.ViewModel.Projects.Single().Items.Last();

                explorer.SelectFirstCustomFolder();
                var collapsed = explorer.ViewModel.SelectedItem;
                explorer.ViewModel.CollapseAllSubnodesCommand.Execute(collapsed);

                Assert.IsTrue(GetNodeExpandedStates(expanded).All(state => state));
                Assert.IsFalse(GetNodeExpandedStates(collapsed).All(state => state));
            }
        }

        [Category("Code Explorer")]
        [Test]
        [TestCase(false, true)]
        [TestCase(true, false)]
        [TestCase(false, false)]
        [TestCase(true, true)]
        public void SetSortByNameCommand_LinkedToViewModel(bool name, bool code)
        {
            using (var explorer = new MockedCodeExplorer(ProjectType.HostProject, new[] { ComponentType.StandardModule, ComponentType.ClassModule })
                .SelectFirstCustomFolder())
            {
                var view = explorer.ViewModel;

                var settings = explorer.WindowSettings;
                settings.CodeExplorer_SortByName = name;
                settings.CodeExplorer_SortByCodeOrder = code;

                view.SetNameSortCommand.Execute(true);

                Assert.IsTrue(view.SortByName);
                Assert.IsFalse(view.SortByCodeOrder);
            }
        }

        [Category("Code Explorer")]
        [Test]
        [TestCase(false, true)]
        [TestCase(true, false)]
        [TestCase(false, false)]
        [TestCase(true, true)]
        public void SetSortByCodeOrder_LinkedToViewModel(bool name, bool code)
        {
            using (var explorer = new MockedCodeExplorer(ProjectType.HostProject, new[] { ComponentType.StandardModule, ComponentType.ClassModule })
                .SelectFirstCustomFolder())
            {
                var view = explorer.ViewModel;

                var settings = explorer.WindowSettings;
                settings.CodeExplorer_SortByName = name;
                settings.CodeExplorer_SortByCodeOrder = code;

                view.SetCodeOrderSortCommand.Execute(true);

                Assert.IsTrue(view.SortByCodeOrder);
                Assert.IsFalse(view.SortByName);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void CompareByName_ReturnsZeroForIdenticalNodes()
        {
            var folderNode = new CodeExplorerCustomFolderViewModel(null, "Name", "Name", null, null);
            Assert.AreEqual(0, new CompareByName().Compare(folderNode, folderNode));
        }

        [Category("Code Explorer")]
        [Test]
        public void CompareByName_ReturnsZeroForIdenticalNames()
        {
            // this won't happen, but just to be thorough...--besides, it is good for the coverage
            var folderNode1 = new CodeExplorerCustomFolderViewModel(null, "Name", "Name", null, null);
            var folderNode2 = new CodeExplorerCustomFolderViewModel(null, "Name", "Name", null, null);

            Assert.AreEqual(0, new CompareByName().Compare(folderNode1, folderNode2));
        }

        [Category("Code Explorer")]
        [Test]
        public void CompareByName_ReturnsCorrectOrdering()
        {
            // this won't happen, but just to be thorough...--besides, it is good for the coverage
            var folderNode1 = new CodeExplorerCustomFolderViewModel(null, "Name1", "Name1", null, null);
            var folderNode2 = new CodeExplorerCustomFolderViewModel(null, "Name2", "Name2", null, null);

            Assert.IsTrue(new CompareByName().Compare(folderNode1, folderNode2) < 0);
        }

        [Category("Code Explorer")]
        [Test]
        public void CompareByType_ReturnsZeroForIdenticalNodes()
        {
            var errorNode = new CodeExplorerCustomFolderViewModel(null, "Name", "folder1.folder2", null, null);
            Assert.AreEqual(0, new CompareByName().Compare(errorNode, errorNode));
        }

        [Category("Code Explorer")]
        [Test]
        public void CompareByType_ReturnsEventAboveConst()
        {
            const string inputCode =
@"Public Event Foo(ByVal arg1 As Integer, ByVal arg2 As String)
Public Const Bar = 0";

            using (var explorer = new MockedCodeExplorer(inputCode)
                .SelectFirstModule())
            {
                var module = explorer.ViewModel.SelectedItem;
                var eventNode = module.Items.Single(s => s.Name == "Foo");
                var constNode = module.Items.Single(s => s.Name == "Bar = 0");

                Assert.AreEqual(-1, new CompareByType().Compare(eventNode, constNode));
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void CompareByType_ReturnsConstAboveField()
        {
            const string inputCode =
@"Public Const Foo = 0
Public Bar As Boolean";

            using (var explorer = new MockedCodeExplorer(inputCode)
                .SelectFirstModule())
            {
                var module = explorer.ViewModel.SelectedItem;
                var constNode = module.Items.Single(s => s.Name == "Foo = 0");
                var fieldNode = module.Items.Single(s => s.Name == "Bar");

                Assert.AreEqual(-1, new CompareByType().Compare(constNode, fieldNode));
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void CompareByType_ReturnsFieldAbovePropertyGet()
        {
            const string inputCode =
@"Private Bar As Boolean

Public Property Get Foo() As Variant
End Property
";

            using (var explorer = new MockedCodeExplorer(inputCode)
                .SelectFirstModule())
            {
                var module = explorer.ViewModel.SelectedItem;
                var fieldNode = module.Items.Single(s => s.Name == "Bar");
                var propertyGetNode = module.Items.Single(s => s.Name == "Foo (Get)");

                Assert.AreEqual(-1, new CompareByType().Compare(fieldNode, propertyGetNode));
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void CompareByType_ReturnsPropertyGetEqualToPropertyLet()
        {
            const string inputCode =
@"Public Property Get Foo() As Variant
End Property

Public Property Let Foo(ByVal Value As Variant)
End Property
";

            using (var explorer = new MockedCodeExplorer(inputCode)
                .SelectFirstModule())
            {
                var module = explorer.ViewModel.SelectedItem;
                var propertyGetNode = module.Items.Single(s => s.Name == "Foo (Get)");
                var propertyLetNode = module.Items.Single(s => s.Name == "Foo (Let)");

                Assert.AreEqual(0, new CompareByType().Compare(propertyGetNode, propertyLetNode));
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void CompareByType_ReturnsPropertyGetEqualToPropertySet()
        {
            const string inputCode =
@"Public Property Get Foo() As Variant
End Property

Public Property Set Foo(ByVal Value As Variant)
End Property
";

            using (var explorer = new MockedCodeExplorer(inputCode)
                .SelectFirstModule())
            {
                var module = explorer.ViewModel.SelectedItem;
                var propertyGetNode = module.Items.Single(s => s.Name == "Foo (Get)");
                var propertyLetNode = module.Items.Single(s => s.Name == "Foo (Set)");

                Assert.AreEqual(0, new CompareByType().Compare(propertyGetNode, propertyLetNode));
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void CompareByType_ReturnsPropertyLetEqualToPropertyGet()
        {
            const string inputCode =
@"Public Property Let Foo(ByVal Value As Variant)
End Property

Public Property Get Foo() As Variant
End Property
";

            using (var explorer = new MockedCodeExplorer(inputCode)
                .SelectFirstModule())
            {
                var module = explorer.ViewModel.SelectedItem;
                var propertyLetNode = module.Items.Single(s => s.Name == "Foo (Let)");
                var propertySetNode = module.Items.Single(s => s.Name == "Foo (Get)");

                Assert.AreEqual(0, new CompareByType().Compare(propertyLetNode, propertySetNode));
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void CompareByType_ReturnsPropertyLetEqualToPropertySet()
        {
            const string inputCode =
@"Public Property Let Foo(ByVal Value As Variant)
End Property

Public Property Set Foo(ByVal Value As Variant)
End Property
";

            using (var explorer = new MockedCodeExplorer(inputCode)
                .SelectFirstModule())
            {
                var module = explorer.ViewModel.SelectedItem;
                var propertyLetNode = module.Items.Single(s => s.Name == "Foo (Let)");
                var propertySetNode = module.Items.Single(s => s.Name == "Foo (Set)");

                Assert.AreEqual(0, new CompareByType().Compare(propertyLetNode, propertySetNode));
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void CompareByType_ReturnsPropertySetAboveFunction()
        {
            const string inputCode =
@"Public Property Set Foo(ByVal Value As Variant)
End Property

Public Function Bar() As Boolean
End Function
";

            using (var explorer = new MockedCodeExplorer(inputCode)
                .SelectFirstModule())
            {
                var module = explorer.ViewModel.SelectedItem;
                var propertySetNode = module.Items.Single(s => s.Name == "Foo (Set)");
                var functionNode = module.Items.Single(s => s.Name == "Bar");

                Assert.AreEqual(-1, new CompareByType().Compare(propertySetNode, functionNode));
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void CompareByType_ReturnsSubsAndFunctionsEqual()
        {
            const string inputCode =
@"Public Function Foo() As Boolean
End Function

Public Sub Bar()
End Sub
";

            using (var explorer = new MockedCodeExplorer(inputCode)
                .SelectFirstModule())
            {
                var module = explorer.ViewModel.SelectedItem;
                var functionNode = module.Items.Single(s => s.Name == "Foo");
                var subNode = module.Items.Single(s => s.Name == "Bar");

                Assert.AreEqual(0, new CompareByType().Compare(functionNode, subNode));
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void CompareByType_ReturnsPublicMethodsAbovePrivateMethods()
        {
            const string inputCode =
 @"Private Sub Foo()
End Sub

Public Sub Bar()
End Sub
";

             using (var explorer = new MockedCodeExplorer(inputCode)
                .SelectFirstModule())
            {
                var module = explorer.ViewModel.SelectedItem;
                var privateNode = module.Items.Single(s => s.Name == "Foo");
                var publicNode = module.Items.Single(s => s.Name == "Bar");

                Assert.AreEqual(-1, new CompareByType().Compare(publicNode, privateNode));
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void CompareByType_ReturnsClassModuleBelowDocument()
        {

            using (var explorer = new MockedCodeExplorer(ProjectType.HostProject, new[] { ComponentType.ClassModule, ComponentType.Document })
                .SelectFirstCustomFolder())
            {
                var folder = explorer.ViewModel.SelectedItem;
                var docNode = folder.Items.Single(s => s.Name == "Document");
                var clsNode = folder.Items.Single(s => s.Name == "ClassModule");

                // this tests the logic I wrote to place docs above cls modules even though the parser calls them both cls modules
                Assert.AreEqual(((ICodeExplorerDeclarationViewModel)clsNode).Declaration.DeclarationType,
                    ((ICodeExplorerDeclarationViewModel)docNode).Declaration.DeclarationType);

                Assert.AreEqual(-1, new CompareByType().Compare(docNode, clsNode));
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void CompareBySelection_ReturnsZeroForIdenticalNodes()
        {
            const string inputCode =
@"Sub Foo()
End Sub

Sub Bar()
    Foo
End Sub";

            using (var explorer = new MockedCodeExplorer(inputCode)
                .SelectFirstModule())
            {
                var module = explorer.ViewModel.SelectedItem;
                var node = module.Items.Single(s => s.Name == "Foo");

                Assert.AreEqual(0, new CompareByName().Compare(node, node));
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void CompareByNodeType_ReturnsCorrectMemberFirst_MemberPassedFirst()
        {
            const string inputCode =
@"Sub Foo()
End Sub

Sub Bar()
    Foo
End Sub";

            using (var explorer = new MockedCodeExplorer(inputCode)
                .SelectFirstModule())
            {
                var module = explorer.ViewModel.SelectedItem;
                var memberNode1 = module.Items.Single(s => s.Name == "Foo");
                var memberNode2 = module.Items.Single(s => s.Name == "Bar");

                Assert.AreEqual(-1, new CompareBySelection().Compare(memberNode1, memberNode2));
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void CompareByNodeType_ReturnsZeroForIdenticalNodes()
        {
            const string inputCode =
@"Sub Foo()
End Sub

Sub Bar()
    Foo
End Sub";

            using (var explorer = new MockedCodeExplorer(inputCode)
                .SelectFirstModule())
            {
                var module = explorer.ViewModel.SelectedItem;
                var node = module.Items.Single(s => s.Name == "Foo");

                Assert.AreEqual(0, new CompareByNodeType().Compare(node, node));
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void CompareByNodeType_FoldersAreSortedByName()
        {
            var folderNode1 = new CodeExplorerCustomFolderViewModel(null, "AAA", string.Empty, null, null);
            var folderNode2 = new CodeExplorerCustomFolderViewModel(null, "zzz", string.Empty, null, null);

            Assert.IsTrue(new CompareByNodeType().Compare(folderNode1, folderNode2) < 0);
        }

        protected class MockedCodeExplorer : IDisposable
        {
            private readonly GeneralSettings _generalSettings = new GeneralSettings();

            private readonly Mock<IUiDispatcher> _uiDispatcher = new Mock<IUiDispatcher>();
            private readonly Mock<IConfigProvider<GeneralSettings>> _generalSettingsProvider = new Mock<IConfigProvider<GeneralSettings>>();
            private readonly Mock<IConfigProvider<WindowSettings>> _windowSettingsProvider = new Mock<IConfigProvider<WindowSettings>>();
            private readonly Mock<ConfigurationLoader> _configLoader = new Mock<ConfigurationLoader>(null, null, null, null, null, null, null, null);
            private readonly Mock<IVBEInteraction> _interaction = new Mock<IVBEInteraction>();
            private readonly Mock<IFileSystemBrowserFactory> _browserFactory = new Mock<IFileSystemBrowserFactory>();

            private MockedCodeExplorer()
            {
                _generalSettingsProvider.Setup(s => s.Create()).Returns(_generalSettings);
                _windowSettingsProvider.Setup(s => s.Create()).Returns(WindowSettings);
                _configLoader.Setup(c => c.LoadConfiguration()).Returns(GetDefaultUnitTestConfig());

                SaveDialog = new Mock<ISaveFileDialog>();
                SaveDialog.Setup(o => o.OverwritePrompt);

                OpenDialog = new Mock<IOpenFileDialog>();
                OpenDialog.Setup(o => o.AddExtension);
                OpenDialog.Setup(o => o.AutoUpgradeEnabled);
                OpenDialog.Setup(o => o.CheckFileExists);
                OpenDialog.Setup(o => o.Multiselect);
                OpenDialog.Setup(o => o.ShowHelp);
                OpenDialog.Setup(o => o.Filter);
                OpenDialog.Setup(o => o.CheckFileExists);

                FolderBrowser = new Mock<IFolderBrowser>();
                _browserFactory
                    .Setup(m => m.CreateFolderBrowser(It.IsAny<string>(), true,
                        @"C:\Users\Rubberduck\Documents\Subfolder")).Returns(FolderBrowser.Object);
            }

            public MockedCodeExplorer(string code) : this(ProjectType.HostProject, ComponentType.StandardModule, code) { }

            public MockedCodeExplorer(ProjectType projectType, ComponentType componentType = ComponentType.StandardModule, string code = "") : this()
            {
                var builder = new MockVbeBuilder();
                var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected, projectType)
                    .AddComponent("TestModule", componentType, code);

                VbComponents = project.MockVBComponents;
                VbComponent = project.MockComponents.First();
                VbProject = project.Build();
                Vbe = builder.AddProject(VbProject).Build();

                SetupViewModelAndParse();
            }

            public MockedCodeExplorer(ProjectType projectType,
                IReadOnlyList<ComponentType> componentTypes,
                IReadOnlyList<string> code = null) : this()
            {
                if (code != null && componentTypes.Count != code.Count)
                {
                    Assert.Inconclusive("MockedCodeExplorer Setup Error");
                }

                var builder = new MockVbeBuilder();
                var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected, projectType);

                for (var index = 0; index < componentTypes.Count; index++)
                {
                    var item = componentTypes[index];
                    if (item == ComponentType.UserForm)
                    {
                        project.MockUserFormBuilder(item.ToString(), code is null ? string.Empty : code[index]).AddFormToProjectBuilder();
                    }
                    else
                    {
                        project.AddComponent(item.ToString(), item, code is null ? string.Empty : code[index]);
                    }
                }

                VbComponents = project.MockVBComponents;
                VbComponent = project.MockComponents.First();
                VbProject = project.Build();
                Vbe = builder.AddProject(VbProject).Build();

                SetupViewModelAndParse();

                VbProject.SetupGet(m => m.VBComponents.Count).Returns(componentTypes.Count);
            }

            private void SetupViewModelAndParse()
            {
                var parser = MockParser.Create(Vbe.Object, null, MockVbeEvents.CreateMockVbeEvents(Vbe));
                State = parser.State;

                var removeCommand = new RemoveCommand(SaveDialog.Object, MessageBox.Object, State.ProjectsProvider);

                ViewModel = new CodeExplorerViewModel(new FolderHelper(State, Vbe.Object), State, removeCommand,
                    _generalSettingsProvider.Object,
                    _windowSettingsProvider.Object, _uiDispatcher.Object, Vbe.Object, null);

                parser.Parse(new CancellationTokenSource());
                if (parser.State.Status >= ParserState.Error)
                {
                    Assert.Inconclusive("Parser Error");
                }
            }

            public RubberduckParserState State { get; set; }
            public Mock<IVBE> Vbe { get; }
            public CodeExplorerViewModel ViewModel { get; set; }
            public Mock<IVBProject> VbProject { get; }
            public Mock<IVBComponents> VbComponents { get; }
            public Mock<IVBComponent> VbComponent { get; }
            public Mock<ISaveFileDialog> SaveDialog { get; }
            public Mock<IOpenFileDialog> OpenDialog { get; }
            public Mock<IFolderBrowser> FolderBrowser { get; }
            public Mock<IMessageBox> MessageBox { get; } = new Mock<IMessageBox>();

            public WindowSettings WindowSettings { get; } = new WindowSettings();

            public MockedCodeExplorer ImplementAddStdModuleCommand()
            {
                ViewModel.AddStdModuleCommand = new AddStdModuleCommand(new AddComponentCommand(Vbe.Object));
                return this;
            }

            public void ExecuteAddStdModuleCommand()
            {
                if (ViewModel.AddStdModuleCommand is null)
                {
                    ImplementAddStdModuleCommand();
                }
                ViewModel.AddStdModuleCommand.Execute(ViewModel.SelectedItem);
            }

            public MockedCodeExplorer ImplementAddClassModuleCommand()
            {
                ViewModel.AddClassModuleCommand = new AddClassModuleCommand(new AddComponentCommand(Vbe.Object));
                return this;
            }

            public void ExecuteAddClassModuleCommand()
            {
                if (ViewModel.AddClassModuleCommand is null)
                {
                    ImplementAddClassModuleCommand();
                }
                ViewModel.AddClassModuleCommand.Execute(ViewModel.SelectedItem);
            }

            public MockedCodeExplorer ImplementAddUserFormCommand()
            {
                ViewModel.AddUserFormCommand = new AddUserFormCommand(new AddComponentCommand(Vbe.Object));
                return this;
            }

            public void ExecuteAddUserFormCommand()
            {
                if (ViewModel.AddUserFormCommand is null)
                {
                    ImplementAddUserFormCommand();
                }
                ViewModel.AddUserFormCommand.Execute(ViewModel.SelectedItem);
            }

            public MockedCodeExplorer ImplementAddVbFormCommand()
            {
                ViewModel.AddVBFormCommand = new AddVBFormCommand(new AddComponentCommand(Vbe.Object));
                return this;
            }

            public void ExecuteAddVbFormCommand()
            {
                if (ViewModel.AddVBFormCommand is null)
                {
                    ImplementAddVbFormCommand();
                }
                ViewModel.AddVBFormCommand.Execute(ViewModel.SelectedItem);
            }

            public MockedCodeExplorer ImplementAddMdiFormCommand()
            {
                ViewModel.AddMDIFormCommand = new AddMDIFormCommand(Vbe.Object, new AddComponentCommand(Vbe.Object));
                return this;
            }

            public void ExecuteAddMdiFormCommand()
            {
                if (ViewModel.AddMDIFormCommand is null)
                {
                    ImplementAddMdiFormCommand();
                }
                ViewModel.AddMDIFormCommand.Execute(ViewModel.SelectedItem);
            }

            public MockedCodeExplorer ImplementAddUserControlCommand()
            {
                ViewModel.AddUserControlCommand = new AddUserControlCommand(new AddComponentCommand(Vbe.Object));
                return this;
            }

            public void ExecuteAddUserControlCommand()
            {
                if (ViewModel.AddUserControlCommand is null)
                {
                    ImplementAddUserControlCommand();
                }
                ViewModel.AddUserControlCommand.Execute(ViewModel.SelectedItem);
            }

            public MockedCodeExplorer ImplementAddPropertyPageCommand()
            {
                ViewModel.AddPropertyPageCommand = new AddPropertyPageCommand(new AddComponentCommand(Vbe.Object));
                return this;
            }

            public void ExecuteAddPropertyPageCommand()
            {
                if (ViewModel.AddPropertyPageCommand is null)
                {
                    ImplementAddPropertyPageCommand();
                }
                ViewModel.AddPropertyPageCommand.Execute(ViewModel.SelectedItem);
            }

            public MockedCodeExplorer ImplementAddUserDocumentCommand()
            {
                ViewModel.AddUserDocumentCommand = new AddUserDocumentCommand(new AddComponentCommand(Vbe.Object));
                return this;
            }

            public void ExecuteAddUserDocumentCommand()
            {
                if (ViewModel.AddUserDocumentCommand is null)
                {
                    ImplementAddUserDocumentCommand();
                }
                ViewModel.AddUserDocumentCommand.Execute(ViewModel.SelectedItem);
            }

            public MockedCodeExplorer ImplementAddTestModuleCommand()
            {
                ViewModel.AddTestModuleCommand = new AddTestModuleCommand(Vbe.Object, State, _configLoader.Object, MessageBox.Object, _interaction.Object);
                return this;
            }

            public void ExecuteAddTestModuleCommand()
            {
                if (ViewModel.AddTestModuleCommand is null)
                {
                    ImplementAddTestModuleCommand();
                }
                ViewModel.AddTestModuleCommand.Execute(ViewModel.SelectedItem);
            }

            public MockedCodeExplorer ImplementAddTestModuleWithStubsCommand()
            {
                ImplementAddTestModuleCommand();
                ViewModel.AddTestModuleWithStubsCommand = new AddTestModuleWithStubsCommand(Vbe.Object, ViewModel.AddTestModuleCommand);
                return this;
            }

            public void ExecuteAddTestModuleWithStubsCommand()
            {
                if (ViewModel.AddTestModuleWithStubsCommand is null)
                {
                    ImplementAddTestModuleWithStubsCommand();
                }
                ViewModel.AddTestModuleWithStubsCommand.Execute(ViewModel.SelectedItem);
            }

            public void ExecuteImportCommand()
            {
                ViewModel.ImportCommand = new ImportCommand(Vbe.Object, OpenDialog.Object);
                ViewModel.ImportCommand.Execute(ViewModel.SelectedItem);
            }

            public void ExecuteExportAllCommand()
            {
                if (ViewModel.ExportAllCommand is null)
                {
                    ImplementExportAllCommand();
                }
                ViewModel.ExportAllCommand.Execute(ViewModel.SelectedItem);
            }

            public MockedCodeExplorer ImplementExportAllCommand()
            {
                ViewModel.ExportAllCommand = new ExportAllCommand(Vbe.Object, _browserFactory.Object);
                return this;
            }

            public void ExecuteExportCommand()
            {
                if (ViewModel.ExportCommand is null)
                {
                    ImplementExportCommand();
                }
                ViewModel.ExportCommand.Execute(ViewModel.SelectedItem);
            }

            public MockedCodeExplorer ImplementExportCommand()
            {
                ViewModel.ExportCommand = new ExportCommand(SaveDialog.Object, State.ProjectsProvider);
                return this;
            }

            public void ExecuteOpenDesignerCommand()
            {
                if (ViewModel.OpenDesignerCommand is null)
                {
                    ImplementOpenDesignerCommand();
                }
                ViewModel.OpenDesignerCommand.Execute(ViewModel.SelectedItem);
            }

            public MockedCodeExplorer ImplementOpenDesignerCommand()
            {
                ViewModel.OpenDesignerCommand = new OpenDesignerCommand(State.ProjectsProvider);
                return this;
            }

            public void ExecuteIndenterCommand()
            {
                if (ViewModel.IndenterCommand is null)
                {
                    ImplementIndenterCommand();
                }
                ViewModel.IndenterCommand.Execute(ViewModel.SelectedItem);
            }

            public MockedCodeExplorer ImplementIndenterCommand()
            {
                ViewModel.IndenterCommand = new IndentCommand(State, new Indenter(Vbe.Object, () => Settings.IndenterSettingsTests.GetMockIndenterSettings()), null);
                return this;
            }

            public MockedCodeExplorer ConfigureSaveDialog(string path, DialogResult result)
            {
                SaveDialog.Setup(o => o.FileName).Returns(path);
                SaveDialog.Setup(o => o.ShowDialog()).Returns(result);
                return this;
            }

            public MockedCodeExplorer ConfigureOpenDialog(string[] paths, DialogResult result)
            {
                OpenDialog.Setup(o => o.FileNames).Returns(paths);
                OpenDialog.Setup(o => o.ShowDialog()).Returns(result);
                return this;
            }

            public MockedCodeExplorer ConfigureFolderBrowser(string selected, DialogResult result)
            {
                FolderBrowser.Setup(m => m.SelectedPath).Returns(selected);
                FolderBrowser.Setup(m => m.ShowDialog()).Returns(result);
                return this;
            }

            public MockedCodeExplorer ConfigureMessageBox(ConfirmationOutcome result)
            {
                MessageBox.Setup(m => m.Confirm(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<ConfirmationOutcome>())).Returns(result);
                return this;
            }

            public MockedCodeExplorer SelectFirstProject()
            {
                ViewModel.SelectedItem = ViewModel.Projects.First();
                return this;
            }

            public MockedCodeExplorer SelectFirstCustomFolder()
            {
                ViewModel.SelectedItem = ViewModel.Projects.First().Items.First(node => node is CodeExplorerCustomFolderViewModel);
                return this;
            }

            public MockedCodeExplorer SelectFirstModule()
            {
                ViewModel.SelectedItem = ViewModel.Projects.First().Items.First(node => !(node is CodeExplorerReferenceFolderViewModel)).Items.First();
                return this;
            }

            public MockedCodeExplorer SelectFirstMember()
            {
                ViewModel.SelectedItem = ViewModel.Projects.First().Items.First(node => !(node is CodeExplorerReferenceFolderViewModel)).Items.First().Items.First();
                return this;
            }

            private Configuration GetDefaultUnitTestConfig()
            {
                var unitTestSettings = new UnitTestSettings(BindingMode.LateBinding, AssertMode.StrictAssert, true, true, false);

                var generalSettings = new GeneralSettings
                {
                    EnableExperimentalFeatures = new List<ExperimentalFeatures>
                    {
                        new ExperimentalFeatures()
                    }
                };

                var userSettings = new UserSettings(generalSettings, null, null, null, null, unitTestSettings, null, null);
                return new Configuration(userSettings);
            }

            public void Dispose()
            {
                Dispose(true);
                GC.SuppressFinalize(this);
            }

            private bool _disposed;
            protected virtual void Dispose(bool disposing)
            {
                if (disposing && !_disposed)
                {
                    State?.Dispose();
                }
                _disposed = true;
            }
        }

    }
}
