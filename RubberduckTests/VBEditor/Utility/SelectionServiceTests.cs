using System;
using System.Linq;
using System.Runtime.InteropServices;
using Moq;
using NUnit.Framework;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.Utility;
using RubberduckTests.Mocks;

namespace RubberduckTests.VBEditor.Utility
{
    [TestFixture]
    public class SelectionServiceTests
    {
        [Test]
        public void NoCodePaneOpen_ActiveSelectionReturnsNull()
        {
            var vbe = new Mock<IVBE>().Object;
            var projectsProvider = new Mock<IProjectsProvider>().Object;

            var selectionService = new SelectionService(vbe, projectsProvider);

            var activeSelection = selectionService.ActiveSelection();
            
            Assert.IsNull(activeSelection);
        }

        [Test]
        public void CodePaneOpen_ActiveSelectionReturnsSelectionOfActiveCodePane()
        {
            var vbe = MockVbeBuilder.BuildFromStdModules(new[]
            {
                ("activeModule", $"{Environment.NewLine}{Environment.NewLine}{Environment.NewLine}"),
                ("otherModule", string.Empty)
            }).Object;
            var activeCodePane = vbe.VBProjects.Single().VBComponents.Single(comp => comp.Name.Equals("activeModule")).CodeModule.CodePane;
            var selection = new Selection(2, 1);

            activeCodePane.Selection = selection;
            vbe.ActiveCodePane = activeCodePane;

            var projectsProvider = new Mock<IProjectsProvider>().Object;

            var selectionService = new SelectionService(vbe, projectsProvider);

            var expectedActiveSelection = vbe.GetActiveSelection();
            var actualActiveSelection = selectionService.ActiveSelection();

            Assert.AreEqual(expectedActiveSelection, actualActiveSelection);
        }

        [Test]
        public void OpenModulesReturnsTheModulesOfActiveCodePanes()
        {
            var vbeBuilder = new MockVbeBuilder();
            var project = vbeBuilder.ProjectBuilder("test", ProjectProtection.Unprotected)
                .AddComponent("activeModule", ComponentType.ClassModule, string.Empty)
                .AddComponent("otherActiveModule", ComponentType.ClassModule, string.Empty)
                .AddComponent("otherModule", ComponentType.ClassModule, string.Empty)
                .Build();
            var openCodePanes = project.Object.VBComponents
                .Where(component => component.Name == "activeModule" || component.Name == "otherActiveModule")
                .Select(component => component.CodeModule.CodePane)
                .ToList();
            var vbe = vbeBuilder.AddProject(project)
                .SetOpenCodePanes(openCodePanes)
                .Build()
                .Object;
            var projectsProvider = new Mock<IProjectsProvider>().Object;
            var selectionService = new SelectionService(vbe, projectsProvider);

            var expectedOpenModules = openCodePanes.Select(pane => pane.QualifiedModuleName).ToList();
            var actualOpenModules = selectionService.OpenModules();

            Assert.AreEqual(expectedOpenModules.Count, actualOpenModules.Count);
            foreach (var module in expectedOpenModules)
            {
                Assert.IsTrue(actualOpenModules.Contains(module));
            }
        }

        [Test]
        public void ComponentExists_SelectionReturnsSelection()
        {
            var vbe = MockVbeBuilder.BuildFromStdModules(new[]
            {
                ("someModule", $"{Environment.NewLine}{Environment.NewLine}{Environment.NewLine}"),
                ("otherModule", string.Empty)
            }).Object;
            var someCodePane = vbe.VBProjects.Single().VBComponents.Single(comp => comp.Name.Equals("someModule")).CodeModule.CodePane;
            var selection = new Selection(2, 1);
            someCodePane.Selection = selection;
            var projectsProvider = new ProjectsRepository(vbe);
            projectsProvider.Refresh();
            var module = someCodePane.QualifiedModuleName;

            var selectionService = new SelectionService(vbe, projectsProvider);

            var expectedSelection = someCodePane.Selection;
            var actualSelection = selectionService.Selection(module);

            Assert.AreEqual(expectedSelection, actualSelection);
        }

        [Test]
        public void ComponentDoesNotExist_SelectionReturnsNull()
        {
            var vbe = MockVbeBuilder.BuildFromStdModules(new[]
            {
                ("someModule", $"{Environment.NewLine}{Environment.NewLine}{Environment.NewLine}"),
                ("otherModule", string.Empty)
            }).Object;
            var projectsProvider = new ProjectsRepository(vbe);
            projectsProvider.Refresh();
            var module = new QualifiedModuleName("test", string.Empty, "yetAnotherModule");

            var selectionService = new SelectionService(vbe, projectsProvider);

            var actualSelection = selectionService.Selection(module);

            Assert.IsNull(actualSelection);
        }

        [Test]
        public void ComponentExists_TryActiveActivatesComponentAndReturnsTrue()
        {
            var vbe = MockVbeBuilder.BuildFromStdModules(new[]
            {
                ("activeModule", $"{Environment.NewLine}{Environment.NewLine}{Environment.NewLine}"),
                ("otherModule", string.Empty)
            }).Object;
            var activeCodePane = vbe.VBProjects.Single().VBComponents.Single(comp => comp.Name.Equals("activeModule")).CodeModule.CodePane;
            vbe.ActiveCodePane = activeCodePane;
            var projectsProvider = new ProjectsRepository(vbe);
            projectsProvider.Refresh();
            var selectionService = new SelectionService(vbe, projectsProvider);

            var otherModule = vbe.VBProjects.Single().VBComponents.Single(comp => comp.Name.Equals("otherModule")).QualifiedModuleName;
            var success = selectionService.TryActivate(otherModule);

            var expectedActiveModule = otherModule;
            var actualActiveModule = vbe.ActiveCodePane.QualifiedModuleName;

            Assert.IsTrue(success);
            Assert.AreEqual(expectedActiveModule, actualActiveModule);
        }

        [Test]
        public void ComponentDoesNotExist_TryActiveDoesNotChangeTheActiveModuleAndReturnsFalse()
        {
            var vbe = MockVbeBuilder.BuildFromStdModules(new[]
            {
                ("activeModule", $"{Environment.NewLine}{Environment.NewLine}{Environment.NewLine}"),
                ("otherModule", string.Empty)
            }).Object;
            var activeCodePane = vbe.VBProjects.Single().VBComponents.Single(comp => comp.Name.Equals("activeModule")).CodeModule.CodePane;
            vbe.ActiveCodePane = activeCodePane;
            var projectsProvider = new ProjectsRepository(vbe);
            projectsProvider.Refresh();
            var selectionService = new SelectionService(vbe, projectsProvider);

            var nonExistentModule = new QualifiedModuleName("test", string.Empty, "nonExistentModule");
            var success = selectionService.TryActivate(nonExistentModule);

            var expectedActiveModule = activeCodePane.QualifiedModuleName;
            var actualActiveModule = vbe.ActiveCodePane.QualifiedModuleName;

            Assert.IsFalse(success);
            Assert.AreEqual(expectedActiveModule, actualActiveModule);
        }

        [Test]
        public void VbeThrowsExceptionOnActivation_TryActiveDoesNotChangeTheActiveModuleAndReturnsFalse()
        {
            var vbeMock = MockVbeBuilder.BuildFromStdModules(new[]
            {
                ("activeModule", $"{Environment.NewLine}{Environment.NewLine}{Environment.NewLine}"),
                ("otherModule", string.Empty)
            });
            var vbe = vbeMock.Object;
            var activeCodePane = vbe.VBProjects.Single().VBComponents.Single(comp => comp.Name.Equals("activeModule")).CodeModule.CodePane;
            vbe.ActiveCodePane = activeCodePane;
            var projectsProvider = new ProjectsRepository(vbe);
            projectsProvider.Refresh();
            var selectionService = new SelectionService(vbe, projectsProvider);

            var otherCodePane = vbe.VBProjects.Single().VBComponents.Single(comp => comp.Name.Equals("otherModule")).CodeModule.CodePane;
            var otherModule = otherCodePane.QualifiedModuleName;
            vbeMock.SetupSet(m => m.ActiveCodePane = otherCodePane).Callback(() => throw new COMException());

            var success = selectionService.TryActivate(otherModule);

            var expectedActiveModule = activeCodePane.QualifiedModuleName;
            var actualActiveModule = vbe.ActiveCodePane.QualifiedModuleName;

            Assert.IsFalse(success);
            Assert.AreEqual(expectedActiveModule, actualActiveModule);
        }

        [Test]
        public void ComponentExists_TrySetActiveSelectionSetsActiveSelectionAndReturnsTrue_QualifiedSelection()
        {
            var vbe = MockVbeBuilder.BuildFromStdModules(new[]
            {
                ("activeModule", $"{Environment.NewLine}{Environment.NewLine}{Environment.NewLine}"),
                ("otherModule", $"{Environment.NewLine}{Environment.NewLine}{Environment.NewLine}")
            }).Object;
            var activeCodePane = vbe.VBProjects.Single().VBComponents.Single(comp => comp.Name.Equals("activeModule")).CodeModule.CodePane;
            var activeSelection = new Selection(2, 1);
            activeCodePane.Selection = activeSelection;
            vbe.ActiveCodePane = activeCodePane;
            var projectsProvider = new ProjectsRepository(vbe);
            projectsProvider.Refresh();
            var selectionService = new SelectionService(vbe, projectsProvider);

            var otherModule = vbe.VBProjects.Single().VBComponents.Single(comp => comp.Name.Equals("otherModule")).QualifiedModuleName;
            var newSelection = new Selection(3, 1);
            var newQualifiedSelection = new QualifiedSelection(otherModule, newSelection);

            var success = selectionService.TrySetActiveSelection(newQualifiedSelection);

            var expectedActiveModule = otherModule;
            var actualActiveModule = vbe.ActiveCodePane.QualifiedModuleName;
            var actualSelection = vbe.ActiveCodePane.Selection;

            Assert.IsTrue(success);
            Assert.AreEqual(expectedActiveModule, actualActiveModule);
            Assert.AreEqual(newSelection, actualSelection);
        }

        [Test]
        public void ComponentDoesNotExist_TrySetActiveSelectionDoesNotChangeTheActiveSelectionAndReturnsFalse_QualifiedSelection()
        {
            var vbe = MockVbeBuilder.BuildFromStdModules(new[]
            {
                ("activeModule", $"{Environment.NewLine}{Environment.NewLine}{Environment.NewLine}"),
                ("otherModule", $"{Environment.NewLine}{Environment.NewLine}{Environment.NewLine}")
            }).Object;
            var activeCodePane = vbe.VBProjects.Single().VBComponents.Single(comp => comp.Name.Equals("activeModule")).CodeModule.CodePane;
            var activeSelection = new Selection(2, 1);
            activeCodePane.Selection = activeSelection;
            vbe.ActiveCodePane = activeCodePane;
            var projectsProvider = new ProjectsRepository(vbe);
            projectsProvider.Refresh();
            var selectionService = new SelectionService(vbe, projectsProvider);

            var nonExistentModule = new QualifiedModuleName("test", string.Empty, "nonExistentModule");
            var newSelection = new Selection(3, 1);
            var newQualifiedSelection = new QualifiedSelection(nonExistentModule, newSelection);

            var success = selectionService.TrySetActiveSelection(newQualifiedSelection);

            var expectedActiveModule = activeCodePane.QualifiedModuleName;
            var actualActiveModule = vbe.ActiveCodePane.QualifiedModuleName;
            var actualSelection = vbe.ActiveCodePane.Selection;

            Assert.IsFalse(success);
            Assert.AreEqual(expectedActiveModule, actualActiveModule);
            Assert.AreEqual(activeSelection, actualSelection);
        }

        [Test]
        public void VbeThrowsExceptionOnComponentActivation_TrySetActiveSelectionDoesNotChangeTheActiveSelectionAndReturnsFalse_QualifiedSelection()
        {
            var vbeMock = MockVbeBuilder.BuildFromStdModules(new[]
            {
                ("activeModule", $"{Environment.NewLine}{Environment.NewLine}{Environment.NewLine}"),
                ("otherModule", $"{Environment.NewLine}{Environment.NewLine}{Environment.NewLine}")
            });
            var vbe = vbeMock.Object;
            var activeCodePane = vbe.VBProjects.Single().VBComponents.Single(comp => comp.Name.Equals("activeModule")).CodeModule.CodePane;
            var activeSelection = new Selection(2, 1);
            activeCodePane.Selection = activeSelection;
            vbe.ActiveCodePane = activeCodePane;
            var projectsProvider = new ProjectsRepository(vbe);
            projectsProvider.Refresh();
            var selectionService = new SelectionService(vbe, projectsProvider);

            var otherCodePane = vbe.VBProjects.Single().VBComponents.Single(comp => comp.Name.Equals("otherModule")).CodeModule.CodePane;
            var otherModule = otherCodePane.QualifiedModuleName;
            var newSelection = new Selection(3, 1);
            var newQualifiedSelection = new QualifiedSelection(otherModule, newSelection);
            vbeMock.SetupSet(m => m.ActiveCodePane = otherCodePane).Callback(() => throw new COMException());

            var success = selectionService.TrySetActiveSelection(newQualifiedSelection);

            var expectedActiveModule = activeCodePane.QualifiedModuleName;
            var actualActiveModule = vbe.ActiveCodePane.QualifiedModuleName;
            var actualSelection = vbe.ActiveCodePane.Selection;

            Assert.IsFalse(success);
            Assert.AreEqual(expectedActiveModule, actualActiveModule);
            Assert.AreEqual(activeSelection, actualSelection);
        }

        [Test]
        public void SomeCodePaneOpen_ThrowsExceptionOnSelectionChange_TrySetActiveSelectionDoesNotChangeTheActiveCodePaneAndReturnsFalse_QualifiedSelection()
        {
            var vbeBuilder = new MockVbeBuilder();
            var activeSelection = new Selection(2, 1);
            var projectBuilder = vbeBuilder.ProjectBuilder("test", ProjectProtection.Unprotected)
                .AddComponent("activeModule", ComponentType.StandardModule, $"{Environment.NewLine}{Environment.NewLine}{Environment.NewLine}", activeSelection)
                .AddComponent("otherModule", ComponentType.StandardModule, $"{Environment.NewLine}{Environment.NewLine}{Environment.NewLine}");

            var vbe = projectBuilder.AddProjectToVbeBuilder().Build().Object;
            var activeCodePane = vbe.VBProjects.Single().VBComponents.Single(comp => comp.Name.Equals("activeModule")).CodeModule.CodePane;
            vbe.ActiveCodePane = activeCodePane;
            var projectsProvider = new ProjectsRepository(vbe);
            projectsProvider.Refresh();
            var selectionService = new SelectionService(vbe, projectsProvider);

            var otherCodeModuleMock = projectBuilder.MockCodeModules.Single(mock => mock.Object.Name.Equals("otherModule"));
            var otherPaneMock = new Mock<ICodePane>();
            otherPaneMock.SetupSet(m => m.Selection = It.IsAny<Selection>()).Callback(() => throw new COMException());
            otherCodeModuleMock.SetupGet(m => m.CodePane).Returns(otherPaneMock.Object);
            var otherModule = otherCodeModuleMock.Object.QualifiedModuleName;
            var newSelection = new Selection(3, 1);
            var newQualifiedSelection = new QualifiedSelection(otherModule, newSelection);

            var success = selectionService.TrySetActiveSelection(newQualifiedSelection);

            var expectedActiveModule = activeCodePane.QualifiedModuleName;
            var actualActiveModule = vbe.ActiveCodePane.QualifiedModuleName;
            var actualSelection = vbe.ActiveCodePane.Selection;

            Assert.IsFalse(success);
            Assert.AreEqual(expectedActiveModule, actualActiveModule);
            Assert.AreEqual(activeSelection, actualSelection); ;
        }

        [Test]
        public void ComponentExists_TrySetActiveSelectionSetsActiveSelectionAndReturnsTrue_QualifiedModuleNameAndSelection()
        {
            var vbe = MockVbeBuilder.BuildFromStdModules(new[]
            {
                ("activeModule", $"{Environment.NewLine}{Environment.NewLine}{Environment.NewLine}"),
                ("otherModule", $"{Environment.NewLine}{Environment.NewLine}{Environment.NewLine}")
            }).Object;
            var activeCodePane = vbe.VBProjects.Single().VBComponents.Single(comp => comp.Name.Equals("activeModule")).CodeModule.CodePane;
            var activeSelection = new Selection(2, 1);
            activeCodePane.Selection = activeSelection;
            vbe.ActiveCodePane = activeCodePane;
            var projectsProvider = new ProjectsRepository(vbe);
            projectsProvider.Refresh();
            var selectionService = new SelectionService(vbe, projectsProvider);

            var otherModule = vbe.VBProjects.Single().VBComponents.Single(comp => comp.Name.Equals("otherModule")).QualifiedModuleName;
            var newSelection = new Selection(3, 1);

            var success = selectionService.TrySetActiveSelection(otherModule, newSelection);

            var expectedActiveModule = otherModule;
            var actualActiveModule = vbe.ActiveCodePane.QualifiedModuleName;
            var actualSelection = vbe.ActiveCodePane.Selection;

            Assert.IsTrue(success);
            Assert.AreEqual(expectedActiveModule, actualActiveModule);
            Assert.AreEqual(newSelection, actualSelection);
        }

        [Test]
        public void ComponentDoesNotExist_TrySetActiveSelectionDoesNotChangeTheActiveSelectionAndReturnsFalse_QualifiedModuleNameAndSelection()
        {
            var vbe = MockVbeBuilder.BuildFromStdModules(new[]
            {
                ("activeModule", $"{Environment.NewLine}{Environment.NewLine}{Environment.NewLine}"),
                ("otherModule", $"{Environment.NewLine}{Environment.NewLine}{Environment.NewLine}")
            }).Object;
            var activeCodePane = vbe.VBProjects.Single().VBComponents.Single(comp => comp.Name.Equals("activeModule")).CodeModule.CodePane;
            var activeSelection = new Selection(2, 1);
            activeCodePane.Selection = activeSelection;
            vbe.ActiveCodePane = activeCodePane;
            var projectsProvider = new ProjectsRepository(vbe);
            projectsProvider.Refresh();
            var selectionService = new SelectionService(vbe, projectsProvider);

            var nonExistentModule = new QualifiedModuleName("test", string.Empty, "nonExistentModule");
            var newSelection = new Selection(3, 1);

            var success = selectionService.TrySetActiveSelection(nonExistentModule, newSelection);

            var expectedActiveModule = activeCodePane.QualifiedModuleName;
            var actualActiveModule = vbe.ActiveCodePane.QualifiedModuleName;
            var actualSelection = vbe.ActiveCodePane.Selection;

            Assert.IsFalse(success);
            Assert.AreEqual(expectedActiveModule, actualActiveModule);
            Assert.AreEqual(activeSelection, actualSelection);
        }

        [Test]
        public void VbeThrowsExceptionOnComponentActivation_TrySetActiveSelectionDoesNotChangeTheActiveSelectionAndReturnsFalse_QualifiedModuleNameAndSelection()
        {
            var vbeMock = MockVbeBuilder.BuildFromStdModules(new[]
            {
                ("activeModule", $"{Environment.NewLine}{Environment.NewLine}{Environment.NewLine}"),
                ("otherModule", $"{Environment.NewLine}{Environment.NewLine}{Environment.NewLine}")
            });
            var vbe = vbeMock.Object;
            var activeCodePane = vbe.VBProjects.Single().VBComponents.Single(comp => comp.Name.Equals("activeModule")).CodeModule.CodePane;
            var activeSelection = new Selection(2, 1);
            activeCodePane.Selection = activeSelection;
            vbe.ActiveCodePane = activeCodePane;
            var projectsProvider = new ProjectsRepository(vbe);
            projectsProvider.Refresh();
            var selectionService = new SelectionService(vbe, projectsProvider);

            var otherCodePane = vbe.VBProjects.Single().VBComponents.Single(comp => comp.Name.Equals("otherModule")).CodeModule.CodePane;
            var otherModule = otherCodePane.QualifiedModuleName;
            var newSelection = new Selection(3, 1);
            vbeMock.SetupSet(m => m.ActiveCodePane = otherCodePane).Callback(() => throw new COMException());

            var success = selectionService.TrySetActiveSelection(otherModule, newSelection);

            var expectedActiveModule = activeCodePane.QualifiedModuleName;
            var actualActiveModule = vbe.ActiveCodePane.QualifiedModuleName;
            var actualSelection = vbe.ActiveCodePane.Selection;

            Assert.IsFalse(success);
            Assert.AreEqual(expectedActiveModule, actualActiveModule);
            Assert.AreEqual(activeSelection, actualSelection);
        }

        [Test]
        public void SomeCodePaneOpen_ThrowsExceptionOnSelectionChange_TrySetActiveSelectionDoesNotChangeTheActiveCodePaneAndReturnsFalse_QualifiedModuleNameAndSelection()
        {
            var vbeBuilder = new MockVbeBuilder();
            var activeSelection = new Selection(2, 1);
            var projectBuilder = vbeBuilder.ProjectBuilder("test", ProjectProtection.Unprotected)
                .AddComponent("activeModule", ComponentType.StandardModule, $"{Environment.NewLine}{Environment.NewLine}{Environment.NewLine}", activeSelection)
                .AddComponent("otherModule", ComponentType.StandardModule, $"{Environment.NewLine}{Environment.NewLine}{Environment.NewLine}");

            var vbe = projectBuilder.AddProjectToVbeBuilder().Build().Object;
            var activeCodePane = vbe.VBProjects.Single().VBComponents.Single(comp => comp.Name.Equals("activeModule")).CodeModule.CodePane;
            vbe.ActiveCodePane = activeCodePane;
            var projectsProvider = new ProjectsRepository(vbe);
            projectsProvider.Refresh();
            var selectionService = new SelectionService(vbe, projectsProvider);

            var otherCodeModuleMock = projectBuilder.MockCodeModules.Single(mock => mock.Object.Name.Equals("otherModule"));
            var otherPaneMock = new Mock<ICodePane>();
            otherPaneMock.SetupSet(m => m.Selection = It.IsAny<Selection>()).Callback(() => throw new COMException());
            otherCodeModuleMock.SetupGet(m => m.CodePane).Returns(otherPaneMock.Object);
            var otherModule = otherCodeModuleMock.Object.QualifiedModuleName;
            var newSelection = new Selection(3, 1);

            var success = selectionService.TrySetActiveSelection(otherModule, newSelection);

            var expectedActiveModule = activeCodePane.QualifiedModuleName;
            var actualActiveModule = vbe.ActiveCodePane.QualifiedModuleName;
            var actualSelection = vbe.ActiveCodePane.Selection;

            Assert.IsFalse(success);
            Assert.AreEqual(expectedActiveModule, actualActiveModule);
            Assert.AreEqual(activeSelection, actualSelection); ;
        }

        [Test]
        public void ComponentExists_TrySetSelectionSetsSelectionOfModuleAndReturnsTrue()
        {
            var vbe = MockVbeBuilder.BuildFromStdModules(new[]
            {
                ("someModule", $"{Environment.NewLine}{Environment.NewLine}{Environment.NewLine}"),
                ("otherModule", $"{Environment.NewLine}{Environment.NewLine}{Environment.NewLine}")
            }).Object;
            var projectsProvider = new ProjectsRepository(vbe);
            projectsProvider.Refresh();
            var selectionService = new SelectionService(vbe, projectsProvider);

            var somePane = vbe.VBProjects.Single().VBComponents.Single(comp => comp.Name.Equals("someModule")).CodeModule.CodePane;
            var someModule = somePane.QualifiedModuleName;
            var newSelection = new Selection(3, 1);

            var success = selectionService.TrySetSelection(someModule, newSelection);

            var actualSelection = somePane.Selection;

            Assert.IsTrue(success);
            Assert.AreEqual(newSelection, actualSelection);
        }

        [Test]
        public void ComponentDoesNotExist_TrySetSelectionReturnsFalse()
        {
            var vbe = MockVbeBuilder.BuildFromStdModules(new[]
            {
                ("someModule", $"{Environment.NewLine}{Environment.NewLine}{Environment.NewLine}"),
                ("otherModule", $"{Environment.NewLine}{Environment.NewLine}{Environment.NewLine}")
            }).Object;
            var projectsProvider = new ProjectsRepository(vbe);
            projectsProvider.Refresh();
            var selectionService = new SelectionService(vbe, projectsProvider);

            var nonExistentModule = new QualifiedModuleName("test", string.Empty, "nonExistentModule");
            var newSelection = new Selection(3, 1);

            var success = selectionService.TrySetSelection(nonExistentModule, newSelection);

            Assert.IsFalse(success);
        }

        [Test]
        public void ThrowsExceptionOnSelectionChange_TrySetSelectionReturnsFalse()
        {
            var vbeBuilder = new MockVbeBuilder();
            var oldSelection = new Selection(2, 1);
            var projectBuilder = vbeBuilder.ProjectBuilder("test", ProjectProtection.Unprotected)
                .AddComponent("someModule", ComponentType.StandardModule, $"{Environment.NewLine}{Environment.NewLine}{Environment.NewLine}", oldSelection)
                .AddComponent("otherModule", ComponentType.StandardModule, $"{Environment.NewLine}{Environment.NewLine}{Environment.NewLine}");

            var vbe = projectBuilder.AddProjectToVbeBuilder().Build().Object;
            var projectsProvider = new ProjectsRepository(vbe);
            projectsProvider.Refresh();
            var selectionService = new SelectionService(vbe, projectsProvider);

            var someCodeModuleMock = projectBuilder.MockCodeModules.Single(mock => mock.Object.Name.Equals("someModule"));
            var somePaneMock = new Mock<ICodePane>();
            somePaneMock.SetupSet(m => m.Selection = It.IsAny<Selection>()).Callback(() => throw new COMException());
            someCodeModuleMock.SetupGet(m => m.CodePane).Returns(somePaneMock.Object);
            var someModule = someCodeModuleMock.Object.QualifiedModuleName;
            var newSelection = new Selection(3, 1);

            var success = selectionService.TrySetSelection(someModule, newSelection);

            Assert.IsFalse(success);
        }

        [Test]
        public void ComponentExists_TrySetSelectionSetsSelectionOfModuleAndReturnsTrue_QualifiedSelesction()
        {
            var vbe = MockVbeBuilder.BuildFromStdModules(new[]
            {
                ("someModule", $"{Environment.NewLine}{Environment.NewLine}{Environment.NewLine}"),
                ("otherModule", $"{Environment.NewLine}{Environment.NewLine}{Environment.NewLine}")
            }).Object;
            var projectsProvider = new ProjectsRepository(vbe);
            projectsProvider.Refresh();
            var selectionService = new SelectionService(vbe, projectsProvider);

            var somePane = vbe.VBProjects.Single().VBComponents.Single(comp => comp.Name.Equals("someModule")).CodeModule.CodePane;
            var someModule = somePane.QualifiedModuleName;
            var newSelection = new Selection(3, 1);
            var newQualifiedSelection = new QualifiedSelection(someModule, newSelection);

            var success = selectionService.TrySetSelection(newQualifiedSelection);

            var actualSelection = somePane.Selection;

            Assert.IsTrue(success);
            Assert.AreEqual(newSelection, actualSelection);
        }

        [Test]
        public void ComponentDoesNotExist_TrySetSelectionReturnsFalse_QualifiedSelesction()
        {
            var vbe = MockVbeBuilder.BuildFromStdModules(new[]
            {
                ("someModule", $"{Environment.NewLine}{Environment.NewLine}{Environment.NewLine}"),
                ("otherModule", $"{Environment.NewLine}{Environment.NewLine}{Environment.NewLine}")
            }).Object;
            var projectsProvider = new ProjectsRepository(vbe);
            projectsProvider.Refresh();
            var selectionService = new SelectionService(vbe, projectsProvider);

            var nonExistentModule = new QualifiedModuleName("test", string.Empty, "nonExistentModule");
            var newSelection = new Selection(3, 1);
            var newQualifiedSelection = new QualifiedSelection(nonExistentModule, newSelection);

            var success = selectionService.TrySetSelection(newQualifiedSelection);

            Assert.IsFalse(success);
        }

        [Test]
        public void ThrowsExceptionOnSelectionChange_TrySetSelectionReturnsFalse_QualifiedSelesction()
        {
            var vbeBuilder = new MockVbeBuilder();
            var oldSelection = new Selection(2, 1);
            var projectBuilder = vbeBuilder.ProjectBuilder("test", ProjectProtection.Unprotected)
                .AddComponent("someModule", ComponentType.StandardModule, $"{Environment.NewLine}{Environment.NewLine}{Environment.NewLine}", oldSelection)
                .AddComponent("otherModule", ComponentType.StandardModule, $"{Environment.NewLine}{Environment.NewLine}{Environment.NewLine}");

            var vbe = projectBuilder.AddProjectToVbeBuilder().Build().Object;
            var projectsProvider = new ProjectsRepository(vbe);
            projectsProvider.Refresh();
            var selectionService = new SelectionService(vbe, projectsProvider);

            var someCodeModuleMock = projectBuilder.MockCodeModules.Single(mock => mock.Object.Name.Equals("someModule"));
            var somePaneMock = new Mock<ICodePane>();
            somePaneMock.SetupSet(m => m.Selection = It.IsAny<Selection>()).Callback(() => throw new COMException());
            someCodeModuleMock.SetupGet(m => m.CodePane).Returns(somePaneMock.Object);
            var someModule = someCodeModuleMock.Object.QualifiedModuleName;
            var newSelection = new Selection(3, 1);
            var newQualifiedSelection = new QualifiedSelection(someModule, newSelection);

            var success = selectionService.TrySetSelection(newQualifiedSelection);

            Assert.IsFalse(success);
        }
    }
} 