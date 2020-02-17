using System.Collections.Generic;
using NUnit.Framework;
using System.Linq;
using Rubberduck.AddRemoveReferences;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace RubberduckTests.CodeExplorer
{
    [TestFixture]
    public class CodeExplorerProjectViewModelTests
    {
        [Test]
        [Category("Code Explorer")]
        public void Constructor_SetsDeclaration()
        {
            var declarations = CodeExplorerTestSetup.TestProjectOneDeclarations;
            var projectDeclaration = declarations.First(declaration => declaration.DeclarationType == DeclarationType.Project);

            var project = new CodeExplorerProjectViewModel(projectDeclaration, ref declarations, null, null, CodeExplorerTestSetup.ProjectOneProvider);

            Assert.AreSame(projectDeclaration, project.Declaration);
        }

        [Test]
        [Category("Code Explorer")]
        public void Constructor_SetsName()
        {
            var declarations = CodeExplorerTestSetup.TestProjectOneDeclarations;
            var projectDeclaration = declarations.First(declaration => declaration.DeclarationType == DeclarationType.Project);

            var project = new CodeExplorerProjectViewModel(projectDeclaration, ref declarations, null, null, CodeExplorerTestSetup.ProjectOneProvider);

            Assert.AreEqual(CodeExplorerTestSetup.TestProjectOneName, project.Name);
        }

        [Test]
        [Category("Code Explorer")]
        public void Constructor_NameWithSignatureIsSet()
        {
            var declarations = CodeExplorerTestSetup.TestProjectOneDeclarations;
            var projectDeclaration = declarations.First(declaration => declaration.DeclarationType == DeclarationType.Project);

            var project = new CodeExplorerProjectViewModel(projectDeclaration, ref declarations, null, null, CodeExplorerTestSetup.ProjectOneProvider);

            Assert.IsFalse(string.IsNullOrEmpty(project.NameWithSignature));
        }

        [Test]
        [Category("Code Explorer")]
        public void Constructor_PanelTitleIsSet()
        {
            var declarations = CodeExplorerTestSetup.TestProjectOneDeclarations;
            var projectDeclaration = declarations.First(declaration => declaration.DeclarationType == DeclarationType.Project);

            var project = new CodeExplorerProjectViewModel(projectDeclaration, ref declarations, null, null, CodeExplorerTestSetup.ProjectOneProvider);

            Assert.IsFalse(string.IsNullOrEmpty(project.PanelTitle));
        }

        [Test]
        [Category("Code Explorer")]
        public void Constructor_SetsIsExpandedTrue()
        {
            var declarations = CodeExplorerTestSetup.TestProjectOneDeclarations;
            var projectDeclaration = declarations.First(declaration => declaration.DeclarationType == DeclarationType.Project);

            var project = new CodeExplorerProjectViewModel(projectDeclaration, ref declarations, null, null, CodeExplorerTestSetup.ProjectOneProvider);

            Assert.IsTrue(project.IsExpanded);
        }

        [Test]
        [Category("Code Explorer")]
        public void Constructor_ToolTipIsSet()
        {
            var declarations = CodeExplorerTestSetup.TestProjectOneDeclarations;
            var projectDeclaration = declarations.First(declaration => declaration.DeclarationType == DeclarationType.Project);

            var project = new CodeExplorerProjectViewModel(projectDeclaration, ref declarations, null, null, CodeExplorerTestSetup.ProjectOneProvider);

            Assert.IsFalse(string.IsNullOrEmpty(project.ToolTip));
        }

        [Test]
        [Category("Code Explorer")]
        [TestCase(CodeExplorerSortOrder.Undefined)]
        [TestCase(CodeExplorerSortOrder.Name)]
        [TestCase(CodeExplorerSortOrder.CodeLine)]
        [TestCase(CodeExplorerSortOrder.DeclarationType)]
        [TestCase(CodeExplorerSortOrder.DeclarationTypeThenName)]
        [TestCase(CodeExplorerSortOrder.DeclarationTypeThenCodeLine)]
        public void SortComparerIsNodeType(CodeExplorerSortOrder order)
        {
            var declarations = CodeExplorerTestSetup.TestProjectOneDeclarations;
            var projectDeclaration = declarations.First(declaration => declaration.DeclarationType == DeclarationType.Project);

            var project = new CodeExplorerProjectViewModel(projectDeclaration, ref declarations, null, null, CodeExplorerTestSetup.ProjectOneProvider)
            {
                SortOrder = order
            };

            Assert.AreEqual(CodeExplorerItemComparer.NodeType.GetType(), project.SortComparer.GetType());
        }

        [Test]
        [Category("Code Explorer")]
        [TestCase(null)]
        [TestCase("")]
        [TestCase(CodeExplorerTestSetup.TestProjectOneName)]
        public void IsNotFiltered(string filter)
        {
            var declarations = CodeExplorerTestSetup.TestProjectOneDeclarations;
            var projectDeclaration = declarations.First(declaration => declaration.DeclarationType == DeclarationType.Project);

            var project = new CodeExplorerProjectViewModel(projectDeclaration, ref declarations, null, null, CodeExplorerTestSetup.ProjectOneProvider)
            {
                Filter = filter
            };

            Assert.IsFalse(project.Filtered);
        }

        [Test]
        [Category("Code Explorer")]
        [TestCase(DeclarationType.Parameter)]
        [TestCase(DeclarationType.LineLabel)]
        [TestCase(DeclarationType.UnresolvedMember, Ignore = "Pending test setup that will actually create one")]
        [TestCase(DeclarationType.BracketedExpression, Ignore = "This causes a parser error in testing due to no host application.")]
        public void TrackedDeclarations_ExcludesNonNodeTypes(DeclarationType excluded)
        {
            var declarations = CodeExplorerTestSetup.TestProjectOneDeclarations;
            // Validate test setup:
            if (declarations.All(declaration => declaration.DeclarationType != excluded))
            {
                Assert.Inconclusive("DeclarationType under test not found in test declarations.");
            }

            var projectDeclaration = declarations.First(declaration => declaration.DeclarationType == DeclarationType.Project);
            var tracked = CodeExplorerProjectViewModel.ExtractTrackedDeclarationsForProject(projectDeclaration, ref declarations);

            Assert.IsFalse(tracked.Any(declaration => declaration.DeclarationType == excluded));
        }

        [Test]
        [Category("Code Explorer")]
        [TestCase(DeclarationType.Variable)]
        [TestCase(DeclarationType.Control, Ignore = "Pending test setup that will actually create one.")]
        [TestCase(DeclarationType.Constant)]
        public void TrackedDeclarations_ExcludesMemberEnclosedTypes(DeclarationType excluded)
        {
            var declarations = CodeExplorerTestSetup.TestProjectOneDeclarations;
            // Validate test setup:
            if (declarations.All(declaration => declaration.DeclarationType != excluded))
            {
                Assert.Inconclusive("DeclarationType under test not found in test declarations.");
            }

            var projectDeclaration = declarations.First(declaration => declaration.DeclarationType == DeclarationType.Project);
            var tracked = CodeExplorerProjectViewModel.ExtractTrackedDeclarationsForProject(projectDeclaration, ref declarations);

            Assert.IsFalse(tracked.Any(declaration => declaration.DeclarationType == excluded && !declaration.ParentDeclaration.DeclarationType.HasFlag(DeclarationType.Module)));
        }

        [Test]
        [Category("Code Explorer")]
        public void Constructor_CreatesDefaultProjectFolder()
        {
            var declarations = CodeExplorerTestSetup.TestProjectOneDeclarations;
            var projectDeclaration = declarations.First(declaration => declaration.DeclarationType == DeclarationType.Project);

            var project = new CodeExplorerProjectViewModel(projectDeclaration, ref declarations, null, null, CodeExplorerTestSetup.ProjectOneProvider);
            var folder = project.Children.OfType<CodeExplorerCustomFolderViewModel>().Single();

            Assert.AreEqual(projectDeclaration.IdentifierName, folder.Name);
        }

        [Test]
        [Category("Code Explorer")]
        public void Constructor_ClearsDeclarationList()
        {
            var declarations = CodeExplorerTestSetup.TestProjectOneDeclarations;
            var projectDeclaration = declarations.First(declaration => declaration.DeclarationType == DeclarationType.Project);

            // ReSharper disable once UnusedVariable
            var project = new CodeExplorerProjectViewModel(projectDeclaration, ref declarations, null, null, CodeExplorerTestSetup.ProjectOneProvider);

            Assert.AreEqual(0, declarations.Count);
        }

        [Test]
        [Category("Code Explorer")]
        public void Synchronize_UsesDefaultProjectFolder_NoChanges()
        {
            var declarations = CodeExplorerTestSetup.TestProjectOneDeclarations;
            var projectDeclaration = declarations.First(declaration => declaration.DeclarationType == DeclarationType.Project);
            var updates = declarations.ToList();

            var project = new CodeExplorerProjectViewModel(projectDeclaration, ref declarations, null, null, CodeExplorerTestSetup.ProjectOneProvider);

            project.Synchronize(ref updates);

            var folder = project.Children.OfType<CodeExplorerCustomFolderViewModel>().Single();

            Assert.AreEqual(projectDeclaration.IdentifierName, folder.Name);
        }

        [Test]
        [Category("Code Explorer")]
        public void Synchronize_ClearsPassedDeclarationList_NoChanges()
        {
            var declarations = CodeExplorerTestSetup.TestProjectOneDeclarations;
            var projectDeclaration = declarations.First(declaration => declaration.DeclarationType == DeclarationType.Project);

            var project = new CodeExplorerProjectViewModel(projectDeclaration, ref declarations, null, null, CodeExplorerTestSetup.ProjectOneProvider);
            var updates = CodeExplorerTestSetup.TestProjectOneDeclarations;
            project.Synchronize(ref updates);

            Assert.AreEqual(0, updates.Count);
        }

        [Test]
        [Category("Code Explorer")]
        public void Synchronize_DoesNotAlterDeclarationList_DifferentProject()
        {
            var declarations = CodeExplorerTestSetup.TestProjectOneDeclarations;
            var projectDeclaration = declarations.First(declaration => declaration.DeclarationType == DeclarationType.Project);

            var project = new CodeExplorerProjectViewModel(projectDeclaration, ref declarations, null, null, CodeExplorerTestSetup.ProjectOneProvider);
            var updates = CodeExplorerTestSetup.TestProjectTwoDeclarations;
            project.Synchronize(ref updates);

            var expected = CodeExplorerTestSetup.TestProjectTwoDeclarations
                .Select(declaration => declaration.QualifiedName.ToString())
                .OrderBy(_ => _);

            var actual = updates.Select(declaration => declaration.QualifiedName.ToString())
                .OrderBy(_ => _);

            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Code Explorer")]
        public void Constructor_PlacesAllTrackedDeclarations()
        {
            var declarations = CodeExplorerTestSetup.TestProjectOneDeclarations;
            var projectDeclaration = declarations.First(declaration => declaration.DeclarationType == DeclarationType.Project);
            var results = declarations.ToList();

            var project = new CodeExplorerProjectViewModel(projectDeclaration, ref declarations, null, null, CodeExplorerTestSetup.ProjectOneProvider);

            var expected = 
                CodeExplorerProjectViewModel.ExtractTrackedDeclarationsForProject(projectDeclaration, ref results)
                .Select(declaration => declaration.QualifiedName.ToString())
                .OrderBy(_ => _);

            var actual = project.GetAllChildDeclarations()
                .Select(declaration => declaration.QualifiedName.ToString())
                .OrderBy(_ => _);

            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Code Explorer")]
        public void Synchronize_PlacesAllTrackedDeclarations_NoChanges()
        {
            var declarations = CodeExplorerTestSetup.TestProjectOneDeclarations;
            var projectDeclaration = declarations.First(declaration => declaration.DeclarationType == DeclarationType.Project);

            var project = new CodeExplorerProjectViewModel(projectDeclaration, ref declarations, null, null, CodeExplorerTestSetup.ProjectOneProvider);
            var updates = CodeExplorerTestSetup.TestProjectOneDeclarations;
            var results = updates.ToList();

            project.Synchronize(ref updates);

            var expected =
                CodeExplorerProjectViewModel.ExtractTrackedDeclarationsForProject(projectDeclaration, ref results)
                .Select(declaration => declaration.QualifiedName.ToString())
                .OrderBy(_ => _);

            var actual = project.GetAllChildDeclarations()
                .Select(declaration => declaration.QualifiedName.ToString())
                .OrderBy(_ => _);

            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Code Explorer")]
        [TestCase(CodeExplorerTestSetup.TestDocumentName, CodeExplorerTestSetup.TestModuleName)]
        [TestCase(CodeExplorerTestSetup.TestModuleName, CodeExplorerTestSetup.TestClassName)]
        [TestCase(CodeExplorerTestSetup.TestClassName, CodeExplorerTestSetup.TestUserFormName)]
        [TestCase(CodeExplorerTestSetup.TestUserFormName, CodeExplorerTestSetup.TestDocumentName)]
        public void Synchronize_PlacesAllTrackedDeclarations_AddedComponent(string component, string added)
        {
            var declarations = CodeExplorerTestSetup.TestProjectOneDeclarations
                .TestProjectWithComponentDeclarations(new[] { component },out var projectDeclaration);

            var project = new CodeExplorerProjectViewModel(projectDeclaration, ref declarations, null, null, CodeExplorerTestSetup.ProjectOneProvider);

            var updates = CodeExplorerTestSetup.TestProjectOneDeclarations
                .TestProjectWithComponentDeclarations(new[] { component, added }, out _).ToList();

            var results = updates.ToList();
            var expected =
                CodeExplorerProjectViewModel.ExtractTrackedDeclarationsForProject(projectDeclaration, ref results)
                    .Select(declaration => declaration.QualifiedName.ToString())
                    .OrderBy(_ => _)
                    .ToList();

            project.Synchronize(ref updates);

            var children = project.GetAllChildDeclarations();
            var actual = children
                .Select(declaration => declaration.QualifiedName.ToString())
                .OrderBy(_ => _);

            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Code Explorer")]
        [TestCase(CodeExplorerTestSetup.TestDocumentName, CodeExplorerTestSetup.TestModuleName)]
        [TestCase(CodeExplorerTestSetup.TestModuleName, CodeExplorerTestSetup.TestClassName)]
        [TestCase(CodeExplorerTestSetup.TestClassName, CodeExplorerTestSetup.TestUserFormName)]
        [TestCase(CodeExplorerTestSetup.TestUserFormName, CodeExplorerTestSetup.TestDocumentName)]
        public void Synchronize_AddedComponent_SingleProjectFolderExists(string component, string added)
        {
            var declarations = CodeExplorerTestSetup.TestProjectOneDeclarations
                .TestProjectWithComponentDeclarations(new[] { component }, out var projectDeclaration);

            var project = new CodeExplorerProjectViewModel(projectDeclaration, ref declarations, null, null, CodeExplorerTestSetup.ProjectOneProvider);

            var updates = CodeExplorerTestSetup.TestProjectOneDeclarations
                .TestProjectWithComponentDeclarations(new[] { component, added }, out _);

            project.Synchronize(ref updates);

            var actual = project.Children.OfType<CodeExplorerCustomFolderViewModel>()
                .Count(folder => folder.Name.Equals(CodeExplorerTestSetup.TestProjectOneName));

            Assert.AreEqual(1, actual);
        }

        [Test]
        [Category("Code Explorer")]
        [TestCase(CodeExplorerTestSetup.TestDocumentName)]
        [TestCase(CodeExplorerTestSetup.TestModuleName)]
        [TestCase(CodeExplorerTestSetup.TestClassName)]
        [TestCase(CodeExplorerTestSetup.TestUserFormName)]
        public void Synchronize_RemovesComponent(string removed)
        {
            var declarations = CodeExplorerTestSetup.TestProjectOneDeclarations;
            var projectDeclaration = declarations.First(declaration => declaration.DeclarationType == DeclarationType.Project);

            var project = new CodeExplorerProjectViewModel(projectDeclaration, ref declarations, null, null, CodeExplorerTestSetup.ProjectOneProvider);

            var updates = CodeExplorerTestSetup.TestProjectOneDeclarations.TestProjectWithComponentRemoved(removed);
            var expected = updates.Select(declaration => declaration.QualifiedName.ToString())
                .OrderBy(_ => _)
                .ToList();

            project.Synchronize(ref updates);

            var actual = project.GetAllChildDeclarations()
                .Select(declaration => declaration.QualifiedName.ToString())
                .OrderBy(_ => _);

            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Code Explorer")]
        public void Synchronize_SetsDeclarationNull_NoDeclarationsForProject()
        {
            var declarations = CodeExplorerTestSetup.TestProjectOneDeclarations;
            var projectDeclaration = declarations.First(declaration => declaration.DeclarationType == DeclarationType.Project);

            var project = new CodeExplorerProjectViewModel(projectDeclaration, ref declarations, null, null, CodeExplorerTestSetup.ProjectOneProvider);
            if (project.Declaration is null)
            {
                Assert.Inconclusive("Project declaration is null. Fix test setup and see why no other tests failed.");
            }

            var updates = new List<Declaration>();
            project.Synchronize(ref updates);

            Assert.IsNull(project.Declaration);
        }

        [Test]
        [Category("Code Explorer")]
        public void Synchronize_SetsDeclarationNull_DeclarationsForDifferentProject()
        {
            var declarations = CodeExplorerTestSetup.TestProjectOneDeclarations;
            var projectDeclaration = declarations.First(declaration => declaration.DeclarationType == DeclarationType.Project);

            var project = new CodeExplorerProjectViewModel(projectDeclaration, ref declarations, null, null, CodeExplorerTestSetup.ProjectOneProvider);
            if (project.Declaration is null)
            {
                Assert.Inconclusive("Project declaration is null. Fix test setup and see why no other tests failed.");
            }

            var updates = CodeExplorerTestSetup.TestProjectTwoDeclarations;
            project.Synchronize(ref updates);

            Assert.IsNull(project.Declaration);
        }

        [Test]
        [Category("Code Explorer")]
        [TestCase(true, false, TestName = "Constructor_CreatesReferenceFolders_LibrariesOnly")]
        [TestCase(true, true, TestName = "Constructor_CreatesReferenceFolders_LibrariesAndProjects")]
        [TestCase(false, false, TestName = "Constructor_CreatesReferenceFolders_NoReferences")]
        [TestCase(false, true, TestName = "Constructor_CreatesReferenceFolders_ProjectsOnly")]
        public void Constructor_CreatesReferenceFolders(bool libraries, bool projects)
        {
            var declarations = CodeExplorerTestSetup.GetProjectDeclarationsWithReferences(libraries, projects, out var state);
            using (state)
            {
                var projectDeclaration = declarations.First(declaration => declaration.DeclarationType == DeclarationType.Project);

                var project = new CodeExplorerProjectViewModel(projectDeclaration, ref declarations, state, null, state.ProjectsProvider);

                var libraryFolder = project.Children.OfType<CodeExplorerReferenceFolderViewModel>()
                    .SingleOrDefault(folder => folder.ReferenceKind == ReferenceKind.TypeLibrary);
                var projectFolder = project.Children.OfType<CodeExplorerReferenceFolderViewModel>()
                    .SingleOrDefault(folder => folder.ReferenceKind == ReferenceKind.Project);

                Assert.AreEqual(libraries, libraryFolder != null);
                Assert.AreEqual(projects, projectFolder != null);
            }
        }

        [Test]
        [Category("Code Explorer")]
        [TestCase(true, false, TestName = "Synchronize_ReferenceFolders_NoChanges_LibrariesOnly")]
        [TestCase(true, true, TestName = "Synchronize_ReferenceFolders_NoChanges_LibrariesAndProjects")]
        [TestCase(false, false, TestName = "Synchronize_ReferenceFolders_NoChanges_NoReferences")]
        [TestCase(false, true, TestName = "Synchronize_ReferenceFolders_NoChanges_ProjectsOnly")]
        public void Synchronize_ReferenceFolders_NoChanges(bool libraries, bool projects)
        {
            var declarations = CodeExplorerTestSetup.GetProjectDeclarationsWithReferences(libraries, projects, out var state);
            using (state)
            {
                var updates = declarations.ToList();
                var projectDeclaration = declarations.First(declaration => declaration.DeclarationType == DeclarationType.Project);

                var expected = GetReferencesFromProjectDeclaration(projectDeclaration, state.ProjectsProvider).Select(reference => reference.Name).ToList();

                var project = new CodeExplorerProjectViewModel(projectDeclaration, ref declarations, state, null, state.ProjectsProvider);
                project.Synchronize(ref updates);

                var actual = GetReferencesFromProjectViewModel(project).OrderBy(reference => reference.Priority).Select(reference => reference.Name);

                Assert.IsTrue(expected.SequenceEqual(actual));
            }
        }

        [Test]
        [Ignore("References.Remove is not mocked yet. Remove this annotation when it is working.")]
        [Category("Code Explorer")]
        [TestCase(true, false, TestName = "Synchronize_ReferenceFolderRemoved_Libraries")]
        [TestCase(false, true, TestName = "Synchronize_ReferenceFolderRemoved_Projects")]
        [TestCase(true, true, TestName = "Synchronize_ReferenceFolderRemoved_Both")]
        public void Synchronize_ReferenceFolderRemoved(bool libraries, bool projects)
        {
            var declarations = CodeExplorerTestSetup.GetProjectDeclarationsWithReferences(libraries, projects, out var state);
            using (state)
            {
                var updates = declarations.ToList();
                var projectDeclaration = declarations.First(declaration => declaration.DeclarationType == DeclarationType.Project);

                var project = new CodeExplorerProjectViewModel(projectDeclaration, ref declarations, state, null, state.ProjectsProvider);

                var references = state.ProjectsProvider.Project(projectDeclaration.ProjectId).References;
                foreach (var reference in references.ToList())
                {
                    references.Remove(reference);
                }

                project.Synchronize(ref updates);

                var libraryFolder = project.Children.OfType<CodeExplorerReferenceFolderViewModel>()
                    .SingleOrDefault(folder => folder.ReferenceKind == ReferenceKind.TypeLibrary);
                var projectFolder = project.Children.OfType<CodeExplorerReferenceFolderViewModel>()
                    .SingleOrDefault(folder => folder.ReferenceKind == ReferenceKind.Project);

                Assert.IsNull(libraryFolder);
                Assert.IsNull(projectFolder);
            }
        }   

        private static List<ReferenceModel> GetReferencesFromProjectViewModel(ICodeExplorerNode viewModel)
        {
            return viewModel.Children
                .OfType<CodeExplorerReferenceFolderViewModel>()
                .SelectMany(folder => folder.Children
                    .OfType<CodeExplorerReferenceViewModel>()
                    .Select(vm => vm.Reference))
                .ToList();
        }

        private static List<IReference> GetReferencesFromProjectDeclaration(Declaration project, IProjectsProvider projectsProvider)
        {
            return projectsProvider.Project(project.ProjectId).References.ToList();
        }
    }
}
