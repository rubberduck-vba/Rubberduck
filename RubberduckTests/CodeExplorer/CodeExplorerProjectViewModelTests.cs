using NUnit.Framework;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Navigation.CodeExplorer;

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

            var project = new CodeExplorerProjectViewModel(projectDeclaration, declarations, null, null);

            Assert.AreSame(projectDeclaration, project.Declaration);
        }

        [Test]
        [Category("Code Explorer")]
        public void Constructor_SetsName()
        {
            var declarations = CodeExplorerTestSetup.TestProjectOneDeclarations;
            var projectDeclaration = declarations.First(declaration => declaration.DeclarationType == DeclarationType.Project);

            var project = new CodeExplorerProjectViewModel(projectDeclaration, declarations, null, null);

            Assert.AreEqual(CodeExplorerTestSetup.TestProjectOneName, project.Name);
        }

        [Test]
        [Category("Code Explorer")]
        public void Constructor_NameWithSignatureIsSet()
        {
            var declarations = CodeExplorerTestSetup.TestProjectOneDeclarations;
            var projectDeclaration = declarations.First(declaration => declaration.DeclarationType == DeclarationType.Project);

            var project = new CodeExplorerProjectViewModel(projectDeclaration, declarations, null, null);

            Assert.IsFalse(string.IsNullOrEmpty(project.NameWithSignature));
        }

        [Test]
        [Category("Code Explorer")]
        public void Constructor_PanelTitleIsSet()
        {
            var declarations = CodeExplorerTestSetup.TestProjectOneDeclarations;
            var projectDeclaration = declarations.First(declaration => declaration.DeclarationType == DeclarationType.Project);

            var project = new CodeExplorerProjectViewModel(projectDeclaration, declarations, null, null);

            Assert.IsFalse(string.IsNullOrEmpty(project.PanelTitle));
        }

        [Test]
        [Category("Code Explorer")]
        public void Constructor_SetsIsExpandedTrue()
        {
            var declarations = CodeExplorerTestSetup.TestProjectOneDeclarations;
            var projectDeclaration = declarations.First(declaration => declaration.DeclarationType == DeclarationType.Project);

            var project = new CodeExplorerProjectViewModel(projectDeclaration, declarations, null, null);

            Assert.IsTrue(project.IsExpanded);
        }

        [Test]
        [Category("Code Explorer")]
        public void Constructor_ToolTipIsSet()
        {
            var declarations = CodeExplorerTestSetup.TestProjectOneDeclarations;
            var projectDeclaration = declarations.First(declaration => declaration.DeclarationType == DeclarationType.Project);

            var project = new CodeExplorerProjectViewModel(projectDeclaration, declarations, null, null);

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

            var project = new CodeExplorerProjectViewModel(projectDeclaration, declarations, null, null)
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

            var project = new CodeExplorerProjectViewModel(projectDeclaration, declarations, null, null)
            {
                Filter = filter
            };

            Assert.IsFalse(project.Filtered);
        }

        [Test]
        [Category("Code Explorer")]
        [TestCase(DeclarationType.Parameter)]
        [TestCase(DeclarationType.LineLabel)]
        [TestCase(DeclarationType.UnresolvedMember)]        // TODO: Inconclusive pending test setup that will actually create one :-/
        //[TestCase(DeclarationType.BracketedExpression)]   // TODO: This causes a parser error in testing due to no host application.
        public void TrackedDeclarations_ExcludesNonNodeTypes(DeclarationType excluded)
        {
            var declarations = CodeExplorerTestSetup.TestProjectOneDeclarations;
            // Validate test setup:
            if (declarations.All(declaration => declaration.DeclarationType != excluded))
            {
                Assert.Inconclusive("DeclarationType under test not found in test declarations.");
            }

            var projectDeclaration = declarations.First(declaration => declaration.DeclarationType == DeclarationType.Project);
            var tracked = CodeExplorerProjectViewModel.ExtractTrackedDeclarationsForProject(projectDeclaration, declarations);

            Assert.IsFalse(tracked.Any(declaration => declaration.DeclarationType == excluded));
        }

        [Test]
        [Category("Code Explorer")]
        [TestCase(DeclarationType.Variable)]
        [TestCase(DeclarationType.Control)]
        [TestCase(DeclarationType.Constant)]    // TODO: Inconclusive pending test setup that will actually create one :-/
        public void TrackedDeclarations_ExcludesMemberEnclosedTypes(DeclarationType excluded)
        {
            var declarations = CodeExplorerTestSetup.TestProjectOneDeclarations;
            // Validate test setup:
            if (declarations.All(declaration => declaration.DeclarationType != excluded))
            {
                Assert.Inconclusive("DeclarationType under test not found in test declarations.");
            }

            var projectDeclaration = declarations.First(declaration => declaration.DeclarationType == DeclarationType.Project);
            var tracked = CodeExplorerProjectViewModel.ExtractTrackedDeclarationsForProject(projectDeclaration, declarations);

            Assert.IsFalse(tracked.Any(declaration => declaration.DeclarationType == excluded && !declaration.ParentDeclaration.DeclarationType.HasFlag(DeclarationType.Module)));
        }

        [Test]
        [Category("Code Explorer")]
        public void Constructor_CreatesDefaultProjectFolder()
        {
            var declarations = CodeExplorerTestSetup.TestProjectOneDeclarations;
            var projectDeclaration = declarations.First(declaration => declaration.DeclarationType == DeclarationType.Project);

            var project = new CodeExplorerProjectViewModel(projectDeclaration, declarations, null, null);
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
            var project = new CodeExplorerProjectViewModel(projectDeclaration, declarations, null, null);

            Assert.AreEqual(0, declarations.Count);
        }

        [Test]
        [Category("Code Explorer")]
        public void Synchronize_UsesDefaultProjectFolder_NoChanges()
        {
            var declarations = CodeExplorerTestSetup.TestProjectOneDeclarations;
            var projectDeclaration = declarations.First(declaration => declaration.DeclarationType == DeclarationType.Project);

            var project = new CodeExplorerProjectViewModel(projectDeclaration, declarations, null, null);
            project.Synchronize(CodeExplorerTestSetup.TestProjectOneDeclarations);

            var folder = project.Children.OfType<CodeExplorerCustomFolderViewModel>().Single();

            Assert.AreEqual(projectDeclaration.IdentifierName, folder.Name);
        }

        [Test]
        [Category("Code Explorer")]
        public void Synchronize_ClearsPassedDeclarationList_NoChanges()
        {
            var declarations = CodeExplorerTestSetup.TestProjectOneDeclarations;
            var projectDeclaration = declarations.First(declaration => declaration.DeclarationType == DeclarationType.Project);

            var project = new CodeExplorerProjectViewModel(projectDeclaration, declarations, null, null);
            var updates = CodeExplorerTestSetup.TestProjectOneDeclarations;
            project.Synchronize(updates);

            Assert.AreEqual(0, updates.Count);
        }

        [Test]
        [Category("Code Explorer")]
        public void Synchronize_DoesNotAlterDeclarationList_DifferentProject()
        {
            var declarations = CodeExplorerTestSetup.TestProjectOneDeclarations;
            var projectDeclaration = declarations.First(declaration => declaration.DeclarationType == DeclarationType.Project);

            var project = new CodeExplorerProjectViewModel(projectDeclaration, declarations, null, null);
            var updates = CodeExplorerTestSetup.TestProjectTwoDeclarations;
            project.Synchronize(updates);

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

            var project = new CodeExplorerProjectViewModel(projectDeclaration, declarations, null, null);

            var expected = 
                CodeExplorerProjectViewModel.ExtractTrackedDeclarationsForProject(projectDeclaration, CodeExplorerTestSetup.TestProjectOneDeclarations)
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

            var project = new CodeExplorerProjectViewModel(projectDeclaration, declarations, null, null);
            var updates = CodeExplorerTestSetup.TestProjectOneDeclarations;
            project.Synchronize(updates);

            var expected =
                CodeExplorerProjectViewModel.ExtractTrackedDeclarationsForProject(projectDeclaration, CodeExplorerTestSetup.TestProjectOneDeclarations)
                .Select(declaration => declaration.QualifiedName.ToString())
                .OrderBy(_ => _);

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

            var project = new CodeExplorerProjectViewModel(projectDeclaration, declarations, null, null);
            if (project.Declaration is null)
            {
                Assert.Inconclusive("Project declaration is null. Fix test setup and see why no other tests failed.");
            }

            project.Synchronize(Enumerable.Empty<Declaration>().ToList());

            Assert.IsNull(project.Declaration);
        }

        [Test]
        [Category("Code Explorer")]
        public void Synchronize_SetsDeclarationNull_DeclarationsForDifferentProject()
        {
            var declarations = CodeExplorerTestSetup.TestProjectOneDeclarations;
            var projectDeclaration = declarations.First(declaration => declaration.DeclarationType == DeclarationType.Project);

            var project = new CodeExplorerProjectViewModel(projectDeclaration, declarations, null, null);
            if (project.Declaration is null)
            {
                Assert.Inconclusive("Project declaration is null. Fix test setup and see why no other tests failed.");
            }

            project.Synchronize(CodeExplorerTestSetup.TestProjectTwoDeclarations);

            Assert.IsNull(project.Declaration);
        }
    }
}
