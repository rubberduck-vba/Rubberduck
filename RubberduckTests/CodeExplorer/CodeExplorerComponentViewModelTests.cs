using System.Collections.Generic;
using NUnit.Framework;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.Annotations;

namespace RubberduckTests.CodeExplorer
{
    [TestFixture]
    public class CodeExplorerComponentViewModelTests
    {
        [Test]
        [Category("Code Explorer")]
        [TestCase(CodeExplorerTestSetup.TestDocumentName)]
        [TestCase(CodeExplorerTestSetup.TestModuleName)]
        [TestCase(CodeExplorerTestSetup.TestClassName)]
        [TestCase(CodeExplorerTestSetup.TestUserFormName)]
        public void Constructor_SetsDeclaration(string name)
        {
            var declarations = CodeExplorerTestSetup.TestProjectOneDeclarations.TestComponentDeclarations(name);
            var componentDeclaration = declarations
                .First(declaration => declaration.DeclarationType.HasFlag(DeclarationType.Module) && declaration.IdentifierName.Equals(name));

            var component = new CodeExplorerComponentViewModel(null, componentDeclaration, ref declarations, null);

            Assert.AreSame(componentDeclaration, component.Declaration);
        }

        [Test]
        [Category("Code Explorer")]
        [TestCase(CodeExplorerTestSetup.TestDocumentName)]
        [TestCase(CodeExplorerTestSetup.TestModuleName)]
        [TestCase(CodeExplorerTestSetup.TestClassName)]
        [TestCase(CodeExplorerTestSetup.TestUserFormName)]
        public void Constructor_SetsName(string name)
        {
            var declarations = CodeExplorerTestSetup.TestProjectOneDeclarations.TestComponentDeclarations(name);
            var componentDeclaration = declarations
                .First(declaration => declaration.DeclarationType.HasFlag(DeclarationType.Module) && declaration.IdentifierName.Equals(name));

            var component = new CodeExplorerComponentViewModel(null, componentDeclaration, ref declarations, null);

            Assert.AreEqual(name, component.Name);
        }

        [Test]
        [Category("Code Explorer")]
        [TestCase(CodeExplorerTestSetup.TestDocumentName)]
        [TestCase(CodeExplorerTestSetup.TestModuleName)]
        [TestCase(CodeExplorerTestSetup.TestClassName)]
        [TestCase(CodeExplorerTestSetup.TestUserFormName)]
        public void Constructor_NameWithSignatureIsSet(string name)
        {
            var declarations = CodeExplorerTestSetup.TestProjectOneDeclarations.TestComponentDeclarations(name);
            var componentDeclaration = declarations
                .First(declaration => declaration.DeclarationType.HasFlag(DeclarationType.Module) && declaration.IdentifierName.Equals(name));

            var component = new CodeExplorerComponentViewModel(null, componentDeclaration, ref declarations, null);

            Assert.IsFalse(string.IsNullOrEmpty(component.NameWithSignature));
        }

        [Test]
        [Category("Code Explorer")]
        [TestCase(CodeExplorerTestSetup.TestDocumentName)]
        [TestCase(CodeExplorerTestSetup.TestModuleName)]
        [TestCase(CodeExplorerTestSetup.TestClassName)]
        [TestCase(CodeExplorerTestSetup.TestUserFormName)]
        public void Constructor_ToolTipIsSet(string name)
        {
            var declarations = CodeExplorerTestSetup.TestProjectOneDeclarations.TestComponentDeclarations(name);
            var componentDeclaration = declarations
                .First(declaration => declaration.DeclarationType.HasFlag(DeclarationType.Module) && declaration.IdentifierName.Equals(name));

            var component = new CodeExplorerComponentViewModel(null, componentDeclaration, ref declarations, null);

            Assert.IsFalse(string.IsNullOrEmpty(component.ToolTip));
        }

        [Test]
        [Category("Code Explorer")]
        [TestCase(CodeExplorerTestSetup.TestDocumentName)]
        [TestCase(CodeExplorerTestSetup.TestModuleName)]
        [TestCase(CodeExplorerTestSetup.TestClassName)]
        [TestCase(CodeExplorerTestSetup.TestUserFormName)]
        public void Constructor_SetsIsExpandedFalse(string name)
        {
            var declarations = CodeExplorerTestSetup.TestProjectOneDeclarations.TestComponentDeclarations(name);
            var componentDeclaration = declarations
                .First(declaration => declaration.DeclarationType.HasFlag(DeclarationType.Module) && declaration.IdentifierName.Equals(name));

            var component = new CodeExplorerComponentViewModel(null, componentDeclaration, ref declarations, null);

            Assert.IsFalse(component.IsExpanded);
        }

        [Test]
        [Category("Code Explorer")]
        public void PredeclaredClassIsPredeclared()
        {
            var projectDeclaration = CodeExplorerTestSetup.TestProjectOneDeclarations
                .First(declaration => declaration.DeclarationType == DeclarationType.Project);
            var componentDeclaration = PredeclaredClassDeclaration(projectDeclaration);

            var declarations = new List<Declaration>();
            var component = new CodeExplorerComponentViewModel(null, componentDeclaration, ref declarations, null);

            Assert.IsTrue(component.IsPredeclared);
        }

        [Test]
        [Category("Code Explorer")]
        public void PredeclaredClassSignatureEndsWithPredeclared()
        {
            var projectDeclaration = CodeExplorerTestSetup.TestProjectOneDeclarations
                .First(declaration => declaration.DeclarationType == DeclarationType.Project);
            var componentDeclaration = PredeclaredClassDeclaration(projectDeclaration);

            var declarations = new List<Declaration>();
            var component = new CodeExplorerComponentViewModel(null, componentDeclaration, ref declarations, null);

            Assert.IsTrue(component.NameWithSignature.EndsWith("(Predeclared)"));
        }

        [Test]
        [Category("Code Explorer")]
        [TestCase(CodeExplorerSortOrder.Undefined)]
        [TestCase(CodeExplorerSortOrder.Name)]
        [TestCase(CodeExplorerSortOrder.CodeLine)]
        [TestCase(CodeExplorerSortOrder.DeclarationType)]
        [TestCase(CodeExplorerSortOrder.DeclarationTypeThenName)]
        [TestCase(CodeExplorerSortOrder.DeclarationTypeThenCodeLine)]
        public void SortComparerIsComponentType(CodeExplorerSortOrder order)
        {
            var declarations = CodeExplorerTestSetup.TestProjectOneDeclarations.TestComponentDeclarations(CodeExplorerTestSetup.TestModuleName);
            var componentDeclaration = declarations
                .First(declaration => declaration.DeclarationType.HasFlag(DeclarationType.Module) &&
                    declaration.IdentifierName.Equals(CodeExplorerTestSetup.TestModuleName));

            var component = new CodeExplorerComponentViewModel(null, componentDeclaration, ref declarations, null);

            Assert.AreEqual(CodeExplorerItemComparer.ComponentType.GetType(), component.SortComparer.GetType());
        }

        [Test]
        [Category("Code Explorer")]
        [TestCase(CodeExplorerTestSetup.TestDocumentName)]
        [TestCase(CodeExplorerTestSetup.TestModuleName)]
        [TestCase(CodeExplorerTestSetup.TestClassName)]
        [TestCase(CodeExplorerTestSetup.TestUserFormName)]
        public void FilteredIsFalseForSubsetsOfName(string name)
        {
            var componentDeclaration = CodeExplorerTestSetup.TestProjectOneDeclarations
                .First(declaration => declaration.DeclarationType.HasFlag(DeclarationType.Module) && declaration.IdentifierName.Equals(name));

            var declarations = new List<Declaration> { componentDeclaration };
            var component = new CodeExplorerComponentViewModel(null, componentDeclaration, ref declarations, null);

            for (var characters = 1; characters <= name.Length; characters++)
            {
                component.Filter = name.Substring(0, characters);
                Assert.IsFalse(component.Filtered);
            }

            for (var position = name.Length - 2; position > 0; position--)
            {
                component.Filter = name.Substring(position);
                Assert.IsFalse(component.Filtered);
            }
        }

        [Test]
        [Category("Code Explorer")]
        [TestCase(CodeExplorerTestSetup.TestDocumentName)]
        [TestCase(CodeExplorerTestSetup.TestModuleName)]
        [TestCase(CodeExplorerTestSetup.TestClassName)]
        [TestCase(CodeExplorerTestSetup.TestUserFormName)]
        public void FilteredIsFalseIfMemberMatches(string name)
        {
            var declarations = CodeExplorerTestSetup.TestProjectOneDeclarations.TestComponentDeclarations(name);
            var componentDeclaration = declarations
                .First(declaration => declaration.DeclarationType.HasFlag(DeclarationType.Module) && declaration.IdentifierName.Equals(name));

            var component = new CodeExplorerComponentViewModel(null, componentDeclaration, ref declarations, null);
            var childName = component.Children.First().Name;

            for (var characters = 1; characters <= childName.Length; characters++)
            {
                component.Filter = childName.Substring(0, characters);
                Assert.IsFalse(component.Filtered);
            }

            for (var position = childName.Length - 2; position > 0; position--)
            {
                component.Filter = childName.Substring(position);
                Assert.IsFalse(component.Filtered);
            }
        }

        [Test]
        [Category("Code Explorer")]
        [TestCase(CodeExplorerTestSetup.TestDocumentName)]
        [TestCase(CodeExplorerTestSetup.TestModuleName)]
        [TestCase(CodeExplorerTestSetup.TestClassName)]
        [TestCase(CodeExplorerTestSetup.TestUserFormName)]
        public void FilteredIsTrueForCharactersNotInName(string name)
        {
            const string testCharacters = "abcdefghijklmnopqrstuwxyz";

            var componentDeclaration = CodeExplorerTestSetup.TestProjectOneDeclarations
                .First(declaration => declaration.DeclarationType.HasFlag(DeclarationType.Module) && declaration.IdentifierName.Equals(name));

            var declarations = new List<Declaration> { componentDeclaration };
            var component = new CodeExplorerComponentViewModel(null, componentDeclaration, ref declarations, null);

            var nonMatching = testCharacters.ToCharArray().Except(name.ToLowerInvariant().ToCharArray());

            foreach (var character in nonMatching.Select(letter => letter.ToString()))
            {
                component.Filter = character;
                Assert.IsTrue(component.Filtered);
            }
        }

        [Test]
        [Category("Code Explorer")]
        [TestCase(CodeExplorerTestSetup.TestDocumentName)]
        [TestCase(CodeExplorerTestSetup.TestModuleName)]
        [TestCase(CodeExplorerTestSetup.TestClassName)]
        [TestCase(CodeExplorerTestSetup.TestUserFormName)]
        public void UnfilteredStateIsRestored(string name)
        {
            var declarations = CodeExplorerTestSetup.TestProjectOneDeclarations.TestComponentDeclarations(name);
            var componentDeclaration = declarations
                .First(declaration => declaration.DeclarationType.HasFlag(DeclarationType.Module) && declaration.IdentifierName.Equals(name));

            var component = new CodeExplorerComponentViewModel(null, componentDeclaration, ref declarations, null);
            var childName = component.Children.First().Name;

            component.IsExpanded = false;
            component.Filter = childName;
            Assert.IsTrue(component.IsExpanded);

            component.Filter = string.Empty;
            Assert.IsFalse(component.IsExpanded);
        }

        [Test]
        [Category("Code Explorer")]
        [TestCase(CodeExplorerTestSetup.TestDocumentName)]
        [TestCase(CodeExplorerTestSetup.TestModuleName)]
        [TestCase(CodeExplorerTestSetup.TestClassName)]
        [TestCase(CodeExplorerTestSetup.TestUserFormName)]
        public void Constructor_PlacesAllTrackedDeclarations(string name)
        {
            var declarations = CodeExplorerTestSetup.TestProjectOneDeclarations.TestComponentDeclarations(name);
            var componentDeclaration = declarations
                .First(declaration => declaration.DeclarationType.HasFlag(DeclarationType.Module) && declaration.IdentifierName.Equals(name));

            var component = new CodeExplorerComponentViewModel(null, componentDeclaration, ref declarations, null);

            var expected = CodeExplorerTestSetup.TestProjectOneDeclarations.TestComponentDeclarations(name)
                .Select(declaration => declaration.QualifiedName.ToString())
                .OrderBy(_ => _);

            var actual = component.GetAllChildDeclarations()
                .Select(declaration => declaration.QualifiedName.ToString())
                .OrderBy(_ => _);

            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Code Explorer")]
        [TestCase(CodeExplorerTestSetup.TestDocumentName)]
        [TestCase(CodeExplorerTestSetup.TestModuleName)]
        [TestCase(CodeExplorerTestSetup.TestClassName)]
        [TestCase(CodeExplorerTestSetup.TestUserFormName)]
        public void Synchronize_ClearsPassedDeclarationList_NoChanges(string name)
        {
            var declarations = CodeExplorerTestSetup.TestProjectOneDeclarations.TestComponentDeclarations(name);
            var componentDeclaration = declarations
                .First(declaration => declaration.DeclarationType.HasFlag(DeclarationType.Module) && declaration.IdentifierName.Equals(name));

            var component = new CodeExplorerComponentViewModel(null, componentDeclaration, ref declarations, null);

            var updates = CodeExplorerTestSetup.TestProjectOneDeclarations.TestComponentDeclarations(name);
            component.Synchronize(ref updates);

            Assert.AreEqual(0, updates.Count);
        }

        [Test]
        [Category("Code Explorer")]
        [TestCase(CodeExplorerTestSetup.TestDocumentName, CodeExplorerTestSetup.TestModuleName)]
        [TestCase(CodeExplorerTestSetup.TestModuleName, CodeExplorerTestSetup.TestClassName)]
        [TestCase(CodeExplorerTestSetup.TestClassName, CodeExplorerTestSetup.TestModuleName)]
        [TestCase(CodeExplorerTestSetup.TestUserFormName, CodeExplorerTestSetup.TestModuleName)]
        public void Synchronize_DoesNotAlterDeclarationList_DifferentComponent(string name, string other)
        {
            var declarations = CodeExplorerTestSetup.TestProjectOneDeclarations.TestComponentDeclarations(name);
            var componentDeclaration = declarations
                .First(declaration => declaration.DeclarationType.HasFlag(DeclarationType.Module) && declaration.IdentifierName.Equals(name));

            var component = new CodeExplorerComponentViewModel(null, componentDeclaration, ref declarations, null);

            var updates = CodeExplorerTestSetup.TestProjectOneDeclarations.TestComponentDeclarations(other);
            component.Synchronize(ref updates);

            var expected = CodeExplorerTestSetup.TestProjectOneDeclarations.TestComponentDeclarations(other)
                .Select(declaration => declaration.QualifiedName.ToString()).OrderBy(_ => _);
            var actual = updates.Select(declaration => declaration.QualifiedName.ToString()).OrderBy(_ => _);

            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Code Explorer")]
        [TestCase(CodeExplorerTestSetup.TestDocumentName)]
        [TestCase(CodeExplorerTestSetup.TestModuleName)]
        [TestCase(CodeExplorerTestSetup.TestClassName)]
        [TestCase(CodeExplorerTestSetup.TestUserFormName)]
        public void Synchronize_PlacesAllTrackedDeclarations_NoChanges(string name)
        {
            var declarations = CodeExplorerTestSetup.TestProjectOneDeclarations.TestComponentDeclarations(name);
            var componentDeclaration = declarations
                .First(declaration => declaration.DeclarationType.HasFlag(DeclarationType.Module) && declaration.IdentifierName.Equals(name));

            var component = new CodeExplorerComponentViewModel(null, componentDeclaration, ref declarations, null);

            var updates = CodeExplorerTestSetup.TestProjectOneDeclarations.TestComponentDeclarations(name);
            component.Synchronize(ref updates);

            var expected = CodeExplorerTestSetup.TestProjectOneDeclarations.TestComponentDeclarations(name)
                .Select(declaration => declaration.QualifiedName.ToString()).OrderBy(_ => _);

            var actual = component.GetAllChildDeclarations()
                .Select(declaration => declaration.QualifiedName.ToString())
                .OrderBy(_ => _);

            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Code Explorer")]
        [TestCase(CodeExplorerTestSetup.TestDocumentName)]
        [TestCase(CodeExplorerTestSetup.TestModuleName)]
        [TestCase(CodeExplorerTestSetup.TestClassName)]
        [TestCase(CodeExplorerTestSetup.TestUserFormName)]
        public void Synchronize_SetsDeclarationNull_NoDeclarationsForComponent(string name)
        {
            var declarations = CodeExplorerTestSetup.TestProjectOneDeclarations;
            var componentDeclaration = declarations
                .First(declaration => declaration.DeclarationType.HasFlag(DeclarationType.Module) && declaration.IdentifierName.Equals(name));

            var original = declarations.TestComponentDeclarations(name);
            var updates = declarations.Except(original).ToList();

            var component = new CodeExplorerComponentViewModel(null, componentDeclaration, ref original, null);
            if (component.Declaration is null)
            {
                Assert.Inconclusive("Component declaration is null. Fix test setup and see why no other tests failed.");
            }

            component.Synchronize(ref updates);

            Assert.IsNull(component.Declaration);
        }

        [Test]
        [Category("Code Explorer")]
        [TestCase(CodeExplorerTestCode.TestSubName)]
        [TestCase(CodeExplorerTestCode.TestSubWithLineLabelName)]
        [TestCase(CodeExplorerTestCode.TestSubWithUnresolvedMemberName)]
        [TestCase(CodeExplorerTestCode.TestFunctionName)]
        [TestCase(CodeExplorerTestCode.TestPropertyName, DeclarationType.PropertyGet)]
        [TestCase(CodeExplorerTestCode.TestPropertyName, DeclarationType.PropertyLet)]
        [TestCase(CodeExplorerTestCode.TestPropertyName, DeclarationType.PropertySet)]
        [TestCase(CodeExplorerTestCode.TestTypeName)]
        [TestCase(CodeExplorerTestCode.TestEnumName)]
        [TestCase(CodeExplorerTestCode.TestConstantName)]
        [TestCase(CodeExplorerTestCode.TestFieldName)]
        [TestCase(CodeExplorerTestCode.TestLibraryFunctionName)]
        [TestCase(CodeExplorerTestCode.TestLibraryProcedureName)]
        [TestCase(CodeExplorerTestCode.TestEventName)]
        public void Synchronize_PlacesAllTrackedDeclarations_AddedMember(string added, DeclarationType type = DeclarationType.Member)
        {
            var declarations = CodeExplorerTestSetup.TestProjectOneDeclarations
                .TestProjectWithMemberRemoved(added, out var componentDeclaration, type);

            var component = new CodeExplorerComponentViewModel(null, componentDeclaration, ref declarations, null);
            var updates = CodeExplorerTestSetup.TestProjectOneDeclarations.TestComponentDeclarations(componentDeclaration.IdentifierName);

            var expected = updates.Select(declaration => declaration.QualifiedName.ToString())
                    .OrderBy(_ => _)
                    .ToList();

            component.Synchronize(ref updates);

            var actual = component.GetAllChildDeclarations()
                .Select(declaration => declaration.QualifiedName.ToString())
                .OrderBy(_ => _);

            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Code Explorer")]
        [TestCase(CodeExplorerTestCode.TestSubName)]
        [TestCase(CodeExplorerTestCode.TestSubWithLineLabelName)]
        [TestCase(CodeExplorerTestCode.TestSubWithUnresolvedMemberName)]
        [TestCase(CodeExplorerTestCode.TestFunctionName)]
        [TestCase(CodeExplorerTestCode.TestPropertyName, DeclarationType.PropertyGet)]
        [TestCase(CodeExplorerTestCode.TestPropertyName, DeclarationType.PropertyLet)]
        [TestCase(CodeExplorerTestCode.TestPropertyName, DeclarationType.PropertySet)]
        [TestCase(CodeExplorerTestCode.TestTypeName)]
        [TestCase(CodeExplorerTestCode.TestEnumName)]
        [TestCase(CodeExplorerTestCode.TestConstantName)]
        [TestCase(CodeExplorerTestCode.TestFieldName)]
        [TestCase(CodeExplorerTestCode.TestLibraryFunctionName)]
        [TestCase(CodeExplorerTestCode.TestLibraryProcedureName)]
        [TestCase(CodeExplorerTestCode.TestEventName)]
        public void Synchronize_RemovesMember(string removed, DeclarationType type = DeclarationType.Member)
        {
            var updates = CodeExplorerTestSetup.TestProjectOneDeclarations
                .TestProjectWithMemberRemoved(removed, out var componentDeclaration, type)
                .TestComponentDeclarations(componentDeclaration.IdentifierName);

            var declarations = CodeExplorerTestSetup.TestProjectOneDeclarations
                .TestComponentDeclarations(componentDeclaration.IdentifierName);

            var component = new CodeExplorerComponentViewModel(null, componentDeclaration, ref declarations, null);

            var expected = updates.Select(declaration => declaration.QualifiedName.ToString())
                    .OrderBy(_ => _)
                    .ToList();

            component.Synchronize(ref updates);

            var actual = component.GetAllChildDeclarations()
                .Select(declaration => declaration.QualifiedName.ToString())
                .OrderBy(_ => _);

            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        private const string PredeclaredClassName = "PredeclaredClass";

        private static Declaration PredeclaredClassDeclaration(Declaration project)
        {
            var attributes = new Attributes();
            attributes.AddPredeclaredIdTypeAttribute();

            return new ClassModuleDeclaration(
                  project.QualifiedName.QualifiedModuleName.QualifyMemberName(PredeclaredClassName),
                  project,
                  PredeclaredClassName,
                  true,
                  Enumerable.Empty<IParseTreeAnnotation>(),
                  attributes);
        }
    }
}
