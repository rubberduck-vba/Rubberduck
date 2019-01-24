using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NUnit.Framework;
using Rubberduck.AddRemoveReferences;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Navigation.Folders;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Resources.CodeExplorer;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.AddRemoveReferences;

namespace RubberduckTests.CodeExplorer
{
    [TestFixture]
    public class CodeExplorerReferenceFolderViewModelTests
    {
        [Test]
        [Category("Code Explorer")]
        [TestCase(ReferenceKind.TypeLibrary)]
        [TestCase(ReferenceKind.Project)]
        public void Constructor_SetsReferenceKind(ReferenceKind type)
        {
            var folder = new CodeExplorerReferenceFolderViewModel(null, null, Enumerable.Empty<ReferenceModel>().ToList(), type);

            Assert.AreEqual(type, folder.ReferenceKind);
        }

        [Test]
        [Category("Code Explorer")]
        public void Constructor_SetsFolderName_TypeLibraries()
        {
            var references = AddRemoveReferencesSetup.DummyReferencesList;
            var folder = new CodeExplorerReferenceFolderViewModel(null, null, references, ReferenceKind.TypeLibrary);

            Assert.AreEqual(CodeExplorerUI.CodeExplorer_LibraryReferences, folder.Name);
        }

        [Test]
        [Category("Code Explorer")]
        public void Constructor_SetsFolderName_Projects()
        {
            var references = AddRemoveReferencesSetup.DummyProjectsList;
            var folder = new CodeExplorerReferenceFolderViewModel(null, null, references, ReferenceKind.Project);

            Assert.AreEqual(CodeExplorerUI.CodeExplorer_ProjectReferences, folder.Name);
        }

        [Test]
        [Category("Code Explorer")]
        public void Constructor_NameWithSignatureIsName_TypeLibraries()
        {
            var references = AddRemoveReferencesSetup.DummyReferencesList;
            var folder = new CodeExplorerReferenceFolderViewModel(null, null, references, ReferenceKind.TypeLibrary);

            Assert.AreEqual(CodeExplorerUI.CodeExplorer_LibraryReferences, folder.NameWithSignature);
        }

        [Test]
        [Category("Code Explorer")]
        public void Constructor_NameWithSignatureIsName_Projects()
        {
            var references = AddRemoveReferencesSetup.DummyProjectsList;
            var folder = new CodeExplorerReferenceFolderViewModel(null, null, references, ReferenceKind.Project);

            Assert.AreEqual(CodeExplorerUI.CodeExplorer_ProjectReferences, folder.NameWithSignature);
        }

        [Test]
        [Category("Code Explorer")]
        public void Constructor_PanelTitleIsName_TypeLibraries()
        {
            var references = AddRemoveReferencesSetup.DummyReferencesList;
            var folder = new CodeExplorerReferenceFolderViewModel(null, null, references, ReferenceKind.TypeLibrary);

            Assert.AreEqual(CodeExplorerUI.CodeExplorer_LibraryReferences, folder.PanelTitle);
        }

        [Test]
        [Category("Code Explorer")]
        public void Constructor_PanelTitleIsName_Projects()
        {
            var references = AddRemoveReferencesSetup.DummyProjectsList;
            var folder = new CodeExplorerReferenceFolderViewModel(null, null, references, ReferenceKind.Project);

            Assert.AreEqual(CodeExplorerUI.CodeExplorer_ProjectReferences, folder.PanelTitle);
        }

        [Test]
        [Category("Code Explorer")]
        [TestCase(ReferenceKind.TypeLibrary)]
        [TestCase(ReferenceKind.Project)]
        public void Constructor_DescriptionIsEmpty(ReferenceKind type)
        {
            var folder = new CodeExplorerReferenceFolderViewModel(null, null, Enumerable.Empty<ReferenceModel>().ToList(), type);

            Assert.AreEqual(string.Empty, folder.Description);
        }

        [Test]
        [Category("Code Explorer")]
        [TestCase(ReferenceKind.TypeLibrary)]
        [TestCase(ReferenceKind.Project)]
        public void Constructor_ToolTipIsSet(ReferenceKind type)
        {
            var folder = new CodeExplorerReferenceFolderViewModel(null, null, Enumerable.Empty<ReferenceModel>().ToList(), type);

            Assert.IsFalse(string.IsNullOrEmpty(folder.ToolTip));
        }

        [Test]
        [Category("Code Explorer")]
        [TestCase(ReferenceKind.TypeLibrary)]
        [TestCase(ReferenceKind.Project)]
        public void Constructor_SetsIsExpandedFalse(ReferenceKind type)
        {
            var folder = new CodeExplorerReferenceFolderViewModel(null, null, Enumerable.Empty<ReferenceModel>().ToList(), type);

            Assert.IsFalse(folder.IsExpanded);
        }

        [Test]
        [Category("Code Explorer")]
        [TestCase(CodeExplorerSortOrder.Undefined)]
        [TestCase(CodeExplorerSortOrder.Name)]
        [TestCase(CodeExplorerSortOrder.CodeLine)]
        [TestCase(CodeExplorerSortOrder.DeclarationType)]
        [TestCase(CodeExplorerSortOrder.DeclarationTypeThenName)]
        [TestCase(CodeExplorerSortOrder.DeclarationTypeThenCodeLine)]
        public void SortComparerIsReferenceType_TypeLibraries(CodeExplorerSortOrder order)
        {
            var references = AddRemoveReferencesSetup.DummyReferencesList;
            var folder = new CodeExplorerReferenceFolderViewModel(null, null, references, ReferenceKind.TypeLibrary)
            {
                SortOrder = order
            };

            Assert.AreEqual(CodeExplorerItemComparer.ReferenceType.GetType(), folder.SortComparer.GetType());
        }

        [Test]
        [Category("Code Explorer")]
        [TestCase(CodeExplorerSortOrder.Undefined)]
        [TestCase(CodeExplorerSortOrder.Name)]
        [TestCase(CodeExplorerSortOrder.CodeLine)]
        [TestCase(CodeExplorerSortOrder.DeclarationType)]
        [TestCase(CodeExplorerSortOrder.DeclarationTypeThenName)]
        [TestCase(CodeExplorerSortOrder.DeclarationTypeThenCodeLine)]
        public void SortComparerIsReferenceType_Projects(CodeExplorerSortOrder order)
        {
            var references = AddRemoveReferencesSetup.DummyProjectsList;
            var folder = new CodeExplorerReferenceFolderViewModel(null, null, references, ReferenceKind.Project)
            {
                SortOrder = order
            };

            Assert.AreEqual(CodeExplorerItemComparer.ReferenceType.GetType(), folder.SortComparer.GetType());
        }

        [Test]
        [Category("Code Explorer")]
        [TestCase(ReferenceKind.TypeLibrary)]
        [TestCase(ReferenceKind.Project)]
        public void FilteredIsFalseForAnyCharacter(ReferenceKind type)
        {
            const string testCharacters = "abcdefghijklmnopqrstuwxyz";

            var folder = new CodeExplorerReferenceFolderViewModel(null, null, Enumerable.Empty<ReferenceModel>().ToList(), type);

            foreach (var character in testCharacters.ToCharArray().Select(letter => letter.ToString()))
            {
                folder.Filter = character;
                Assert.IsFalse(folder.Filtered);
            }
        }

        [Test]
        [Category("Code Explorer")]
        public void UnfilteredStateIsRestored_TypeLibraries()
        {
            var references = AddRemoveReferencesSetup.DummyReferencesList;
            var folder = new CodeExplorerReferenceFolderViewModel(null, null, references, ReferenceKind.TypeLibrary);
            var childName = folder.Children.First().Name;

            folder.IsExpanded = false;
            folder.Filter = childName;
            Assert.IsTrue(folder.IsExpanded);

            folder.Filter = string.Empty;
            Assert.IsFalse(folder.IsExpanded);
        }

        [Test]
        [Category("Code Explorer")]
        public void UnfilteredStateIsRestored_Projects()
        {
            var references = AddRemoveReferencesSetup.DummyProjectsList;
            var folder = new CodeExplorerReferenceFolderViewModel(null, null, references, ReferenceKind.Project);
            var childName = folder.Children.First().Name;

            folder.IsExpanded = false;
            folder.Filter = childName;
            Assert.IsTrue(folder.IsExpanded);

            folder.Filter = string.Empty;
            Assert.IsFalse(folder.IsExpanded);
        }

        [Test]
        [Category("Code Explorer")]
        [TestCase(ReferenceKind.TypeLibrary)]
        [TestCase(ReferenceKind.Project)]
        public void ErrorStateCanNotBeSet(ReferenceKind type)
        {
            var folder = new CodeExplorerReferenceFolderViewModel(null, null, Enumerable.Empty<ReferenceModel>().ToList(), type)
            {
                IsErrorState = true
            };

            Assert.IsFalse(folder.IsErrorState);
        }

        [Test]
        [Category("Code Explorer")]
        public void Constructor_PlacesAllReferences_TypeLibraries()
        {
            var references = AddRemoveReferencesSetup.DummyReferencesList;
            var expected = references.Count;

            var folder = new CodeExplorerReferenceFolderViewModel(null, null, references, ReferenceKind.TypeLibrary);

            Assert.AreEqual(expected, folder.Children.Count);
        }

        [Test]
        [Category("Code Explorer")]
        public void Constructor_PlacesAllReferences_Projects()
        {
            var references = AddRemoveReferencesSetup.DummyProjectsList;
            var expected = references.Count;

            var folder = new CodeExplorerReferenceFolderViewModel(null, null, references, ReferenceKind.Project);

            Assert.AreEqual(expected, folder.Children.Count);
        }

        [Test]
        [Category("Code Explorer")]
        public void Synchronize_PlacesAllReferences_TypeLibraries_NoChanges()
        {
            var references = AddRemoveReferencesSetup.DummyReferencesList;
            var folder = new CodeExplorerReferenceFolderViewModel(null, null, references, ReferenceKind.TypeLibrary);

            var updates = AddRemoveReferencesSetup.DummyReferencesList;
            var expected = updates.Select(reference => reference.Name).OrderBy(_ => _).ToList();

            folder.Synchronize(null, updates);
            var actual = folder.Children.Cast<CodeExplorerReferenceViewModel>().Select(reference => reference.Reference.Name).OrderBy(_ => _);

            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Code Explorer")]
        public void Synchronize_PlacesAllReferences_Projects_NoChanges()
        {
            var references = AddRemoveReferencesSetup.DummyProjectsList;
            // ReSharper disable once UnusedVariable
            var folder = new CodeExplorerReferenceFolderViewModel(null, null, references, ReferenceKind.Project);

            var updates = AddRemoveReferencesSetup.DummyProjectsList;
            var expected = updates.Select(reference => reference.Name).OrderBy(_ => _).ToList();

            folder.Synchronize(null, updates);
            var actual = folder.Children.Cast<CodeExplorerReferenceViewModel>().Select(reference => reference.Reference.Name).OrderBy(_ => _);

            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Code Explorer")]
        public void Synchronize_PlacesAllReferences_TypeLibraryAdded()
        {
            var references = AddRemoveReferencesSetup.DummyReferencesList;
            var folder = new CodeExplorerReferenceFolderViewModel(null, null, references.Take(references.Count - 1).ToList(), ReferenceKind.TypeLibrary);

            var updates = AddRemoveReferencesSetup.DummyReferencesList;
            var expected = updates.Select(reference => reference.Name).OrderBy(_ => _).ToList();

            folder.Synchronize(null, updates);
            var actual = folder.Children.Cast<CodeExplorerReferenceViewModel>().Select(reference => reference.Reference.Name).OrderBy(_ => _);

            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Code Explorer")]
        public void Synchronize_PlacesAllReferences_ProjectAdded()
        {
            var references = AddRemoveReferencesSetup.DummyProjectsList;
            // ReSharper disable once UnusedVariable
            var folder = new CodeExplorerReferenceFolderViewModel(null, null, references.Take(references.Count - 1).ToList(), ReferenceKind.Project);

            var updates = AddRemoveReferencesSetup.DummyProjectsList;
            var expected = updates.Select(reference => reference.Name).OrderBy(_ => _).ToList();

            folder.Synchronize(null, updates);
            var actual = folder.Children.Cast<CodeExplorerReferenceViewModel>().Select(reference => reference.Reference.Name).OrderBy(_ => _);

            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Code Explorer")]
        public void Synchronize_RemovesReference_TypeLibraryRemoved()
        {
            var references = AddRemoveReferencesSetup.DummyReferencesList.ToList();
            var updates = AddRemoveReferencesSetup.DummyReferencesList.Take(references.Count - 1).ToList();
            var expected = updates.Select(reference => reference.Name).OrderBy(_ => _).ToList();

            var folder = new CodeExplorerReferenceFolderViewModel(null, null, references, ReferenceKind.TypeLibrary);
            folder.Synchronize(null, updates);
            var actual = folder.Children.Cast<CodeExplorerReferenceViewModel>().Select(reference => reference.Reference.Name).OrderBy(_ => _);

            Assert.IsTrue(expected.SequenceEqual(actual));
        }

        [Test]
        [Category("Code Explorer")]
        public void Synchronize_RemovesReference_ProjectRemoved()
        {
            var references = AddRemoveReferencesSetup.DummyProjectsList.ToList();
            var updates = AddRemoveReferencesSetup.DummyProjectsList.Take(references.Count - 1).ToList();
            var expected = updates.Select(reference => reference.Name).OrderBy(_ => _).ToList();

            var folder = new CodeExplorerReferenceFolderViewModel(null, null, references, ReferenceKind.Project);
            folder.Synchronize(null, updates);
            var actual = folder.Children.Cast<CodeExplorerReferenceViewModel>().Select(reference => reference.Reference.Name).OrderBy(_ => _);

            Assert.IsTrue(expected.SequenceEqual(actual));
        }
    }
}
