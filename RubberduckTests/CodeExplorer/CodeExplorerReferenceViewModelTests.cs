using System.Linq;
using NUnit.Framework;
using Rubberduck.AddRemoveReferences;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.AddRemoveReferences;

namespace RubberduckTests.CodeExplorer
{
    [TestFixture]
    public class CodeExplorerReferenceViewModelTests
    {
        private static readonly ReferenceModel LibraryReference = AddRemoveReferencesSetup.DummyReferencesList.First();
        private static readonly ReferenceModel ProjectReference = AddRemoveReferencesSetup.DummyProjectsList.First();

        [Test]
        [Category("Code Explorer")]
        [TestCase(ReferenceKind.TypeLibrary)]
        [TestCase(ReferenceKind.Project)]
        public void Constructor_SetsReferenceKind(ReferenceKind type)
        {
            var reference = type == ReferenceKind.TypeLibrary ? LibraryReference : ProjectReference;
            var viewModel = new CodeExplorerReferenceViewModel(null, reference);

            Assert.AreSame(reference, viewModel.Reference);
        }

        [Test]
        [Category("Code Explorer")]
        [TestCase(ReferenceKind.TypeLibrary)]
        [TestCase(ReferenceKind.Project)]
        public void Constructor_SetsName(ReferenceKind type)
        {
            var reference = type == ReferenceKind.TypeLibrary ? LibraryReference : ProjectReference;
            var viewModel = new CodeExplorerReferenceViewModel(null, reference);

            Assert.AreEqual(reference.Name, viewModel.Name);
        }

        [Test]
        [Category("Code Explorer")]
        [TestCase(ReferenceKind.TypeLibrary)]
        [TestCase(ReferenceKind.Project)]
        public void Constructor_SetsNameWithSignature(ReferenceKind type)
        {
            var reference = type == ReferenceKind.TypeLibrary ? LibraryReference : ProjectReference;
            var viewModel = new CodeExplorerReferenceViewModel(null, reference);

            Assert.IsFalse(string.IsNullOrEmpty(viewModel.NameWithSignature));
        }

        [Test]
        [Category("Code Explorer")]
        [TestCase(ReferenceKind.TypeLibrary)]
        [TestCase(ReferenceKind.Project)]
        public void Constructor_SetsDescription(ReferenceKind type)
        {
            var reference = type == ReferenceKind.TypeLibrary ? LibraryReference : ProjectReference;
            var viewModel = new CodeExplorerReferenceViewModel(null, reference);

            Assert.AreEqual(reference.FullPath, viewModel.Description);
        }

        [Test]
        [Category("Code Explorer")]
        [TestCase(ReferenceKind.TypeLibrary)]
        [TestCase(ReferenceKind.Project)]
        public void Constructor_ToolTip_IsDescription(ReferenceKind type)
        {
            var reference = type == ReferenceKind.TypeLibrary ? LibraryReference : ProjectReference;
            var viewModel = new CodeExplorerReferenceViewModel(null, reference);

            Assert.AreEqual(reference.Description, viewModel.ToolTip);
        }

        [Test]
        [Category("Code Explorer")]
        [TestCase(ReferenceKind.TypeLibrary)]
        [TestCase(ReferenceKind.Project)]
        public void UnusedReferenceIsDimmed(ReferenceKind type)
        {
            var reference = type == ReferenceKind.TypeLibrary ? LibraryReference : ProjectReference;
            var viewModel = new CodeExplorerReferenceViewModel(null, reference);

            reference.IsUsed = false;

            Assert.IsTrue(viewModel.IsDimmed);
        }

        [Test]
        [Category("Code Explorer")]
        [TestCase(ReferenceKind.TypeLibrary)]
        [TestCase(ReferenceKind.Project)]
        public void Constructor_SetsIsExpandedFalse(ReferenceKind type)
        {
            var reference = type == ReferenceKind.TypeLibrary ? LibraryReference : ProjectReference;
            var viewModel = new CodeExplorerReferenceViewModel(null, reference);

            Assert.IsFalse(viewModel.IsExpanded);
        }

        [Test]
        [Category("Code Explorer")]
        [TestCase(ReferenceKind.TypeLibrary)]
        [TestCase(ReferenceKind.Project)]
        public void UsedReferenceIsNotDimmed(ReferenceKind type)
        {
            var reference = type == ReferenceKind.TypeLibrary ? LibraryReference : ProjectReference;
            var viewModel = new CodeExplorerReferenceViewModel(null, reference);

            reference.IsUsed = true;

            Assert.IsFalse(viewModel.IsDimmed);
        }

        [Test]
        [Category("Code Explorer")]
        [TestCase(ReferenceKind.TypeLibrary)]
        [TestCase(ReferenceKind.Project)]
        public void ErrorStateCanNotBeSet(ReferenceKind type)
        {
            var reference = type == ReferenceKind.TypeLibrary ? LibraryReference : ProjectReference;
            var viewModel = new CodeExplorerReferenceViewModel(null, reference)
            {
                IsErrorState = true
            };

            Assert.IsFalse(viewModel.IsErrorState);
        }

        [Test]
        [Category("Code Explorer")]
        [TestCase(ReferenceKind.TypeLibrary, true)]
        [TestCase(ReferenceKind.TypeLibrary, false)]
        [TestCase(ReferenceKind.Project, true)]
        [TestCase(ReferenceKind.Project, false)]
        public void ReferenceLockedStateMatchesIsBuiltIn(ReferenceKind type, bool builtIn)
        {
            var reference = type == ReferenceKind.TypeLibrary ? LibraryReference : ProjectReference;
            reference.IsBuiltIn = builtIn;

            var viewModel = new CodeExplorerReferenceViewModel(null, reference);

            Assert.AreEqual(builtIn, viewModel.Locked);
        }

        [Test]
        [Category("Code Explorer")]
        [TestCase(ReferenceKind.TypeLibrary)]
        [TestCase(ReferenceKind.Project)]
        public void FilteredIsTrueForCharactersNotInName(ReferenceKind type)
        {
            const string testCharacters = "abcdefghijklmnopqrstuwxyz";

            var reference = type == ReferenceKind.TypeLibrary ? LibraryReference : ProjectReference;
            var viewModel = new CodeExplorerReferenceViewModel(null, reference);
            var name = viewModel.Name;

            var nonMatching = testCharacters.ToCharArray().Except(name.ToLowerInvariant().ToCharArray());

            foreach (var character in nonMatching.Select(letter => letter.ToString()))
            {
                viewModel.Filter = character;
                Assert.IsTrue(viewModel.Filtered);
            }
        }

        [Test]
        [Category("Code Explorer")]
        [TestCase(ReferenceKind.TypeLibrary)]
        [TestCase(ReferenceKind.Project)]
        public void FilteredIsFalseForSubsetsOfName(ReferenceKind type)
        {
            var reference = type == ReferenceKind.TypeLibrary ? LibraryReference : ProjectReference;
            var viewModel = new CodeExplorerReferenceViewModel(null, reference);
            var name = viewModel.Name;

            for (var characters = 1; characters <= name.Length; characters++)
            {
                viewModel.Filter = name.Substring(0, characters);
                Assert.IsFalse(viewModel.Filtered);
            }

            for (var position = name.Length - 2; position > 0; position--)
            {
                viewModel.Filter = name.Substring(position);
                Assert.IsFalse(viewModel.Filtered);
            }
        }


        [Test]
        [Category("Code Explorer")]
        [TestCase(CodeExplorerSortOrder.Undefined)]
        [TestCase(CodeExplorerSortOrder.Name)]
        [TestCase(CodeExplorerSortOrder.CodeLine)]
        [TestCase(CodeExplorerSortOrder.DeclarationType)]
        [TestCase(CodeExplorerSortOrder.DeclarationTypeThenName)]
        [TestCase(CodeExplorerSortOrder.DeclarationTypeThenCodeLine)]
        public void SortComparerIsPriority_LibraryReference(CodeExplorerSortOrder order)
        {
            var viewModel = new CodeExplorerReferenceViewModel(null, LibraryReference)
            {
                SortOrder = order
            };

            Assert.AreEqual(CodeExplorerItemComparer.ReferencePriority.GetType(), viewModel.SortComparer.GetType());
        }

        [Test]
        [Category("Code Explorer")]
        [TestCase(CodeExplorerSortOrder.Undefined)]
        [TestCase(CodeExplorerSortOrder.Name)]
        [TestCase(CodeExplorerSortOrder.CodeLine)]
        [TestCase(CodeExplorerSortOrder.DeclarationType)]
        [TestCase(CodeExplorerSortOrder.DeclarationTypeThenName)]
        [TestCase(CodeExplorerSortOrder.DeclarationTypeThenCodeLine)]
        public void SortComparerIsPriority_ProjectReference(CodeExplorerSortOrder order)
        {
            var viewModel = new CodeExplorerReferenceViewModel(null, ProjectReference)
            {
                SortOrder = order
            };

            Assert.AreEqual(CodeExplorerItemComparer.ReferencePriority.GetType(), viewModel.SortComparer.GetType());
        }

        [Test]
        [Category("Code Explorer")]
        [TestCase(ReferenceKind.TypeLibrary)]
        [TestCase(ReferenceKind.Project)]
        public void Synchronize_RemovesReferenceFromList(ReferenceKind type)
        {
            var references = type == ReferenceKind.TypeLibrary
                ? AddRemoveReferencesSetup.DummyReferencesList
                : AddRemoveReferencesSetup.DummyProjectsList;

            var synching = references.First();

            var viewModel = new CodeExplorerReferenceViewModel(null, synching);
            viewModel.Synchronize(null, references);

            Assert.IsFalse(references.Contains(synching));
        }

        [Test]
        [Category("Code Explorer")]
        [TestCase(ReferenceKind.TypeLibrary)]
        [TestCase(ReferenceKind.Project)]
        public void Synchronize_Unmatched_SetsReferenceNull(ReferenceKind type)
        {
            var references = type == ReferenceKind.TypeLibrary
                ? AddRemoveReferencesSetup.DummyReferencesList
                : AddRemoveReferencesSetup.DummyProjectsList;

            var removing = references.First();

            var viewModel = new CodeExplorerReferenceViewModel(null, removing);

            references.Remove(removing);
            viewModel.Synchronize(null, references);

            Assert.IsNull(viewModel.Reference);
        }
    }
}
