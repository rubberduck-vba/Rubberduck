using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Moq;
using NUnit.Framework;
using Rubberduck.AddRemoveReferences;
using Rubberduck.UI;
using Rubberduck.UI.AddRemoveReferences;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;

namespace RubberduckTests.AddRemoveReferences
{
    [TestFixture]
    public class AddRemoveReferencesViewModelTests
    {
        [Test]
        [Category("AddRemoveReferences")]
        public void ViewModelCtor_AddsProjectReferencesInOrder()
        {
            var viewModel = AddRemoveReferencesSetup.ArrangeViewModel(out _, out var projectReferences, out _);

            var actual = new List<ReferenceModel>(viewModel.ProjectReferences.Cast<ReferenceModel>());

            Assert.IsTrue(projectReferences.OrderBy(reference => reference.Priority).SequenceEqual(actual));
        }

        [Test]
        [Category("AddRemoveReferences")]
        public void ViewModelCtor_AvailableDoesNotIncludeProjectReferences()
        {
            var viewModel = AddRemoveReferencesSetup.ArrangeViewModel(out _, out var projectReferences, out _);

            var actual = new List<ReferenceModel>(viewModel.AvailableReferences.Cast<ReferenceModel>());

            Assert.IsTrue(!actual.Any(reference => projectReferences.Contains(reference)));
        }

        [Test]
        [Category("AddRemoveReferences")]
        public void ViewModelCtor_SetsBuiltInCount()
        {
            var viewModel = AddRemoveReferencesSetup.ArrangeViewModel(out _, out var projectReferences, out _);

            var expected = projectReferences.Count(reference => reference.IsBuiltIn);

            Assert.AreEqual(expected, viewModel.BuiltInReferenceCount);
        }

        [Test]
        [Category("AddRemoveReferences")]
        public void ViewModel_NewViewModelIsNotDirty()
        {
            var viewModel = AddRemoveReferencesSetup.ArrangeViewModel();
            Assert.IsFalse(viewModel.IsProjectDirty);
        }

        [Test]
        [Category("AddRemoveReferences")]
        public void ViewModelFilters_RecentIsAppliedToAvailableReferences()
        {
            var viewModel = AddRemoveReferencesSetup.ArrangeViewModel(out var allReferences, out _, out _, true);
            viewModel.SelectedFilter = "Recent";

            var expected = allReferences.Count(reference => !reference.IsReferenced && reference.IsRecent);
            var actual = viewModel.AvailableReferences.Cast<ReferenceModel>().Count();

            Assert.AreEqual(expected, actual);
        }

        [Test]
        [Category("AddRemoveReferences")]
        public void ViewModelFilters_PinnedIsAppliedToAvailableReferences()
        {
            var viewModel = AddRemoveReferencesSetup.ArrangeViewModel(out var allReferences, out _, out _, true);
            viewModel.SelectedFilter = "Pinned";

            var expected = allReferences.Count(reference => !reference.IsReferenced && reference.IsPinned);
            var actual = viewModel.AvailableReferences.Cast<ReferenceModel>().Count();

            Assert.AreEqual(expected, actual);
        }

        [Test]
        [Category("AddRemoveReferences")]
        public void ViewModelFilters_ComTypesIsAppliedToAvailableReferences()
        {
            var viewModel = AddRemoveReferencesSetup.ArrangeViewModel(out var allReferences, out _, out _, true);
            viewModel.SelectedFilter = "ComTypes";

            var expected = allReferences.Count(reference => !reference.IsReferenced && reference.Type == ReferenceKind.TypeLibrary);
            var actual = viewModel.AvailableReferences.Cast<ReferenceModel>().Count();

            Assert.AreEqual(expected, actual);
        }

        [Test]
        [Category("AddRemoveReferences")]
        public void ViewModelFilters_ProjectsIsAppliedToAvailableReferences()
        {
            var viewModel = AddRemoveReferencesSetup.ArrangeViewModel(out var allReferences, out _, out _, true);
            viewModel.SelectedFilter = "Projects";

            var expected = allReferences.Count(reference => !reference.IsReferenced && reference.Type == ReferenceKind.Project);
            var actual = viewModel.AvailableReferences.Cast<ReferenceModel>().Count();

            Assert.AreEqual(expected, actual);
        }

        private static readonly List<ReferenceModel> SearchReferencesList = new List<ReferenceModel>
        {
            new ReferenceModel(new ReferenceInfo(Guid.NewGuid(), "Scripting", @"C:\Shortcut\scripting.dll", 1, 0), ReferenceKind.TypeLibrary),
            new ReferenceModel(new ReferenceInfo(Guid.NewGuid(), "Scripting Regular Expressions", @"C:\Office\regex.dll", 2, 0), ReferenceKind.TypeLibrary),
            new ReferenceModel(new ReferenceInfo(Guid.NewGuid(), "ReferenceOne", @"C:\Libs\reference1.dll", 3, 0), ReferenceKind.TypeLibrary),
            new ReferenceModel(new ReferenceInfo(Guid.NewGuid(), "ReferenceTwo", @"C:\Libs\reference2.dll", 4, 0), ReferenceKind.TypeLibrary),
            new ReferenceModel(new ReferenceInfo(Guid.NewGuid(), string.Empty, @"C:\Libs\empty.dll", 5, 0), ReferenceKind.TypeLibrary),
        };

        [Test]
        [Category("AddRemoveReferences")]
        [TestCase("", 5)]
        [TestCase("scri", 2)]
        [TestCase("Regu", 1)]
        [TestCase("REFERENCE", 2)]
        [TestCase("empty", 1)]
        [TestCase("shortcut", 1)]
        [TestCase("libs", 3)]
        [TestCase("1", 1)]
        [TestCase(null, 5)]
        public void ViewModelFilters_SearchInputFiltersList(string input, int expected)
        {
            var declaration = AddRemoveReferencesSetup.ArrangeMocksAndGetProject();
            var settings = AddRemoveReferencesSetup.GetNonDefaultReferenceSettings();
            var model = new AddRemoveReferencesModel(null, declaration, SearchReferencesList, settings);
            var reconciler = AddRemoveReferencesSetup.ArrangeReferenceReconciler(settings);

            var viewModel = new AddRemoveReferencesViewModel(model, reconciler, new Mock<IFileSystemBrowserFactory>().Object, null);
            viewModel.SelectedFilter = ReferenceFilter.ComTypes.ToString();
            viewModel.Search = input;

            var actual = viewModel.AvailableReferences.OfType<ReferenceModel>().Count();

            Assert.AreEqual(expected, actual);
        }

        [Test]
        [Category("AddRemoveReferences")]
        public void ViewModelAddCommand_AddsSelectedLibrary()
        {
            var viewModel = AddRemoveReferencesSetup.ArrangeViewModel();

            var adding = viewModel.AvailableReferences.OfType<ReferenceModel>().First();
            viewModel.SelectedFilter = ReferenceFilter.ComTypes.ToString();
            viewModel.SelectedLibrary = adding;
            viewModel.AddCommand.Execute(null);

            Assert.IsTrue(viewModel.ProjectReferences.OfType<ReferenceModel>().Contains(adding));
        }

        [Test]
        [Category("AddRemoveReferences")]
        public void ViewModelAddCommand_AddedReferenceIsLast()
        {
            var viewModel = AddRemoveReferencesSetup.ArrangeViewModel();

            var adding = viewModel.AvailableReferences.OfType<ReferenceModel>().First();
            viewModel.SelectedFilter = ReferenceFilter.ComTypes.ToString();
            viewModel.SelectedLibrary = adding;
            viewModel.AddCommand.Execute(null);

            var expected = viewModel.ProjectReferences.OfType<ReferenceModel>().Count();
            Assert.AreEqual(expected, adding.Priority);
        }

        [Test]
        [Category("AddRemoveReferences")]
        public void ViewModelAddCommand_AddedReferenceIsNotAvailable()
        {
            var viewModel = AddRemoveReferencesSetup.ArrangeViewModel();

            var adding = viewModel.AvailableReferences.OfType<ReferenceModel>().First();
            viewModel.SelectedFilter = ReferenceFilter.ComTypes.ToString();
            viewModel.SelectedLibrary = adding;
            viewModel.AddCommand.Execute(null);

            Assert.IsFalse(viewModel.AvailableReferences.OfType<ReferenceModel>().Contains(adding));
        }

        [Test]
        [Category("AddRemoveReferences")]
        public void ViewModelRemoveCommand_RemovedReferenceIsNowAvailable()
        {
            var viewModel = AddRemoveReferencesSetup.ArrangeViewModel();

            var removing = viewModel.ProjectReferences.OfType<ReferenceModel>().First(reference => !reference.IsBuiltIn);
            viewModel.SelectedFilter = ReferenceFilter.ComTypes.ToString();
            viewModel.SelectedReference = removing;
            viewModel.RemoveCommand.Execute(null);

            Assert.IsTrue(viewModel.AvailableReferences.OfType<ReferenceModel>().Contains(removing));
        }

        [Test]
        [Category("AddRemoveReferences")]
        public void ViewModelRemoveCommand_ClearsPriority()
        {
            var viewModel = AddRemoveReferencesSetup.ArrangeViewModel();

            var removing = viewModel.ProjectReferences.OfType<ReferenceModel>().First(reference => !reference.IsBuiltIn);
            viewModel.SelectedReference = removing;
            viewModel.RemoveCommand.Execute(null);

            Assert.IsNull(removing.Priority);
        }

        [Test]
        [Category("AddRemoveReferences")]
        public void ViewModelRemoveCommand_RemovedReferenceIsNotInProject()
        {
            var viewModel = AddRemoveReferencesSetup.ArrangeViewModel();

            var removing = viewModel.ProjectReferences.OfType<ReferenceModel>().First(reference => !reference.IsBuiltIn);
            viewModel.SelectedReference = removing;
            viewModel.RemoveCommand.Execute(null);

            Assert.IsFalse(viewModel.ProjectReferences.OfType<ReferenceModel>().Contains(removing));
        }

        [Test]
        [Category("AddRemoveReferences")]
        public void ViewModelBrowseCommand_ShowsDialog()
        {
            var viewModel = AddRemoveReferencesSetup.ArrangeViewModel(out var dialogFactory, out _);

            dialogFactory.SetupMockedOpenDialog(@"C:\Foo\bar.dll", DialogResult.Cancel, out var dialog);
            viewModel.BrowseCommand.Execute(null);

            dialogFactory.Verify(m => m.CreateOpenFileDialog(), Times.Once);
            dialog.Verify(m => m.ShowDialog(), Times.Once);
        }

        [Test]
        [Category("AddRemoveReferences")]
        public void ViewModelBrowseCommand_CallsLoadLibrary()
        {
            const string filename = @"C:\Foo\bar.dll";
            var viewModel = AddRemoveReferencesSetup.ArrangeViewModel(out var dialogFactory, out var libraryProvider);

            dialogFactory.SetupMockedOpenDialog(filename, DialogResult.OK);
            viewModel.BrowseCommand.Execute(null);

            libraryProvider.Verify(m => m.LoadTypeLibrary(filename), Times.Once);
        }

        [Test]
        [Category("AddRemoveReferences")]
        public void ViewModelBrowseCommand_DoesNotCallLoadLibraryOnCancel()
        {
            const string filename = @"C:\Foo\bar.dll";
            var viewModel = AddRemoveReferencesSetup.ArrangeViewModel(out var dialogFactory, out var libraryProvider);

            dialogFactory.SetupMockedOpenDialog(filename, DialogResult.Cancel);
            viewModel.BrowseCommand.Execute(null);

            libraryProvider.Verify(m => m.LoadTypeLibrary(filename), Times.Never);
        }

        [Test]
        [Category("AddRemoveReferences")]
        public void ViewModelBrowseCommand_DoesNotCallLoadLibraryOnEmptyResult()
        {
            var viewModel = AddRemoveReferencesSetup.ArrangeViewModel(out var dialogFactory, out var libraryProvider);

            dialogFactory.SetupMockedOpenDialog(string.Empty, DialogResult.OK);
            viewModel.BrowseCommand.Execute(null);

            libraryProvider.Verify(m => m.LoadTypeLibrary(string.Empty), Times.Never);
        }

        [Test]
        [Category("AddRemoveReferences")]
        public void ViewModelBrowseCommand_BrowsedLibraryAddedToProject()
        {
            const string path = @"C:\Windows\System32\library.dll";
            const string name = "Library";
            const string description = "Library 1.1";

            var viewModel = AddRemoveReferencesSetup.ArrangeViewModel(out var dialogFactory, out var libraryProvider);
            var info = new ReferenceInfo(AddRemoveReferencesSetup.DummyGuidOne, name, path, 1, 1);
            AddRemoveReferencesSetup.SetupIComLibraryProvider(libraryProvider, info, path, description);

            dialogFactory.SetupMockedOpenDialog(path, DialogResult.OK);
            viewModel.BrowseCommand.Execute(null);

            var expected = viewModel.ProjectReferences.OfType<ReferenceModel>().Last();

            Assert.IsTrue(expected.FullPath.Equals(path));
        }

        [Test]
        [Category("AddRemoveReferences")]
        public void ViewModelBrowseCommand_BrowsedLibraryAddedIfBroken()
        {
            const string path = @"C:\Windows\System32\borked.dll";

            var viewModel = AddRemoveReferencesSetup.ArrangeViewModel(out var dialogFactory, out var libraryProvider);
            libraryProvider.Setup(m => m.LoadTypeLibrary(path)).Throws(new COMException());

            dialogFactory.SetupMockedOpenDialog(path, DialogResult.OK);
            viewModel.BrowseCommand.Execute(null);

            var expected = viewModel.ProjectReferences.OfType<ReferenceModel>().Last();

            Assert.IsTrue(expected.FullPath.Equals(path));
            Assert.IsTrue(expected.IsBroken);
        }

        [Test]
        [Category("AddRemoveReferences")]
        public void ViewModelBrowseCommand_BrowseMatchesAvailableLibraries()
        {
            var viewModel = AddRemoveReferencesSetup.ArrangeViewModel(out var allReferences, out _, out var dialogFactory, out var libraryProvider);

            var expected = allReferences.First(reference => !reference.IsReferenced);
            var browsed = expected.FullPath;
            dialogFactory.SetupMockedOpenDialog(browsed, DialogResult.OK);
            viewModel.BrowseCommand.Execute(null);

            var actual = viewModel.ProjectReferences.OfType<ReferenceModel>().Last();

            libraryProvider.Verify(m => m.LoadTypeLibrary(browsed), Times.Never);
            Assert.AreSame(expected, actual);
        }

        [Test]
        [Category("AddRemoveReferences")]
        public void ViewModelBrowseCommand_BrowseMatchRemovesFromAvailableLibraries()
        {
            var viewModel = AddRemoveReferencesSetup.ArrangeViewModel(out var allReferences, out _, out var dialogFactory, out _);

            var browsed = allReferences.First(reference => !reference.IsReferenced);

            dialogFactory.SetupMockedOpenDialog(browsed.FullPath, DialogResult.OK);
            viewModel.BrowseCommand.Execute(null);

            Assert.IsFalse(viewModel.AvailableReferences.OfType<ReferenceModel>().Contains(browsed));
        }

        [Test]
        [Category("AddRemoveReferences")]
        public void ViewModelMoveUpCommand_MovesPriorityUp()
        {
            var viewModel = AddRemoveReferencesSetup.ArrangeViewModel();
            var referenced = viewModel.ProjectReferences.OfType<ReferenceModel>().ToDictionary(model => model.Priority.GetValueOrDefault());

            var last = referenced.Count;
            var moving = referenced[last];
            var switching = referenced[last - 1];
            viewModel.SelectedReference = moving;
            viewModel.MoveUpCommand.Execute(null);

            Assert.AreEqual(last - 1, moving.Priority);
            Assert.AreEqual(last, switching.Priority);
        }

        [Test]
        [Category("AddRemoveReferences")]
        public void ViewModelMoveUpCommand_DoesNotMoveBeforeLocked()
        {
            var viewModel = AddRemoveReferencesSetup.ArrangeViewModel();
            var referenced = viewModel.ProjectReferences.OfType<ReferenceModel>().ToDictionary(model => model.Priority.GetValueOrDefault());

            var startingPriority = viewModel.BuiltInReferenceCount + 1;
            var moving = referenced[startingPriority];

            viewModel.SelectedReference = moving;
            viewModel.MoveUpCommand.Execute(null);

            Assert.AreEqual(startingPriority, moving.Priority);
        }

        [Test]
        [Category("AddRemoveReferences")]
        public void ViewModelMoveDownCommand_MovesPriorityDown()
        {
            var viewModel = AddRemoveReferencesSetup.ArrangeViewModel();
            var referenced = viewModel.ProjectReferences.OfType<ReferenceModel>().ToDictionary(model => model.Priority.GetValueOrDefault());

            var startingPriority = viewModel.BuiltInReferenceCount + 1;
            var moving = referenced[startingPriority];
            var switching = referenced[startingPriority + 1];

            viewModel.SelectedReference = moving;
            viewModel.MoveDownCommand.Execute(null);

            Assert.AreEqual(startingPriority + 1, moving.Priority);
            Assert.AreEqual(startingPriority, switching.Priority);

        }

        [Test]
        [Category("AddRemoveReferences")]
        public void ViewModelMoveDownCommand_DoesNotMoveLastReference()
        
        {
            var viewModel = AddRemoveReferencesSetup.ArrangeViewModel();
            var referenced = viewModel.ProjectReferences.OfType<ReferenceModel>().ToDictionary(model => model.Priority.GetValueOrDefault());

            var last = referenced.Count;
            var moving = referenced[last];

            viewModel.SelectedReference = moving;
            viewModel.MoveDownCommand.Execute(null);

            Assert.AreEqual(last, moving.Priority);
        }

        [Test]
        [Category("AddRemoveReferences")]
        [TestCase(true, false)]
        [TestCase(false, true)]
        public void ViewModelPinLibraryCommand_TogglesLibraryPin(bool starting, bool ending)
        {
            var viewModel = AddRemoveReferencesSetup.ArrangeViewModel();

            var pinning = viewModel.AvailableReferences.OfType<ReferenceModel>().First();
            pinning.IsPinned = starting;
            viewModel.SelectedLibrary = pinning;

            viewModel.PinLibraryCommand.Execute(null);

            Assert.AreEqual(ending, pinning.IsPinned);
        }

        [Test]
        [Category("AddRemoveReferences")]
        [TestCase(true, false)]
        [TestCase(false, true)]
        public void ViewModelPinReferenceCommand_TogglesReferencePin(bool starting, bool ending)
        {
            var viewModel = AddRemoveReferencesSetup.ArrangeViewModel();

            var pinning = viewModel.ProjectReferences.OfType<ReferenceModel>().First();
            pinning.IsPinned = starting;
            viewModel.SelectedReference = pinning;

            viewModel.PinReferenceCommand.Execute(null);

            Assert.AreEqual(ending, pinning.IsPinned);
        }
    }
}
