using System;
using System.Linq;
using System.Runtime.InteropServices;
using Moq;
using NUnit.Framework;
using Rubberduck.AddRemoveReferences;
using Rubberduck.Resources;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;


namespace RubberduckTests.AddRemoveReferences
{
    [TestFixture]
    public class ReferenceReconcilerTests
    {
        private static readonly ReferenceInfo DummyReferenceInfo = new ReferenceInfo(Guid.Empty, "RecentProject", @"C:\Users\Rubberduck\Documents\RecentBook.xlsm", 0, 0);

        [Test]
        [Category("AddRemoveReferences")]       
        public void GetLibraryInfoFromPath_HandlesProjects()
        {
            const string path = @"C:\Users\Rubberduck\Documents\Book1.xlsm";

            var reconciler = AddRemoveReferencesSetup.ArrangeReferenceReconciler();
            var model = reconciler.GetLibraryInfoFromPath(path);

            Assert.AreEqual(path, model.FullPath);
            Assert.IsFalse(model.IsBroken);
        }

        [Test]
        [Category("AddRemoveReferences")]
        public void GetLibraryInfoFromPath_ProjectsDoNotCallLoadLibrary()
        {
            const string path = @"C:\Users\Rubberduck\Documents\Book1.xlsm";

            var reconciler = AddRemoveReferencesSetup.ArrangeReferenceReconciler(null, out _, out var provider);
            reconciler.GetLibraryInfoFromPath(path);

            provider.Verify(m => m.LoadTypeLibrary(It.IsAny<string>()), Times.Never);
        }

        [Test]
        [TestCase(".olb")]
        [TestCase(".tlb")]
        [TestCase(".dll")]
        [TestCase(".ocx")]
        [TestCase(".exe")]
        [Category("AddRemoveReferences")]
        public void GetLibraryInfoFromPath_LoadLibraryCalledOnTypeExtensions(string extension)
        {
            var reconciler = AddRemoveReferencesSetup.ArrangeReferenceReconciler(null, out _, out var provider);
            var path = $@"C:\Windows\System32\library{extension}";

            AddRemoveReferencesSetup.SetupIComLibraryProvider(provider, new ReferenceInfo(Guid.Empty, "Library", path, 1, 1), path, "Library 1.1");
            reconciler.GetLibraryInfoFromPath(path);

            provider.Verify(m => m.LoadTypeLibrary(It.IsAny<string>()), Times.Once);
        }

        [Test]
        [Category("AddRemoveReferences")]
        public void GetLibraryInfoFromPath_NoExtensionReturnsNull()
        {
            const string path = @"C:\Users\Rubberduck\Documents\Book1";

            var reconciler = AddRemoveReferencesSetup.ArrangeReferenceReconciler();
            var model = reconciler.GetLibraryInfoFromPath(path);

            Assert.IsNull(model);
        }

        [Test]
        [Category("AddRemoveReferences")]
        public void GetLibraryInfoFromPath_GivesBrokenReferenceOnThrow()
        {
            const string path = @"C:\Windows\System32\bad.dll";

            var reconciler = AddRemoveReferencesSetup.ArrangeReferenceReconciler(null, out _, out var provider);
            provider.Setup(m => m.LoadTypeLibrary(path)).Throws(new COMException());
            var model = reconciler.GetLibraryInfoFromPath(path);

            Assert.IsTrue(model.IsBroken);
        }

        [Test]
        [Category("AddRemoveReferences")]
        public void GetLibraryInfoFromPath_LoadLibraryLoadsModel()
        {
            const string path = @"C:\Windows\System32\library.dll";
            const string name = "Library";
            const string description = "Library 1.1";

            var reconciler = AddRemoveReferencesSetup.ArrangeReferenceReconciler(null, out _, out var provider);
            var info = new ReferenceInfo(AddRemoveReferencesSetup.DummyGuidOne, name, path, 1, 1);
            AddRemoveReferencesSetup.SetupIComLibraryProvider(provider, info, path, description);
        
            var model = reconciler.GetLibraryInfoFromPath(path);

            Assert.Multiple(() =>
            {
                Assert.AreEqual(model.Guid, AddRemoveReferencesSetup.DummyGuidOne);
                Assert.AreEqual(model.Name, name);
                Assert.AreEqual(model.Description, description);
                Assert.AreEqual(model.FullPath, path);
                Assert.AreEqual(model.Major, 1);
                Assert.AreEqual(model.Minor, 1);
                Assert.IsFalse(model.IsBroken);
            });
        }

        [Test]
        [TestCase(true)]
        [TestCase(false)]
        [Category("AddRemoveReferences")]
        public void UpdateSettings_UpdatesRecentLibrariesBasedOnFlag(bool updating)
        {
            var settings = AddRemoveReferencesSetup.GetNonDefaultReferenceSettings();
            var reconciler = AddRemoveReferencesSetup.ArrangeReferenceReconciler(settings);

            var input = settings.GetRecentReferencesForHost(null).Select(info =>
                new ReferenceModel(info, ReferenceKind.TypeLibrary) { IsRecent = true }).ToList();

            var added = new ReferenceInfo(AddRemoveReferencesSetup.DummyGuidOne, "Reference", @"C:\Windows\System32\reference.dll", 1, 0);
            var output = input.Union(new []{ new ReferenceModel(added, ReferenceKind.TypeLibrary) { IsReferenced = true } }).ToList();

            var model = AddRemoveReferencesSetup.ArrangeAddRemoveReferencesModel(output, null, settings);

            reconciler.UpdateSettings(model.Object, updating);

            var actual = settings.GetRecentReferencesForHost(null);
            var expected = (updating ? output : input).Select(reference => reference.ToReferenceInfo()).ToList();

            Assert.AreEqual(expected.Count, actual.Count);
            Assert.IsTrue(expected.All(info => actual.Contains(info)));
        }

        [Test]
        [TestCase(true)]
        [TestCase(false)]
        [Category("AddRemoveReferences")]
        public void UpdateSettings_UpdatesRecentProjectsBasedOnFlag(bool updating)
        {
            var settings = AddRemoveReferencesSetup.GetNonDefaultReferenceSettings();
            var reconciler = AddRemoveReferencesSetup.ArrangeReferenceReconciler(settings);

            var input = settings.GetRecentReferencesForHost("EXCEL.EXE").Select(info =>
                new ReferenceModel(info, ReferenceKind.Project) { IsRecent = true }).ToList();

            var added = DummyReferenceInfo;
            var output = input.Union(new[] { new ReferenceModel(added, ReferenceKind.TypeLibrary) { IsReferenced = true } }).ToList();

            var model = AddRemoveReferencesSetup.ArrangeAddRemoveReferencesModel(output, null, settings);

            reconciler.UpdateSettings(model.Object, updating);

            var actual = settings.GetRecentReferencesForHost("EXCEL.EXE");
            var expected = (updating ? output : input).Select(reference => reference.ToReferenceInfo()).ToList();

            Assert.AreEqual(updating ? expected.Count : input.Count, actual.Count);
            Assert.IsTrue(expected.All(info => actual.Contains(info)));
        }

        [Test]
        [Category("AddRemoveReferences")]
        public void UpdateSettings_AddsPinnedLibraries()
        {
            var settings = AddRemoveReferencesSetup.GetDefaultReferenceSettings();
            var reconciler = AddRemoveReferencesSetup.ArrangeReferenceReconciler(settings);

            var input = settings.GetPinnedReferencesForHost(null).Select(info =>
                new ReferenceModel(info, ReferenceKind.TypeLibrary) {IsPinned = true}).ToList();

            var output = input.Union(AddRemoveReferencesSetup.LibraryReferenceInfoList.Select(info =>
                new ReferenceModel(info, ReferenceKind.TypeLibrary) {IsPinned = true})).ToList();

            var model = AddRemoveReferencesSetup.ArrangeAddRemoveReferencesModel(output, null, settings);

            reconciler.UpdateSettings(model.Object);

            var actual = settings.GetPinnedReferencesForHost(null);
            var expected = output.Select(reference => reference.ToReferenceInfo()).ToList();

            Assert.AreEqual(expected.Count, actual.Count);
            Assert.IsTrue(expected.All(info => actual.Contains(info)));
        }

        [Test]
        [Category("AddRemoveReferences")]
        public void UpdateSettings_RemovesPinnedLibraries()
        {
            var settings = AddRemoveReferencesSetup.GetDefaultReferenceSettings();
            var reconciler = AddRemoveReferencesSetup.ArrangeReferenceReconciler(settings);

            var input = settings.GetPinnedReferencesForHost(null).Select(info =>
                new ReferenceModel(info, ReferenceKind.TypeLibrary) { IsPinned = true }).ToList();

            var output = input.Take(1).ToList();

            var model = AddRemoveReferencesSetup.ArrangeAddRemoveReferencesModel(output, null, settings);

            reconciler.UpdateSettings(model.Object);

            var actual = settings.GetPinnedReferencesForHost(null);
            var expected = output.Select(reference => reference.ToReferenceInfo()).ToList();

            Assert.AreEqual(expected.Count, actual.Count);
            Assert.IsTrue(expected.All(info => actual.Contains(info)));
        }

        [Test]
        [Category("AddRemoveReferences")]
        public void TryAddReferenceString_CallsAddFromFile()
        {
            const string file = @"C:\Windows\System32\reference.dll";

            var reconciler = AddRemoveReferencesSetup.ArrangeReferenceReconciler();
            var references = AddRemoveReferencesSetup.GetReferencesMock(out var project, out _);

            reconciler.TryAddReference(project.Object, file);

            references.Verify(m => m.AddFromFile(file), Times.Once);
        }

        [Test]
        [Category("AddRemoveReferences")]
        public void TryAddReferenceString_ReturnsNullOnThrow()
        {
            const string file = @"C:\Windows\System32\reference.dll";

            var reconciler = AddRemoveReferencesSetup.ArrangeReferenceReconciler();
            var references = AddRemoveReferencesSetup.GetReferencesMock(out var project, out _);
            references.Setup(m => m.AddFromFile(file)).Throws(new COMException());

            var model = reconciler.TryAddReference(project.Object, file);

            Assert.IsNull(model);
        }

        [Test]
        [Category("AddRemoveReferences")]
        public void TryAddReferenceString_DisplaysMessageOnThrow()
        {
            const string file = @"C:\Windows\System32\reference.dll";
            const string exception = "Don't mock me.";

            var reconciler = AddRemoveReferencesSetup.ArrangeReferenceReconciler(null, out var messageBox, out _);
            var references = AddRemoveReferencesSetup.GetReferencesMock(out var project, out _);
            references.Setup(m => m.AddFromFile(file)).Throws(new COMException(exception));

            reconciler.TryAddReference(project.Object, file);

            messageBox.Verify(m => m.NotifyWarn(exception, RubberduckUI.References_AddFailedCaption));
        }

        [Test]
        [Category("AddRemoveReferences")]
        public void TryAddReferenceString_ReturnedReferenceIsRecent()
        {
            const string file = @"C:\Windows\System32\reference.dll";

            var reconciler = AddRemoveReferencesSetup.ArrangeReferenceReconciler();
            var references = AddRemoveReferencesSetup.GetReferencesMock(out var project, out var builder);

            var returned = builder.CreateReferenceMock("Reference", file, 1, 1, false).Object;
            references.Setup(m => m.AddFromFile(file)).Returns(returned);

            var model = reconciler.TryAddReference(project.Object, file);

            Assert.IsTrue(model.IsRecent);
        }

        [Test]
        [Category("AddRemoveReferences")]
        public void TryAddReferenceReferenceModel_ReturnedReferenceIsRecent()
        {
            var input = new ReferenceModel(DummyReferenceInfo, 0);
            var reconciler = AddRemoveReferencesSetup.ArrangeReferenceReconciler();
            AddRemoveReferencesSetup.GetReferencesMock(out var project, out _);

            var model = reconciler.TryAddReference(project.Object, input);

            Assert.IsTrue(model.IsRecent);
        }

        [Test]
        [Category("AddRemoveReferences")]
        public void TryAddReferenceReferenceModel_ReturnedReferenceIsLastPriority()
        {
            var input = new ReferenceModel(DummyReferenceInfo, 0);
            var reconciler = AddRemoveReferencesSetup.ArrangeReferenceReconciler();
            var references = AddRemoveReferencesSetup.GetReferencesMock(out var project, out _).Object;

            var priority = reconciler.TryAddReference(project.Object, input).Priority;

            Assert.IsTrue(priority > 0);
        }

        [Test]
        [Category("AddRemoveReferences")]
        public void TryAddReferenceReferenceModel_DisplaysMessageOnThrow()
        {
            var input = new ReferenceModel(DummyReferenceInfo, 0);
            const string exception = "Don't mock me.";

            var reconciler = AddRemoveReferencesSetup.ArrangeReferenceReconciler(null, out var messageBox, out _);
            var references = AddRemoveReferencesSetup.GetReferencesMock(out var project, out _);
            references.Setup(m => m.AddFromFile(input.FullPath)).Throws(new COMException(exception));

            reconciler.TryAddReference(project.Object, input);

            messageBox.Verify(m => m.NotifyWarn(exception, RubberduckUI.References_AddFailedCaption));
        }

        [Test]
        [Category("AddRemoveReferences")]
        public void TryAddReferenceReferenceModel_ReturnsNullOnThrow()
        {
            var input = new ReferenceModel(DummyReferenceInfo, 0);

            var reconciler = AddRemoveReferencesSetup.ArrangeReferenceReconciler();
            var references = AddRemoveReferencesSetup.GetReferencesMock(out var project, out _);
            references.Setup(m => m.AddFromFile(input.FullPath)).Throws(new COMException());

            var model = reconciler.TryAddReference(project.Object, input);

            Assert.IsNull(model);
        }

        [Test]
        [Category("AddRemoveReferences")]
        public void ReconcileReferences_ReturnsEmptyWithoutNewReferences()
        {
            var model = AddRemoveReferencesSetup.ArrangeParsedAddRemoveReferencesModel(null, null, null, out _, out _, out var projectsProvider);
            var reconciler = AddRemoveReferencesSetup.ArrangeReferenceReconciler(projectsProvider);

            var output = reconciler.ReconcileReferences(model.Object);

            Assert.IsEmpty(output);
        }

        [Test]
        [Category("AddRemoveReferences")]
        public void ReconcileReferencesOverload_ReturnsEmptyWithoutNewReferences()
        {
            var model = AddRemoveReferencesSetup.ArrangeParsedAddRemoveReferencesModel(null, null, null, out _, out _, out var projectsProvider);
            var reconciler = AddRemoveReferencesSetup.ArrangeReferenceReconciler(projectsProvider);
            var output = reconciler.ReconcileReferences(model.Object, null);

            Assert.IsEmpty(output);
        }

        [Test]
        [Category("AddRemoveReferences")]
        public void ReconcileReferences_UpdatesSettingsPinned()
        {
            var newReferences = AddRemoveReferencesSetup.LibraryReferenceInfoList
                .Select(reference => new ReferenceModel(reference, ReferenceKind.TypeLibrary)).ToList();

            var model = AddRemoveReferencesSetup.ArrangeParsedAddRemoveReferencesModel(null, newReferences, null, out _, out _, out var projectsProvider);
            var reconciler = AddRemoveReferencesSetup.ArrangeReferenceReconciler(projectsProvider);

            var pinned = newReferences.First();
            pinned.IsPinned = true;

            reconciler.ReconcileReferences(model.Object, newReferences);
            var result = model.Object.Settings.GetPinnedReferencesForHost(null).Exists(info => pinned.Matches(info));

            Assert.IsTrue(result);
        }

        [Test]
        [Category("AddRemoveReferences")]
        public void ReconcileReferences_AllReferencesAreAdded()
        {
            var newReferences = AddRemoveReferencesSetup.LibraryReferenceInfoList
                .Select(reference => new ReferenceModel(reference, ReferenceKind.TypeLibrary)).ToList();

            var model = AddRemoveReferencesSetup.ArrangeParsedAddRemoveReferencesModel(newReferences, newReferences, newReferences, out var references, out var builder, out var projectsProvider);

            var reconciler = AddRemoveReferencesSetup.ArrangeReferenceReconciler(projectsProvider);

            var priority = references.Object.Count;
            foreach (var item in newReferences)
            {
                item.Priority = ++priority;
                var result = builder.CreateReferenceMock(item.Name, item.FullPath, item.Major, item.Minor);
                references.Setup(m => m.AddFromFile(item.FullPath)).Returns(result.Object);
            }

            var added = reconciler.ReconcileReferences(model.Object, newReferences);

            Assert.IsTrue(newReferences.All(reference => added.Contains(reference)));
        }

        [Test]
        [Category("AddRemoveReferences")]
        public void ReconcileReferences_RemoveNotCalledOnBuiltIn()
        {
            var registered = AddRemoveReferencesSetup.DummyReferencesList;
            var model = AddRemoveReferencesSetup.ArrangeParsedAddRemoveReferencesModel(registered, registered, registered, out var references, out _, out var projectsProvider);
            var reconciler = AddRemoveReferencesSetup.ArrangeReferenceReconciler(projectsProvider);

            var vba = references.Object.First(lib => lib.Name.Equals("VBA"));
            var excel = references.Object.First(lib => lib.Name.Equals("Excel"));

            reconciler.ReconcileReferences(model.Object, registered);
            references.Verify(m => m.Remove(vba), Times.Never);
            references.Verify(m => m.Remove(excel), Times.Never);
        }
    }
}