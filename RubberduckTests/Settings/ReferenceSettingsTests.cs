using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using Moq;
using NUnit.Framework;
using Rubberduck.Settings;
using Rubberduck.UI;
using Rubberduck.UI.Settings;
using Rubberduck.VBEditor;
using RubberduckTests.AddRemoveReferences;

namespace RubberduckTests.Settings
{
    [TestFixture]
    public class ReferenceSettingsTests
    {
        private static AddRemoveReferencesUserSettingsViewModel GetSettingsViewModel(ReferenceSettings settings)
        {
            return new AddRemoveReferencesUserSettingsViewModel(AddRemoveReferencesSetup.GetReferenceSettingsProvider(settings), new Mock<IFileSystemBrowserFactory>().Object, null);
        }

        [Test]
        [Category("Settings")]
        public void CopyCtorCopiesAllValues()
        {
            var settings = AddRemoveReferencesSetup.GetNonDefaultReferenceSettings();
            var copy = new ReferenceSettings(settings);

            Assert.IsTrue(settings.Equals(copy));
        }

        [Test]
        [Category("Settings")]
        public void PinReference_RejectsDuplicateLibraries()
        {
            var library = new ReferenceInfo(AddRemoveReferencesSetup.DummyGuidOne, "Reference", @"C:\Windows\System32\reference.dll", 1, 0);

            var settings = new ReferenceSettings();
            settings.PinReference(library);
            settings.PinReference(library);

            Assert.AreEqual(1, settings.GetPinnedReferencesForHost(null).Count);
        }

        [Test]
        [Category("Settings")]
        public void PinReference_RejectsDuplicateProjects()
        {
            const string host = "EXCEL.EXE";
            var project = new ReferenceInfo(Guid.Empty, "RecentProject", @"C:\Users\Rubberduck\Documents\RecentBook.xlsm", 0, 0);

            var settings = new ReferenceSettings();
            settings.PinReference(project, host);
            settings.PinReference(project, host);

            Assert.AreEqual(1, settings.GetPinnedReferencesForHost(host).Count);
        }

        [Test]
        [Category("Settings")]
        public void UpdatePinnedReferencesForHost_RejectsDuplicateLibraries()
        {
            var library = new ReferenceInfo(AddRemoveReferencesSetup.DummyGuidOne, "Reference", @"C:\Windows\System32\reference.dll", 1, 0);

            var settings = new ReferenceSettings();
            settings.UpdatePinnedReferencesForHost(null, new List<ReferenceInfo> { library });
            settings.UpdatePinnedReferencesForHost(null, new List<ReferenceInfo> { library });

            Assert.AreEqual(1, settings.GetPinnedReferencesForHost(null).Count);
        }

        [Test]
        [Category("Settings")]
        public void UpdatePinnedReferencesForHost_RejectsDuplicateProjects()
        {
            const string host = "EXCEL.EXE";
            var project = new ReferenceInfo(Guid.Empty, "RecentProject", @"C:\Users\Rubberduck\Documents\RecentBook.xlsm", 0, 0);

            var settings = new ReferenceSettings();
            settings.UpdatePinnedReferencesForHost(host, new List<ReferenceInfo> { project });
            settings.UpdatePinnedReferencesForHost(host, new List<ReferenceInfo> { project });

            Assert.AreEqual(1, settings.GetPinnedReferencesForHost(host).Count);
        }

        [Test]
        [Category("Settings")]
        public void TrackUsage_RejectsDuplicateLibraries()
        {
            var library = new ReferenceInfo(AddRemoveReferencesSetup.DummyGuidOne, "Reference", @"C:\Windows\System32\reference.dll", 1, 0);

            var settings = new ReferenceSettings { RecentReferencesTracked = 20 };
            settings.TrackUsage(library);
            settings.TrackUsage(library);

            Assert.AreEqual(1, settings.GetRecentReferencesForHost(null).Count);
        }

        [Test]
        [Category("Settings")]
        public void TrackUsage_RejectsDuplicateProjects()
        {
            const string host = "EXCEL.EXE";
            var project = new ReferenceInfo(Guid.Empty, "RecentProject", @"C:\Users\Rubberduck\Documents\RecentBook.xlsm", 0, 0);

            var settings = new ReferenceSettings { RecentReferencesTracked = 20 };
            settings.TrackUsage(project, host);
            settings.TrackUsage(project, host);

            Assert.AreEqual(1, settings.GetRecentReferencesForHost(host).Count);
        }

        [Test]
        [Category("Settings")]
        public void UpdateRecentReferencesForHost_RejectsDuplicateLibraries()
        {
            var library = new ReferenceInfo(AddRemoveReferencesSetup.DummyGuidOne, "Reference", @"C:\Windows\System32\reference.dll", 1, 0);

            var settings = new ReferenceSettings { RecentReferencesTracked = 20 };
            settings.UpdateRecentReferencesForHost(null, new List<ReferenceInfo> { library });
            settings.UpdateRecentReferencesForHost(null, new List<ReferenceInfo> { library });

            Assert.AreEqual(1, settings.GetRecentReferencesForHost(null).Count);
        }

        [Test]
        [Category("Settings")]
        public void UpdateRecentReferencesForHost_RejectsDuplicateProjects()
        {
            const string host = "EXCEL.EXE";
            var project = new ReferenceInfo(Guid.Empty, "RecentProject", @"C:\Users\Rubberduck\Documents\RecentBook.xlsm", 0, 0);

            var settings = new ReferenceSettings { RecentReferencesTracked = 20 };
            settings.UpdateRecentReferencesForHost(host, new List<ReferenceInfo> { project });
            settings.UpdateRecentReferencesForHost(host, new List<ReferenceInfo> { project });

            Assert.AreEqual(1, settings.GetRecentReferencesForHost(host).Count);
        }

        [Test]
        [Category("Settings")]
        public void TrackUsage_KeepsNewestLibraries()
        {
            var settings = new ReferenceSettings { RecentReferencesTracked = AddRemoveReferencesSetup.LibraryReferenceInfoList.Count };
            settings.UpdateRecentReferencesForHost(null, AddRemoveReferencesSetup.LibraryReferenceInfoList);

            var expected = AddRemoveReferencesSetup.LibraryReferenceInfoList.First();
            settings.TrackUsage(expected);

            var actual = settings.GetRecentReferencesForHost(null).Last();

            Assert.IsTrue(expected.Equals(actual));
        }

        [Test]
        [Category("Settings")]
        public void TrackUsage_KeepsNewestProjects()
        {
            const string host = "EXCEL.EXE";

            var settings = new ReferenceSettings { RecentReferencesTracked = AddRemoveReferencesSetup.ProjectReferenceInfoList.Count };
            settings.UpdateRecentReferencesForHost(host, AddRemoveReferencesSetup.ProjectReferenceInfoList);

            var expected = AddRemoveReferencesSetup.ProjectReferenceInfoList.First();
            settings.TrackUsage(expected, host);

            var actual = settings.GetRecentReferencesForHost(host).Last();

            Assert.IsTrue(expected.Equals(actual));
        }

        [Test]
        [Category("Settings")]
        [TestCase(10, 10)]
        [TestCase(-1, 0)]
        [TestCase(100, ReferenceSettings.RecentTrackingLimit)]
        public void RecentReferencesTracked_LimitedToRange(int input, int expected)
        {
            var settings = new ReferenceSettings { RecentReferencesTracked = input };
            Assert.AreEqual(expected, settings.RecentReferencesTracked);
        }

        [Test]
        [Category("Settings")]
        public void GetRecentReferencesForHostLibraries_LimitedByRecentReferencesTracked()
        {
            const int tracked = 3;

            var settings = new ReferenceSettings { RecentReferencesTracked = tracked };
            settings.UpdateRecentReferencesForHost(null, AddRemoveReferencesSetup.LibraryReferenceInfoList);

            Assert.AreEqual(tracked, settings.GetRecentReferencesForHost(null).Count);
        }

        [Test]
        [Category("Settings")]
        public void GetRecentReferencesForHostProjects_LimitedByRecentReferencesTracked()
        {
            const string host = "EXCEL.EXE";
            const int tracked = 3;

            var settings = new ReferenceSettings { RecentReferencesTracked = tracked };
            settings.UpdateRecentReferencesForHost(host, AddRemoveReferencesSetup.RecentProjectReferenceInfoList);

            Assert.AreEqual(tracked, settings.GetRecentReferencesForHost(host).Count);
        }

        [Test]
        [Category("Settings")]
        public void GetRecentReferencesForHostCombined_LimitedByRecentReferencesTracked()
        {
            const string host = "EXCEL.EXE";
            const int tracked = 7;

            var settings = new ReferenceSettings { RecentReferencesTracked = tracked };
            settings.UpdateRecentReferencesForHost(null, AddRemoveReferencesSetup.RecentLibraryReferenceInfoList);
            settings.UpdateRecentReferencesForHost(host, AddRemoveReferencesSetup.RecentProjectReferenceInfoList);

            Assert.AreEqual(tracked, settings.GetRecentReferencesForHost(host).Count);
        }

        [Test]
        [Category("Settings")]
        [TestCase(@"C:\Foo\bar.xlsm", @"C:\Foo\bar.xlsm", "EXCEL.EXE", "EXCEL.EXE", true)]
        [TestCase(@"C:\FOO\BAR.XLSM", @"c:\foo\bar.xlsm", "EXCEL.EXE", "EXCEL.EXE", true)]
        [TestCase(@"c:\foo\bar.xlsm", @"C:\FOO\BAR.XLSM", "EXCEL.EXE", "EXCEL.EXE", true)]
        [TestCase(@"C:\Foo\bar.xlsm", @"C:\Bar\foo.xlsm", "EXCEL.EXE", "EXCEL.EXE", false)]
        [TestCase(@"C:\Foo\bar.xlsm", @"X:\Foo\bar.xlsm", "EXCEL.EXE", "EXCEL.EXE", false)]
        [TestCase(@"C:\Foo\bar.xlsm", @"C:\Foo\bar.xlsm", "WINWORD.EXE", "EXCEL.EXE", false)]
        [TestCase(@"C:\Foo\bar.xlsm", @"C:\Foo\bar.xlsm", "EXCEL.EXE", "WINWORD.EXE", false)]
        public void IsPinnedProject_CorrectResult(string pinned, string tested, string host1, string host2, bool expected)
        {
            var settings = new ReferenceSettings();
            settings.PinReference(new ReferenceInfo(Guid.Empty, string.Empty, pinned, 0, 0), host1);

            Assert.AreEqual(expected, settings.IsPinnedProject(tested, host2));
        }

        [Test]
        [Category("Settings")]
        [TestCase(@"C:\Foo\bar.xlsm", @"C:\Foo\bar.xlsm", "EXCEL.EXE", "EXCEL.EXE", true)]
        [TestCase(@"C:\FOO\BAR.XLSM", @"c:\foo\bar.xlsm", "EXCEL.EXE", "EXCEL.EXE", true)]
        [TestCase(@"c:\foo\bar.xlsm", @"C:\FOO\BAR.XLSM", "EXCEL.EXE", "EXCEL.EXE", true)]
        [TestCase(@"C:\Foo\bar.xlsm", @"C:\Bar\foo.xlsm", "EXCEL.EXE", "EXCEL.EXE", false)]
        [TestCase(@"C:\Foo\bar.xlsm", @"X:\Foo\bar.xlsm", "EXCEL.EXE", "EXCEL.EXE", false)]
        [TestCase(@"C:\Foo\bar.xlsm", @"C:\Foo\bar.xlsm", "WINWORD.EXE", "EXCEL.EXE", false)]
        [TestCase(@"C:\Foo\bar.xlsm", @"C:\Foo\bar.xlsm", "EXCEL.EXE", "WINWORD.EXE", false)]
        public void IsRecentProject_CorrectResult(string pinned, string tested, string host1, string host2, bool expected)
        {
            var settings = new ReferenceSettings { RecentReferencesTracked = 20 };
            settings.TrackUsage(new ReferenceInfo(Guid.Empty, string.Empty, pinned, 0, 0), host1);

            Assert.AreEqual(expected, settings.IsRecentProject(tested, host2));
        }
       
        [Test]
        [Category("Settings")]
        public void UpdateConfig_CallsSave()
        {
            var clean = AddRemoveReferencesSetup.GetDefaultReferenceSettings();
            var provider = AddRemoveReferencesSetup.GetMockReferenceSettingsProvider(clean);
            var viewModel = new AddRemoveReferencesUserSettingsViewModel(provider.Object, new Mock<IFileSystemBrowserFactory>().Object, null);

            viewModel.UpdateConfig(null);
            provider.Verify(m => m.Save(It.IsAny<ReferenceSettings>()), Times.Once);
        }

        [Test]
        [Category("Settings")]
        public void UpdateConfig_UsesLoadedSettingsInstance()
        {
            var clean = AddRemoveReferencesSetup.GetDefaultReferenceSettings();
            var provider = AddRemoveReferencesSetup.GetMockReferenceSettingsProvider(clean);
            var viewModel = new AddRemoveReferencesUserSettingsViewModel(provider.Object, new Mock<IFileSystemBrowserFactory>().Object, null);

            viewModel.UpdateConfig(null);
            provider.Verify(m => m.Save(clean), Times.Once);
        }

        [Test]
        [Category("Settings")]
        [TestCase("EXCEL.EXE")]
        [TestCase(null)]
        public void UpdateConfig_DoesNotChangePinned(string host)
        {
            var clean = AddRemoveReferencesSetup.GetNonDefaultReferenceSettings();
            var expected = clean.GetPinnedReferencesForHost(host);

            var viewModel = GetSettingsViewModel(clean);
            viewModel.UpdateConfig(null);
            var actual = clean.GetPinnedReferencesForHost(host);

            Assert.AreEqual(expected.Count, actual.Count);
            Assert.IsTrue(expected.All(reference => actual.Contains(reference)));
        }

        [Test]
        [Category("Settings")]
        [TestCase("EXCEL.EXE")]
        [TestCase(null)]
        public void UpdateConfig_DoesNotChangeRecent(string host)
        {
            var clean = AddRemoveReferencesSetup.GetNonDefaultReferenceSettings();
            var expected = clean.GetRecentReferencesForHost(host);

            var viewModel = GetSettingsViewModel(clean);
            viewModel.UpdateConfig(null);
            var actual = clean.GetRecentReferencesForHost(host);

            Assert.AreEqual(expected.Count, actual.Count);
            Assert.IsTrue(expected.All(reference => actual.Contains(reference)));
        }

        [Test]
        [Category("Settings")]
        [TestCase("EXCEL.EXE")]
        [TestCase(null)]
        public void SetDefaults_DoesNotChangePinned(string host)
        {
            var clean = AddRemoveReferencesSetup.GetNonDefaultReferenceSettings();
            var expected = clean.GetPinnedReferencesForHost(host);

            var viewModel = GetSettingsViewModel(clean);
            viewModel.SetToDefaults(null);
            var actual = clean.GetPinnedReferencesForHost(host);

            Assert.AreEqual(expected.Count, actual.Count);
            Assert.IsTrue(expected.All(reference => actual.Contains(reference)));
        }

        [Test]
        [Category("Settings")]
        [TestCase("EXCEL.EXE")]
        [TestCase(null)]
        public void SetDefaults_DoesNotChangeRecent(string host)
        {
            var clean = AddRemoveReferencesSetup.GetNonDefaultReferenceSettings();
            var expected = clean.GetRecentReferencesForHost(host);

            var viewModel = GetSettingsViewModel(clean);
            viewModel.SetToDefaults(null);
            var actual = clean.GetRecentReferencesForHost(host);

            Assert.AreEqual(expected.Count, actual.Count);
            Assert.IsTrue(expected.All(reference => actual.Contains(reference)));
        }

        [Test]
        [Category("Settings")]
        public void SettingsTransferToViewModel()
        {
            var clean = AddRemoveReferencesSetup.GetNonDefaultReferenceSettings();
            var viewModel = GetSettingsViewModel(clean);

            Assert.Multiple(() =>
            {
                Assert.AreEqual(clean.RecentReferencesTracked, viewModel.RecentReferencesTracked);
                Assert.AreEqual(clean.FixBrokenReferences, viewModel.FixBrokenReferences);
                Assert.AreEqual(clean.AddToRecentOnReferenceEvents, viewModel.AddToRecentOnReferenceEvents);
                Assert.IsTrue(clean.ProjectPaths.SequenceEqual(viewModel.ProjectPaths));
            });
        }

        [Test]
        [Category("Settings")]
        public void ViewModelTransfersToSettings()
        {
            var clean = AddRemoveReferencesSetup.GetDefaultReferenceSettings();
            var viewModel = GetSettingsViewModel(clean);

            viewModel.RecentReferencesTracked = 42;
            viewModel.FixBrokenReferences = true;
            viewModel.AddToRecentOnReferenceEvents = true;

            var paths = new List<string> { @"C:\Foo" };
            viewModel.ProjectPaths = new ObservableCollection<string>(paths);
            viewModel.UpdateConfig(null);

            Assert.Multiple(() =>
            {
                Assert.AreEqual(clean.RecentReferencesTracked, 42);
                Assert.AreEqual(clean.FixBrokenReferences, true);
                Assert.AreEqual(clean.AddToRecentOnReferenceEvents, true);
                Assert.IsTrue(clean.ProjectPaths.SequenceEqual(paths));
            });
        }
    }
}
