using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;
using System.Windows.Forms;
using Moq;
using Rubberduck.AddRemoveReferences;
using Rubberduck.Interaction;
using Rubberduck.Parsing.ComReflection;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Resources.Registration;
using Rubberduck.Settings;
using Rubberduck.SettingsProvider;
using Rubberduck.UI;
using Rubberduck.UI.AddRemoveReferences;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;

namespace RubberduckTests.AddRemoveReferences
{
    public static class AddRemoveReferencesSetup
    {
        // Note that these are just random for tests, don't use these for vbe7.dll or excel.exe in reality...
        public static Guid VbaGuid = new Guid("c331e9a5-9f55-45d8-ab1c-3a6cb9b4e3c9");
        public static Guid ExcelGuid = new Guid("e58523e5-ad69-48fe-990c-712df2180ebc");
        public static Guid DummyGuidOne = new Guid(Enumerable.Range(1, 16).Select(x => (byte)x).ToArray());
        public static Guid DummyGuidTwo = new Guid(Enumerable.Range(2, 16).Select(x => (byte)x).ToArray());

        public static List<ReferenceInfo> LibraryReferenceInfoList =>
            Enumerable.Range(1, 5)
                .Select(info =>
                    new ReferenceInfo(new Guid(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, (byte)info), $"Reference{info}", $@"C:\Windows\System32\ref{info}.dll", 1, 0))
                .ToList();

        public static List<ReferenceInfo> ProjectReferenceInfoList =>
            Enumerable.Range(1, 5)
                .Select(info =>
                    new ReferenceInfo(Guid.Empty, $"VBProject{info}", $@"C:\Users\Rubberduck\Documents\Book{info}.xlsm", 0, 0))
                .ToList();

        public static List<ReferenceModel> DummyProjectsList => ProjectReferenceInfoList
            .Select(proj => new ReferenceModel(proj, ReferenceKind.Project)).ToList();

        public static List<ReferenceInfo> RecentProjectReferenceInfoList =>
            Enumerable.Range(1, 3)
                .Select(info =>
                    new ReferenceInfo(Guid.Empty, $"RecentProject{info}", $@"C:\Users\Rubberduck\Documents\RecentBook{info}.xlsm", 0, 0))
                .ToList();

        public static List<ReferenceInfo> RecentLibraryReferenceInfoList =>
            Enumerable.Range(1, 5)
                .Select(info =>
                    new ReferenceInfo(new Guid(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, (byte)info), $"Recent{info}", $@"C:\Windows\System32\recent{info}.dll", 1, 0))
                .ToList();

        public static List<ReferenceModel> DummyReferencesList => new List<ReferenceModel>
        {
            new ReferenceModel(new ReferenceInfo(VbaGuid, "VBA", @"C:\Shortcut\VBE7.DLL", 4, 2), ReferenceKind.TypeLibrary) {IsBuiltIn = true, IsReferenced = true, Priority = 1 },
            new ReferenceModel(new ReferenceInfo(ExcelGuid, "Excel", @"C:\Office\EXCEL.EXE", 15, 0), ReferenceKind.TypeLibrary) {IsBuiltIn = true, IsReferenced = true, Priority = 2},
            new ReferenceModel(new ReferenceInfo(DummyGuidOne, "ReferenceOne", @"C:\Libs\reference1.dll", 1, 1), ReferenceKind.TypeLibrary) {IsReferenced = true, Priority = 3 },
            new ReferenceModel(new ReferenceInfo(DummyGuidTwo, "ReferenceTwo", @"C:\Libs\reference2.dll", 2, 2), ReferenceKind.TypeLibrary) {IsReferenced = true, Priority = 4 }
        };

        public static ReferenceSettings GetDefaultReferenceSettings()
        {
            var defaults = new ReferenceSettings
            {
                RecentReferencesTracked = 20
            };
            defaults.PinReference(new ReferenceInfo(new Guid(RubberduckGuid.RubberduckTypeLibGuid), string.Empty, string.Empty, 2, 4));
            defaults.PinReference(new ReferenceInfo(new Guid(RubberduckGuid.RubberduckApiTypeLibGuid), string.Empty, string.Empty, 2, 4));
            defaults.ProjectPaths.Add(@"C:\Users\Rubberduck\Documents");

            return defaults;
        }

        public static ReferenceSettings GetNonDefaultReferenceSettings()
        {
            var settings = new ReferenceSettings
            {
                RecentReferencesTracked = 42,
                FixBrokenReferences = true,
                AddToRecentOnReferenceEvents = true,
                ProjectPaths = new List<string> { @"C:\Users\SomeOtherUser\Documents" }
            };

            settings.UpdatePinnedReferencesForHost(null, LibraryReferenceInfoList);
            settings.UpdatePinnedReferencesForHost("EXCEL.EXE", ProjectReferenceInfoList);
            settings.UpdateRecentReferencesForHost(null, RecentLibraryReferenceInfoList);
            settings.UpdateRecentReferencesForHost("EXCEL.EXE", RecentProjectReferenceInfoList);

            return settings;
        }

        public static IConfigurationService<ReferenceSettings> GetReferenceSettingsProvider(ReferenceSettings settings = null)
        {
            return GetMockReferenceSettingsProvider(settings).Object;
        }

        public static Mock<IConfigurationService<ReferenceSettings>> GetMockReferenceSettingsProvider(ReferenceSettings settings = null)
        {
            var output = new Mock<IConfigurationService<ReferenceSettings>>();

            output.Setup(m => m.Read()).Returns(() => settings ?? GetDefaultReferenceSettings());
            output.Setup(m => m.ReadDefaults()).Returns(GetDefaultReferenceSettings);
            output.Setup(m => m.Save(It.IsAny<ReferenceSettings>()));

            return output;
        }

        public static ReferenceReconciler ArrangeReferenceReconciler(ReferenceSettings settings = null)
        {
            return ArrangeReferenceReconciler(settings, null, out _, out _);
        }

        public static ReferenceReconciler ArrangeReferenceReconciler(IProjectsProvider projectsProvider, ReferenceSettings settings = null)
        {
            return ArrangeReferenceReconciler(settings, projectsProvider, out _, out _);
        }

        public static ReferenceReconciler ArrangeReferenceReconciler(
            ReferenceSettings settings,
            out Mock<IMessageBox> messageBox,
            out Mock<IComLibraryProvider> libraryProvider)
        {
            return ArrangeReferenceReconciler(settings, null, out messageBox, out libraryProvider);
        }

        public static ReferenceReconciler ArrangeReferenceReconciler(
            ReferenceSettings settings,
            IProjectsProvider projectsprovider,
            out Mock<IMessageBox> messageBox,
            out Mock<IComLibraryProvider> libraryProvider)
        {
            messageBox = new Mock<IMessageBox>();
            libraryProvider = new Mock<IComLibraryProvider>();
            return new ReferenceReconciler(messageBox.Object, GetReferenceSettingsProvider(settings), libraryProvider.Object, projectsprovider);
        }

        public static void SetupIComLibraryProvider(Mock<IComLibraryProvider> provider, ReferenceInfo reference, string path, string description = "")
        {
            var documentation = new Mock<IComDocumentation>();
            documentation.Setup(p => p.DocString).Returns(description);
            documentation.Setup(p => p.Name).Returns(reference.Name);
            documentation.Setup(p => p.HelpContext).Returns(0);
            documentation.Setup(p => p.HelpFile).Returns(string.Empty);

            provider.Setup(m => m.GetComDocumentation(It.IsAny<ITypeLib>())).Returns(documentation.Object);
            provider.Setup(m => m.GetReferenceInfo(It.IsAny<ITypeLib>(), reference.Name, path)).Returns(reference);
        }

        public static Mock<IReferences> GetReferencesMock(out Mock<IVBProject> project, out MockProjectBuilder builder)
        {
            builder = new MockProjectBuilder("TestBook", @"C:\TestBook.xlsm", ProjectProtection.Unprotected, ProjectType.HostProject, null, null);
            var references = builder
                .AddReference("VBA", @"C:\Shortcut\VBE7.DLL", 4, 2, true)
                .AddReference("Excel", @"C:\Office\EXCEL.EXE", 15, 0, true)
                .AddReference("ReferenceOne", @"C:\Libs\reference1.dll", 1, 1)
                .AddReference("ReferenceTwo", @"C:\Libs\reference2.dll", 2, 2)
                .GetMockedReferences(out project);

            return references;
        }

        public static Mock<IAddRemoveReferencesModel> ArrangeAddRemoveReferencesModel(List<ReferenceModel> input, List<ReferenceModel> output, ReferenceSettings settings = null)
        {
            var model = new Mock<IAddRemoveReferencesModel>();

            model.Setup(p => p.HostApplication).Returns("EXCEL.EXE");
            model.Setup(p => p.Settings).Returns(settings);
            model.Setup(p => p.References).Returns(input);
            model.Setup(p => p.NewReferences).Returns(output);

            return model;
        }

        public static ProjectDeclaration ArrangeMocksAndGetProject()
        {
            return ArrangeMocksAndGetProject(out _, out _, out _);
        }

        public static ProjectDeclaration ArrangeMocksAndGetProject(out MockProjectBuilder projectBuilder, out Mock<IReferences> references, out IProjectsProvider projectsProvider)
        {
            var builder = new MockVbeBuilder();

            projectBuilder = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("TestModule", ComponentType.StandardModule, string.Empty);

            references = projectBuilder
                .AddReference("VBA", @"C:\Shortcut\VBE7.DLL", 4, 2, true)
                .AddReference("Excel", @"C:\Office\EXCEL.EXE", 15, 0, true)
                .AddReference("Library One", @"C:\Libs\library1.dll", 1, 1)
                .AddReference("Library Two", @"C:\Libs\library2.dll", 2, 2)
                .GetMockedReferences(out _);

            builder.AddProject(projectBuilder.Build());

            var state = MockParser.CreateAndParse(builder.Build().Object);

            projectsProvider = state.ProjectsProvider;

            return state.DeclarationFinder
                .UserDeclarations(DeclarationType.Project)
                .OfType<ProjectDeclaration>()
                .Single();
        }

        public static Mock<IAddRemoveReferencesModel> ArrangeParsedAddRemoveReferencesModel(
            List<ReferenceModel> input,
            List<ReferenceModel> output, 
            List<ReferenceModel> registered, 
            out Mock<IReferences> references,
            out MockProjectBuilder projectBuilder,
            out IProjectsProvider projectsProvider)
        {
            var declaration = ArrangeMocksAndGetProject(out projectBuilder, out references, out projectsProvider);

            var model = ArrangeAddRemoveReferencesModel(input, output, GetDefaultReferenceSettings());
            model.Setup(m => m.Project).Returns(declaration);

            return model;
        }

        public static AddRemoveReferencesViewModel ArrangeViewModel()
        {
            return ArrangeViewModel(out _, out _, out _);
        }

        public static AddRemoveReferencesViewModel ArrangeViewModel(
            out List<ReferenceModel> allReferences,
            out List<ReferenceModel> projectReferences,
            out Mock<IFileSystemBrowserFactory> browserFactory,
            bool addHostProjects = false)
        {
            return ArrangeViewModel(out allReferences, out projectReferences, out browserFactory, out _, addHostProjects);
        }

        public static AddRemoveReferencesViewModel ArrangeViewModel(
            out Mock<IFileSystemBrowserFactory> browserFactory,
            out Mock<IComLibraryProvider> libraryProvider,
            bool addHostProjects = false)
        {
            return ArrangeViewModel(out _, out _, out browserFactory, out libraryProvider, addHostProjects);
        }

        public static AddRemoveReferencesViewModel ArrangeViewModel(
            out List<ReferenceModel> allReferences, 
            out List<ReferenceModel> projectReferences, 
            out Mock<IFileSystemBrowserFactory> browserFactory,
            out Mock<IComLibraryProvider> libraryProvider,
            bool addHostProjects = false)
        {
            var registered = LibraryReferenceInfoList.Select(reference => new ReferenceModel(reference, ReferenceKind.TypeLibrary)).ToList();

            var declaration = ArrangeMocksAndGetProject(out _, out var references, out var projectsProvider);
            var settings = GetNonDefaultReferenceSettings();

            var priority = 1;
            projectReferences = references.Object.Select(item => new ReferenceModel(item, priority++) { IsRegistered = true }).ToList();

            allReferences = registered.ToList();
            var pinnedLibrary = registered.First();
            pinnedLibrary.IsPinned = true;
            pinnedLibrary.IsRecent = true;

            allReferences.AddRange(projectReferences.Where(proj => !registered.Any(item =>
                    item.FullPath.Equals(proj.FullPath, StringComparison.OrdinalIgnoreCase))));

            if (addHostProjects) //RecentProjectReferenceInfoList
            {
                var projects = RecentProjectReferenceInfoList.Select(project => new ReferenceModel(project, ReferenceKind.Project)).ToList();
                var pinnedProject = projects.First();
                pinnedProject.IsPinned = true;
                pinnedProject.IsRecent = true;

                allReferences.AddRange(projects);
            }

            var model = new AddRemoveReferencesModel(null, declaration, allReferences, settings);
            var reconciler = ArrangeReferenceReconciler(settings, projectsProvider, out _, out libraryProvider);
            browserFactory = new Mock<IFileSystemBrowserFactory>();

            return new AddRemoveReferencesViewModel(model, reconciler, browserFactory.Object, projectsProvider);
        }

        public static void SetupMockedOpenDialog(this Mock<IFileSystemBrowserFactory> factory, string filename, DialogResult result)
        {
            factory.SetupMockedOpenDialog(filename, result, out _);
        }

        public static void SetupMockedOpenDialog(this Mock<IFileSystemBrowserFactory> factory, string filename, DialogResult result, out Mock<IOpenFileDialog> dialog)
        {
            dialog = new Mock<IOpenFileDialog>();
            dialog.Setup(m => m.FileName).Returns(filename);
            dialog.Setup(m => m.ShowDialog()).Returns(result);
            factory.Setup(m => m.CreateOpenFileDialog()).Returns(dialog.Object);
        }
    }
}
