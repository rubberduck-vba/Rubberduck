using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using NUnit.Framework;
using Rubberduck.JunkDrawer.Extensions;
using Rubberduck.Navigation.CodeExplorer;

namespace RubberduckTests.CodeExplorer
{
    [TestFixture]
    [SuppressMessage("ReSharper", "PossibleNullReferenceException")]
    public class CodeExplorerCustomFolderViewModelTests
    {
        [Test]
        [Category("Code Explorer")]
        [TestCase(new object[] { CodeExplorerTestSetup.TestModuleName, "Foo"}, TestName = "Constructor_SetsFolderName_RootFolder")]
        [TestCase(new object[] { CodeExplorerTestSetup.TestModuleName, "Foo.Bar" }, TestName = "Constructor_SetsFolderName_SubFolder")]
        [TestCase(new object[] { CodeExplorerTestSetup.TestModuleName, "Foo.Bar.Baz" }, TestName = "Constructor_SetsFolderName_SubSubFolder")]
        public void Constructor_SetsFolderName(object[] parameters)
        {
            var structure = ToFolderStructure(parameters.Cast<string>());
            var folderPath = structure.First().Folder;
            var path = folderPath.Split(FolderExtensions.FolderDelimiter);

            var declarations = CodeExplorerTestSetup.TestProjectWithFolderStructure(structure, out _, out var state);
            using (state)
            {
                var folder = new CodeExplorerCustomFolderViewModel(null, path.First(), path.First(), null, ref declarations);

                foreach (var name in path)
                {
                    Assert.AreEqual(name, folder.Name);
                    folder = folder.Children.OfType<CodeExplorerCustomFolderViewModel>().FirstOrDefault();
                }
            }
        }

        [Test]
        [Category("Code Explorer")]
        [TestCase(new object[] { CodeExplorerTestSetup.TestModuleName, "Foo" }, TestName = "Constructor_SetsFullPath_RootFolder")]
        [TestCase(new object[] { CodeExplorerTestSetup.TestModuleName, "Foo.Bar" }, TestName = "Constructor_SetsFullPath_SubFolder")]
        [TestCase(new object[] { CodeExplorerTestSetup.TestModuleName, "Foo.Bar.Baz" }, TestName = "Constructor_SetsFullPath_SubSubFolder")]
        public void Constructor_SetsFullPath(object[] parameters)
        {
            var structure = ToFolderStructure(parameters.Cast<string>());
            var folderPath = structure.First().Folder;
            var path = folderPath.Split(FolderExtensions.FolderDelimiter);

            var declarations = CodeExplorerTestSetup.TestProjectWithFolderStructure(structure, out _, out var state);
            using (state)
            {
                var folder = new CodeExplorerCustomFolderViewModel(null, path.First(), path.First(), null, ref declarations);

                var depth = 1;
                foreach (var _ in path)
                {
                    Assert.AreEqual(string.Join(FolderExtensions.FolderDelimiter.ToString(), path.Take(depth++)), folder.FullPath);
                    folder = folder.Children.OfType<CodeExplorerCustomFolderViewModel>().FirstOrDefault();
                }
            }
        }

        [Test]
        [Category("Code Explorer")]
        [TestCase(new object[] { CodeExplorerTestSetup.TestModuleName, "Foo" }, TestName = "Constructor_PanelTitleIsFullPath_RootFolder")]
        [TestCase(new object[] { CodeExplorerTestSetup.TestModuleName, "Foo.Bar" }, TestName = "Constructor_PanelTitleIsFullPath_SubFolder")]
        [TestCase(new object[] { CodeExplorerTestSetup.TestModuleName, "Foo.Bar.Baz" }, TestName = "Constructor_PanelTitleIsFullPath_SubSubFolder")]
        public void Constructor_PanelTitleIsFullPath(object[] parameters)
        {
            var structure = ToFolderStructure(parameters.Cast<string>());
            var folderPath = structure.First().Folder;
            var path = folderPath.Split(FolderExtensions.FolderDelimiter);

            var declarations = CodeExplorerTestSetup.TestProjectWithFolderStructure(structure, out _, out var state);
            using (state)
            {
                var folder = new CodeExplorerCustomFolderViewModel(null, path.First(), path.First(), null, ref declarations);

                foreach (var _ in path)
                {
                    Assert.AreEqual(folder.FullPath, folder.PanelTitle);
                    folder = folder.Children.OfType<CodeExplorerCustomFolderViewModel>().FirstOrDefault();
                }
            }
        }

        [Test]
        [Category("Code Explorer")]
        [TestCase(new object[] { CodeExplorerTestSetup.TestModuleName, "Foo" }, TestName = "Constructor_FolderAttributeIsCorrect_RootFolder")]
        [TestCase(new object[] { CodeExplorerTestSetup.TestModuleName, "Foo.Bar" }, TestName = "Constructor_FolderAttributeIsCorrect_SubFolder")]
        [TestCase(new object[] { CodeExplorerTestSetup.TestModuleName, "Foo.Bar.Baz" }, TestName = "Constructor_FolderAttributeIsCorrect_SubSubFolder")]
        public void Constructor_FolderAttributeIsCorrect(object[] parameters)
        {
            var structure = ToFolderStructure(parameters.Cast<string>());
            var folderPath = structure.First().Folder;
            var path = folderPath.Split(FolderExtensions.FolderDelimiter);

            var declarations = CodeExplorerTestSetup.TestProjectWithFolderStructure(structure, out _, out var state);
            using (state)
            {
                var folder =
                    new CodeExplorerCustomFolderViewModel(null, path.First(), path.First(), null, ref declarations);

                foreach (var _ in path)
                {
                    Assert.AreEqual($"'@Folder(\"{folder.FullPath}\")", folder.FolderAttribute);
                    folder = folder.Children.OfType<CodeExplorerCustomFolderViewModel>().FirstOrDefault();
                }
            }
        }

        [Test]
        [Category("Code Explorer")]
        [TestCase(new object[] { CodeExplorerTestSetup.TestModuleName, "Foo" }, TestName = "Constructor_DescriptionIsFolderAttribute_RootFolder")]
        [TestCase(new object[] { CodeExplorerTestSetup.TestModuleName, "Foo.Bar" }, TestName = "Constructor_DescriptionIsFolderAttribute_SubFolder")]
        [TestCase(new object[] { CodeExplorerTestSetup.TestModuleName, "Foo.Bar.Baz" }, TestName = "Constructor_DescriptionIsFolderAttribute_SubSubFolder")]
        public void Constructor_DescriptionIsFolderAttribute(object[] parameters)
        {
            var structure = ToFolderStructure(parameters.Cast<string>());
            var folderPath = structure.First().Folder;
            var path = folderPath.Split(FolderExtensions.FolderDelimiter);

            var declarations = CodeExplorerTestSetup.TestProjectWithFolderStructure(structure, out _, out var state);
            using (state)
            {
                var folder =
                    new CodeExplorerCustomFolderViewModel(null, path.First(), path.First(), null, ref declarations);

                foreach (var _ in path)
                {
                    Assert.AreEqual(folder.FolderAttribute, folder.Description);
                    folder = folder.Children.OfType<CodeExplorerCustomFolderViewModel>().FirstOrDefault();
                }
            }
        }

        [Test]
        [Category("Code Explorer")]
        [TestCase(new object[] { CodeExplorerTestSetup.TestModuleName, "Foo" }, TestName = "Constructor_SetsFolderDepth_RootFolder")]
        [TestCase(new object[] { CodeExplorerTestSetup.TestModuleName, "Foo.Bar" }, TestName = "Constructor_SetsFolderDepth_SubFolder")]
        [TestCase(new object[] { CodeExplorerTestSetup.TestModuleName, "Foo.Bar.Baz" }, TestName = "Constructor_SetsFolderDepth_SubSubFolder")]
        public void Constructor_SetsFolderDepth(object[] parameters)
        {
            var structure = ToFolderStructure(parameters.Cast<string>());
            var folderPath = structure.First().Folder;
            var path = folderPath.Split(FolderExtensions.FolderDelimiter);

            var declarations = CodeExplorerTestSetup.TestProjectWithFolderStructure(structure, out _, out var state);
            using (state)
            {
                var folder =
                    new CodeExplorerCustomFolderViewModel(null, path.First(), path.First(), null, ref declarations);

                var depth = 1;
                foreach (var _ in path)
                {
                    Assert.AreEqual(depth++, folder.FolderDepth);
                    folder = folder.Children.OfType<CodeExplorerCustomFolderViewModel>().FirstOrDefault();
                }
            }
        }

        [Test]
        [Category("Code Explorer")]
        public void FilteredIsTrueForCharactersNotInName()
        {
            const string testCharacters = "abcdefghijklmnopqrstuwxyz";
            const string folderName = "Asdf";

            var testFolder = (Name: CodeExplorerTestSetup.TestModuleName, Folder: folderName);
            var declarations = CodeExplorerTestSetup.TestProjectWithFolderStructure(new List<(string Name, string Folder)> { testFolder }, out _, out var state);
            using (state)
            {
                var children = declarations.SelectMany(declaration => declaration.IdentifierName.ToCharArray())
                    .Distinct().ToList();

                var folder =
                    new CodeExplorerCustomFolderViewModel(null, folderName, folderName, null, ref declarations);

                var nonMatching = testCharacters.ToCharArray()
                    .Except(folderName.ToLowerInvariant().ToCharArray().Union(children));

                foreach (var character in nonMatching.Select(letter => letter.ToString()))
                {
                    folder.Filter = character;
                    Assert.IsTrue(folder.Filtered);
                }
            }
        }

        [Test]
        [Category("Code Explorer")]
        public void FilteredIsFalseForSubsetsOfName()
        {
            const string folderName = "Foobar";

            var testFolder = (Name: CodeExplorerTestSetup.TestModuleName, Folder: folderName);
            var declarations = CodeExplorerTestSetup.TestProjectWithFolderStructure(new List<(string Name, string Folder)> { testFolder }, out _, out var state);
            using (state)
            {

                var folder =
                    new CodeExplorerCustomFolderViewModel(null, folderName, folderName, null, ref declarations);

                for (var characters = 1; characters <= folderName.Length; characters++)
                {
                    folder.Filter = folderName.Substring(0, characters);
                    Assert.IsFalse(folder.Filtered);
                }

                for (var position = folderName.Length - 2; position > 0; position--)
                {
                    folder.Filter = folderName.Substring(position);
                    Assert.IsFalse(folder.Filtered);
                }
            }
        }

        [Test]
        [Category("Code Explorer")]
        public void FilteredIsFalseIfChildMatches()
        {
            const string folderName = "Foobar";

            var testFolder = (Name: CodeExplorerTestSetup.TestModuleName, Folder: folderName);
            var declarations = CodeExplorerTestSetup.TestProjectWithFolderStructure(new List<(string Name, string Folder)> { testFolder }, out _, out var state);
            using (state)
            {

                var folder =
                    new CodeExplorerCustomFolderViewModel(null, folderName, folderName, null, ref declarations);
                var childName = folder.Children.First().Name;

                for (var characters = 1; characters <= childName.Length; characters++)
                {
                    folder.Filter = childName.Substring(0, characters);
                    Assert.IsFalse(folder.Filtered);
                }

                for (var position = childName.Length - 2; position > 0; position--)
                {
                    folder.Filter = childName.Substring(position);
                    Assert.IsFalse(folder.Filtered);
                }
            }
        }

        [Test]
        [Category("Code Explorer")]
        public void UnfilteredStateIsRestored()
        {
            const string folderName = "Foobar";

            var testFolder = (Name: CodeExplorerTestSetup.TestModuleName, Folder: folderName);
            var declarations = CodeExplorerTestSetup.TestProjectWithFolderStructure(new List<(string Name, string Folder)> { testFolder }, out _, out var state);
            using (state)
            {

                var folder =
                    new CodeExplorerCustomFolderViewModel(null, folderName, folderName, null, ref declarations);
                var childName = folder.Children.First().Name;

                folder.IsExpanded = false;
                folder.Filter = childName;
                Assert.IsTrue(folder.IsExpanded);

                folder.Filter = string.Empty;
                Assert.IsFalse(folder.IsExpanded);
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
        public void SortComparerIsName(CodeExplorerSortOrder order)
        {
            const string folderName = "Foo";

            var testFolder = (Name: CodeExplorerTestSetup.TestModuleName, Folder: folderName);
            var declarations = CodeExplorerTestSetup.TestProjectWithFolderStructure(new List<(string Name, string Folder)> { testFolder }, out _, out var state);
            using (state)
            {

                var folder = new CodeExplorerCustomFolderViewModel(null, folderName, folderName, null, ref declarations)
                {
                    SortOrder = order
                };

                Assert.AreEqual(CodeExplorerItemComparer.Name.GetType(), folder.SortComparer.GetType());
            }
        }

        [Test]
        [Category("Code Explorer")]
        public void ErrorStateCanNotBeSet()
        {
            const string folderName = "Foo";

            var testFolder = (Name: CodeExplorerTestSetup.TestModuleName, Folder: folderName);
            var declarations = CodeExplorerTestSetup.TestProjectWithFolderStructure(new List<(string Name, string Folder)> { testFolder }, out _, out var state);
            using (state)
            {

                var folder = new CodeExplorerCustomFolderViewModel(null, folderName, folderName, null, ref declarations)
                {
                    IsErrorState = true
                };

                Assert.IsFalse(folder.IsErrorState);
            }
        }

        [Test]
        [Category("Code Explorer")]
        [TestCase(CodeExplorerTestSetup.TestModuleName, "Foo.Modules", 
                  CodeExplorerTestSetup.TestClassName, "Foo.Classes", 
                  CodeExplorerTestSetup.TestDocumentName, "Foo.Docs", 
                  CodeExplorerTestSetup.TestUserFormName, "Foo.Forms")]
        [TestCase(CodeExplorerTestSetup.TestModuleName, "Foo",
                  CodeExplorerTestSetup.TestClassName, "Foo.Bar",
                  CodeExplorerTestSetup.TestDocumentName, "Foo.Bar",
                  CodeExplorerTestSetup.TestUserFormName, "Foo.Bar")]
        [TestCase(CodeExplorerTestSetup.TestModuleName, "Foo",
                  CodeExplorerTestSetup.TestClassName, "Foo.Bar",
                  CodeExplorerTestSetup.TestDocumentName, "Foo.Bar.Baz",
                  CodeExplorerTestSetup.TestUserFormName, "Foo.Bar")]
        [TestCase(CodeExplorerTestSetup.TestModuleName, "Foo.Bar.Baz",
                  CodeExplorerTestSetup.TestClassName, "Foo.Bar.Baz",
                  CodeExplorerTestSetup.TestDocumentName, "Foo.Bar.Baz",
                  CodeExplorerTestSetup.TestUserFormName, "Foo.Bar.Baz")]
        [TestCase(CodeExplorerTestSetup.TestModuleName, "Foo.Baz",
                  CodeExplorerTestSetup.TestClassName, "Foo.Baz",
                  CodeExplorerTestSetup.TestDocumentName, "Foo.Bar",
                  CodeExplorerTestSetup.TestUserFormName, "Foo.Bar")]
        [TestCase(CodeExplorerTestSetup.TestModuleName, "Foo",
                  CodeExplorerTestSetup.TestClassName, "Foo.Bar.Baz",
                  CodeExplorerTestSetup.TestDocumentName, "Foo.Bar.Baz",
                  CodeExplorerTestSetup.TestUserFormName, "Foo.Foo.Foo")]
        [TestCase(CodeExplorerTestSetup.TestModuleName, "Foo Bar",
                  CodeExplorerTestSetup.TestClassName, "Foo Bar.Baz Baz",
                  CodeExplorerTestSetup.TestDocumentName, "Foo Bar.Baz",
                  CodeExplorerTestSetup.TestUserFormName, "Foo Bar.Foo Foo")]
        public void Constructor_CreatesCorrectSubFolderStructure(params object[] parameters)
        {
            var structure = ToFolderStructure(parameters.Cast<string>());
            var root = structure.First().Folder;
            var path = root.Split(FolderExtensions.FolderDelimiter);

            var declarations = CodeExplorerTestSetup.TestProjectWithFolderStructure(structure, out var projectDeclaration, out var state);
            using (state)
            {
                var contents =
                    CodeExplorerProjectViewModel.ExtractTrackedDeclarationsForProject(projectDeclaration,
                        ref declarations);

                var folder =
                    new CodeExplorerCustomFolderViewModel(null, path.First(), path.First(), null, ref contents);

                AssertFolderStructureIsCorrect(folder, structure);
            }
        }

        [Test]
        [Category("Code Explorer")]
        [TestCase(CodeExplorerTestSetup.TestModuleName, "Foo.Modules",
                  CodeExplorerTestSetup.TestClassName, "Foo.Classes",
                  CodeExplorerTestSetup.TestDocumentName, "Foo.Docs",
                  CodeExplorerTestSetup.TestUserFormName, "Foo.Forms")]
        [TestCase(CodeExplorerTestSetup.TestModuleName, "Foo",
                  CodeExplorerTestSetup.TestDocumentName, "Foo.Bar",
                  CodeExplorerTestSetup.TestUserFormName, "Foo.Bar",
                  CodeExplorerTestSetup.TestClassName, "Foo.Bar")]
        [TestCase(CodeExplorerTestSetup.TestClassName, "Foo.Bar",
                  CodeExplorerTestSetup.TestDocumentName, "Foo.Bar.Baz",
                  CodeExplorerTestSetup.TestUserFormName, "Foo.Bar",
                  CodeExplorerTestSetup.TestModuleName, "Foo")]
        [TestCase(CodeExplorerTestSetup.TestModuleName, "Foo.Bar.Baz",
                  CodeExplorerTestSetup.TestClassName, "Foo.Bar.Baz",
                  CodeExplorerTestSetup.TestUserFormName, "Foo.Bar.Baz",
                  CodeExplorerTestSetup.TestDocumentName, "Foo.Bar.Baz")]
        [TestCase(CodeExplorerTestSetup.TestModuleName, "Foo.Baz",
                  CodeExplorerTestSetup.TestDocumentName, "Foo.Bar",
                  CodeExplorerTestSetup.TestUserFormName, "Foo.Bar",
                  CodeExplorerTestSetup.TestClassName, "Foo.Baz")]
        [TestCase(CodeExplorerTestSetup.TestModuleName, "Foo",
                  CodeExplorerTestSetup.TestClassName, "Foo.Bar.Baz",
                  CodeExplorerTestSetup.TestUserFormName, "Foo.Foo.Foo",
                  CodeExplorerTestSetup.TestDocumentName, "Foo.Bar.Baz")]
        [TestCase(CodeExplorerTestSetup.TestClassName, "Foo Bar.Baz Baz",
                  CodeExplorerTestSetup.TestDocumentName, "Foo Bar.Baz",
                  CodeExplorerTestSetup.TestUserFormName, "Foo Bar.Foo Foo",
                  CodeExplorerTestSetup.TestModuleName, "Foo Bar")]
        public void Synchronize_AddedComponent_HasCorrectSubFolderStructure(params object[] parameters)
        {
            var structure = ToFolderStructure(parameters.Cast<string>());
            var root = structure.First().Folder;
            var path = root.Split(FolderExtensions.FolderDelimiter);

            var declarations = CodeExplorerTestSetup.TestProjectWithFolderStructure(structure, out var projectDeclaration, out var state);
            using (state)
            {
                var synchronizing =
                    CodeExplorerProjectViewModel.ExtractTrackedDeclarationsForProject(projectDeclaration,
                        ref declarations);
                var component = synchronizing.TestComponentDeclarations(structure.Last().Name);
                var contents = synchronizing.Except(component).ToList();

                var project = new CodeExplorerProjectViewModel(projectDeclaration, ref contents, state, null, state.ProjectsProvider);
                var folder = project.Children.OfType<CodeExplorerCustomFolderViewModel>()
                    .Single(item => item.Name.Equals(path.First()));

                project.Synchronize(ref synchronizing);

                AssertFolderStructureIsCorrect(folder, structure);
            }
        }

        [Test]
        [Category("Code Explorer")]
        [TestCase(CodeExplorerTestSetup.TestModuleName, "Foo.Modules",
                  CodeExplorerTestSetup.TestClassName, "Foo.Classes",
                  CodeExplorerTestSetup.TestDocumentName, "Foo.Docs",
                  CodeExplorerTestSetup.TestUserFormName, "Foo.Forms")]
        [TestCase(CodeExplorerTestSetup.TestModuleName, "Foo",
                  CodeExplorerTestSetup.TestClassName, "Foo.Bar",
                  CodeExplorerTestSetup.TestDocumentName, "Foo.Bar",
                  CodeExplorerTestSetup.TestUserFormName, "Foo.Bar")]
        [TestCase(CodeExplorerTestSetup.TestModuleName, "Foo",
                  CodeExplorerTestSetup.TestClassName, "Foo.Bar",
                  CodeExplorerTestSetup.TestDocumentName, "Foo.Bar.Baz",
                  CodeExplorerTestSetup.TestUserFormName, "Foo.Bar")]
        [TestCase(CodeExplorerTestSetup.TestModuleName, "Foo.Bar.Baz",
                  CodeExplorerTestSetup.TestClassName, "Foo.Bar.Baz",
                  CodeExplorerTestSetup.TestDocumentName, "Foo.Bar.Baz",
                  CodeExplorerTestSetup.TestUserFormName, "Foo.Bar.Baz")]
        [TestCase(CodeExplorerTestSetup.TestModuleName, "Foo.Baz",
                  CodeExplorerTestSetup.TestClassName, "Foo.Baz",
                  CodeExplorerTestSetup.TestDocumentName, "Foo.Bar",
                  CodeExplorerTestSetup.TestUserFormName, "Foo.Bar")]
        [TestCase(CodeExplorerTestSetup.TestModuleName, "Foo",
                  CodeExplorerTestSetup.TestClassName, "Foo.Bar.Baz",
                  CodeExplorerTestSetup.TestDocumentName, "Foo.Bar.Baz",
                  CodeExplorerTestSetup.TestUserFormName, "Foo.Foo.Foo")]
        [TestCase(CodeExplorerTestSetup.TestModuleName, "Foo Bar",
                  CodeExplorerTestSetup.TestClassName, "Foo Bar.Baz Baz",
                  CodeExplorerTestSetup.TestDocumentName, "Foo Bar.Baz",
                  CodeExplorerTestSetup.TestUserFormName, "Foo Bar.Foo Foo")]
        public void Synchronize_RemovedComponent_HasCorrectSubFolderStructure(params object[] parameters)
        {
            var structure = ToFolderStructure(parameters.Cast<string>());
            var root = structure.First().Folder;
            var path = root.Split(FolderExtensions.FolderDelimiter);

            var declarations = CodeExplorerTestSetup.TestProjectWithFolderStructure(structure, out var projectDeclaration, out var state);
            using (state)
            {
                var contents =
                    CodeExplorerProjectViewModel.ExtractTrackedDeclarationsForProject(projectDeclaration,
                        ref declarations);
                var component = contents.TestComponentDeclarations(structure.Last().Name);
                var synchronizing = contents.Except(component).ToList();

                var project = new CodeExplorerProjectViewModel(projectDeclaration, ref contents, state, null, state.ProjectsProvider);
                var folder = project.Children.OfType<CodeExplorerCustomFolderViewModel>()
                    .Single(item => item.Name.Equals(path.First()));

                project.Synchronize(ref synchronizing);

                AssertFolderStructureIsCorrect(folder, structure.Take(structure.Count - 1).ToList());
            }
        }

        private static void AssertFolderStructureIsCorrect(CodeExplorerCustomFolderViewModel underTest, List<(string Name, string Folder)> structure)
        {
            foreach (var (name, fullPath) in structure)
            {
                var folder = underTest;
                var path = fullPath.Split(FolderExtensions.FolderDelimiter);
                var depth = path.Length;

                for (var sub = 1; sub < depth; sub++)
                {
                    folder = folder.Children.OfType<CodeExplorerCustomFolderViewModel>()
                        .SingleOrDefault(subFolder =>
                            subFolder.FullPath.Equals(string.Join(FolderExtensions.FolderDelimiter.ToString(),
                                path.Take(folder.FolderDepth + 1))));
                }

                Assert.IsNotNull(folder, $"Folder {fullPath} was not found.");

                var components = folder.Children.OfType<CodeExplorerComponentViewModel>().ToList();
                var component = components.SingleOrDefault(subFolder => subFolder.Name.Equals(name));

                Assert.IsNotNull(component, $"Component {name} was not found in folder {fullPath}.");

                var expected = structure.Where(item => item.Folder.Equals(fullPath)).Select(item => item.Name).OrderBy(_ => _);
                var actual = components.Select(item => item.Declaration.IdentifierName).OrderBy(_ => _);

                Assert.IsTrue(expected.SequenceEqual(actual), $"Folder {fullPath} does not contain expected components.");
            }
        }

        private static List<(string Name, string Folder)> ToFolderStructure(IEnumerable<string> structure)
        {
            var input = structure.ToArray();
            var output = new List<(string Name, string Folder)>();

            for (var module = 0; module < input.Length; module += 2)
            {
                output.Add((Name: input[module], Folder: input[module + 1]));
            }

            return output;
        }
    }
}
