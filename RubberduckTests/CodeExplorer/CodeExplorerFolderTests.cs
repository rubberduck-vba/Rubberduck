using NUnit.Framework;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.VBEditor.SafeComWrappers;
using System.Linq;

namespace RubberduckTests.CodeExplorer
{
    [TestFixture]
    public class CodeExplorerFolderTests
    {
        [Category("Code Explorer")]
        [Test]
        public void DefaultProjectFolderIsCreated()
        {
            const string inputCode =
@"Sub Foo()
Dim d As Boolean
d = True
End Sub";

            using (var explorer = new MockedCodeExplorer(ProjectType.HostProject, ComponentType.StandardModule, inputCode)
                .SelectFirstProject())
            {
                var project = explorer.ViewModel.SelectedItem;
                var folder = project.Children.OfType<CodeExplorerCustomFolderViewModel>().Single();

                Assert.NotNull(folder);
                Assert.AreEqual(project.Declaration.IdentifierName, folder.Name);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void SubFolderAnnotationCreatesSubFolders()
        {
            const string inputCode =
@"'@Folder(""First.Second.Third"")

Sub Foo()
Dim d As Boolean
d = True
End Sub";

            using (var explorer = new MockedCodeExplorer(ProjectType.HostProject, new[] { ComponentType.StandardModule }, new[] { inputCode })
                .SelectFirstCustomFolder())
            {
                var folder = (CodeExplorerCustomFolderViewModel)explorer.ViewModel.SelectedItem;
                var subfolder = folder.Children.OfType<CodeExplorerCustomFolderViewModel>().Single();
                var subsubfolder = subfolder.Children.OfType<CodeExplorerCustomFolderViewModel>().Single();

                Assert.AreEqual("First", folder.Name);
                Assert.AreEqual("Second", subfolder.Name);
                Assert.AreEqual("Third", subsubfolder.Name);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void SubFolderModuleIsChildOfDeepestSubFolder()
        {
            const string inputCode =
@"'@Folder(""First.Second.Third"")

Sub Foo()
Dim d As Boolean
d = True
End Sub";

            using (var explorer = new MockedCodeExplorer(ProjectType.HostProject, ComponentType.StandardModule, inputCode)
                .SelectFirstCustomFolder())
            {
                var folder = explorer.ViewModel.SelectedItem;
                while (folder.Children.OfType<CodeExplorerCustomFolderViewModel>().Any())
                {
                    folder = folder.Children.OfType<CodeExplorerCustomFolderViewModel>().Single();
                }

                var component = folder.Children.OfType<CodeExplorerComponentViewModel>().Single();

                Assert.AreEqual("TestModule", component.Name);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void SubFoldersForkOnCorrectDepth()
        {
            string[] folders = new[] 
            {
                "Foo.Bar.Baz",
                "Foo.Bar",
                "Bar.Bar",
                "Bar.Baz"
            };

            var modules = folders.Select(folder => $@"'@Folder(""{folder}"")").ToArray();
            var components = folders.Select(_ => ComponentType.StandardModule).ToArray();

            using (var explorer = new MockedCodeExplorer(ProjectType.HostProject, components, modules)
                .SelectFirstProject())
            {
                var project = explorer.ViewModel.SelectedItem;
                var custom = project.Children.OfType<CodeExplorerCustomFolderViewModel>().Where(folder => !folder.Name.Equals(project.Name)).ToList();

                var foo = custom.Last();
                var foobar = foo.Children.OfType<CodeExplorerCustomFolderViewModel>().Single();
                var foobarbaz = foobar.Children.OfType<CodeExplorerCustomFolderViewModel>().Single();

                var bar = custom.First();
                var barbar = bar.Children.OfType<CodeExplorerCustomFolderViewModel>().First();
                var barbaz = bar.Children.OfType<CodeExplorerCustomFolderViewModel>().Last();

                Assert.AreEqual(3, project.Children.OfType<CodeExplorerCustomFolderViewModel>().Count());
                Assert.AreEqual(1, foo.Children.OfType<CodeExplorerCustomFolderViewModel>().Count());
                Assert.AreEqual(1, foobar.Children.OfType<CodeExplorerCustomFolderViewModel>().Count());
                Assert.AreEqual(2, bar.Children.OfType<CodeExplorerCustomFolderViewModel>().Count());

                Assert.AreEqual("Foo", foo.Name);              
                Assert.AreEqual("Bar", foobar.Name);            
                Assert.AreEqual("Baz", foobarbaz.Name);
                Assert.AreEqual("Bar", bar.Name);
                Assert.AreEqual("Bar", barbar.Name);
                Assert.AreEqual("Baz", barbaz.Name);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void FoldersSortByName()
        {
            string[] folders = new[]
            {
                "AFolder",
                "BFolder",
                "CFolder",
            };

            var modules = folders.Select(folder => $@"'@Folder(""{folder}"")").ToArray();
            var components = folders.Select(_ => ComponentType.StandardModule).ToArray();

            using (var explorer = new MockedCodeExplorer(ProjectType.HostProject, components, modules)
                .SelectFirstProject())
            {
                var project = explorer.ViewModel.SelectedItem;
                var custom = project.Children.OfType<CodeExplorerCustomFolderViewModel>().Where(folder => !folder.Name.Equals(project.Name)).Select(folder => folder.Name).ToList();
                Assert.IsTrue(custom.OrderBy(_ => _).SequenceEqual(folders.OrderBy(_ => _)));
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void FoldersNamesAreCaseSensitive()
        {
            string[] folders = new[]
            {
                "foo",
                "Foo"
            };

            var modules = folders.Select(folder => $@"'@Folder(""{folder}"")").ToArray();
            var components = folders.Select(_ => ComponentType.StandardModule).ToArray();

            using (var explorer = new MockedCodeExplorer(ProjectType.HostProject, components, modules)
                .SelectFirstProject())
            {
                var project = explorer.ViewModel.SelectedItem;
                Assert.AreEqual(3, project.Children.Count);
            }
        }
    }
}
