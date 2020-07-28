using Moq;
using NUnit.Framework;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Annotations.Concrete;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using System.Collections.Generic;
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
        public void AddedModuleIsAtCorrectDepth_DefaultNode()
        {
            const string inputCode =
@"'@Folder(""First.Second.Third"")

Sub Foo()
Dim d As Boolean
d = True
End Sub";

            using (var explorer = new MockedCodeExplorer(ProjectType.HostProject, new[] { ComponentType.StandardModule }, new[] { inputCode })
                .SelectFirstProject())
            {
                var project = (CodeExplorerProjectViewModel)explorer.ViewModel.SelectedItem;
                var folder = project.Children.OfType<CodeExplorerCustomFolderViewModel>().First(node => node.Name.Equals(project.Declaration.IdentifierName));
                var declarations = project.State.AllUserDeclarations.ToList();
                declarations.Add(GetNewClassDeclaration(project.Declaration, "Foo"));

                project.Synchronize(ref declarations);
                var added = folder.Children.OfType<CodeExplorerComponentViewModel>().Single();

                Assert.AreEqual(DeclarationType.ClassModule, added.Declaration.DeclarationType);
                Assert.AreEqual(project.Declaration.IdentifierName, added.Declaration.CustomFolder);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void AddedModuleIsAtCorrectDepth_FirstAnnotation()
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
                var project = explorer.ViewModel.Projects.OfType<CodeExplorerProjectViewModel>().First();
                var folder = (CodeExplorerCustomFolderViewModel)explorer.ViewModel.SelectedItem;
                var declarations = project.State.AllUserDeclarations.ToList();

                var mockedAnnotation = new Mock<IParseTreeAnnotation>();
                mockedAnnotation.Setup(m => m.Annotation).Returns(new FolderAnnotation());
                mockedAnnotation.Setup(m => m.QualifiedSelection).Returns(new QualifiedSelection(project.Declaration.QualifiedModuleName, new Selection(1, 1)));
                // returns unquoted argument because the FolderAnnotation's argument processing is never invoked in the mock
                mockedAnnotation.Setup(m => m.AnnotationArguments).Returns(new[] { "First" }.ToList());
                var annotation = mockedAnnotation.Object;
                var predeclared = new ParseTreeAnnotation(new PredeclaredIdAnnotation(), new QualifiedSelection(project.Declaration.QualifiedModuleName, new Selection(2, 1)), (VBAParser.AnnotationContext)null);

                declarations.Add(GetNewClassDeclaration(project.Declaration, "Foo", new IParseTreeAnnotation[] { annotation, predeclared }));

                project.Synchronize(ref declarations);
                var added = folder.Children.OfType<CodeExplorerComponentViewModel>().Single();

                Assert.AreEqual(DeclarationType.ClassModule, added.Declaration.DeclarationType);
                Assert.AreEqual("First", added.Declaration.CustomFolder);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void AddedModuleIsAtCorrectDepth_NotFirstAnnotation()
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
                var project = explorer.ViewModel.Projects.OfType<CodeExplorerProjectViewModel>().First();
                var folder = (CodeExplorerCustomFolderViewModel)explorer.ViewModel.SelectedItem;
                var declarations = project.State.AllUserDeclarations.ToList();

                var mockedAnnotation = new Mock<IParseTreeAnnotation>();
                mockedAnnotation.Setup(m => m.Annotation).Returns(new FolderAnnotation());
                mockedAnnotation.Setup(m => m.QualifiedSelection).Returns(new QualifiedSelection(project.Declaration.QualifiedModuleName, new Selection(2, 1)));
                // returns unquoted argument because the FolderAnnotation's argument processing is never invoked in the mock
                mockedAnnotation.Setup(m => m.AnnotationArguments).Returns(new[] { "First" }.ToList());
                var annotation = mockedAnnotation.Object;
                var predeclared = new ParseTreeAnnotation(new PredeclaredIdAnnotation(), new QualifiedSelection(project.Declaration.QualifiedModuleName, new Selection(1, 1)), (VBAParser.AnnotationContext)null);

                declarations.Add(GetNewClassDeclaration(project.Declaration, "Foo", new IParseTreeAnnotation[] { predeclared, annotation }));

                project.Synchronize(ref declarations);
                var added = folder.Children.OfType<CodeExplorerComponentViewModel>().Single();

                Assert.AreEqual(DeclarationType.ClassModule, added.Declaration.DeclarationType);
                Assert.AreEqual("First", added.Declaration.CustomFolder);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void AddedModuleIsAtCorrectDepth_RootNode()
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
                var project = explorer.ViewModel.Projects.OfType<CodeExplorerProjectViewModel>().First();
                var folder = (CodeExplorerCustomFolderViewModel)explorer.ViewModel.SelectedItem;
                var declarations = project.State.AllUserDeclarations.ToList();
                declarations.Add(GetNewClassDeclaration(project.Declaration, "Foo", "\"First\""));

                project.Synchronize(ref declarations);
                var added = folder.Children.OfType<CodeExplorerComponentViewModel>().Single();

                Assert.AreEqual(DeclarationType.ClassModule, added.Declaration.DeclarationType);
                Assert.AreEqual("First", added.Declaration.CustomFolder);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void AddedModuleIsAtCorrectDepth_SubNode()
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
                var project = explorer.ViewModel.Projects.OfType<CodeExplorerProjectViewModel>().First();
                var folder = (CodeExplorerCustomFolderViewModel)explorer.ViewModel.SelectedItem;
                var subfolder = folder.Children.OfType<CodeExplorerCustomFolderViewModel>().Single();
                var declarations = project.State.AllUserDeclarations.ToList();
                declarations.Add(GetNewClassDeclaration(project.Declaration, "Foo", "\"First.Second\""));

                project.Synchronize(ref declarations);
                var added = subfolder.Children.OfType<CodeExplorerComponentViewModel>().Single();

                Assert.AreEqual(DeclarationType.ClassModule, added.Declaration.DeclarationType);
                Assert.AreEqual("First.Second", added.Declaration.CustomFolder);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void AddedModuleIsAtCorrectDepth_TerminalNode()
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
                var project = explorer.ViewModel.Projects.OfType<CodeExplorerProjectViewModel>().First();
                var folder = (CodeExplorerCustomFolderViewModel)explorer.ViewModel.SelectedItem;
                var subfolder = folder.Children.OfType<CodeExplorerCustomFolderViewModel>().Single()
                                      .Children.OfType<CodeExplorerCustomFolderViewModel>().Single();
                var declarations = project.State.AllUserDeclarations.ToList();
                declarations.Add(GetNewClassDeclaration(project.Declaration, "Foo", "\"First.Second.Third\""));

                project.Synchronize(ref declarations);

                var added = subfolder.Children.OfType<CodeExplorerComponentViewModel>()                   
                    .SingleOrDefault(node => node.Declaration.DeclarationType == DeclarationType.ClassModule);

                Assert.IsNotNull(added);
                Assert.AreEqual("First.Second.Third", added.Declaration.CustomFolder);
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

        private static Declaration GetNewClassDeclaration(Declaration project, string name, string folder = "")
        {
            var mockedFolderAnnotation = new Mock<IParseTreeAnnotation>();
            mockedFolderAnnotation.Setup(m => m.Annotation).Returns(new FolderAnnotation());
            mockedFolderAnnotation.Setup(m => m.AnnotationArguments).Returns(new[] { folder }.ToList());
            mockedFolderAnnotation.Setup(m => m.QualifiedSelection).Returns(new QualifiedSelection(project.QualifiedModuleName, new Selection(1, 1)));

            var annotations = string.IsNullOrEmpty(folder)
                ? Enumerable.Empty<IParseTreeAnnotation>()
                : new[] { mockedFolderAnnotation.Object };

            return GetNewClassDeclaration(project, name, annotations);
        }

        private static Declaration GetNewClassDeclaration(Declaration project, string name, IEnumerable<IParseTreeAnnotation> annotations)
        {
            var declaration =
                new ClassModuleDeclaration(new QualifiedMemberName(project.QualifiedModuleName, name), project, name, true, annotations, new Attributes());

            return declaration;
        }
    }
}
