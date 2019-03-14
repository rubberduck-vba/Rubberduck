using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.Parsing.Symbols;

namespace RubberduckTests.CodeExplorer
{
    // TODO: These tests should probably be refactored to use test cases.
    [TestFixture]
    public class CodeExplorerComparerTests
    {
        [Category("Code Explorer")]
        [Test]
        public void CompareByName_ReturnsZeroForIdenticalNodes()
        {
            var declarations = new List<Declaration>();
            var folderNode = new CodeExplorerCustomFolderViewModel(null, "Name", "Name", null, ref declarations);
            Assert.AreEqual(0, new CompareByName().Compare(folderNode, folderNode));
        }

        [Category("Code Explorer")]
        [Test]
        public void CompareByName_ReturnsZeroForIdenticalNames()
        {
            var declarations = new List<Declaration>();
            // this won't happen, but just to be thorough...--besides, it is good for the coverage
            var folderNode1 = new CodeExplorerCustomFolderViewModel(null, "Name", "Name", null, ref declarations);
            var folderNode2 = new CodeExplorerCustomFolderViewModel(null, "Name", "Name", null, ref declarations);

            Assert.AreEqual(0, new CompareByName().Compare(folderNode1, folderNode2));
        }

        [Category("Code Explorer")]
        [Test]
        public void CompareByName_ReturnsCorrectOrdering()
        {
            var declarations = new List<Declaration>();
            var folderNode1 = new CodeExplorerCustomFolderViewModel(null, "Name1", "Name1", null, ref declarations);
            var folderNode2 = new CodeExplorerCustomFolderViewModel(null, "Name2", "Name2", null, ref declarations);

            Assert.IsTrue(new CompareByName().Compare(folderNode1, folderNode2) < 0);
        }

        [Category("Code Explorer")]
        [Test]
        public void CompareByType_ReturnsZeroForIdenticalNodes()
        {
            var declarations = new List<Declaration>();
            var errorNode = new CodeExplorerCustomFolderViewModel(null, "Name", "folder1.folder2", null, ref declarations);
            Assert.AreEqual(0, new CompareByName().Compare(errorNode, errorNode));
        }

        [Category("Code Explorer")]
        [Test]
        public void CompareByType_ReturnsEventAboveConst()
        {
            const string inputCode =
@"Public Event Foo(ByVal arg1 As Integer, ByVal arg2 As String)
Public Const Bar = 0";

            using (var explorer = new MockedCodeExplorer(inputCode)
                .SelectFirstModule())
            {
                var module = explorer.ViewModel.SelectedItem;
                var eventNode = module.Children.Single(s => s.Name == "Foo");
                var constNode = module.Children.Single(s => s.Name == "Bar");

                Assert.AreEqual(-1, new CompareByDeclarationType().Compare(eventNode, constNode));
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void CompareByType_ReturnsConstAboveField()
        {
            const string inputCode =
@"Public Const Foo = 0
Public Bar As Boolean";

            using (var explorer = new MockedCodeExplorer(inputCode)
                .SelectFirstModule())
            {
                var module = explorer.ViewModel.SelectedItem;
                var constNode = module.Children.Single(s => s.Name == "Foo");
                var fieldNode = module.Children.Single(s => s.Name == "Bar");

                Assert.AreEqual(-1, new CompareByDeclarationType().Compare(constNode, fieldNode));
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void CompareByType_ReturnsFieldAbovePropertyGet()
        {
            const string inputCode =
@"Private Bar As Boolean

Public Property Get Foo() As Variant
End Property
";

            using (var explorer = new MockedCodeExplorer(inputCode)
                .SelectFirstModule())
            {
                var module = explorer.ViewModel.SelectedItem;
                var fieldNode = module.Children.Single(s => s.Name == "Bar");
                var propertyGetNode = module.Children.Single(s => s.Name == "Foo (Get)");

                Assert.AreEqual(-1, new CompareByDeclarationType().Compare(fieldNode, propertyGetNode));
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void CompareByType_ReturnsPropertyGetEqualToPropertyLet()
        {
            const string inputCode =
@"Public Property Get Foo() As Variant
End Property

Public Property Let Foo(ByVal Value As Variant)
End Property
";

            using (var explorer = new MockedCodeExplorer(inputCode)
                .SelectFirstModule())
            {
                var module = explorer.ViewModel.SelectedItem;
                var propertyGetNode = module.Children.Single(s => s.Name == "Foo (Get)");
                var propertyLetNode = module.Children.Single(s => s.Name == "Foo (Let)");

                Assert.AreEqual(0, new CompareByDeclarationType().Compare(propertyGetNode, propertyLetNode));
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void CompareByType_ReturnsPropertyGetEqualToPropertySet()
        {
            const string inputCode =
@"Public Property Get Foo() As Variant
End Property

Public Property Set Foo(ByVal Value As Variant)
End Property
";

            using (var explorer = new MockedCodeExplorer(inputCode)
                .SelectFirstModule())
            {
                var module = explorer.ViewModel.SelectedItem;
                var propertyGetNode = module.Children.Single(s => s.Name == "Foo (Get)");
                var propertyLetNode = module.Children.Single(s => s.Name == "Foo (Set)");

                Assert.AreEqual(0, new CompareByDeclarationType().Compare(propertyGetNode, propertyLetNode));
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void CompareByType_ReturnsPropertyLetEqualToPropertyGet()
        {
            const string inputCode =
@"Public Property Let Foo(ByVal Value As Variant)
End Property

Public Property Get Foo() As Variant
End Property
";

            using (var explorer = new MockedCodeExplorer(inputCode)
                .SelectFirstModule())
            {
                var module = explorer.ViewModel.SelectedItem;
                var propertyLetNode = module.Children.Single(s => s.Name == "Foo (Let)");
                var propertySetNode = module.Children.Single(s => s.Name == "Foo (Get)");

                Assert.AreEqual(0, new CompareByDeclarationType().Compare(propertyLetNode, propertySetNode));
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void CompareByType_ReturnsPropertyLetEqualToPropertySet()
        {
            const string inputCode =
@"Public Property Let Foo(ByVal Value As Variant)
End Property

Public Property Set Foo(ByVal Value As Variant)
End Property
";

            using (var explorer = new MockedCodeExplorer(inputCode)
                .SelectFirstModule())
            {
                var module = explorer.ViewModel.SelectedItem;
                var propertyLetNode = module.Children.Single(s => s.Name == "Foo (Let)");
                var propertySetNode = module.Children.Single(s => s.Name == "Foo (Set)");

                Assert.AreEqual(0, new CompareByDeclarationType().Compare(propertyLetNode, propertySetNode));
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void CompareByType_ReturnsPropertySetAboveFunction()
        {
            const string inputCode =
@"Public Property Set Foo(ByVal Value As Variant)
End Property

Public Function Bar() As Boolean
End Function
";

            using (var explorer = new MockedCodeExplorer(inputCode)
                .SelectFirstModule())
            {
                var module = explorer.ViewModel.SelectedItem;
                var propertySetNode = module.Children.Single(s => s.Name == "Foo (Set)");
                var functionNode = module.Children.Single(s => s.Name == "Bar");

                Assert.AreEqual(-1, new CompareByDeclarationType().Compare(propertySetNode, functionNode));
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void CompareByType_ReturnsSubsAndFunctionsEqual()
        {
            const string inputCode =
@"Public Function Foo() As Boolean
End Function

Public Sub Bar()
End Sub
";

            using (var explorer = new MockedCodeExplorer(inputCode)
                .SelectFirstModule())
            {
                var module = explorer.ViewModel.SelectedItem;
                var functionNode = module.Children.Single(s => s.Name == "Foo");
                var subNode = module.Children.Single(s => s.Name == "Bar");

                Assert.AreEqual(0, new CompareByDeclarationType().Compare(functionNode, subNode));
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void CompareByAccessibility_ReturnsPublicMethodsAbovePrivateMethods()
        {
            const string inputCode =
 @"Private Sub Foo()
End Sub

Public Sub Bar()
End Sub
";

            using (var explorer = new MockedCodeExplorer(inputCode)
               .SelectFirstModule())
            {
                var module = explorer.ViewModel.SelectedItem;
                var privateNode = module.Children.Single(s => s.Name == "Foo");
                var publicNode = module.Children.Single(s => s.Name == "Bar");

                Assert.AreEqual(-1, new CompareByAccessibility().Compare(publicNode, privateNode));
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void CompareByType_ReturnsClassModuleBelowDocument()
        {

            using (var explorer = new MockedCodeExplorer(ProjectType.HostProject, new[] { ComponentType.ClassModule, ComponentType.Document })
                .SelectFirstCustomFolder())
            {
                var folder = explorer.ViewModel.SelectedItem;

                var clsNode = folder.Children.Single(s => s.Name == "ClassModule0");
                var docNode = folder.Children.Single(s => s.Name == "Document1");

                // this tests the logic I wrote to place docs above cls modules even though the parser calls them both cls modules
                Assert.AreEqual(-1, new CompareByDeclarationType().Compare(docNode, clsNode));
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void CompareBySelection_ReturnsZeroForIdenticalNodes()
        {
            const string inputCode =
@"Sub Foo()
End Sub

Sub Bar()
    Foo
End Sub";

            using (var explorer = new MockedCodeExplorer(inputCode)
                .SelectFirstModule())
            {
                var module = explorer.ViewModel.SelectedItem;
                var node = module.Children.Single(s => s.Name == "Foo");

                Assert.AreEqual(0, new CompareByName().Compare(node, node));
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void CompareByNodeType_ReturnsCorrectMemberFirst_MemberPassedFirst()
        {
            const string inputCode =
@"Sub Foo()
End Sub

Sub Bar()
    Foo
End Sub";

            using (var explorer = new MockedCodeExplorer(inputCode)
                .SelectFirstModule())
            {
                var module = explorer.ViewModel.SelectedItem;
                var memberNode1 = module.Children.Single(s => s.Name == "Foo");
                var memberNode2 = module.Children.Single(s => s.Name == "Bar");

                Assert.AreEqual(-1, new CompareByCodeLine().Compare(memberNode1, memberNode2));
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void CompareByNodeType_ReturnsZeroForIdenticalNodes()
        {
            const string inputCode =
@"Sub Foo()
End Sub

Sub Bar()
    Foo
End Sub";

            using (var explorer = new MockedCodeExplorer(inputCode)
                .SelectFirstModule())
            {
                var module = explorer.ViewModel.SelectedItem;
                var node = module.Children.Single(s => s.Name == "Foo");

                Assert.AreEqual(0, new CompareByNodeType().Compare(node, node));
            }
        }
    }
}
