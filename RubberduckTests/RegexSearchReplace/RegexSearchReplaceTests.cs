using Microsoft.Vbe.Interop;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Navigations.RegexSearchReplace;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;
using RubberduckTests.Mocks;

namespace RubberduckTests.RegexSearchReplace
{
    [TestClass]
    public class RegexSearchReplaceTests
    {
        [TestMethod]
        public void RegexSearch_ExactMatch_CurrentFile()
        {
            const string inputCode = @"
Private Sub Foo()
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode)
                .Build().Object;
            var vbe = builder.Build();

            var codePaneFactory = new RubberduckCodePaneFactory();
            var parseResult = new RubberduckParser(codePaneFactory).Parse(project);

            var regexSearchReplace =
                new Rubberduck.Navigations.RegexSearchReplace.RegexSearchReplace(
                    new RegexSearchReplaceModel(vbe.Object, parseResult, new QualifiedSelection()), codePaneFactory);
            var results = regexSearchReplace.Find("Foo", RegexSearchReplaceScope.CurrentFile);

            //assert
            Assert.AreEqual(1, results.Count);
        }

        [TestMethod]
        public void RegexSearch_MatchUSPhoneNumber_CurrentFile()
        {
            const string inputCode = @"
Private Sub Foo()
    Dim phoneNumber1 As String
    phoneNumber1 = ""123-456-7890""

    Dim phoneNumber2 As String
    phoneNumber2 = ""987-654-3210""

    Dim phoneNumber3 As String
    phoneNumber3 = ""1-222-333-4444""
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode)
                .Build().Object;
            var vbe = builder.Build();

            var codePaneFactory = new RubberduckCodePaneFactory();
            var parseResult = new RubberduckParser(codePaneFactory).Parse(project);

            var regexSearchReplace =
                new Rubberduck.Navigations.RegexSearchReplace.RegexSearchReplace(
                    new RegexSearchReplaceModel(vbe.Object, parseResult, new QualifiedSelection()), codePaneFactory);
            var results = regexSearchReplace.Find("(1-)?\\p{N}{3}-\\p{N}{3}-\\p{N}{4}\\b", RegexSearchReplaceScope.CurrentFile);

            //assert
            Assert.AreEqual(3, results.Count);
        }

        [TestMethod]
        public void RegexSearchReplace_RemoveUSPhoneNumber_CurrentFile()
        {
            const string inputCode = @"
Private Sub Foo()
    Dim phoneNumber1 As String
    phoneNumber1 = ""123-456-7890""

    Dim phoneNumber2 As String
    phoneNumber2 = ""987-654-3210""

    Dim phoneNumber3 As String
    phoneNumber3 = ""1-222-333-4444""
End Sub";

            const string expectedCode = @"
Private Sub Foo()
    Dim phoneNumber1 As String
    phoneNumber1 = ""hi""

    Dim phoneNumber2 As String
    phoneNumber2 = ""987-654-3210""

    Dim phoneNumber3 As String
    phoneNumber3 = ""1-222-333-4444""
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode)
                .Build().Object;
            var vbe = builder.Build();
            var module = project.VBComponents.Item(0).CodeModule;

            var codePaneFactory = new RubberduckCodePaneFactory();
            var parseResult = new RubberduckParser(codePaneFactory).Parse(project);

            var regexSearchReplace =
                new Rubberduck.Navigations.RegexSearchReplace.RegexSearchReplace(
                    new RegexSearchReplaceModel(vbe.Object, parseResult, new QualifiedSelection()), codePaneFactory);
            regexSearchReplace.Replace("(1-)?\\p{N}{3}-\\p{N}{3}-\\p{N}{4}\\b", "hi", RegexSearchReplaceScope.CurrentFile);

            //assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void RegexSearchReplace_RemoveUSPhoneNumber_CurrentFile_NoResults()
        {
            const string inputCode = @"
Private Sub Foo()
End Sub";

            const string expectedCode = @"
Private Sub Foo()
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode)
                .Build().Object;
            var vbe = builder.Build();
            var module = project.VBComponents.Item(0).CodeModule;

            var codePaneFactory = new RubberduckCodePaneFactory();
            var parseResult = new RubberduckParser(codePaneFactory).Parse(project);

            var regexSearchReplace =
                new Rubberduck.Navigations.RegexSearchReplace.RegexSearchReplace(
                    new RegexSearchReplaceModel(vbe.Object, parseResult, new QualifiedSelection()), codePaneFactory);
            regexSearchReplace.Replace("Goo", "Hoo", RegexSearchReplaceScope.CurrentFile);

            //assert
            Assert.AreEqual(expectedCode, module.Lines());
        }


        [TestMethod]
        public void RegexSearchReplaceAll_RemoveUSPhoneNumber_CurrentFile()
        {
            const string inputCode = @"
Private Sub Foo()
    Dim phoneNumber1 As String
    phoneNumber1 = ""123-456-7890""

    Dim phoneNumber2 As String
    phoneNumber2 = ""987-654-3210""

    Dim phoneNumber3 As String
    phoneNumber3 = ""1-222-333-4444""
End Sub";

            const string expectedCode = @"
Private Sub Foo()
    Dim phoneNumber1 As String
    phoneNumber1 = ""hi""

    Dim phoneNumber2 As String
    phoneNumber2 = ""hi""

    Dim phoneNumber3 As String
    phoneNumber3 = ""hi""
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode)
                .Build().Object;
            var vbe = builder.Build();
            var module = project.VBComponents.Item(0).CodeModule;

            var codePaneFactory = new RubberduckCodePaneFactory();
            var parseResult = new RubberduckParser(codePaneFactory).Parse(project);

            var regexSearchReplace =
                new Rubberduck.Navigations.RegexSearchReplace.RegexSearchReplace(
                    new RegexSearchReplaceModel(vbe.Object, parseResult, new QualifiedSelection()), codePaneFactory);
            regexSearchReplace.ReplaceAll("(1-)?\\p{N}{3}-\\p{N}{3}-\\p{N}{4}\\b", "hi", RegexSearchReplaceScope.CurrentFile);

            //assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void RegexSearch_MatchUSPhoneNumber_Selection()
        {
            const string inputCode = @"
Private Sub Foo()
    Dim phoneNumber1 As String
    phoneNumber1 = ""123-456-7890""

    Dim phoneNumber2 As String
    phoneNumber2 = ""987-654-3210""

    Dim phoneNumber3 As String
    phoneNumber3 = ""1-222-333-4444""
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode)
                .Build().Object;
            var vbe = builder.Build();

            var codePaneFactory = new RubberduckCodePaneFactory();
            var parseResult = new RubberduckParser(codePaneFactory).Parse(project);

            var regexSearchReplace =
                new Rubberduck.Navigations.RegexSearchReplace.RegexSearchReplace(new RegexSearchReplaceModel(
                    vbe.Object, parseResult,
                    new QualifiedSelection(new QualifiedModuleName(), new Selection(5, 1, 8, 1))), codePaneFactory);
            var results = regexSearchReplace.Find("(1-)?\\p{N}{3}-\\p{N}{3}-\\p{N}{4}\\b", RegexSearchReplaceScope.Selection);

            //assert
            Assert.AreEqual(1, results.Count);
        }

        [TestMethod]
        public void RegexSearchReplace_RemoveUSPhoneNumber_Selection()
        {
            const string inputCode = @"
Private Sub Foo()
    Dim phoneNumber1 As String
    phoneNumber1 = ""123-456-7890""

    Dim phoneNumber2 As String
    phoneNumber2 = ""987-654-3210""

    Dim phoneNumber3 As String
    phoneNumber3 = ""1-222-333-4444""
End Sub";

            const string expectedCode = @"
Private Sub Foo()
    Dim phoneNumber1 As String
    phoneNumber1 = ""123-456-7890""

    Dim phoneNumber2 As String
    phoneNumber2 = ""hi""

    Dim phoneNumber3 As String
    phoneNumber3 = ""1-222-333-4444""
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode)
                .Build().Object;
            var vbe = builder.Build();
            var module = project.VBComponents.Item(0).CodeModule;

            var codePaneFactory = new RubberduckCodePaneFactory();
            var parseResult = new RubberduckParser(codePaneFactory).Parse(project);

            var regexSearchReplace =
                new Rubberduck.Navigations.RegexSearchReplace.RegexSearchReplace(new RegexSearchReplaceModel(
                    vbe.Object, parseResult,
                    new QualifiedSelection(new QualifiedModuleName(), new Selection(5, 1, 8, 1))), codePaneFactory);
            regexSearchReplace.Replace("(1-)?\\p{N}{3}-\\p{N}{3}-\\p{N}{4}\\b", "hi", RegexSearchReplaceScope.Selection);

            //assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void RegexSearchReplaceAll_RemoveUSPhoneNumber_Selection()
        {
            const string inputCode = @"
Private Sub Foo()
    Dim phoneNumber1 As String
    phoneNumber1 = ""123-456-7890""

    Dim phoneNumber2 As String
    phoneNumber2 = ""987-654-3210""

    Dim phoneNumber3 As String
    phoneNumber3 = ""1-222-333-4444""
End Sub";

            const string expectedCode = @"
Private Sub Foo()
    Dim phoneNumber1 As String
    phoneNumber1 = ""123-456-7890""

    Dim phoneNumber2 As String
    phoneNumber2 = ""hi""

    Dim phoneNumber3 As String
    phoneNumber3 = ""hi""
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode)
                .Build().Object;
            var vbe = builder.Build();
            var module = project.VBComponents.Item(0).CodeModule;

            var codePaneFactory = new RubberduckCodePaneFactory();
            var parseResult = new RubberduckParser(codePaneFactory).Parse(project);

            var regexSearchReplace =
                new Rubberduck.Navigations.RegexSearchReplace.RegexSearchReplace(new RegexSearchReplaceModel(
                    vbe.Object, parseResult,
                    new QualifiedSelection(new QualifiedModuleName(), new Selection(5, 1, 12, 1))), codePaneFactory);
            regexSearchReplace.ReplaceAll("(1-)?\\p{N}{3}-\\p{N}{3}-\\p{N}{4}\\b", "hi", RegexSearchReplaceScope.Selection);

            //assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void RegexSearch_MatchUSPhoneNumber_CurrentProject()
        {
            const string inputCode1 = @"
Private Sub Foo()
    Dim phoneNumber1 As String
    phoneNumber1 = ""123-456-7890""

    Dim phoneNumber2 As String
    phoneNumber2 = ""987-654-3210""

    Dim phoneNumber3 As String
    phoneNumber3 = ""1-222-333-4444""
End Sub";

            const string inputCode2 = @"
Private Sub Foo()
    Dim phoneNumber1 As String
    phoneNumber1 = ""123-456-7890""

    Dim phoneNumber2 As String
    phoneNumber2 = ""987-654-3210""

    Dim phoneNumber3 As String
    phoneNumber3 = ""1-222-333-4444""
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode1)
                .AddComponent("Class2", vbext_ComponentType.vbext_ct_ClassModule, inputCode2)
                .Build().Object;
            var vbe = builder.Build();
            vbe.Setup(v => v.ActiveVBProject).Returns(project);

            var codePaneFactory = new RubberduckCodePaneFactory();
            var parseResult = new RubberduckParser(codePaneFactory).Parse(project);

            var regexSearchReplace =
                new Rubberduck.Navigations.RegexSearchReplace.RegexSearchReplace(
                    new RegexSearchReplaceModel(vbe.Object, parseResult, new QualifiedSelection()), codePaneFactory);
            var results = regexSearchReplace.Find("(1-)?\\p{N}{3}-\\p{N}{3}-\\p{N}{4}\\b", RegexSearchReplaceScope.CurrentProject);

            //assert
            Assert.AreEqual(6, results.Count);
        }

        [TestMethod]
        public void RegexSearchReplace_RemoveUSPhoneNumber_CurrentProject()
        {
            const string inputCode1 = @"
Private Sub Foo()
    Dim phoneNumber1 As String
    phoneNumber1 = ""123-456-7890""

    Dim phoneNumber2 As String
    phoneNumber2 = ""987-654-3210""

    Dim phoneNumber3 As String
    phoneNumber3 = ""1-222-333-4444""
End Sub";

            const string inputCode2 = @"
Private Sub Foo()
    Dim phoneNumber1 As String
    phoneNumber1 = ""123-456-7890""

    Dim phoneNumber2 As String
    phoneNumber2 = ""987-654-3210""

    Dim phoneNumber3 As String
    phoneNumber3 = ""1-222-333-4444""
End Sub";

            const string expectedCode1 = @"
Private Sub Foo()
    Dim phoneNumber1 As String
    phoneNumber1 = ""hi""

    Dim phoneNumber2 As String
    phoneNumber2 = ""987-654-3210""

    Dim phoneNumber3 As String
    phoneNumber3 = ""1-222-333-4444""
End Sub";

            const string expectedCode2 = @"
Private Sub Foo()
    Dim phoneNumber1 As String
    phoneNumber1 = ""123-456-7890""

    Dim phoneNumber2 As String
    phoneNumber2 = ""987-654-3210""

    Dim phoneNumber3 As String
    phoneNumber3 = ""1-222-333-4444""
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode1)
                .AddComponent("Class2", vbext_ComponentType.vbext_ct_ClassModule, inputCode2)
                .Build().Object;
            var vbe = builder.Build();
            vbe.Setup(v => v.ActiveVBProject).Returns(project);
            var module1 = project.VBComponents.Item(0).CodeModule;
            var module2 = project.VBComponents.Item(1).CodeModule;

            var codePaneFactory = new RubberduckCodePaneFactory();
            var parseResult = new RubberduckParser(codePaneFactory).Parse(project);

            var regexSearchReplace =
                new Rubberduck.Navigations.RegexSearchReplace.RegexSearchReplace(
                    new RegexSearchReplaceModel(vbe.Object, parseResult, new QualifiedSelection()), codePaneFactory);
            regexSearchReplace.Replace("(1-)?\\p{N}{3}-\\p{N}{3}-\\p{N}{4}\\b", "hi", RegexSearchReplaceScope.CurrentProject);

            //assert
            Assert.AreEqual(expectedCode1, module1.Lines());
            Assert.AreEqual(expectedCode2, module2.Lines());
        }

        [TestMethod]
        public void RegexSearchReplaceAll_RemoveUSPhoneNumber_CurrentProject()
        {
            const string inputCode1 = @"
Private Sub Foo()
    Dim phoneNumber1 As String
    phoneNumber1 = ""123-456-7890""

    Dim phoneNumber2 As String
    phoneNumber2 = ""987-654-3210""

    Dim phoneNumber3 As String
    phoneNumber3 = ""1-222-333-4444""
End Sub";

            const string inputCode2 = @"
Private Sub Foo()
    Dim phoneNumber1 As String
    phoneNumber1 = ""123-456-7890""

    Dim phoneNumber2 As String
    phoneNumber2 = ""987-654-3210""

    Dim phoneNumber3 As String
    phoneNumber3 = ""1-222-333-4444""
End Sub";

            const string expectedCode1 = @"
Private Sub Foo()
    Dim phoneNumber1 As String
    phoneNumber1 = ""hi""

    Dim phoneNumber2 As String
    phoneNumber2 = ""hi""

    Dim phoneNumber3 As String
    phoneNumber3 = ""hi""
End Sub";

            const string expectedCode2 = @"
Private Sub Foo()
    Dim phoneNumber1 As String
    phoneNumber1 = ""hi""

    Dim phoneNumber2 As String
    phoneNumber2 = ""hi""

    Dim phoneNumber3 As String
    phoneNumber3 = ""hi""
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode1)
                .AddComponent("Class2", vbext_ComponentType.vbext_ct_ClassModule, inputCode2)
                .Build().Object;
            var vbe = builder.Build();
            vbe.Setup(v => v.ActiveVBProject).Returns(project);
            var module1 = project.VBComponents.Item(0).CodeModule;
            var module2 = project.VBComponents.Item(1).CodeModule;

            var codePaneFactory = new RubberduckCodePaneFactory();
            var parseResult = new RubberduckParser(codePaneFactory).Parse(project);

            var regexSearchReplace =
                new Rubberduck.Navigations.RegexSearchReplace.RegexSearchReplace(
                    new RegexSearchReplaceModel(vbe.Object, parseResult, new QualifiedSelection()), codePaneFactory);
            regexSearchReplace.ReplaceAll("(1-)?\\p{N}{3}-\\p{N}{3}-\\p{N}{4}\\b", "hi", RegexSearchReplaceScope.CurrentProject);

            //assert
            Assert.AreEqual(expectedCode1, module1.Lines());
            Assert.AreEqual(expectedCode2, module2.Lines());
        }

        [TestMethod]
        public void RegexSearch_MatchUSPhoneNumber_AllOpenProjects ()
        {
            const string inputCode1 = @"
Private Sub Foo()
    Dim phoneNumber1 As String
    phoneNumber1 = ""123-456-7890""

    Dim phoneNumber2 As String
    phoneNumber2 = ""987-654-3210""

    Dim phoneNumber3 As String
    phoneNumber3 = ""1-222-333-4444""
End Sub";

            const string inputCode2 = @"
Private Sub Foo()
    Dim phoneNumber1 As String
    phoneNumber1 = ""123-456-7890""

    Dim phoneNumber2 As String
    phoneNumber2 = ""987-654-3210""

    Dim phoneNumber3 As String
    phoneNumber3 = ""1-222-333-4444""
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            var project1 = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode1)
                .AddComponent("Class2", vbext_ComponentType.vbext_ct_ClassModule, inputCode2)
                .Build();
            var project2 = builder.ProjectBuilder("TestProject2", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode1)
                .AddComponent("Class2", vbext_ComponentType.vbext_ct_ClassModule, inputCode2)
                .Build();

            builder.AddProject(project1);
            builder.AddProject(project2);
            var vbe = builder.Build();

            var codePaneFactory = new RubberduckCodePaneFactory();
            var regexSearchReplace =
                new Rubberduck.Navigations.RegexSearchReplace.RegexSearchReplace(
                    new RegexSearchReplaceModel(vbe.Object, null, new QualifiedSelection()), codePaneFactory);
            var results = regexSearchReplace.Find("(1-)?\\p{N}{3}-\\p{N}{3}-\\p{N}{4}\\b", RegexSearchReplaceScope.AllOpenProjects);

            //assert
            Assert.AreEqual(12, results.Count);
        }

        [TestMethod]
        public void RegexSearchReplace_RemoveUSPhoneNumber_AllOpenProjects ()
        {
            const string inputCode1 = @"
Private Sub Foo()
    Dim phoneNumber1 As String
    phoneNumber1 = ""123-456-7890""

    Dim phoneNumber2 As String
    phoneNumber2 = ""987-654-3210""

    Dim phoneNumber3 As String
    phoneNumber3 = ""1-222-333-4444""
End Sub";

            const string inputCode2 = @"
Private Sub Foo()
    Dim phoneNumber1 As String
    phoneNumber1 = ""123-456-7890""

    Dim phoneNumber2 As String
    phoneNumber2 = ""987-654-3210""

    Dim phoneNumber3 As String
    phoneNumber3 = ""1-222-333-4444""
End Sub";

            const string expectedCode1 = @"
Private Sub Foo()
    Dim phoneNumber1 As String
    phoneNumber1 = ""hi""

    Dim phoneNumber2 As String
    phoneNumber2 = ""987-654-3210""

    Dim phoneNumber3 As String
    phoneNumber3 = ""1-222-333-4444""
End Sub";

            const string expectedCode2 = @"
Private Sub Foo()
    Dim phoneNumber1 As String
    phoneNumber1 = ""123-456-7890""

    Dim phoneNumber2 As String
    phoneNumber2 = ""987-654-3210""

    Dim phoneNumber3 As String
    phoneNumber3 = ""1-222-333-4444""
End Sub";

            const string expectedCode3 = @"
Private Sub Foo()
    Dim phoneNumber1 As String
    phoneNumber1 = ""123-456-7890""

    Dim phoneNumber2 As String
    phoneNumber2 = ""987-654-3210""

    Dim phoneNumber3 As String
    phoneNumber3 = ""1-222-333-4444""
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            var project1 = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode1)
                .AddComponent("Class2", vbext_ComponentType.vbext_ct_ClassModule, inputCode2)
                .Build();
            var project2 = builder.ProjectBuilder("TestProject2", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode1)
                .AddComponent("Class2", vbext_ComponentType.vbext_ct_ClassModule, inputCode2)
                .Build();

            builder.AddProject(project1);
            builder.AddProject(project2);
            var vbe = builder.Build();
            vbe.Setup(v => v.ActiveVBProject).Returns(project1.Object);
            var module1 = project1.Object.VBComponents.Item(0).CodeModule;
            var module2 = project1.Object.VBComponents.Item(1).CodeModule;
            var module3 = project2.Object.VBComponents.Item(0).CodeModule;
            var module4 = project2.Object.VBComponents.Item(1).CodeModule;

            var codePaneFactory = new RubberduckCodePaneFactory();
            var parseResult = new RubberduckParser(codePaneFactory).Parse(project1.Object);

            var regexSearchReplace =
                new Rubberduck.Navigations.RegexSearchReplace.RegexSearchReplace(
                    new RegexSearchReplaceModel(vbe.Object, parseResult, new QualifiedSelection()), codePaneFactory);
            regexSearchReplace.Replace("(1-)?\\p{N}{3}-\\p{N}{3}-\\p{N}{4}\\b", "hi", RegexSearchReplaceScope.AllOpenProjects);

            //assert
            Assert.AreEqual(expectedCode1, module1.Lines());
            Assert.AreEqual(expectedCode2, module2.Lines());
            Assert.AreEqual(expectedCode3, module3.Lines());
            Assert.AreEqual(expectedCode2, module4.Lines());
        }

        [TestMethod]
        public void RegexSearchReplaceAll_RemoveUSPhoneNumber_AllOpenProjects ()
        {
            const string inputCode1 = @"
Private Sub Foo()
    Dim phoneNumber1 As String
    phoneNumber1 = ""123-456-7890""

    Dim phoneNumber2 As String
    phoneNumber2 = ""987-654-3210""

    Dim phoneNumber3 As String
    phoneNumber3 = ""1-222-333-4444""
End Sub";

            const string inputCode2 = @"
Private Sub Foo()
    Dim phoneNumber1 As String
    phoneNumber1 = ""123-456-7890""

    Dim phoneNumber2 As String
    phoneNumber2 = ""987-654-3210""

    Dim phoneNumber3 As String
    phoneNumber3 = ""1-222-333-4444""
End Sub";

            const string expectedCode1 = @"
Private Sub Foo()
    Dim phoneNumber1 As String
    phoneNumber1 = ""hi""

    Dim phoneNumber2 As String
    phoneNumber2 = ""hi""

    Dim phoneNumber3 As String
    phoneNumber3 = ""hi""
End Sub";

            const string expectedCode2 = @"
Private Sub Foo()
    Dim phoneNumber1 As String
    phoneNumber1 = ""hi""

    Dim phoneNumber2 As String
    phoneNumber2 = ""hi""

    Dim phoneNumber3 As String
    phoneNumber3 = ""hi""
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            var project1 = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode1)
                .AddComponent("Class2", vbext_ComponentType.vbext_ct_ClassModule, inputCode2)
                .Build();
            var project2 = builder.ProjectBuilder("TestProject2", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode1)
                .AddComponent("Class2", vbext_ComponentType.vbext_ct_ClassModule, inputCode2)
                .Build();

            builder.AddProject(project1);
            builder.AddProject(project2);
            var vbe = builder.Build();
            vbe.Setup(v => v.ActiveVBProject).Returns(project1.Object);
            var module1 = project1.Object.VBComponents.Item(0).CodeModule;
            var module2 = project1.Object.VBComponents.Item(1).CodeModule;
            var module3 = project2.Object.VBComponents.Item(0).CodeModule;
            var module4 = project2.Object.VBComponents.Item(1).CodeModule;

            var codePaneFactory = new RubberduckCodePaneFactory();
            var parseResult = new RubberduckParser(codePaneFactory).Parse(project1.Object);

            var regexSearchReplace =
                new Rubberduck.Navigations.RegexSearchReplace.RegexSearchReplace(
                    new RegexSearchReplaceModel(vbe.Object, parseResult, new QualifiedSelection()), codePaneFactory);
            regexSearchReplace.ReplaceAll("(1-)?\\p{N}{3}-\\p{N}{3}-\\p{N}{4}\\b", "hi", RegexSearchReplaceScope.AllOpenProjects);

            //assert
            Assert.AreEqual(expectedCode1, module1.Lines());
            Assert.AreEqual(expectedCode2, module2.Lines());
            Assert.AreEqual(expectedCode1, module3.Lines());
            Assert.AreEqual(expectedCode2, module4.Lines());
        }

/*        [TestMethod]
        public void RegexSearch_MatchUSPhoneNumber_CurrentBlock()
        {
            const string inputCode = @"
Private Sub Foo()
    Dim phoneNumber1 As String
    phoneNumber1 = ""123-456-7890""
End Sub

Private Sub Goo()
    Dim phoneNumber1 As String
    phoneNumber1 = ""123-456-7890""
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode)
                .Build().Object;
            var vbe = builder.Build();
            vbe.Setup(v => v.ActiveVBProject).Returns(project);

            var codePaneFactory = new RubberduckCodePaneFactory();
            var parseResult = new RubberduckParser(codePaneFactory).Parse(project);

            var regexSearchReplace =
                new Rubberduck.Navigations.RegexSearchReplace.RegexSearchReplace(new RegexSearchReplaceModel(
                    vbe.Object, parseResult, new QualifiedSelection(new QualifiedModuleName(), new Selection(3, 1, 3, 1))));
            var results = regexSearchReplace.Find("(1-)?\\p{N}{3}-\\p{N}{3}-\\p{N}{4}\\b", RegexSearchReplaceScope.CurrentBlock);

            //assert
            Assert.AreEqual(1, results.Count);
        }

        [TestMethod]
        public void RegexSearchReplace_RemoveUSPhoneNumber_CurrentBlock()
        {
            const string inputCode = @"
Private Sub Foo()
    Dim phoneNumber1 As String
    phoneNumber1 = ""123-456-7890""
End Sub

Private Sub Goo()
    Dim phoneNumber1 As String
    phoneNumber1 = ""123-456-7890""
End Sub";

            const string expectedCode = @"
Private Sub Foo()
    Dim phoneNumber1 As String
    phoneNumber1 = ""hi""
End Sub

Private Sub Goo()
    Dim phoneNumber1 As String
    phoneNumber1 = ""123-456-7890""
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode)
                .Build().Object;
            var vbe = builder.Build();
            vbe.Setup(v => v.ActiveVBProject).Returns(project);
            var module = project.VBComponents.Item(0).CodeModule;

            var codePaneFactory = new RubberduckCodePaneFactory();
            var parseResult = new RubberduckParser(codePaneFactory).Parse(project);

            var regexSearchReplace =
                new Rubberduck.Navigations.RegexSearchReplace.RegexSearchReplace(new RegexSearchReplaceModel(
                    vbe.Object, parseResult, new QualifiedSelection(new QualifiedModuleName(), new Selection(3, 1, 3, 1))));
            regexSearchReplace.Replace("(1-)?\\p{N}{3}-\\p{N}{3}-\\p{N}{4}\\b", "hi", RegexSearchReplaceScope.CurrentBlock);

            //assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void RegexSearchReplaceAll_RemoveUSPhoneNumber_CurrentBlock()
        {
            const string inputCode = @"
Private Sub Foo()
    Dim phoneNumber1 As String
    phoneNumber1 = ""123-456-7890""

    Dim phoneNumber2 As String
    phoneNumber2 = ""1-123-456-7890""
End Sub

Private Sub Goo()
    Dim phoneNumber1 As String
    phoneNumber1 = ""123-456-7890""
End Sub";

            const string expectedCode = @"
Private Sub Foo()
    Dim phoneNumber1 As String
    phoneNumber1 = ""hi""

    Dim phoneNumber2 As String
    phoneNumber2 = ""hi""
End Sub

Private Sub Goo()
    Dim phoneNumber1 As String
    phoneNumber1 = ""123-456-7890""
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode)
                .Build().Object;
            var vbe = builder.Build();
            vbe.Setup(v => v.ActiveVBProject).Returns(project);
            var module = project.VBComponents.Item(0).CodeModule;

            var codePaneFactory = new RubberduckCodePaneFactory();
            var parseResult = new RubberduckParser(codePaneFactory).Parse(project);

            var regexSearchReplace =
                new Rubberduck.Navigations.RegexSearchReplace.RegexSearchReplace(new RegexSearchReplaceModel(vbe.Object, parseResult, new QualifiedSelection()));
            regexSearchReplace.ReplaceAll("(1-)?\\p{N}{3}-\\p{N}{3}-\\p{N}{4}\\b", "hi", RegexSearchReplaceScope.CurrentBlock);

            //assert
            Assert.AreEqual(expectedCode, module.Lines());
        }*/
    }
}
