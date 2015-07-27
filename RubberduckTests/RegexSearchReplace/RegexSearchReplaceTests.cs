using Microsoft.Vbe.Interop;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Navigations.RegexSearchReplace;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Extensions;
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

            var regexSearchReplace = new Rubberduck.Navigations.RegexSearchReplace.RegexSearchReplace(new RegexSearchReplaceModel(vbe.Object, new Selection()));
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

            var regexSearchReplace = new Rubberduck.Navigations.RegexSearchReplace.RegexSearchReplace(new RegexSearchReplaceModel(vbe.Object, new Selection()));
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

            var regexSearchReplace = new Rubberduck.Navigations.RegexSearchReplace.RegexSearchReplace(new RegexSearchReplaceModel(vbe.Object, new Selection()));
            regexSearchReplace.Replace("(1-)?\\p{N}{3}-\\p{N}{3}-\\p{N}{4}\\b", "hi", RegexSearchReplaceScope.CurrentFile);

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

            var regexSearchReplace = new Rubberduck.Navigations.RegexSearchReplace.RegexSearchReplace(new RegexSearchReplaceModel(vbe.Object, new Selection()));
            regexSearchReplace.ReplaceAll("(1-)?\\p{N}{3}-\\p{N}{3}-\\p{N}{4}\\b", "hi", RegexSearchReplaceScope.CurrentFile);

            //assert
            Assert.AreEqual(expectedCode, module.Lines());
        }
    }
}
