using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.API;
using Rubberduck.Inspections;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Application;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;
using Accessibility = Rubberduck.Parsing.Symbols.Accessibility;
using ParserState = Rubberduck.Parsing.VBA.ParserState;

namespace RubberduckTests.Inspections
{
    [TestClass]
    public class MemberNotOnInterfaceInspectionTests
    {
        private static void AddTestBuiltInLibrary(RubberduckParserState state)
        {
            var projectName = new QualifiedModuleName("TestLib", string.Empty, "TestLib");
            var library = new ProjectDeclaration(projectName.QualifyMemberName("TestLib"), "TestLib", true);
            state.AddDeclaration(library);

            var attributes = new Attributes();
            attributes.AddPredeclaredIdTypeAttribute();
            attributes.AddGlobalClassAttribute();

            var extensible = new ClassModuleDeclaration(projectName.QualifyMemberName("Extensible"), 
                                                        library, 
                                                        "Extensible", 
                                                        true, 
                                                        null,
                                                        attributes,
                                                        true) { IsExtensible = true };
            state.AddDeclaration(extensible);
            
            var nonExtensible = new ClassModuleDeclaration(projectName.QualifyMemberName("NonExtensible"), 
                                                           library, 
                                                           "NonExtensible", 
                                                           true, 
                                                           null,
                                                           attributes,
                                                           true);
            state.AddDeclaration(nonExtensible);

            var member = new SubroutineDeclaration(extensible.QualifiedName.QualifiedModuleName.QualifyMemberName("ExtensibleMember"), 
                                                   extensible, 
                                                   extensible,
                                                   string.Empty, 
                                                   Accessibility.Global, 
                                                   null, 
                                                   Selection.Home, 
                                                   true, 
                                                   null,
                                                   new Attributes());
            state.AddDeclaration(member);
            state.CoClasses.TryAdd(new List<string> { member.IdentifierName }, extensible);

            member = new SubroutineDeclaration(nonExtensible.QualifiedName.QualifiedModuleName.QualifyMemberName("NonExtensibleMember"),
                                               nonExtensible,
                                               nonExtensible,
                                               string.Empty,
                                               Accessibility.Global,
                                               null,
                                               Selection.Home,
                                               true,
                                               null,
                                               new Attributes());
            state.AddDeclaration(member);
            state.CoClasses.TryAdd(new List<string> { member.IdentifierName }, nonExtensible);
        }

//        [TestMethod]
//        [TestCategory("Inspections")]
//        public void MemberNotOnInterface_ReturnsResult_GlobalReference()
//        {
//            const string inputCode =
//@"Sub MemberTest()
//    Extensible.Foo
//End Sub";

//            //Arrange
//            var builder = new MockVbeBuilder();
//            var project = builder.ProjectBuilder("VBAProject", ProjectProtection.Unprotected)
//                .AddComponent("Codez", ComponentType.StandardModule, inputCode)
//                .Build();
//            var vbe = builder.AddProject(project).Build();

//            var mockHost = new Mock<IHostApplication>();
//            mockHost.SetupAllProperties();
//            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(new Mock<ISinks>().Object));

//            AddTestBuiltInLibrary(parser.State);

//            parser.Parse(new CancellationTokenSource());
//            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

//            var inspection = new MemberNotOnInterfaceInspection(parser.State);
//            var inspectionResults = inspection.GetInspectionResults();

//            Assert.AreEqual(1, inspectionResults.Count());
//        }

//        [TestMethod]
//        [TestCategory("Inspections")]
//        public void MemberNotOnInterface_DoesNotReturnResult_DeclaredMember()
//        {
//            const string inputCode =
//@"Sub MemberTest()
//    Extensible.ExtensibleMember
//End Sub";

//            //Arrange
//            var builder = new MockVbeBuilder();
//            var project = builder.ProjectBuilder("VBAProject", ProjectProtection.Unprotected)
//                .AddComponent("Codez", ComponentType.StandardModule, inputCode)
//                .Build();
//            var vbe = builder.AddProject(project).Build();

//            var mockHost = new Mock<IHostApplication>();
//            mockHost.SetupAllProperties();
//            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(new Mock<ISinks>().Object));

//            parser.Parse(new CancellationTokenSource());
//            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

//            AddTestBuiltInLibrary(parser.State);

//            var inspection = new MemberNotOnInterfaceInspection(parser.State);
//            var inspectionResults = inspection.GetInspectionResults();

//            Assert.IsFalse(inspectionResults.Any());
//        }

//        [TestMethod]
//        [TestCategory("Inspections")]
//        public void MemberNotOnInterface_DoesNotReturnResult_NonExtensible()
//        {
//            const string inputCode =
//@"Sub MemberTest()
//    NonExtensible.Foo
//End Sub";

//            //Arrange
//            var builder = new MockVbeBuilder();
//            var project = builder.ProjectBuilder("VBAProject", ProjectProtection.Unprotected)
//                .AddComponent("Codez", ComponentType.StandardModule, inputCode)
//                .Build();
//            var vbe = builder.AddProject(project).Build();

//            var mockHost = new Mock<IHostApplication>();
//            mockHost.SetupAllProperties();
//            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(new Mock<ISinks>().Object));

//            parser.Parse(new CancellationTokenSource());
//            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

//            AddTestBuiltInLibrary(parser.State);

//            var inspection = new MemberNotOnInterfaceInspection(parser.State);
//            var inspectionResults = inspection.GetInspectionResults();

//            Assert.IsFalse(inspectionResults.Any());
//        }
    }
}
