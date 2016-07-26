using System.Threading;
using Microsoft.Vbe.Interop;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Parsing;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command.Refactorings;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.VBEHost;
using RubberduckTests.Mocks;

namespace RubberduckTests.Commands
{
    [TestClass]
    public class RefactorCommandTests
    {
        [TestMethod]
        public void EncapsulateField_CanExecute_NullActiveCodePane()
        {
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule("", out component);
            vbe.Setup(v => v.ActiveCodePane).Returns((CodePane)null);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(new Mock<ISinks>().Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error)
            {
                Assert.Inconclusive("Parser Error");
            }

            var encapsulateFieldCommand = new RefactorEncapsulateFieldCommand(vbe.Object, parser.State, null);
            Assert.IsFalse(encapsulateFieldCommand.CanExecute(null));
        }

        [TestMethod]
        public void EncapsulateField_CanExecute_NonReadyState()
        {
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule("", out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(new Mock<ISinks>().Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error)
            {
                Assert.Inconclusive("Parser Error");
            }
            parser.State.SetStatusAndFireStateChanged(ParserState.ResolvedDeclarations);

            var encapsulateFieldCommand = new RefactorEncapsulateFieldCommand(vbe.Object, parser.State, null);
            Assert.IsFalse(encapsulateFieldCommand.CanExecute(null));
        }

        [TestMethod]
        public void EncapsulateField_CanExecute_LocalVariable()
        {
            var input =
@"Sub Foo()
    Dim d As Boolean
End Sub";

            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(input, out component, new Selection(2, 9, 2, 9));
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(new Mock<ISinks>().Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error)
            {
                Assert.Inconclusive("Parser Error");
            }

            var encapsulateFieldCommand = new RefactorEncapsulateFieldCommand(vbe.Object, parser.State, null);
            Assert.IsFalse(encapsulateFieldCommand.CanExecute(null));
        }

        [TestMethod]
        public void EncapsulateField_CanExecute_Proc()
        {
            var input =
@"Dim d As Boolean
Sub Foo()
End Sub";

            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(input, out component, new Selection(2, 7, 2, 7));
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(new Mock<ISinks>().Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error)
            {
                Assert.Inconclusive("Parser Error");
            }

            var encapsulateFieldCommand = new RefactorEncapsulateFieldCommand(vbe.Object, parser.State, null);
            Assert.IsFalse(encapsulateFieldCommand.CanExecute(null));
        }

        [TestMethod]
        public void EncapsulateField_CanExecute_Field()
        {
            var input =
@"Dim d As Boolean
Sub Foo()
End Sub";

            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(input, out component, new Selection(1, 5, 1, 5));
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(new Mock<ISinks>().Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error)
            {
                Assert.Inconclusive("Parser Error");
            }

            var encapsulateFieldCommand = new RefactorEncapsulateFieldCommand(vbe.Object, parser.State, null);
            Assert.IsTrue(encapsulateFieldCommand.CanExecute(null));
        }

        [TestMethod]
        public void ExtractInterface_CanExecute_NullActiveCodePane()
        {
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule("", out component);
            vbe.Setup(v => v.ActiveCodePane).Returns((CodePane)null);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(new Mock<ISinks>().Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error)
            {
                Assert.Inconclusive("Parser Error");
            }

            var extractInterfaceCommand = new RefactorExtractInterfaceCommand(vbe.Object, parser.State, null);
            Assert.IsFalse(extractInterfaceCommand.CanExecute(null));
        }

        [TestMethod]
        public void ExtractInterface_CanExecute_NonReadyState()
        {
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleModule("", vbext_ComponentType.vbext_ct_ClassModule, out component, new Selection());
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(new Mock<ISinks>().Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error)
            {
                Assert.Inconclusive("Parser Error");
            }
            parser.State.SetStatusAndFireStateChanged(ParserState.ResolvedDeclarations);

            var extractInterfaceCommand = new RefactorExtractInterfaceCommand(vbe.Object, parser.State, null);
            Assert.IsFalse(extractInterfaceCommand.CanExecute(null));
        }

        [TestMethod]
        public void ExtractInterface_CanExecute_NoMembers()
        {
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleModule("Option Explicit", vbext_ComponentType.vbext_ct_ClassModule, out component, new Selection());
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(new Mock<ISinks>().Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error)
            {
                Assert.Inconclusive("Parser Error");
            }

            var extractInterfaceCommand = new RefactorExtractInterfaceCommand(vbe.Object, parser.State, null);
            Assert.IsFalse(extractInterfaceCommand.CanExecute(null));
        }

        [TestMethod]
        public void ExtractInterface_CanExecute_Proc_StdModule()
        {
            var input =
@"Sub foo()
End Sub";

            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(input, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(new Mock<ISinks>().Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error)
            {
                Assert.Inconclusive("Parser Error");
            }

            var extractInterfaceCommand = new RefactorExtractInterfaceCommand(vbe.Object, parser.State, null);
            Assert.IsFalse(extractInterfaceCommand.CanExecute(null));
        }

        [TestMethod]
        public void ExtractInterface_CanExecute_Field()
        {
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleModule("Dim d As Boolean", vbext_ComponentType.vbext_ct_ClassModule, out component, new Selection());
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(new Mock<ISinks>().Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error)
            {
                Assert.Inconclusive("Parser Error");
            }

            var extractInterfaceCommand = new RefactorExtractInterfaceCommand(vbe.Object, parser.State, null);
            Assert.IsFalse(extractInterfaceCommand.CanExecute(null));
        }

        [TestMethod]
        public void ExtractInterface_CanExecute_ClassWithoutMembers_SameNameAsClassWithMembers()
        {
            var input =
@"Sub foo()
End Sub";

            var builder = new MockVbeBuilder();
            var proj1 = builder.ProjectBuilder("TestProj1", vbext_ProjectProtection.vbext_pp_none).AddComponent("Comp1", vbext_ComponentType.vbext_ct_ClassModule, input).Build();
            var proj2 = builder.ProjectBuilder("TestProj2", vbext_ProjectProtection.vbext_pp_none).AddComponent("Comp1", vbext_ComponentType.vbext_ct_ClassModule, "").Build();

            var vbe = builder.AddProject(proj1).AddProject(proj2).Build();
            vbe.Setup(s => s.ActiveCodePane).Returns(proj2.Object.VBComponents.Item(0).CodeModule.CodePane);

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(new Mock<ISinks>().Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error)
            {
                Assert.Inconclusive("Parser Error");
            }

            var extractInterfaceCommand = new RefactorExtractInterfaceCommand(vbe.Object, parser.State, null);
            Assert.IsFalse(extractInterfaceCommand.CanExecute(null));
        }

        [TestMethod]
        public void ExtractInterface_CanExecute_ClassWithMembers_SameNameAsClassWithMembers()
        {
            var input =
@"Sub foo()
End Sub";

            var builder = new MockVbeBuilder();
            var proj1 = builder.ProjectBuilder("TestProj1", vbext_ProjectProtection.vbext_pp_none).AddComponent("Comp1", vbext_ComponentType.vbext_ct_ClassModule, input).Build();
            var proj2 = builder.ProjectBuilder("TestProj2", vbext_ProjectProtection.vbext_pp_none).AddComponent("Comp1", vbext_ComponentType.vbext_ct_ClassModule, "").Build();

            var vbe = builder.AddProject(proj1).AddProject(proj2).Build();
            vbe.Setup(s => s.ActiveCodePane).Returns(proj1.Object.VBComponents.Item(0).CodeModule.CodePane);

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(new Mock<ISinks>().Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error)
            {
                Assert.Inconclusive("Parser Error");
            }

            var extractInterfaceCommand = new RefactorExtractInterfaceCommand(vbe.Object, parser.State, null);
            Assert.IsTrue(extractInterfaceCommand.CanExecute(null));
        }

        [TestMethod]
        public void ExtractInterface_CanExecute_Proc()
        {
            var input =
@"Sub foo()
End Sub";

            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleModule(input, vbext_ComponentType.vbext_ct_ClassModule, out component, new Selection());
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(new Mock<ISinks>().Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error)
            {
                Assert.Inconclusive("Parser Error");
            }

            var extractInterfaceCommand = new RefactorExtractInterfaceCommand(vbe.Object, parser.State, null);
            Assert.IsTrue(extractInterfaceCommand.CanExecute(null));
        }

        [TestMethod]
        public void ExtractInterface_CanExecute_Function()
        {
            var input =
@"Function foo() As Integer
End Function";

            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleModule(input, vbext_ComponentType.vbext_ct_ClassModule, out component, new Selection());
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(new Mock<ISinks>().Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error)
            {
                Assert.Inconclusive("Parser Error");
            }

            var extractInterfaceCommand = new RefactorExtractInterfaceCommand(vbe.Object, parser.State, null);
            Assert.IsTrue(extractInterfaceCommand.CanExecute(null));
        }

        [TestMethod]
        public void ExtractInterface_CanExecute_PropertyGet()
        {
            var input =
@"Property Get foo() As Boolean
End Property";

            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleModule(input, vbext_ComponentType.vbext_ct_ClassModule, out component, new Selection());
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(new Mock<ISinks>().Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error)
            {
                Assert.Inconclusive("Parser Error");
            }

            var extractInterfaceCommand = new RefactorExtractInterfaceCommand(vbe.Object, parser.State, null);
            Assert.IsTrue(extractInterfaceCommand.CanExecute(null));
        }

        [TestMethod]
        public void ExtractInterface_CanExecute_PropertyLet()
        {
            var input =
@"Property Let foo(value)
End Property";

            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleModule(input, vbext_ComponentType.vbext_ct_ClassModule, out component, new Selection());
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(new Mock<ISinks>().Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error)
            {
                Assert.Inconclusive("Parser Error");
            }

            var extractInterfaceCommand = new RefactorExtractInterfaceCommand(vbe.Object, parser.State, null);
            Assert.IsTrue(extractInterfaceCommand.CanExecute(null));
        }

        [TestMethod]
        public void ExtractInterface_CanExecute_PropertySet()
        {
            var input =
@"Property Set foo(value)
End Property";

            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleModule(input, vbext_ComponentType.vbext_ct_ClassModule, out component, new Selection());
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(new Mock<ISinks>().Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error)
            {
                Assert.Inconclusive("Parser Error");
            }

            var extractInterfaceCommand = new RefactorExtractInterfaceCommand(vbe.Object, parser.State, null);
            Assert.IsTrue(extractInterfaceCommand.CanExecute(null));
        }

        [TestMethod]
        public void ImplementInterface_CanExecute_NullActiveCodePane()
        {
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule("", out component);
            vbe.Setup(v => v.ActiveCodePane).Returns((CodePane)null);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(new Mock<ISinks>().Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error)
            {
                Assert.Inconclusive("Parser Error");
            }

            var implementInterfaceCommand = new RefactorImplementInterfaceCommand(vbe.Object, parser.State);
            Assert.IsFalse(implementInterfaceCommand.CanExecute(null));
        }

        [TestMethod]
        public void ImplementInterface_CanExecute_NonReadyState()
        {
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule("", out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(new Mock<ISinks>().Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error)
            {
                Assert.Inconclusive("Parser Error");
            }
            parser.State.SetStatusAndFireStateChanged(ParserState.ResolvedDeclarations);

            var implementInterfaceCommand = new RefactorImplementInterfaceCommand(vbe.Object, parser.State);
            Assert.IsFalse(implementInterfaceCommand.CanExecute(null));
        }

        [TestMethod]
        public void ImplementInterface_CanExecute_ImplementsInterfaceNotSelected()
        {
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleModule("", vbext_ComponentType.vbext_ct_ClassModule, out component, new Selection());
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(new Mock<ISinks>().Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error)
            {
                Assert.Inconclusive("Parser Error");
            }

            var implementInterfaceCommand = new RefactorImplementInterfaceCommand(vbe.Object, parser.State);
            Assert.IsFalse(implementInterfaceCommand.CanExecute(null));
        }

        [TestMethod]
        public void ImplementInterface_CanExecute_ImplementsInterfaceSelected()
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("IClass1", vbext_ComponentType.vbext_ct_ClassModule, "")
                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, "Implements IClass1", new Selection(1, 1, 1, 1))
                .Build();

            var vbe = builder.AddProject(project).Build();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(new Mock<ISinks>().Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error)
            {
                Assert.Inconclusive("Parser Error");
            }

            var implementInterfaceCommand = new RefactorImplementInterfaceCommand(vbe.Object, parser.State);
            Assert.IsTrue(implementInterfaceCommand.CanExecute(null));
        }
    }
}