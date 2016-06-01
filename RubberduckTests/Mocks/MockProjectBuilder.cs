using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Vbe.Interop;
using Moq;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace RubberduckTests.Mocks
{
    /// <summary>
    /// Builds a mock <see cref="VBProject"/>.
    /// </summary>
    public class MockProjectBuilder
    {
        private readonly Func<VBE> _getVbe;
        private readonly MockVbeBuilder _mockVbeBuilder;
        private readonly Mock<VBProject> _project;
        private readonly Mock<VBComponents> _vbComponents;
        private readonly Mock<References> _vbReferences;

        private readonly List<Mock<VBComponent>> _componentsMock = new List<Mock<VBComponent>>();
        private readonly List<Reference> _references = new List<Reference>();

        public Mock<VBComponents> MockVBComponents
        {
            get { return _vbComponents; }
        }

        public List<Mock<VBComponent>> MockComponents
        {
            get { return _componentsMock; }
        }

        private List<VBComponent> Components
        {
            get { return _componentsMock.Select(m => m.Object).ToList(); }
        }

        public MockProjectBuilder(string name, string filename, vbext_ProjectProtection protection, Func<VBE> getVbe, MockVbeBuilder mockVbeBuilder)
        {
            _getVbe = getVbe;
            _mockVbeBuilder = mockVbeBuilder;

            _project = CreateProjectMock(name, filename, protection);

            _project.SetupProperty(m => m.HelpFile);

            _vbComponents = CreateComponentsMock();
            _project.SetupGet(m => m.VBComponents).Returns(_vbComponents.Object);

            _vbReferences = CreateReferencesMock();
            _project.SetupGet(m => m.References).Returns(_vbReferences.Object);
        }

        /// <summary>
        /// Adds a new component to the project.
        /// </summary>
        /// <param name="name">The name of the new component.</param>
        /// <param name="type">The type of component to create.</param>
        /// <param name="content">The VBA code associated to the component.</param>
        /// <param name="selection"></param>
        /// <returns>Returns the <see cref="MockProjectBuilder"/> instance.</returns>
        public MockProjectBuilder AddComponent(string name, vbext_ComponentType type, string content, Selection selection = new Selection())
        {
            var component = CreateComponentMock(name, type, content, selection);
            return AddComponent(component);
        }

        /// <summary>
        /// Adds a new mock component to the project.
        /// Use the <see cref="AddComponent(string,vbext_ComponentType,string,Selection)"/> overload to add module components.
        /// Use this overload to add user forms created with a <see cref="RubberduckTests.Mocks.MockUserFormBuilder"/> instance.
        /// </summary>
        /// <param name="component">The component to add.</param>
        /// <returns>Returns the <see cref="MockProjectBuilder"/> instance.</returns>
        public MockProjectBuilder AddComponent(Mock<VBComponent> component)
        {
            _componentsMock.Add(component);
            _getVbe().ActiveCodePane = component.Object.CodeModule.CodePane;
            return this;
        }

        /// <summary>
        /// Adds a mock reference to the project.
        /// </summary>
        /// <param name="name">The name of the referenced library.</param>
        /// <param name="filePath">The path to the referenced library.</param>
        /// <param name="isBuiltIn">Indicates whether the reference is a built-in reference.</param>
        /// <returns>Returns the <see cref="MockProjectBuilder"/> instance.</returns>
        public MockProjectBuilder AddReference(string name, string filePath, bool isBuiltIn = false)
        {
            var reference = CreateReferenceMock(name, filePath, isBuiltIn);
            _references.Add(reference.Object);
            return this;
        }

        /// <summary>
        /// Builds the project, adds it to the VBE,
        /// and returns a <see cref="MockVbeBuilder"/>
        /// to continue adding projects to the VBE.
        /// </summary>
        /// <returns></returns>
        public MockVbeBuilder MockVbeBuilder()
        {
            _mockVbeBuilder.AddProject(Build());
            return _mockVbeBuilder;
        }

        /// <summary>
        /// Creates a <see cref="RubberduckTests.Mocks.MockUserFormBuilder"/> to build a new form component.
        /// </summary>
        /// <param name="name">The name of the component.</param>
        /// <param name="content">The VBA code associated to the component.</param>
        public MockUserFormBuilder MockUserFormBuilder(string name, string content)
        {
            var component = CreateComponentMock(name, vbext_ComponentType.vbext_ct_MSForm, content, new Selection());
            return new MockUserFormBuilder(component, this);
        }

        /// <summary>
        /// Gets the mock <see cref="VBProject"/> instance.
        /// </summary>
        public Mock<VBProject> Build()
        {
            return _project;
        }

        private Mock<VBProject> CreateProjectMock(string name, string filename, vbext_ProjectProtection protection)
        {
            var result = new Mock<VBProject>();

            result.SetupProperty(m => m.Name, name);
            result.SetupGet(m => m.FileName).Returns(() => filename);
            result.SetupGet(m => m.Protection).Returns(() => protection);
            result.SetupGet(m => m.VBE).Returns(_getVbe);

            return result;
        }

        private Mock<VBComponents> CreateComponentsMock()
        {
            var result = new Mock<VBComponents>();

            result.SetupGet(m => m.Parent).Returns(() => _project.Object);
            result.SetupGet(m => m.VBE).Returns(_getVbe);

            result.Setup(c => c.GetEnumerator()).Returns(() => Components.GetEnumerator());
            result.As<IEnumerable>().Setup(c => c.GetEnumerator()).Returns(() => Components.GetEnumerator());

            result.Setup(m => m.Item(It.IsAny<int>())).Returns<int>(index => Components.ElementAt(index));
            result.Setup(m => m.Item(It.IsAny<string>())).Returns<string>(name => Components.Single(item => item.Name == name));
            result.SetupGet(m => m.Count).Returns(Components.Count);

            result.Setup(m => m.Add(It.IsAny<vbext_ComponentType>())).Callback((vbext_ComponentType c) =>
            {
                _componentsMock.Add(CreateComponentMock("test", c, string.Empty, new Selection()));
            });

            result.Setup(m => m.Remove(It.IsAny<VBComponent>())).Callback((VBComponent c) =>
            {
                _componentsMock.Remove(_componentsMock.First(m => m.Object == c));
            });

            result.Setup(m => m.Import(It.IsAny<string>())).Callback((string s) =>
            {
                var parts = s.Split('.').ToList();
                var types = new Dictionary<string, vbext_ComponentType>
                {
                    {"bas", vbext_ComponentType.vbext_ct_StdModule},
                    {"cls", vbext_ComponentType.vbext_ct_ClassModule},
                    {"frm", vbext_ComponentType.vbext_ct_MSForm}
                };

                vbext_ComponentType type;
                types.TryGetValue(parts.Last(), out type);

                _componentsMock.Add(CreateComponentMock(s.Split('\\').Last(), type, string.Empty, new Selection()));
            });

            return result;
        }

        private Mock<References> CreateReferencesMock()
        {
            var result = new Mock<References>();
            result.SetupGet(m => m.Parent).Returns(() => _project.Object);
            result.SetupGet(m => m.VBE).Returns(_getVbe);
            result.Setup(m => m.GetEnumerator()).Returns(() => _references.GetEnumerator());
            result.As<IEnumerable>().Setup(m => m.GetEnumerator()).Returns(() => _references.GetEnumerator());
            result.Setup(m => m.Item(It.IsAny<int>())).Returns<int>(index => _references.ElementAt(index - 1));
            result.SetupGet(m => m.Count).Returns(() => _references.Count);
            result.Setup(m => m.AddFromFile(It.IsAny<string>()));
            return result;
        }

        private Mock<Reference> CreateReferenceMock(string name, string filePath, bool isBuiltIn = true)
        {
            var result = new Mock<Reference>();

            result.SetupGet(m => m.VBE).Returns(_getVbe);
            result.SetupGet(m => m.Collection).Returns(() => _vbReferences.Object);

            result.SetupGet(m => m.Name).Returns(() => name);
            result.SetupGet(m => m.FullPath).Returns(() => filePath);

            result.SetupGet(m => m.BuiltIn).Returns(isBuiltIn);

            return result;
        }

        private Mock<VBComponent> CreateComponentMock(string name, vbext_ComponentType type, string content, Selection selection)
        {
            var result = new Mock<VBComponent>();

            result.SetupGet(m => m.VBE).Returns(_getVbe);
            result.SetupGet(m => m.Collection).Returns(() => _vbComponents.Object);
            result.SetupGet(m => m.Type).Returns(() => type);
            result.SetupProperty(m => m.Name, name);

            var module = CreateCodeModuleMock(name, content, selection);
            module.SetupGet(m => m.Parent).Returns(() => result.Object);
            result.SetupGet(m => m.CodeModule).Returns(() => module.Object);

            result.Setup(m => m.Activate());

            return result;
        }

        private Mock<CodeModule> CreateCodeModuleMock(string name, string content, Selection selection)
        {
            var codePane = CreateCodePaneMock(name, selection);
            codePane.SetupGet(m => m.VBE).Returns(_getVbe);

            var result = CreateCodeModuleMock(content);
            result.SetupGet(m => m.VBE).Returns(_getVbe);
            result.SetupGet(m => m.CodePane).Returns(() => codePane.Object);
            result.SetupProperty(m => m.Name, name);

            codePane.SetupGet(m => m.CodeModule).Returns(() => result.Object);

            result.Setup(m => m.AddFromFile(It.IsAny<string>()));
            result.Setup(m => m.AddFromString(It.IsAny<string>()));
            return result;
        }

        private static readonly string[] ModuleBodyTokens =
        {
            Tokens.Sub, Tokens.Function, Tokens.Property
        };

        private Mock<CodeModule> CreateCodeModuleMock(string content)
        {
            var lines = content.Split(new[] { Environment.NewLine }, StringSplitOptions.None).ToList();

            var codeModule = new Mock<CodeModule>();
            codeModule.SetupGet(c => c.CountOfLines).Returns(() => lines.Count);
            codeModule.SetupGet(c => c.CountOfDeclarationLines).Returns(() =>
                lines.TakeWhile(line => !ModuleBodyTokens.Any(line.Contains)).Count());

            // ReSharper disable once UseIndexedProperty
            codeModule.Setup(m => m.get_Lines(It.IsAny<int>(), It.IsAny<int>()))
                .Returns<int, int>((start, count) => String.Join(Environment.NewLine, lines.Skip(start - 1).Take(count)));

            codeModule.Setup(m => m.ReplaceLine(It.IsAny<int>(), It.IsAny<string>()))
                .Callback<int, string>((index, str) => lines[index - 1] = str);

            codeModule.Setup(m => m.DeleteLines(It.IsAny<int>(), It.IsAny<int>()))
                .Callback<int, int>((index, count) => lines.RemoveRange(index - 1, count));

            codeModule.Setup(m => m.InsertLines(It.IsAny<int>(), It.IsAny<string>()))
                .Callback<int, string>((index, newLine) =>
                {
                    if (index - 1 >= lines.Count)
                    {
                        lines.Add(newLine);
                    }
                    else
                    {
                        lines.Insert(index - 1, newLine);
                    }
                });

            return codeModule;
        }

        private Mock<CodePane> CreateCodePaneMock(string name, Selection selection)
        {
            var windows = _getVbe().Windows as MockWindowsCollection;
            if (windows == null)
            {
                throw new InvalidOperationException("VBE.Windows collection must be a MockWindowsCollection object.");
            }

            var codePane = new Mock<CodePane>();
            var window = windows.CreateWindow(name);
            windows.Add(window);

            codePane.Setup(p => p.SetSelection(It.IsAny<int>(), It.IsAny<int>(), It.IsAny<int>(), It.IsAny<int>()));
            codePane.Setup(p => p.Show());

            var sLine = selection.StartLine;
            var sCol = selection.StartColumn;
            var eLine = selection.EndLine;
            var eCol = selection.EndColumn;

            codePane.Setup(p => p.GetSelection(out sLine, out sCol, out eLine, out eCol));

            codePane.SetupGet(p => p.VBE).Returns(_getVbe);
            codePane.SetupGet(p => p.Window).Returns(() => window);

            return codePane;
        }
    }
}
