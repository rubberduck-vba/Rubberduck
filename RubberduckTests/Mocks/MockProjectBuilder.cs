using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Moq;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace RubberduckTests.Mocks
{
    /// <summary>
    /// Builds a mock <see cref="IVBProject"/>.
    /// </summary>
    public class MockProjectBuilder
    {
        private readonly Func<IVBE> _getVbe;
        private readonly MockVbeBuilder _mockVbeBuilder;
        private readonly Mock<IVBProject> _project;
        private readonly Mock<IVBComponents> _vbComponents;
        private readonly Mock<IReferences> _vbReferences;
        private readonly ProjectType _projectType;

        private readonly List<Mock<IVBComponent>> _componentsMock = new List<Mock<IVBComponent>>();
        private readonly List<Mock<ICodeModule>> _codeModuleMocks = new List<Mock<ICodeModule>>();
        private readonly List<IReference> _references = new List<IReference>();

        public Mock<IVBComponents> MockVBComponents => _vbComponents;

        public List<Mock<IVBComponent>> MockComponents => _componentsMock;
        public List<Mock<ICodeModule>> MockCodeModules => _codeModuleMocks;

        private List<IVBComponent> Components
        {
            get { return _componentsMock.Select(m => m.Object).ToList(); }
        }

        public void RemoveComponent(Mock<IVBComponent> component)
        {
            _componentsMock.Remove(component);
        }

        public MockProjectBuilder(string name, string filename, ProjectProtection protection, ProjectType projectType, Func<IVBE> getVbe, MockVbeBuilder mockVbeBuilder)
        : this(
            name,
            filename,
            Guid.NewGuid().ToString(),
            protection,
            projectType,
            getVbe,
            mockVbeBuilder
            )
        { }

        public MockProjectBuilder(string name, string filename, string projectId, ProjectProtection protection, ProjectType projectType, Func<IVBE> getVbe, MockVbeBuilder mockVbeBuilder)
        {
            _getVbe = getVbe;
            _mockVbeBuilder = mockVbeBuilder;
            _projectType = projectType;

            _project = CreateProjectMock(name, filename, protection);

            _project.SetupProperty(m => m.HelpFile);
            _project.SetupGet(m => m.ProjectId).Returns(() => _project.Object.HelpFile);
            _project.SetupGet(m => m.Type).Returns(_projectType);
            _project.Setup(m => m.AssignProjectId())
                .Callback(() => _project.Object.HelpFile = projectId);

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
        /// <param name="properties">A collection of properties that will be available when calling the component's Properties property.</param>
        /// <returns>Returns the <see cref="MockProjectBuilder"/> instance.</returns>
        public MockProjectBuilder AddComponent(string name, ComponentType type, string content, Selection selection = new Selection(), IEnumerable<IProperty> properties = null)
        {
            var component = CreateComponentMock(name, type, content, selection, properties, out var codeModule);
            return AddComponent(component, codeModule);
        }

        /// <summary>
        /// Adds a new mock component to the project.
        /// </summary>
        /// <param name="component">The component to add.</param>
        /// <param name="codeModule">The codeModule of the component to add.</param>
        /// <returns>Returns the <see cref="MockProjectBuilder"/> instance.</returns>
        public MockProjectBuilder AddComponent(Mock<IVBComponent> component, Mock<ICodeModule> codeModule)
        {
            _componentsMock.Add(component);
            _codeModuleMocks.Add(codeModule);
            _getVbe().ActiveCodePane = component.Object.CodeModule.CodePane;
            return this;
        }

        /// <summary>
        /// Adds a mock reference to the project.
        /// </summary>
        /// <param name="referenceLibrary">The reference library's enum.</param>
        /// <returns>Returns the <see cref="MockProjectBuilder"/> instance.</returns>
        public MockProjectBuilder AddReference(ReferenceLibrary referenceLibrary)
        {
            var (name, path, versionMajor, versionMinor, isBuiltIn) = MockVbeBuilder.ReferenceLibraries[referenceLibrary];
            return AddReference(name, path, versionMajor, versionMinor, isBuiltIn);
        }

        /// <summary>
        /// Adds a mock reference to the project.
        /// </summary>
        /// <param name="name">The name of the referenced library.</param>
        /// <param name="filePath">The path to the referenced library.</param>
        /// <param name="isBuiltIn">Indicates whether the reference is a built-in reference.</param>
        /// <returns>Returns the <see cref="MockProjectBuilder"/> instance.</returns>
        public MockProjectBuilder AddReference(string name, string filePath, int major, int minor, bool isBuiltIn = false, ReferenceKind type = ReferenceKind.TypeLibrary)
        {
            var reference = CreateReferenceMock(name, filePath, major, minor, isBuiltIn, type);
            _references.Add(reference.Object);
            return this;
        }

        /// <summary>
        /// Builds the project, adds it to the VBE,
        /// and returns a <see cref="MockVbeBuilder"/>
        /// to continue adding projects to the VBE.
        /// </summary>
        /// <returns></returns>
        public MockVbeBuilder AddProjectToVbeBuilder()
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
            var component = CreateComponentMock(name, ComponentType.UserForm, content, new Selection(), null, out var codeModule);
            return new MockUserFormBuilder(component, codeModule, this);
        }

        /// <summary>
        /// Gets the mock <see cref="IVBProject"/> instance.
        /// </summary>
        public Mock<IVBProject> Build()
        {
            return _project;
        }

        /// <summary>
        /// Gets the mock <see cref="IVBProject"/> instance after assigning the projectId.
        /// </summary>
        public Mock<IVBProject> BuildWithAssignedProjectId()
        {
            _project.Object.AssignProjectId();
            return _project;
        }

        private Mock<IVBProject> CreateProjectMock(string name, string filename, ProjectProtection protection)
        {
            var result = new Mock<IVBProject>();

            result.Setup(m => m.Dispose());
            result.SetupReferenceEqualityIncludingHashCode();
            result.Setup(m => m.Equals(It.IsAny<IVBProject>()))
                .Returns((IVBProject other) => ReferenceEquals(result.Object, other));
            result.SetupProperty(m => m.Name, name);
            result.SetupGet(m => m.FileName).Returns(() => filename);
            result.SetupGet(m => m.Protection).Returns(() => protection);
            result.SetupGet(m => m.VBE).Returns(_getVbe);
            result.Setup(m => m.ComponentNames()).Returns(() => _vbComponents.Object.Select(component => component.Name).ToArray());

            return result;
        }

        private Mock<IVBComponents> CreateComponentsMock()
        {
            var result = new Mock<IVBComponents>();

            result.Setup(m => m.Dispose());
            result.SetupReferenceEqualityIncludingHashCode();

            result.SetupGet(m => m.Parent).Returns(() => _project.Object);
            result.SetupGet(m => m.VBE).Returns(_getVbe);

            result.Setup(c => c.GetEnumerator()).Returns(() => Components.GetEnumerator());
            result.As<IEnumerable>().Setup(c => c.GetEnumerator()).Returns(() => Components.GetEnumerator());

            result.Setup(m => m[It.IsAny<int>()]).Returns<int>(index => Components.ElementAt(index));
            result.Setup(m => m[It.IsAny<string>()]).Returns<string>(name => Components.Single(item => item.Name == name));
            result.SetupGet(m => m.Count).Returns(() => Components.Count);

            result.Setup(m => m.Add(It.IsAny<ComponentType>()))
                .Callback((ComponentType c) =>
                {
                    _componentsMock.Add(CreateComponentMock("test", c, string.Empty, new Selection(), null, out var codeModule));
                    _codeModuleMocks.Add(codeModule);
                })
                .Returns(() =>
                {
                    var lastComponent = _componentsMock.LastOrDefault();
                    return lastComponent?.Object;
                });

            result.Setup(m => m.Remove(It.IsAny<IVBComponent>())).Callback((IVBComponent c) =>
            {
                _componentsMock.Remove(_componentsMock.First(m => m.Object.QualifiedModuleName == c.QualifiedModuleName));
                _codeModuleMocks.Remove(_codeModuleMocks.First(m => m.Object.Parent.QualifiedModuleName == c.QualifiedModuleName));
            });

            result.Setup(m => m.Import(It.IsAny<string>())).Callback((string s) =>
            {
                var extension = Path.GetExtension(s);
                var types = new Dictionary<string, ComponentType>
                {
                    {".bas", ComponentType.StandardModule},
                    {".cls", ComponentType.ClassModule},
                    {".frm", ComponentType.UserForm}
                };

                ComponentType type;
                types.TryGetValue(extension, out type);

                _componentsMock.Add(CreateComponentMock(Path.GetFileNameWithoutExtension(s), type, string.Empty, new Selection(), null, out var codeModule));
                _codeModuleMocks.Add(codeModule);
            });

            return result;
        }

        public Mock<IReferences> GetMockedReferences(out Mock<IVBProject> project)
        {
            project = Build();
            return _vbReferences;
        }

        private Mock<IReferences> CreateReferencesMock()
        {
            var result = new Mock<IReferences>();
            result.Setup(m => m.Dispose());
            result.SetupReferenceEqualityIncludingHashCode();
            result.SetupGet(m => m.Parent).Returns(() => _project.Object);
            result.SetupGet(m => m.VBE).Returns(_getVbe);
            result.Setup(m => m.GetEnumerator()).Returns(() => _references.GetEnumerator());
            result.As<IEnumerable>().Setup(m => m.GetEnumerator()).Returns(() => _references.GetEnumerator());
            result.Setup(m => m[It.IsAny<int>()]).Returns<int>(index => _references.ElementAt(index - 1));
            result.SetupGet(m => m.Count).Returns(() => _references.Count);
            result.Setup(m => m.AddFromFile(It.IsAny<string>()));
            return result;
        }

        public Mock<IReference> CreateReferenceMock(string name, string filePath, int major, int minor, bool isBuiltIn = true, ReferenceKind type = ReferenceKind.TypeLibrary)
        {
            var result = new Mock<IReference>();

            result.Setup(m => m.Dispose());
            result.SetupReferenceEqualityIncludingHashCode();

            result.SetupGet(m => m.VBE).Returns(_getVbe);
            result.SetupGet(m => m.Collection).Returns(() => _vbReferences.Object);

            result.SetupGet(m => m.Name).Returns(() => name);
            result.SetupGet(m => m.FullPath).Returns(() => filePath);
            result.SetupGet(m => m.Major).Returns(() => major);
            result.SetupGet(m => m.Minor).Returns(() => minor);
            result.SetupGet(m => m.Type).Returns(() => type);

            result.SetupGet(m => m.IsBuiltIn).Returns(isBuiltIn);

            return result;
        }

        private Mock<IVBComponent> CreateComponentMock(string name, ComponentType type, string content, Selection selection,
            IEnumerable<IProperty> properties, out Mock<ICodeModule> moduleMock)
        {
            var result = new Mock<IVBComponent>();

            result.Setup(m => m.Dispose());
            result.SetupReferenceEqualityIncludingHashCode();
            result.Setup(m => m.Equals(It.IsAny<IVBComponent>()))
                .Returns((IVBComponent other) => ReferenceEquals(result.Object, other));

            result.SetupGet(m => m.VBE).Returns(_getVbe);
            result.SetupGet(m => m.Collection).Returns(() => _vbComponents.Object);
            result.SetupGet(m => m.Type).Returns(() => type);
            result.SetupGet(m => m.HasCodeModule).Returns(true);
            if (type == ComponentType.UserForm)
            {
                result.Setup(m => m.HasDesigner).Returns(true);
            }
            result.SetupProperty(m => m.Name, name);
            result.SetupGet(m => m.QualifiedModuleName).Returns(() => new QualifiedModuleName(result.Object));
            result.SetupGet(m => m.QualifiedModuleName).Returns(() => new QualifiedModuleName(result.Object));

            var propertiesMock = new Mock<IProperties>();
            propertiesMock.Setup(m => m.GetEnumerator()).Returns(() => properties?.GetEnumerator());
            propertiesMock.SetupGet(m => m.Count).Returns(properties?.Count() ?? 0);
            propertiesMock.Setup(m => m[It.IsAny<int>()]).Returns<int>(index => properties.ElementAt(index));
            result.SetupGet(m => m.Properties).Returns(propertiesMock.Object);

            var module = CreateCodeModuleMock(name, content, selection, result);
            result.SetupGet(m => m.CodeModule).Returns(() => module.Object);
            // Note that this setup does not account for hashing behavior of designers. See https://github.com/rubberduck-vba/Rubberduck/issues/3387
            result.Setup(m => m.ContentHash()).Returns(() => result.Object.CodeModule.ContentHash());

            result.Setup(m => m.Activate());

            moduleMock = module;
            return result;
        }

        private Mock<ICodeModule> CreateCodeModuleMock(string name, string content, Selection selection, Mock<IVBComponent> component)
        {
            var codePane = CreateCodePaneMock(name, selection, component);

            var result = CreateCodeModuleMock(content);
            result.SetupReferenceEqualityIncludingHashCode();
            result.Setup(m => m.Equals(It.IsAny<ICodeModule>()))
                .Returns((ICodeModule other) => ReferenceEquals(result.Object, other));
            result.SetupGet(m => m.VBE).Returns(_getVbe);
            result.SetupGet(m => m.Parent).Returns(() => component.Object);
            result.SetupGet(m => m.CodePane).Returns(() => codePane.Object);
            result.SetupGet(m => m.QualifiedModuleName).Returns(() => new QualifiedModuleName(component.Object));
            result.Setup(m => m.AddFromFile(It.IsAny<string>()));

            result.SetupGet(m => m.Name).Returns(() => component.Object.Name);
            result.SetupSet(m => m.Name = It.IsAny<string>()).Callback<string>(value => component.Object.Name = value);

            codePane.SetupGet(m => m.CodeModule).Returns(() => result.Object);
            return result;
        }

        private static readonly string[] ModuleBodyTokens =
        {
            Tokens.Sub + ' ', Tokens.Function + ' ', Tokens.Property + ' '
        };

        private Mock<ICodeModule> CreateCodeModuleMock(string content)
        {
            var lines = content.Split(new[] { Environment.NewLine }, StringSplitOptions.None).ToList();

            var codeModule = new Mock<ICodeModule>();
            codeModule.Setup(m => m.Dispose());
            codeModule.Setup(m => m.Clear()).Callback(() => lines = new List<string>());
            codeModule.SetupGet(c => c.CountOfLines).Returns(() => lines.Count);
            codeModule.SetupGet(c => c.CountOfDeclarationLines).Returns(() =>
                lines.TakeWhile(line => line.Contains(Tokens.Declare + ' ') || !ModuleBodyTokens.Any(line.Contains)).Count());

            codeModule.Setup(m => m.Content()).Returns(() => string.Join(Environment.NewLine, lines));

            codeModule.Setup(m => m.ContentHash()).Returns(() => string.IsNullOrEmpty(codeModule.Object.Content())
                ? 0
                : codeModule.Object.Content().GetHashCode());

            codeModule.Setup(m => m.GetLines(It.IsAny<Selection>()))
                .Returns((Selection selection) => string.Join(Environment.NewLine, lines.Skip(selection.StartLine - 1).Take(selection.LineCount)));

            codeModule.Setup(m => m.GetLines(It.IsAny<int>(), It.IsAny<int>()))
                .Returns<int, int>((start, count) => string.Join(Environment.NewLine, lines.Skip(start - 1).Take(count)));

            codeModule.Setup(m => m.ReplaceLine(It.IsAny<int>(), It.IsAny<string>()))
                .Callback<int, string>((index, str) => lines[index - 1] = str);

            codeModule.Setup(m => m.DeleteLines(It.IsAny<Selection>()))
                .Callback<Selection>(selection => lines.RemoveRange(selection.StartLine - 1, selection.LineCount));

            codeModule.Setup(m => m.DeleteLines(It.IsAny<int>(), It.IsAny<int>()))
                .Callback<int, int>((index, count) => lines.RemoveRange(index - 1, count));

            codeModule.Setup(m => m.InsertLines(It.IsAny<int>(), It.IsAny<string>()))
                .Callback<int, string>((index, newLine) =>
                {
                    if (index - 1 >= lines.Count)
                    {
                        lines.AddRange(newLine.Split(new[] { Environment.NewLine }, StringSplitOptions.None));
                    }
                    else
                    {
                        lines.InsertRange(index - 1, newLine.Split(new[] { Environment.NewLine }, StringSplitOptions.None));
                    }
                });

            codeModule.Setup(m => m.AddFromString(It.IsAny<string>()))
                .Callback<string>(newLine =>
                {
                    lines.AddRange(newLine.Split(new[] { Environment.NewLine }, StringSplitOptions.None));
                });

            codeModule.Setup(m => m.Equals(It.IsAny<ICodeModule>()))
                .Returns((ICodeModule other) => codeModule.Object.Name.Equals(other.Name) && codeModule.Object.Content().Equals(other.Content()));
            codeModule.Setup(m => m.GetHashCode()).Returns(() => codeModule.Object.Target.GetHashCode());

            return codeModule;
        }

        private Mock<ICodePane> CreateCodePaneMock(string name, Selection selection, Mock<IVBComponent> component)
        {
            var windows = _getVbe().Windows as Windows;
            if (windows == null)
            {
                throw new InvalidOperationException("VBE.Windows collection must be a MockWindowsCollection object.");
            }

            var codePane = new Mock<ICodePane>();
            var window = windows.CreateWindow(name);
            windows.Add(window);

            codePane.Setup(m => m.Dispose());
            codePane.SetupReferenceEqualityIncludingHashCode();
            codePane.Setup(p => p.GetQualifiedSelection()).Returns(() => {
                if (selection.IsEmpty()) { return null; }
                return new QualifiedSelection(new QualifiedModuleName(component.Object), codePane.Object.Selection);
            });
            codePane.SetupProperty(p => p.Selection, selection);
            codePane.Setup(p => p.Show());

            codePane.SetupGet(p => p.VBE).Returns(_getVbe);
            codePane.SetupGet(p => p.Window).Returns(() => window);
            codePane.SetupGet(m => m.QualifiedModuleName).Returns(() => new QualifiedModuleName(component.Object));

            return codePane;
        }
    }
}