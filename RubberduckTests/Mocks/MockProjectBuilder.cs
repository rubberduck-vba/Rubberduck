using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Vbe.Interop;
using Moq;

namespace RubberduckTests.Mocks
{
    public class MockProjectBuilder
    {
        private readonly Func<VBE> _getVbe;
        private readonly Mock<VBProject> _project;
        private readonly Mock<VBComponents> _vbComponents;
        private readonly Mock<References> _vbReferences;

        private readonly List<Mock<VBComponent>> _components = new List<Mock<VBComponent>>();
        private readonly List<Mock<Reference>> _references = new List<Mock<Reference>>(); 

        public MockProjectBuilder(string name, vbext_ProjectProtection protection, Func<VBE> getVbe)
        {
            _getVbe = getVbe;
            
            _project = CreateProjectMock(name, protection);

            _vbComponents = CreateComponentsMock();
            _project.SetupGet(m => m.VBComponents).Returns(_vbComponents.Object);
            
            _vbReferences = CreateReferencesMock();
            _project.SetupGet(m => m.References).Returns(_vbReferences.Object);
        }

        public MockProjectBuilder AddComponent(string name, vbext_ComponentType type, string content)
        {
            var component = CreateComponentMock(name, type, content);
            _components.Add(component);

            return this;
        }

        public MockProjectBuilder AddReference(string name, string filePath)
        {
            _references.Add(CreateReferenceMock(name, filePath));
            return this;
        }

        public Mock<VBProject> Build()
        {
            return _project;
        }

        private Mock<VBProject> CreateProjectMock(string name, vbext_ProjectProtection protection)
        {
            var result = new Mock<VBProject>();

            result.SetupProperty(m => m.Name, name);
            result.SetupGet(m => m.Protection).Returns(() => protection);
            result.SetupGet(m => m.VBE).Returns(_getVbe);

            return result;
        }

        private Mock<VBComponents> CreateComponentsMock()
        {
            var result = new Mock<VBComponents>();
            
            result.SetupGet(m => m.Parent).Returns(() => _project.Object);
            result.SetupGet(m => m.VBE).Returns(_getVbe);
            
            result.Setup(c => c.GetEnumerator()).Returns(() => _components.GetEnumerator());
            result.As<IEnumerable>().Setup(c => c.GetEnumerator()).Returns(() => _components.GetEnumerator());

            result.Setup(m => m.Item(It.IsAny<int>())).Returns<int>(index => _components[index].Object);
            result.Setup(m => m.Item(It.IsAny<string>())).Returns<string>(name => _components.Single(item => item.Object.Name == name).Object);

            return result;
        }

        private Mock<References> CreateReferencesMock()
        {
            var result = new Mock<References>();
            
            result.SetupGet(m => m.Parent).Returns(() => _project.Object);
            result.SetupGet(m => m.VBE).Returns(_getVbe);

            result.Setup(m => m.GetEnumerator()).Returns(() => _references.GetEnumerator());
            result.As<IEnumerable>().Setup(m => m.GetEnumerator()).Returns(() => _references.GetEnumerator());

            result.Setup(m => m.Item(It.IsAny<int>())).Returns<int>(index => _references[index].Object);

            return result;
        }

        private Mock<Reference> CreateReferenceMock(string name, string filePath)
        {
            var result = new Mock<Reference>();

            result.SetupGet(m => m.VBE).Returns(_getVbe);
            result.SetupGet(m => m.Collection).Returns(() => _vbReferences.Object);

            result.SetupGet(m => m.Name).Returns(() => name);
            result.SetupGet(m => m.FullPath).Returns(() => filePath);

            return result;
        }

        private Mock<VBComponent> CreateComponentMock(string name, vbext_ComponentType type, string content)
        {
            var result = new Mock<VBComponent>();

            result.SetupGet(m => m.VBE).Returns(_getVbe);
            result.SetupGet(m => m.Collection).Returns(() => _vbComponents.Object);
            result.SetupGet(m => m.Type).Returns(() => type);
            result.SetupProperty(m => m.Name, name);

            var module = CreateCodeModuleMock(name, content);
            result.SetupGet(m => m.CodeModule).Returns(() => module.Object);

            result.Setup(m => m.Activate());

            return result;
        }

        private Mock<CodeModule> CreateCodeModuleMock(string name, string content)
        {
            var codePane = CreateCodePaneMock(name);
            codePane.SetupGet(m => m.VBE).Returns(_getVbe);

            var result = CreateCodeModuleMock(content);
            result.SetupGet(m => m.VBE).Returns(_getVbe);
            result.SetupGet(m => m.CodePane).Returns(() => codePane.Object);

            codePane.SetupGet(m => m.CodeModule).Returns(() => result.Object);
            return result;
        }

        private Mock<CodeModule> CreateCodeModuleMock(string content)
        {
            var lines = content.Split(new[] { Environment.NewLine }, StringSplitOptions.None).ToList();

            var codeModule = new Mock<CodeModule>();
            codeModule.SetupGet(c => c.CountOfLines).Returns(() => lines.Count);

            // ReSharper disable once UseIndexedProperty
            codeModule.Setup(m => m.get_Lines(It.IsAny<int>(), It.IsAny<int>()))
                .Returns<int, int>((start, count) => String.Join(Environment.NewLine, lines.Skip(start - 1).Take(count)));

            codeModule.Setup(m => m.ReplaceLine(It.IsAny<int>(), It.IsAny<string>()))
                .Callback<int, string>((index, str) => lines[index - 1] = str);

            codeModule.Setup(m => m.DeleteLines(It.IsAny<int>(), It.IsAny<int>()))
                .Callback<int, int>((index, count) => lines.RemoveRange(index - 1, count));

            codeModule.Setup(m => m.InsertLines(It.IsAny<int>(), It.IsAny<string>()))
                .Callback<int, string>((index, newLine) => lines.Insert(index - 1, newLine));

            return codeModule;
        }

        private Mock<CodePane> CreateCodePaneMock(string name)
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

            codePane.SetupGet(p => p.VBE).Returns(_getVbe);
            codePane.SetupGet(p => p.Window).Returns(() => window);
            
            return codePane;
        }
    }
}