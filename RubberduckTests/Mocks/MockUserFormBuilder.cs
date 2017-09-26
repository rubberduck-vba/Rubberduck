using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using Moq;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.SafeComWrappers.Office.Core.Abstract;

namespace RubberduckTests.Mocks
{
    /// <summary>
    /// Builds a mock UserForm component.
    /// </summary>
    public class MockUserFormBuilder
    {
        private readonly Mock<IVBComponent> _component;
        private readonly MockProjectBuilder _mockProjectBuilder;
        private readonly Mock<IControls> _vbControls;
        private readonly ICollection<IControl> _controls = new List<IControl>();

        public MockUserFormBuilder(Mock<IVBComponent> component, MockProjectBuilder mockProjectBuilder)
        {
            if (component.Object.Type != ComponentType.UserForm)
            {
                throw new InvalidOperationException("Component type must be 'ComponentType.UserForm'.");
            }

            _component = component;
            _mockProjectBuilder = mockProjectBuilder;
            _vbControls = CreateControlsMock();
        }

        /// <summary>
        /// Adds a <see cref="IControl"/> to the form.
        /// </summary>
        /// <param name="name">The name of the control to add.</param>
        /// <returns></returns>
        public MockUserFormBuilder AddControl(string name)
        {
            var control = new Mock<IControl>();
            control.SetupProperty(m => m.Name, name);

            _controls.Add(control.Object);
            return this;
        }

        /// <summary>
        /// Builds the UserForm, adds it to the project,
        /// and returns a <see cref="MockProjectBuilder"/>
        /// to continue adding components to the project.
        /// </summary>
        /// <returns></returns>
        public MockProjectBuilder MockProjectBuilder()
        {
            _mockProjectBuilder.AddComponent(Build());
            return _mockProjectBuilder;
        }

        /// <summary>
        /// Gets the mock UserForm component.
        /// </summary>
        /// <returns></returns>
        public Mock<IVBComponent> Build()
        {
            //var designer = CreateMockDesigner();
            //_component.SetupGet(m => m.Designer).Returns(() => designer.Object);

            var window = new Mock<IWindow>();
            window.SetupProperty(w => w.IsVisible, false);
            _component.Setup(m => m.Controls).Returns(_vbControls.Object);
            _component.Setup(m => m.DesignerWindow()).Returns(window.Object);

            return _component;
        }

        //private Mock<UserForm> CreateMockDesigner()
        //{
        //    var result = new Mock<UserForm>();

        //    result.SetupGet(m => m.Controls).Returns(() => _vbControls.Object);

        //    return result;
        //}

        private Mock<IControls> CreateControlsMock()
        {
            var result = new Mock<IControls>();
            result.Setup(m => m.GetEnumerator()).Returns(() => _controls.GetEnumerator());
            result.As<IEnumerable>().Setup(m => m.GetEnumerator()).Returns(() => _controls.GetEnumerator());

            result.Setup(m => m[It.IsAny<int>()]).Returns<int>(index => _controls.ElementAt(index));
            result.SetupGet(m => m.Count).Returns(_controls.Count);
            return result;
        }
    }
}
