using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Vbe.Interop;
using Microsoft.Vbe.Interop.Forms;
using Moq;

namespace RubberduckTests.Mocks
{
    /// <summary>
    /// Builds a mock <see cref="UserForm"/> component.
    /// </summary>
    public class MockUserFormBuilder
    {
        private readonly Mock<VBComponent> _component;
        private readonly MockProjectBuilder _mockProjectBuilder;
        private readonly Mock<Controls> _vbControls;
        private readonly ICollection<Control> _controls = new List<Control>();

        public MockUserFormBuilder(Mock<VBComponent> component, MockProjectBuilder mockProjectBuilder)
        {
            if (component.Object.Type != vbext_ComponentType.vbext_ct_MSForm)
            {
                throw new InvalidOperationException("Component type must be 'vbext_ComponentType.vbext_ct_MSForm'.");
            }

            _component = component;
            _mockProjectBuilder = mockProjectBuilder;
            _vbControls = CreateControlsMock();
        }

        /// <summary>
        /// Adds a <see cref="Control"/> to the form.
        /// </summary>
        /// <param name="name">The name of the control to add.</param>
        /// <returns></returns>
        public MockUserFormBuilder AddControl(string name)
        {
            var control = new Mock<Control>();
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
        /// Gets the mock <see cref="UserForm"/> component.
        /// </summary>
        /// <returns></returns>
        public Mock<VBComponent> Build()
        {
            var designer = CreateMockDesigner();
            _component.SetupGet(m => m.Designer).Returns(() => designer.Object);

            var window = new Mock<Window>();
            window.SetupProperty(w => w.Visible, false);
            _component.Setup(m => m.DesignerWindow()).Returns(window.Object);

            return _component;
        }

        private Mock<UserForm> CreateMockDesigner()
        {
            var result = new Mock<UserForm>();

            result.SetupGet(m => m.Controls).Returns(() => _vbControls.Object);

            return result;
        }

        private Mock<Controls> CreateControlsMock()
        {
            var result = new Mock<Controls>();
            result.Setup(m => m.GetEnumerator()).Returns(() => _controls.GetEnumerator());
            result.As<IEnumerable>().Setup(m => m.GetEnumerator()).Returns(() => _controls.GetEnumerator());

            result.Setup(m => m.Item(It.IsAny<int>())).Returns<int>(index => _controls.ElementAt(index));
            result.SetupGet(m => m.Count).Returns(_controls.Count);
            return result;
        }
    }
}