using Moq;
using Rubberduck.UI;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace RubberduckTests.Mocks
{
    static class MockFactory
    {
        internal static Mock<IWindow> CreateWindowMock()
        {
            var window = new Mock<IWindow>();
            window.SetupProperty(w => w.IsVisible, false);
            window.SetupGet(w => w.LinkedWindows).Returns((ILinkedWindows) null);
            window.SetupProperty(w => w.Height);
            window.SetupProperty(w => w.Width);

            return window;
        }

        /// <summary>
        /// Creates a mock <see cref="IWindow"/> with it's <see cref="IWindow.Caption"/> propery set up.
        /// </summary>
        /// <param name="caption">The value to return from <see cref="IWindow.Caption"/>.</param>
        /// <returns>
        /// A <see cref="Mock{Window}"/>that has all the properties needed for <see cref="DockableToolwindowPresenter"/> pre-setup.
        /// </returns>
        internal static Mock<IWindow> CreateWindowMock(string caption)
        {
            var window = CreateWindowMock();
            window.SetupGet(w => w.Caption).Returns(caption);

            return window;
        }

        internal static Mock<IVBE> CreateVbeMock(Windows windows)
        {
            var vbe = new Mock<IVBE>();
            windows.VBE = vbe.Object;
            vbe.Setup(m => m.Windows).Returns(windows);
            vbe.SetupProperty(m => m.ActiveCodePane);
            vbe.SetupProperty(m => m.ActiveVBProject);
            vbe.SetupGet(m => m.SelectedVBComponent).Returns(() => vbe.Object.ActiveCodePane.CodeModule.Parent);
            vbe.SetupGet(m => m.ActiveWindow).Returns(() => vbe.Object.ActiveCodePane.Window);

            //setting up a main window lets the native window functions fun
            var mainWindow = new Mock<IWindow>();
            mainWindow.Setup(m => m.HWnd).Returns(0);

            vbe.SetupGet(m => m.MainWindow).Returns(mainWindow.Object);

            return vbe;
        }
    }
}
