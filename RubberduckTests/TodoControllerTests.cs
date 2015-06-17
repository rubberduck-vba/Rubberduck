using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Settings;
using Rubberduck.UI.Settings;

namespace RubberduckTests
{
    [TestClass]
    public class TodoControllerTests
    {
        private static List<ToDoMarker> GetTestMarkers()
        {
            var markers = new List<ToDoMarker>
            {
                new ToDoMarker("Todo:", TodoPriority.Medium),
                new ToDoMarker("Note:", TodoPriority.Low),
                new ToDoMarker("Bug:", TodoPriority.High)
            };
            return markers;
        }

        [TestMethod]
        public void ConstructorWorks()
        {
            //arrange
            var view = new Mock<ITodoSettingsView>();
            var addTodoMarkerView = new Mock<IAddTodoMarkerView>();

            //act
            var controller = new TodoSettingPresenter(view.Object, addTodoMarkerView.Object);

            //assert
            Assert.IsNotNull(controller);
        }

        [TestMethod]
        public void ViewTextIsNotNullOrEmptyAfterControllerConstruction()
        {
            //arrange
            var markers = new List<ToDoMarker> {new ToDoMarker("Todo:", TodoPriority.Medium)};

            var view = new TodoListSettingsUserControl(markers);
            var addTodoMarkerView = new Mock<IAddTodoMarkerView>().Object;
            
            //act
            var controller = new TodoSettingPresenter(view, addTodoMarkerView);

            //assert
            Assert.AreEqual("Todo:", view.ActiveMarkerText);
        }

        [TestMethod]
        public void SetActiveItemChangesViewSelectedIndex()
        {
            //arrange
            var markers = GetTestMarkers();

            var view = new TodoListSettingsUserControl(markers);
            var addTodoMarkerView = new Mock<IAddTodoMarkerView>().Object;

            var controller = new TodoSettingPresenter(view, addTodoMarkerView);

            //act
            controller.SetActiveItem(1);

            Assert.AreEqual(1, view.SelectedIndex);

        }

        [TestMethod]
        public void ViewPriorityMatchesAfterSelectionChange()
        {
            var markers = new List<ToDoMarker>
            {
                new ToDoMarker("Todo:", TodoPriority.Medium),
                new ToDoMarker("Note:", TodoPriority.Low),
                new ToDoMarker("Bug:", TodoPriority.High)
            };

            var view = new TodoListSettingsUserControl(markers);
            var addTodoMarkerView = new Mock<IAddTodoMarkerView>().Object;
            var controller = new TodoSettingPresenter(view, addTodoMarkerView);

            //act
            controller.SetActiveItem(2);

            Assert.AreEqual(TodoPriority.High, view.ActiveMarkerPriority);
        }

        [TestMethod]
        public void ViewTextMatchesAfterSelectionChange()
        {
            var markers = new List<ToDoMarker>
            {
                new ToDoMarker("Todo:", TodoPriority.Medium),
                new ToDoMarker("Note:", TodoPriority.Low)
            };

            var view = new TodoListSettingsUserControl(markers);
            var addTodoMarkerView = new Mock<IAddTodoMarkerView>().Object;
            var controller = new TodoSettingPresenter(view, addTodoMarkerView);

            //act
            controller.SetActiveItem(1);

            Assert.AreEqual("Note:", view.ActiveMarkerText);
        }
    }
}
