using System;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Settings;
using Rubberduck.UI;
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

            var view = new TodoListSettingsUserControl(markers, new Mock<GridViewSort<ToDoMarker>>("", false).Object);
            var addTodoMarkerView = new Mock<IAddTodoMarkerView>().Object;

            //assert
            Assert.AreEqual("Todo:", view.ActiveMarkerText);
        }

        [TestMethod]
        public void SetActiveItemChangesViewSelectedIndex()
        {
            //arrange
            var markers = GetTestMarkers();

            var view = new TodoListSettingsUserControl(markers, new Mock<GridViewSort<ToDoMarker>>("", false).Object);

            //act
            view.SelectedIndex = 1;

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

            var view = new TodoListSettingsUserControl(markers, new Mock<GridViewSort<ToDoMarker>>("", false).Object);

            //act
            view.SelectedIndex = 2;

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

            var view = new TodoListSettingsUserControl(markers, new Mock<GridViewSort<ToDoMarker>>("", false).Object);

            //act
            view.SelectedIndex = 1;

            Assert.AreEqual("Note:", view.ActiveMarkerText);
        }

        [TestMethod]
        public void MarkerChangeSavedOnPriorityChanged()
        {
            var markers = GetTestMarkers();

            var view = new TodoListSettingsUserControl(markers, new Mock<GridViewSort<ToDoMarker>>("", false).Object);

            view.ActiveMarkerPriority = TodoPriority.High;

            Assert.AreEqual(view.ActiveMarkerPriority, view.TodoMarkers[0].Priority);
        }

        [TestMethod]
        public void RemoveReallyDoesRemoveSelectedItem()
        {
            var markers = GetTestMarkers();

            var view = new TodoListSettingsUserControl(markers, new Mock<GridViewSort<ToDoMarker>>("", false).Object);
            view.RemoveMarker += RemoveParam;
            view.SelectedIndex = 2;

            RemoveParam(null, EventArgs.Empty);

            Assert.AreEqual(2, view.TodoMarkers.Count);
        }

        private void RemoveParam(object sender, EventArgs e)
        {
            
        }
    }
}
