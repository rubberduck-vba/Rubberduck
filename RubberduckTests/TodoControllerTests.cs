using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck;
using Rubberduck.Config;
using Rubberduck.ToDoItems;
using Moq;
using Rubberduck.UI.Settings;
using System.Collections.Generic;
using RubberduckTests.Mocks;

namespace RubberduckTests
{
    [TestClass]
    public class TodoControllerTests
    {
        private static List<ToDoMarker> GetTestMarkers()
        {
            var markers = new List<ToDoMarker>();
            markers.Add(new ToDoMarker("Todo:", TodoPriority.Normal));
            markers.Add(new ToDoMarker("Note:", TodoPriority.Low));
            markers.Add(new ToDoMarker("Bug:", TodoPriority.High));
            return markers;
        }

        [TestMethod]
        public void ConstructorWorks()
        {
            //arrange
            Mock<ITodoSettingsView> view = new Mock<ITodoSettingsView>();

            //act
            var controller = new TodoSettingController(view.Object);

            //assert
            Assert.IsNotNull(controller);
        }

        [TestMethod]
        public void ViewTextIsNotNullOrEmptyAfterControllerConstruction()
        {
            //arrange
            var markers = new List<ToDoMarker>();
            markers.Add(new ToDoMarker("Todo:", TodoPriority.Normal));

            ITodoSettingsView view = new MockTodoSettingsView(markers);
            
            //act
            var controller = new TodoSettingController(view);

            //assert
            Assert.AreEqual("Todo:", view.ActiveMarkerText);

        }

        [TestMethod]
        public void SetActiveItemChangesViewSelectedIndex()
        {
            //arrange
            var markers = GetTestMarkers();

            ITodoSettingsView view = new MockTodoSettingsView(markers);

            var controller = new TodoSettingController(view);

            //act
            controller.SetActiveItem(1);

            Assert.AreEqual(1, view.SelectedIndex);

        }

        [TestMethod]
        public void SetActiveItemChangesActiveMarker()
        {
            //arrange
            var markers = GetTestMarkers();

            ITodoSettingsView view = new MockTodoSettingsView(markers);

            var controller = new TodoSettingController(view);

            //act
            controller.SetActiveItem(1);

            Assert.AreEqual(markers[1], controller.ActiveMarker);
        }

        [TestMethod]
        public void ViewPriorityMatchesAfterSelectionChange()
        {
            var markers = new List<ToDoMarker>();
            markers.Add(new ToDoMarker("Todo:", TodoPriority.Normal));
            markers.Add(new ToDoMarker("Note:", TodoPriority.Low));
            markers.Add(new ToDoMarker("Bug:", TodoPriority.High));

            ITodoSettingsView view = new MockTodoSettingsView(markers);
            var controller = new TodoSettingController(view);

            //act
            controller.SetActiveItem(2);

            Assert.AreEqual(TodoPriority.High, view.ActiveMarkerPriority);
        }

        [TestMethod]
        public void ViewTextMatchesAfterSelectionChange()
        {
            var markers = new List<ToDoMarker>();
            markers.Add(new ToDoMarker("Todo:", TodoPriority.Normal));
            markers.Add(new ToDoMarker("Note:", TodoPriority.Low));

            ITodoSettingsView view = new MockTodoSettingsView(markers);
            var controller = new TodoSettingController(view);

            //act
            controller.SetActiveItem(1);

            Assert.AreEqual("Note:", view.ActiveMarkerText);
        }

        [TestMethod]
        public void SaveEnabledAfterTextChange()
        {
            var markers = GetTestMarkers();

            ITodoSettingsView view = new MockTodoSettingsView(markers);
            var controller = new TodoSettingController(view);

            view.ActiveMarkerText = "SomeNewText";

            Assert.IsTrue(view.SaveEnabled);
        }

        [TestMethod]
        public void SaveEnabledAfterPriorityChange()
        {
            var markers = GetTestMarkers();

            ITodoSettingsView view = new MockTodoSettingsView(markers);
            var controller = new TodoSettingController(view);

            view.ActiveMarkerPriority = TodoPriority.High;

            Assert.IsTrue(view.SaveEnabled);
        }

        [TestMethod]
        public void SaveDisabledAfterSelectionChange()
        {
            var markers = GetTestMarkers();

            ITodoSettingsView view = new MockTodoSettingsView(markers);
            var controller = new TodoSettingController(view);

            view.SelectedIndex = 2;

            Assert.IsFalse(view.SaveEnabled);
        }

    }
}
