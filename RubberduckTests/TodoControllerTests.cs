using System;
using System.Collections.Generic;
using System.ComponentModel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Settings;
using Rubberduck.ToDoItems;
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
                new ToDoMarker("Todo", TodoPriority.Medium),
                new ToDoMarker("Note", TodoPriority.Low),
                new ToDoMarker("Bug", TodoPriority.High)
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
            var markers = new List<ToDoMarker> {new ToDoMarker("Todo", TodoPriority.Medium)};

            var view = new TodoListSettingsUserControl(markers, new Mock<GridViewSort<ToDoMarker>>("", false).Object);

            //assert
            Assert.AreEqual("Todo", view.ActiveMarkerText);
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
                new ToDoMarker("Todo", TodoPriority.Medium),
                new ToDoMarker("Note", TodoPriority.Low),
                new ToDoMarker("Bug", TodoPriority.High)
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
                new ToDoMarker("Todo", TodoPriority.Medium),
                new ToDoMarker("Note", TodoPriority.Low)
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

            var view = new Mock<ITodoSettingsView>();
            view.SetupProperty(v => v.TodoMarkers, new BindingList<ToDoMarker>(markers));

            // Shut up R#, I need that to process the event
            // ReSharper disable once UnusedVariable
            var presenter = new TodoSettingPresenter(view.Object, new Mock<IAddTodoMarkerView>().Object);

            view.Raise(v => v.RemoveMarker += null, EventArgs.Empty);

            Assert.AreEqual(markers.Count - 1, view.Object.TodoMarkers.Count);
        }

        [TestMethod]
        public void AddReallyDoesDisplayAddMarkerWindow()
        {
            var markers = GetTestMarkers();

            var addView = new Mock<IAddTodoMarkerView>();

            var view = new Mock<ITodoSettingsView>();
            view.SetupProperty(v => v.TodoMarkers, new BindingList<ToDoMarker>(markers));

            // Shut up R#, I need that to process the event
            // ReSharper disable once UnusedVariable
            var presenter = new TodoSettingPresenter(view.Object, addView.Object);

            view.Raise(v => v.AddMarker += null, EventArgs.Empty);

            addView.Verify(a => a.Show(), Times.Once());
        }

        [TestMethod]
        public void AddMarkerReallyDoesBlockExistingNames()
        {
            var markers = GetTestMarkers();

            var addView = new Mock<IAddTodoMarkerView>();
            addView.SetupProperty(a => a.MarkerText, "TODO");
            addView.SetupProperty(a => a.IsValidMarker);

            var view = new Mock<ITodoSettingsView>();
            view.SetupProperty(v => v.TodoMarkers, new BindingList<ToDoMarker>(markers));

            // Shut up R#, I need that to process the event
            // ReSharper disable once UnusedVariable
            var presenter = new TodoSettingPresenter(view.Object, addView.Object);

            addView.Raise(a => a.TextChanged += null, EventArgs.Empty);

            Assert.IsFalse(addView.Object.IsValidMarker);
        }

        [TestMethod]
        public void AddMarkerReallyDoesBlockEmptyNames()
        {
            var markers = GetTestMarkers();

            var addView = new Mock<IAddTodoMarkerView>();
            addView.SetupProperty(a => a.MarkerText, "");
            addView.SetupProperty(a => a.IsValidMarker);

            var view = new Mock<ITodoSettingsView>();
            view.SetupProperty(v => v.TodoMarkers, new BindingList<ToDoMarker>(markers));

            // Shut up R#, I need that to process the event
            // ReSharper disable once UnusedVariable
            var presenter = new TodoSettingPresenter(view.Object, addView.Object);

            addView.Raise(a => a.TextChanged += null, EventArgs.Empty);

            Assert.IsFalse(addView.Object.IsValidMarker);
        }

        [TestMethod]
        public void AddMarkerReallyDoesAllowExistingNames()
        {
            var markers = GetTestMarkers();

            var addView = new Mock<IAddTodoMarkerView>();
            addView.SetupProperty(a => a.MarkerText, "Hi, guys:");
            addView.SetupProperty(a => a.IsValidMarker);

            var view = new Mock<ITodoSettingsView>();
            view.SetupProperty(v => v.TodoMarkers, new BindingList<ToDoMarker>(markers));

            // Shut up R#, I need that to process the event
            // ReSharper disable once UnusedVariable
            var presenter = new TodoSettingPresenter(view.Object, addView.Object);

            addView.Raise(a => a.TextChanged += null, EventArgs.Empty);

            Assert.IsTrue(addView.Object.IsValidMarker);
        }

        [TestMethod]
        public void AddMarkerReallyDoesAddNewMarker()
        {
            var markers = GetTestMarkers();

            var newMarkerText = "Hi, guys:";
            var newMarkerPriority = TodoPriority.High;

            var addView = new Mock<IAddTodoMarkerView>();
            addView.SetupProperty(a => a.MarkerText, newMarkerText);
            addView.SetupProperty(a => a.MarkerPriority, newMarkerPriority);
            addView.SetupProperty(a => a.TodoMarkers, markers);

            var view = new Mock<ITodoSettingsView>();
            view.SetupProperty(v => v.TodoMarkers, new BindingList<ToDoMarker>(markers));

            // Shut up R#, I need that to process the event
            // ReSharper disable once UnusedVariable
            var presenter = new TodoSettingPresenter(view.Object, addView.Object);

            addView.Raise(a => a.AddMarker += null, EventArgs.Empty);

            Assert.IsTrue(view.Object.TodoMarkers.Contains(new ToDoMarker(newMarkerText, newMarkerPriority)));
        }

        [TestMethod]
        public void AddMarkerHidesOnAdd()
        {
            var markers = GetTestMarkers();

            var addView = new Mock<IAddTodoMarkerView>();
            addView.SetupProperty(a => a.MarkerPriority, TodoPriority.Low);
            addView.SetupProperty(a => a.MarkerText, "new marker");
            addView.SetupProperty(a => a.TodoMarkers, markers);

            var view = new Mock<ITodoSettingsView>();
            view.SetupProperty(v => v.TodoMarkers, new BindingList<ToDoMarker>(markers));

            // Shut up R#, I need that to process the event
            // ReSharper disable once UnusedVariable
            var presenter = new TodoSettingPresenter(view.Object, addView.Object);

            addView.Raise(v => v.AddMarker += null, EventArgs.Empty);

            addView.Verify(a => a.Hide(), Times.Once());
        }

        [TestMethod]
        public void AddMarkerHidesOnCancel()
        {
            var markers = GetTestMarkers();

            var addView = new Mock<IAddTodoMarkerView>();
            addView.SetupProperty(a => a.MarkerPriority, TodoPriority.Low);
            addView.SetupProperty(a => a.MarkerText, "new marker");

            var view = new Mock<ITodoSettingsView>();

            // Shut up R#, I need that to process the event
            // ReSharper disable once UnusedVariable
            var presenter = new TodoSettingPresenter(view.Object, addView.Object);

            addView.Raise(v => v.Cancel += null, EventArgs.Empty);

            addView.Verify(a => a.Hide(), Times.Once());
        }
    }
}
