using System;
using System.Configuration;
using System.Xml.Serialization;
using Rubberduck.UI;
using Rubberduck.UI.CodeExplorer;
using Rubberduck.UI.Inspections;
using Rubberduck.UI.ToDoItems;
using Rubberduck.UI.UnitTesting;

namespace Rubberduck.Settings
{
    public interface IWindowSettings
    {
        bool CodeExplorerVisibleOnStartup { get; set; }
        bool CodeInspectionsVisibleOnStartup { get; set; }
        bool TestExplorerVisibleOnStartup { get; set; }
        bool TodoExplorerVisibleOnStartup { get; set; }

        bool CodeExplorer_SortByName { get; set; }
        bool CodeExplorer_SortByCodeOrder { get; set; }
        bool CodeExplorer_GroupByType { get; set; }

        bool IsWindowVisible(DockableToolwindowPresenter candidate);
    }

    [SettingsSerializeAs(SettingsSerializeAs.Xml)]
    [XmlType(AnonymousType = true)]
    public class WindowSettings : IWindowSettings, IEquatable<WindowSettings>
    {
        /// <Summary>
        /// Default constructor required for XML serialization. Initializes all settings to false.
        /// </Summary>
        public WindowSettings()
        {
        }

        public WindowSettings(bool codeExplorerVisibleOnStartup, bool codeInspectionsVisibleOnStartup, 
            bool testExplorerVisibleOnStartup, bool todoExplorerVisibleOnStartup, 
            bool codeExplorer_SortByName, bool codeExplorer_SortByCodeOrder, bool codeExplorer_GroupByType)
        {
            CodeExplorerVisibleOnStartup = codeExplorerVisibleOnStartup;
            CodeInspectionsVisibleOnStartup = codeInspectionsVisibleOnStartup;
            TestExplorerVisibleOnStartup = testExplorerVisibleOnStartup;
            TodoExplorerVisibleOnStartup = todoExplorerVisibleOnStartup;

            CodeExplorer_SortByName = codeExplorer_SortByName;
            CodeExplorer_SortByCodeOrder = codeExplorer_SortByCodeOrder;
            CodeExplorer_GroupByType = codeExplorer_GroupByType;
        }

        public bool CodeExplorerVisibleOnStartup { get; set; }
        public bool CodeInspectionsVisibleOnStartup { get; set; }
        public bool TestExplorerVisibleOnStartup { get; set; }
        public bool TodoExplorerVisibleOnStartup { get; set; }

        public bool CodeExplorer_SortByName { get; set; }
        public bool CodeExplorer_SortByCodeOrder { get; set; }
        public bool CodeExplorer_GroupByType { get; set; }

        public bool IsWindowVisible(DockableToolwindowPresenter candidate)
        {
            //I'm sure there's a better way to do this, because this is a lazy-ass way to do it.
            //We're injecting into the base class, so check the derived class:
            switch (candidate)
            {
                case CodeExplorerDockablePresenter _:
                    return CodeExplorerVisibleOnStartup;
                case InspectionResultsDockablePresenter _:
                    return CodeInspectionsVisibleOnStartup;
                case TestExplorerDockablePresenter _:
                    return TestExplorerVisibleOnStartup;
                case ToDoExplorerDockablePresenter _:
                    return TodoExplorerVisibleOnStartup;
                default:
                    //Oh. Hello. I have no clue who you are...
                    return false;
            }
        }

        public bool Equals(WindowSettings other)
        {
            return other != null &&
                   CodeExplorerVisibleOnStartup == other.CodeExplorerVisibleOnStartup &&
                   CodeInspectionsVisibleOnStartup == other.CodeInspectionsVisibleOnStartup &&
                   TestExplorerVisibleOnStartup == other.TestExplorerVisibleOnStartup &&
                   TodoExplorerVisibleOnStartup == other.TodoExplorerVisibleOnStartup &&
                   CodeExplorer_SortByName == other.CodeExplorer_SortByName &&
                   CodeExplorer_SortByCodeOrder == other.CodeExplorer_SortByCodeOrder &&
                   CodeExplorer_GroupByType == other.CodeExplorer_GroupByType;
        }
    }
}
