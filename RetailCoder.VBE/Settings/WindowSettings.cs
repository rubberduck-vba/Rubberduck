using System;
using System.Xml.Serialization;
using Rubberduck.UI;
using Rubberduck.UI.CodeExplorer;
using Rubberduck.UI.Inspections;
using Rubberduck.UI.SourceControl;
using Rubberduck.UI.ToDoItems;
using Rubberduck.UI.UnitTesting;

namespace Rubberduck.Settings
{
    public interface IWindowSettings
    {
        bool CodeExplorerVisibleOnStartup { get; set; }
        bool CodeInspectionsVisibleOnStartup { get; set; }
        bool SourceControlVisibleOnStartup { get; set; }
        bool TestExplorerVisibleOnStartup { get; set; }
        bool TodoExplorerVisibleOnStartup { get; set; }

        bool IsWindowVisible(DockableToolwindowPresenter candidate);
    }

    [XmlType(AnonymousType = true)]
    public class WindowSettings : IWindowSettings, IEquatable<WindowSettings>
    {
        public WindowSettings()
            : this(false, false, false, false, false)
        {
            //empty constructor needed for serialization
        }

        public WindowSettings(bool codeExplorerVisibleOnStartup, bool codeInspectionsVisibleOnStartup, bool sourceControlVisibleOnStartup, bool testExplorerVisibleOnStartup, bool todoExplorerVisibleOnStartup)
        {
            CodeExplorerVisibleOnStartup = codeExplorerVisibleOnStartup;
            CodeInspectionsVisibleOnStartup = codeInspectionsVisibleOnStartup;
            SourceControlVisibleOnStartup = sourceControlVisibleOnStartup;
            TestExplorerVisibleOnStartup = testExplorerVisibleOnStartup;
            TodoExplorerVisibleOnStartup = todoExplorerVisibleOnStartup;
        }

        public bool CodeExplorerVisibleOnStartup { get; set; }
        public bool CodeInspectionsVisibleOnStartup { get; set; }
        public bool SourceControlVisibleOnStartup { get; set; }
        public bool TestExplorerVisibleOnStartup { get; set; }
        public bool TodoExplorerVisibleOnStartup { get; set; }

        public bool IsWindowVisible(DockableToolwindowPresenter candidate)
        {
            //I'm sure there's a better way to do this, because this is a lazy-ass way to do it.
            //We're injecting into the base class, so check the derived class:
            if (candidate is CodeExplorerDockablePresenter)
            {
                return CodeExplorerVisibleOnStartup;
            }
            if (candidate is CodeInspectionsDockablePresenter)
            {
                return CodeInspectionsVisibleOnStartup;
            }
            if (candidate is SourceControlDockablePresenter)
            {
                return SourceControlVisibleOnStartup;
            }
            if (candidate is TestExplorerDockablePresenter)
            {
                return TestExplorerVisibleOnStartup;
            }
            if (candidate is ToDoExplorerDockablePresenter)
            {
                return TodoExplorerVisibleOnStartup;
            }
            //Oh. Hello. I have no clue who you are...
            return false;
        }

        public bool Equals(WindowSettings other)
        {
            return other != null &&
                   CodeExplorerVisibleOnStartup == other.CodeExplorerVisibleOnStartup &&
                   CodeInspectionsVisibleOnStartup == other.CodeInspectionsVisibleOnStartup &&
                   SourceControlVisibleOnStartup == other.SourceControlVisibleOnStartup &&
                   TestExplorerVisibleOnStartup == other.TestExplorerVisibleOnStartup &&
                   TodoExplorerVisibleOnStartup == other.TodoExplorerVisibleOnStartup;
        }
    }
}
