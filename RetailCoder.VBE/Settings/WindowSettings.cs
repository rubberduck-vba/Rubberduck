using System.Xml.Serialization;

namespace Rubberduck.Settings
{
    public interface IWindowSettings
    {
        bool CodeExplorerVisibleOnStartup { get; set; }
        bool CodeInspectionsVisibleOnStartup { get; set; }
        bool SourceControlVisibleOnStartup { get; set; }
        bool TestExplorerVisibleOnStartup { get; set; }
        bool TodoExplorerVisibleOnStartup { get; set; }
    }

    [XmlType(AnonymousType = true)]
    public class WindowSettings : IWindowSettings
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
    }
}
