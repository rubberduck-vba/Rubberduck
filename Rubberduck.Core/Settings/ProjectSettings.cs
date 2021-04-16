using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;
using Rubberduck.UI;


namespace Rubberduck.Settings
{
    public interface IProjectSettings
    {
        string ProjectName { get; set; }
        int OpenFileDialogFilterIndex { get; set; }
    }

    [SettingsSerializeAs(SettingsSerializeAs.Xml)]
    [XmlType(AnonymousType = true)]
    public class ProjectSettings : IProjectSettings, IEquatable<ProjectSettings>
    {
        [XmlElement(Type = typeof(string))]
        public string ProjectName { get; set; } = "VBAProject";

        [XmlElement(Type = typeof(int))]
        public int OpenFileDialogFilterIndex { get; set; } = 1;

        /// <Summary>
        /// Default constructor required for XML serialization.
        /// </Summary>
        public ProjectSettings()
        {
        }

        public ProjectSettings(int openFileDialogFilterIndex)
        {
            OpenFileDialogFilterIndex = openFileDialogFilterIndex;
        }

        public bool Equals(ProjectSettings other)
        {
            return ProjectName == other.ProjectName
                && OpenFileDialogFilterIndex== other.OpenFileDialogFilterIndex;
        }
    }
}
