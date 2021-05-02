using System;
using System.Configuration;
using System.Xml.Serialization;

namespace Rubberduck.Settings
{
    public interface IProjectSettings
    {
        int OpenFileDialogFilterIndex { get; set; }
    }

    [SettingsSerializeAs(SettingsSerializeAs.Xml)]
    [XmlType(AnonymousType = true)]
    public class ProjectSettings : IProjectSettings, IEquatable<ProjectSettings>
    {
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
            return OpenFileDialogFilterIndex == other.OpenFileDialogFilterIndex;
        }
    }
}
