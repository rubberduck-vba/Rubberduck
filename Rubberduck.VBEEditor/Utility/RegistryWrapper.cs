using Microsoft.Win32;

namespace Rubberduck.VBEditor.Utility
{
    public interface IRegistryWrapper
    {
        object GetValue(string keyName, string valueName, object defaultValue);
        void SetValue(string keyName, string valueName, object value, RegistryValueKind valueKind);
    }

    public class RegistryWrapper : IRegistryWrapper
    {
        public object GetValue(string keyName, string valueName, object defaultValue)
        {
            return Registry.GetValue(keyName, valueName, defaultValue);
        }

        public void SetValue(string keyName, string valueName, object value, RegistryValueKind valueKind)
        {
            Registry.SetValue(keyName, valueName, value, valueKind);
        }
    }
}
