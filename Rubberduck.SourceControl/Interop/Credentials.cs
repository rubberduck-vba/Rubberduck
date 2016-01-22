using System.Runtime.InteropServices;

namespace Rubberduck.SourceControl.Interop
{
    [ComVisible(true)]
    [Guid("4EDFBFED-F6E7-4AFA-ADB4-B9FCAD21C256")]
    public interface ICredentials 
    {
        string Username { get; set; }
        string Password { get; set; }
    }

    [ComVisible(true)]
    [Guid("AE54B926-49EB-4FB1-9F8A-AFE504A5A569")]
    [ProgId("Rubberduck.Credentials")]
    [ClassInterface(ClassInterfaceType.None)]
    public class Credentials : ICredentials
    {
        public string Username { get; set; }
        public string Password { get; set; }

        public Credentials() { }

        internal Credentials(string username, string password)
            :this()
        {
            Username = username;
            Password = password;
        }
    }
}
