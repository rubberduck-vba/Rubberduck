using System.Security;

namespace Rubberduck.SourceControl
{
    public interface ICredentials<TPassword>
    {
        string Username { get; set; }
        TPassword Password { get; set; }
    }

    public abstract class CredentialsBase<TPassword> : ICredentials<TPassword>
    {
        public string Username { get; set; }
        public TPassword Password { get; set; }

        protected CredentialsBase(string username, TPassword password)
        {
            Username = username;
            Password = password;
        } 
    }

    /// <summary>
    /// Stores user name and password credentials.
    /// </summary>
    /// <remarks>
    /// Do no use internally. For COM Interop only. Use <see cref="SecureCredentials"/> instead./>
    /// </remarks>
    public class Credentials : CredentialsBase<string>
    {
        public Credentials(string username, string password)
            :base(username, password)
        { }
    }

    /// <summary>
    /// Securely stores user name and password credentials.
    /// </summary>
    public class SecureCredentials : CredentialsBase<SecureString>
    {
        public SecureCredentials(string username, SecureString password)
            : base(username, password)
        { }
    }
}
