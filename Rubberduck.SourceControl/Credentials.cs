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
        public virtual string Username { get; set; }
        public virtual TPassword Password { get; set; } 
    }

    /// <summary>
    /// Stores user credentials.
    /// </summary>
    /// <remarks>
    /// Do no use internally. For COM Interop only. Use <see cref="SecureCredentials"/> instead./>
    /// </remarks>
    public sealed class Credentials : CredentialsBase<string>
    {
        public Credentials(string username, string password)
        {
            this.Username = username;
            this.Password = password;
        }
    }

    /// <summary>
    /// Stores user name and password credentials.
    /// </summary>
    public sealed class SecureCredentials : CredentialsBase<SecureString>
    {
        public SecureCredentials(string username, SecureString password)
        {
            this.Username = username;
            this.Password = password;
        }
    }
}
