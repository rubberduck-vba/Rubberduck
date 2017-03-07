using System;

namespace Rubberduck.VBEditor.Application
{
    public interface IHostApplication : IDisposable
    {
        /// <summary>
        /// Runs VBA procedure specified by name. WARNING: The parameter is declared as dynamic to prevent circular referencing.
        /// This should ONLY be passed a Declaration object.
        /// </summary>
        /// <param name="declaration">The Declaration object for the method to be executed.</param>
        void Run(dynamic declaration);

        /// <summary>
        /// Executes a VBA function by name, with specified parameters, and returns a result.
        /// </summary>
        /// <param name="name"></param>
        /// <param name="args"></param>
        /// <returns></returns>
        /// <remarks>
        /// May not be available in all host applications.
        /// </remarks>
        object Run(string name, params object[] args);

        /// <summary>
        /// Gets the name of the application.
        /// </summary>
        /// <remarks>
        /// This is needed only to circumvent the problem that HostApplicationBase-derived types
        /// cannot be used outside assembly boundaries because the type is generic.
        /// </remarks>
        string ApplicationName { get; }
    }
}
