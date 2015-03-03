using System;

namespace Rubberduck.VBEHost
{
    public interface IHostApplication
    {
        /// <summary>
        /// Runs VBA procedure specified by name.
        /// </summary>
        /// <param name="target"> Method Name of the Target VBA Procedure.</param>
        void Run(string target); // note: only implementations of this method are used. does it need to be on this interface?

        /// <summary>
        /// Timed call to Application.Run
        /// </summary>
        /// <param name="projectName">Name of the project containing the method to be run.</param>
        /// <param name="moduleName">Name of the Module containing the method to be run.</param>
        /// <param name="methodName">Name of the method run.</param>
        /// <returns>A TimeSpan object representing the time elapsed during the method call.</returns>
        TimeSpan TimedMethodCall(string projectName, string moduleName, string methodName);
    }
}
