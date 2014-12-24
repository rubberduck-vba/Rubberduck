using System.Runtime.InteropServices;
using System.Diagnostics;
using System;

namespace Rubberduck
{
    [ComVisible(false)]
    public interface IHostApplication
    {
        /// <summary>   Runs VBA procedure specified by name. </summary>
        /// <param name="target"> Method Name of the Target VBA Procedure. </param>
        void Run(string target);

        /// <summary>   Timed call to Application.Run </summary>
        /// <param name="projectName">  Name of the project containing the method to be run. </param>
        /// <param name="moduleName">   Name of the module containing the method to be run. </param>
        /// <param name="methodName">   Name of the method run. </param>
        /// <returns>   Number of milliseconds it took to run the VBA procedure. </returns>
        TimeSpan TimedMethodCall(string projectName, string moduleName, string methodName);
    }
}
