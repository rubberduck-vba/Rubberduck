using System;

namespace Rubberduck.VBEditor.VBEHost
{
    public interface IHostApplication
    {
        /// <summary>
        /// Runs VBA procedure specified by name.
        /// </summary>
        /// <param name="qualifiedMemberName">The method to be executed.</param>
        void Run(QualifiedMemberName qualifiedMemberName);

        /// <summary>
        /// Timed call to Application.Run
        /// </summary>
        /// <param name="qualifiedMemberName">The method to be executed.</param>
        /// <returns>A TimeSpan object representing the time elapsed during the method call.</returns>
        TimeSpan TimedMethodCall(QualifiedMemberName qualifiedMemberName);
    }
}
