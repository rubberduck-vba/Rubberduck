﻿using System;

namespace Rubberduck.VBEditor.VBEHost
{
    public interface IHostApplication : IDisposable
    {
        /// <summary>
        /// Runs VBA procedure specified by name.
        /// </summary>
        /// <param name="qualifiedMemberName">The method to be executed.</param>
        void Run(QualifiedMemberName qualifiedMemberName);

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
