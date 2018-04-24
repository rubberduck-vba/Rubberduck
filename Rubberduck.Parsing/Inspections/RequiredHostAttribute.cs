using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Parsing.Inspections
{
    /// <summary>
    /// This inspection requires a specific host application in order to run.
    /// </summary>
    [AttributeUsage(AttributeTargets.Class)]
    public class RequiredHostAttribute : Attribute
    {
        public IEnumerable<string> HostNames { get; }

        /// <param name="names">Names of hosts for which the inspection should run.</param>
        public RequiredHostAttribute(params string[] names)
        {
            HostNames = names.Select(name => name.ToUpper());
        }
    }
}