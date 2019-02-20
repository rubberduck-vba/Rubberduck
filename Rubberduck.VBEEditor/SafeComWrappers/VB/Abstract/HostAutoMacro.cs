using System.Collections.Generic;

namespace Rubberduck.VBEditor.SafeComWrappers.Abstract
{
    /// <summary>
    /// Provide information about naming schemes that the host
    /// may use to identify a VBA procedure used as part of
    /// automatic execution via macros (e.g. Excel's auto_open or
    /// Word's AutoOpen
    /// </summary>
    public readonly struct HostAutoMacro
    {
        /// <summary>
        /// Enumerates all component types that the host may search for such auto macro
        /// </summary>
        public IEnumerable<ComponentType> ComponentTypes { get; }

        /// <summary>
        /// Indicates whether host will require public access or may ignore the access modifier
        /// </summary>
        public bool MayBePrivate { get; }

        /// <summary>
        /// If the host requires the module to have a specific name, this should be specified.
        /// Otherwise leave null. The procedure name must be specified in this case.
        /// </summary>
        public string ModuleName { get; }

        /// <summary>
        /// If the host requires the procedure to have a specific name, this should be
        /// specified. Otherwise leave null. The module name must be specified in this case.
        /// </summary>
        public string ProcedureName { get; }

        public HostAutoMacro(IEnumerable<ComponentType> componentTypes, bool mayBePrivate, string moduleName,
            string procedureName)
        {
            ComponentTypes = componentTypes;
            MayBePrivate = mayBePrivate;
            ModuleName = moduleName;
            ProcedureName = procedureName;
        }
    }
}
