using System;

namespace Rubberduck.VBEditor.ComManagement.TypeLibs.Abstract
{
    public interface ITypeInfoVBEExtensions
    {
        /// <summary>
        /// Silently compiles the individual VBE component (class/module etc)
        /// </summary>
        /// <returns>true if this module, plus any direct dependent modules compile successfully</returns>
        bool CompileComponent();

        /// <summary>
        /// Provides an accessor object for invoking methods on a standard module in a VBA project
        /// </summary>
        /// <remarks>caller is responsible for calling ReleaseComObject</remarks>
        /// <returns>the accessor object</returns>
        IDispatch GetStdModAccessor();

        /// <summary>
        /// Executes a procedure inside a standard module in a VBA project
        /// </summary>
        /// <param name="name">the name of the procedure to invoke</param>
        /// <param name="args">arguments to pass to the procedure</param>
        /// <remarks>the returned object can be a COM object, and the callee is responsible for releasing it appropriately</remarks>
        /// <returns>an object representing the return value from the procedure, or null if none.</returns>
        object StdModExecute(string name, object[] args = null);
    }
}