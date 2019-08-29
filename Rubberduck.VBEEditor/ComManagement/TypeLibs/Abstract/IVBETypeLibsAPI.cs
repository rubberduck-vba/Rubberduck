using System;
using System.Collections.Generic;
using Rubberduck.VBEditor.ComManagement.TypeLibs.Utility;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.ComManagement.TypeLibs.Abstract
{
    public interface IVBETypeLibsAPI
    {
        /// <summary>
        /// Compile an entire VBE project
        /// </summary>
        /// <param name="ide">Safe-com wrapper representing the VBE</param>
        /// <param name="projectName">The VBA project name</param>
        /// <returns>bool indicating success/failure</returns>
        bool CompileProject(IVBE ide, string projectName);

        /// <summary>
        /// Compile an entire VBA project
        /// </summary>
        /// <param name="project">Safe-com wrapper representing the VBA project</param>
        /// <returns>bool indicating success/failure.</returns>
        bool CompileProject(IVBProject project);

        /// <summary>
        /// Compile an entire VBA project
        /// </summary>
        /// <param name="projectTypeLib">Low-level ITypeLib wrapper representing the VBA project</param>
        /// <returns>bool indicating success/failure</returns>
        bool CompileProject(ITypeLibWrapper projectTypeLib);

        /// <summary>
        /// Compile a single VBA component (e.g. module/class)
        /// </summary>
        /// <remarks>NOTE: This will only return success if ALL components that this component depends on also compile successfully</remarks>
        /// <param name="ide">Safe-com wrapper representing the VBE</param>
        /// <param name="projectName">The VBA project name</param>
        /// <param name="componentName">The name of the component (module/class etc) to compile</param>
        /// <returns>bool indicating success/failure.</returns>
        bool CompileComponent(IVBE ide, string projectName, string componentName);

        /// <summary>
        /// Compile a single VBA component (e.g. module/class)
        /// </summary>
        /// <remarks>NOTE: This will only return success if ALL components that this component depends on also compile successfully</remarks>
        /// <param name="project">Safe-com wrapper representing the VBA project</param>
        /// <param name="componentName">The name of the component (module/class etc) to compile</param>
        /// <returns>bool indicating success/failure.</returns>
        bool CompileComponent(IVBProject project, string componentName);

        /// <summary>
        /// Compile a single VBA component (e.g. module/class)
        /// </summary>
        /// <remarks>NOTE: This will only return success if ALL components that this component depends on also compile successfully</remarks>
        /// <param name="projectTypeLib">Low-level ITypeLib wrapper representing the VBA project</param>
        /// <param name="componentName">The name of the component (module/class etc) to compile</param>
        /// <returns>bool indicating success/failure.</returns>
        bool CompileComponent(ITypeLibWrapper projectTypeLib, string componentName);

        /// <summary>
        /// Compile a single VBA component (e.g. module/class)
        /// </summary>
        /// <remarks>NOTE: This will only return success if ALL components that this component depends on also compile successfully</remarks>
        /// <param name="component">Safe-com wrapper representing the VBA component to compile</param>
        /// <returns>bool indicating success/failure.</returns>
        bool CompileComponent(IVBComponent component);

        /// <summary>
        /// Compile a single VBA component (e.g. module/class)
        /// </summary>
        /// <remarks>NOTE: This will only return success if ALL components that this component depends on also compile successfully</remarks>
        /// <param name="componentTypeInfo">Low-level ITypeInfo wrapper representing the VBA component to compile</param>
        /// <returns>bool indicating success/failure.</returns>
        bool CompileComponent(ITypeInfoWrapper componentTypeInfo);

        /// <summary>
        /// Execute a routine inside a standard VBA code module
        /// </summary>
        /// <remarks>the VBA return value returned here can be a COM object, but needs freeing with Marshal.ReleaseComObject to ensure deterministic behaviour.</remarks>
        /// <param name="ide">Safe-com wrapper representing the VBE</param>
        /// <param name="projectName">VBA Project name, as declared in the VBE</param>
        /// <param name="standardModuleName">Module name, as declared in the VBA project</param>
        /// <param name="procName">Procedure name, as declared in the VBA module</param>
        /// <param name="args">optional array of arguments to pass to the VBA routine</param>
        /// <returns>object representing the VBA return value, if one was provided, or null otherwise.</returns>
        object ExecuteCode(IVBE ide, string projectName, string standardModuleName, string procName, object[] args = null);

        /// <summary>
        /// Execute a routine inside a standard VBA code module
        /// </summary>
        /// <remarks>the VBA return value returned here can be a COM object, but needs freeing with Marshal.ReleaseComObject to ensure deterministic behaviour.</remarks>
        /// <param name="project">Safe-com wrapper representing the VBA project</param>
        /// <param name="standardModuleName">Module name, as declared in the VBA project</param>
        /// <param name="procName">Procedure name, as declared in the VBA module</param>
        /// <param name="args">optional array of arguments to pass to the VBA routine</param>
        /// <returns>object representing the VBA return value, if one was provided, or null otherwise.</returns>
        object ExecuteCode(IVBProject project, string standardModuleName, string procName, object[] args = null);

        /// <summary>
        /// Execute a routine inside a standard VBA code module
        /// </summary>
        /// <remarks>the VBA return value returned here can be a COM object, but needs freeing with Marshal.ReleaseComObject to ensure deterministic behaviour.</remarks>
        /// <param name="projectTypeLib">Low-level ITypeLib wrapper representing the VBA project which contains the routine</param>
        /// <param name="standardModuleName">Module name, as declared in the VBA project</param>
        /// <param name="procName">Procedure name, as declared in the VBA module</param>
        /// <param name="args">optional array of arguments to pass to the VBA routine</param>
        /// <returns>object representing the VBA return value, if one was provided, or null otherwise.</returns>
        object ExecuteCode(ITypeLibWrapper projectTypeLib, string standardModuleName, string procName, object[] args = null);

        /// <summary>
        /// Execute a routine inside a standard VBA code module
        /// </summary>
        /// <remarks>the VBA return value returned here can be a COM object, but needs freeing with Marshal.ReleaseComObject to ensure deterministic behaviour.</remarks>
        /// <param name="component">Safe-com wrapper representing the VBA component where the routine is defined</param>
        /// <param name="procName">Procedure name, as declared in the VBA module</param>
        /// <param name="args">optional array of arguments to pass to the VBA routine</param>
        /// <returns>object representing the VBA return value, if one was provided, or null otherwise.</returns>
        object ExecuteCode(IVBComponent component, string procName, object[] args = null);

        /// <summary>
        /// Execute a routine inside a standard VBA code module
        /// </summary>
        /// <remarks>the VBA return value returned here can be a COM object, but needs freeing with Marshal.ReleaseComObject to ensure deterministic behaviour.</remarks>
        /// <param name="standardModuleTypeInfo">Low-level ITypeInfo wrapper representing the VBA component which contains the routine</param>
        /// <param name="procName">Procedure name, as declared in the VBA module</param>
        /// <param name="args">optional array of arguments to pass to the VBA routine</param>
        /// <returns>object representing the VBA return value, if one was provided, or null otherwise.</returns>
        object ExecuteCode(ITypeInfoWrapper standardModuleTypeInfo, string procName, object[] args = null);

        /// <summary>
        /// Retrieves the developer-defined conditional compilation arguments of a VBA project
        /// </summary>
        /// <remarks>does not expose compiler-defined arguments, such as WIN64, VBA7 etc, which must be determined via the running process</remarks>
        /// <param name="ide">Safe-com wrapper representing the VBE</param>
        /// <param name="projectName">VBA Project name, as declared in the VBE</param>
        /// <returns>returns the raw unparsed conditional arguments string, e.g. "foo = 1 : bar = 2"</returns>
        string GetProjectConditionalCompilationArgsRaw(IVBE ide, string projectName);

        /// <summary>
        /// Retrieves the developer-defined conditional compilation arguments of a VBA project
        /// </summary>
        /// <remarks>does not expose compiler-defined arguments, such as WIN64, VBA7 etc, which must be determined via the running process</remarks>
        /// <param name="project">Safe-com wrapper representing the VBA project</param>
        /// <returns>returns the raw unparsed conditional arguments string, e.g. "foo = 1 : bar = 2"</returns>
        string GetProjectConditionalCompilationArgsRaw(IVBProject project);

        /// <summary>
        /// Retrieves the developer-defined conditional compilation arguments of a VBA project
        /// </summary>
        /// <remarks>does not expose compiler-defined arguments, such as WIN64, VBA7 etc, which must be determined via the running process</remarks>
        /// <param name="projectTypeLib">Low-level ITypeLib wrapper representing the VBA project</param>
        /// <returns>returns the raw unparsed conditional arguments string, e.g. "foo = 1 : bar = 2"</returns>
        string GetProjectConditionalCompilationArgsRaw(ITypeLibWrapper projectTypeLib);

        /// <summary>
        /// Retrieves the developer-defined conditional compilation arguments of a VBA project
        /// </summary>
        /// <remarks>does not expose compiler-defined arguments, such as WIN64, VBA7 etc, which must be determined via the running process</remarks>
        /// <param name="ide">Safe-com wrapper representing the VBE</param>
        /// <param name="projectName">VBA Project name, as declared in the VBE</param>
        /// <returns>returns a Dictionary<string, short>, parsed from the conditional arguments string</returns>
        Dictionary<string, short> GetProjectConditionalCompilationArgs(IVBE ide, string projectName);

        /// <summary>
        /// Retrieves the developer-defined conditional compilation arguments of a VBA project
        /// </summary>
        /// <remarks>does not expose compiler-defined arguments, such as WIN64, VBA7 etc, which must be determined via the running process</remarks>
        /// <param name="project">Safe-com wrapper representing the VBA project</param>
        /// <returns>returns a Dictionary<string, short>, parsed from the conditional arguments string</returns>
        Dictionary<string, short> GetProjectConditionalCompilationArgs(IVBProject project);

        /// <summary>
        /// Retrieves the developer-defined conditional compilation arguments of a VBA project
        /// </summary>
        /// <remarks>does not expose compiler-defined arguments, such as WIN64, VBA7 etc, which must be determined via the running process</remarks>
        /// <param name="projectTypeLib">Low-level ITypeLib wrapper representing the VBA project</param>
        /// <returns>returns a Dictionary<string, short>, parsed from the conditional arguments string</returns>
        Dictionary<string, short> GetProjectConditionalCompilationArgs(ITypeLibWrapper projectTypeLib);

        /// <summary>
        /// Sets the developer-defined conditional compilation arguments of a VBA project
        /// </summary>
        /// <remarks>don't set compiler-defined arguments, such as WIN64, VBA7 etc</remarks>
        /// <param name="ide">Safe-com wrapper representing the VBE</param>
        /// <param name="projectName">VBA Project name, as declared in the VBE</param>
        /// <param name="newConditionalArgs">Raw string representing the arguments, e.g. "foo = 1 : bar = 2"</param>
        void SetProjectConditionalCompilationArgsRaw(IVBE ide, string projectName, string newConditionalArgs);

        /// <summary>
        /// Sets the developer-defined conditional compilation arguments of a VBA project
        /// </summary>
        /// <remarks>don't set compiler-defined arguments, such as WIN64, VBA7 etc</remarks>
        /// <param name="project">Safe-com wrapper representing the VBA project</param>
        /// <param name="newConditionalArgs">Raw string representing the arguments, e.g. "foo = 1 : bar = 2"</param>
        void SetProjectConditionalCompilationArgsRaw(IVBProject project, string newConditionalArgs);

        /// <summary>
        /// Sets the developer-defined conditional compilation arguments of a VBA project
        /// </summary>
        /// <remarks>don't set compiler-defined arguments, such as WIN64, VBA7 etc</remarks>
        /// <param name="projectTypeLib">Low-level ITypeLib wrapper representing the VBA project</param>
        /// <param name="newConditionalArgs">Raw string representing the arguments, e.g. "foo = 1 : bar = 2"</param>
        void SetProjectConditionalCompilationArgsRaw(ITypeLibWrapper projectTypeLib, string newConditionalArgs);

        /// <summary>
        /// Sets the developer-defined conditional compilation arguments of a VBA project
        /// </summary>
        /// <remarks>don't set compiler-defined arguments, such as WIN64, VBA7 etc</remarks>
        /// <param name="ide">Safe-com wrapper representing the VBE</param>
        /// <param name="projectName">VBA Project name, as declared in the VBE</param>
        /// <param name="newConditionalArgs">Dictionary<string, short> representing the argument name-value pairs</param>
        void SetProjectConditionalCompilationArgs(IVBE ide, string projectName, Dictionary<string, short> newConditionalArgs);

        /// <summary>
        /// Sets the developer-defined conditional compilation arguments of a VBA project
        /// </summary>
        /// <remarks>don't set compiler-defined arguments, such as WIN64, VBA7 etc</remarks>
        /// <param name="project">Safe-com wrapper representing the VBA project</param>
        /// <param name="newConditionalArgs">Dictionary<string, short> representing the argument name-value pairs</param>
        void SetProjectConditionalCompilationArgs(IVBProject project, Dictionary<string, short> newConditionalArgs);

        /// <summary>
        /// Sets the developer-defined conditional compilation arguments of a VBA project
        /// </summary>
        /// <remarks>don't set compiler-defined arguments, such as WIN64, VBA7 etc</remarks>
        /// <param name="projectTypeLib">Low-level ITypeLib wrapper representing the VBA project</param>
        /// <param name="newConditionalArgs">Dictionary<string, short> representing the argument name-value pairs</param>
        void SetProjectConditionalCompilationArgs(ITypeLibWrapper projectTypeLib, Dictionary<string, short> newConditionalArgs);

        /// <summary>
        /// Determines whether the specified document class is a known document class type (e.g. Excel._Workbook, Access._Form)
        /// </summary>
        /// <param name="ide">Safe-com wrapper representing the VBE</param>
        /// <param name="projectName">VBA Project name, as declared in the VBE</param>
        /// <param name="className">The name of the class document, as defined in the VBA project</param>
        /// <returns>DocClassType indicating the type of the document class module, or DocType.Unrecognized</returns>
        DocClassType DetermineDocumentClassType(IVBE ide, string projectName, string className);

        /// <summary>
        /// Determines whether the specified document class is a known document class type (e.g. Excel._Workbook, Access._Form)
        /// </summary>
        /// <param name="projectTypeLib">Low-level ITypeLib wrapper representing the VBA project</param>
        /// <param name="className">The name of the class document, as defined in the VBA project</param>
        /// <returns>DocClassType indicating the type of the document class module, or DocType.Unrecognized</returns>
        DocClassType DetermineDocumentClassType(ITypeLibWrapper projectTypeLib, string className);

        /// <summary>
        /// Determines whether the specified document class is a known document class type (e.g. Excel._Workbook, Access._Form)
        /// </summary>
        /// <param name="project">Safe-com wrapper representing the VBA project</param>
        /// <param name="className">The name of the class document, as defined in the VBA project</param>
        /// <returns>DocClassType indicating the type of the document class module, or DocType.Unrecognized</returns>
        DocClassType DetermineDocumentClassType(IVBProject project, string className);

        /// <summary>
        /// Determines whether the specified document class is a known document class type (e.g. Excel._Workbook, Access._Form)
        /// </summary>
        /// <param name="project">Safe-com wrapper representing the VBA component</param>
        /// <returns>DocClassType indicating the type of the document class module, or DocType.Unrecognized</returns>
        DocClassType DetermineDocumentClassType(IVBComponent component);

        /// <summary>
        /// Determines whether the specified document class is a known document class type (e.g. Excel._Workbook, Access._Form)
        /// </summary>
        /// <param name="classTypeInfo">Low-level ITypeInfo wrapper representing the VBA project</param>
        /// <returns>DocClassType indicating the type of the document class module, or DocType.Unrecognized</returns>
        DocClassType DetermineDocumentClassType(ITypeInfoWrapper classTypeInfo);

        /// <summary>
        /// Determines whether the specified VBA class implements a specific interface
        /// </summary>
        /// <param name="ide">Safe-com wrapper representing the VBE</param>
        /// <param name="projectName">VBA Project name, as declared in the VBE</param>
        /// <param name="className">Document class name, as declared in the VBA project</param>
        /// <param name="interfaceProgID">The interface name, preceeded by the library container name, e.g. "Excel._Worksheet"</param>
        /// <returns>bool indicating whether the class does inherit the specified interface</returns>
        bool DoesClassImplementInterface(IVBE ide, string projectName, string className, string interfaceProgID);

        /// <summary>
        /// Determines whether the specified VBA class implements a specific interface
        /// </summary>
        /// <param name="project">Safe-com wrapper representing the VBE project</param>
        /// <param name="className">Document class name, as declared in the VBA project</param>
        /// <param name="interfaceProgID">The interface name, preceeded by the library container name, e.g. "Excel._Worksheet"</param>
        /// <returns>bool indicating whether the class does inherit the specified interface</returns>
        bool DoesClassImplementInterface(IVBProject project, string className, string interfaceProgID);

        /// <summary>
        /// Determines whether the specified VBA class implements a specific interface
        /// </summary>
        /// <param name="projectTypeLib">Low-level ITypeLib wrapper representing the VBA project</param>
        /// <param name="className">Document class name, as declared in the VBA project</param>
        /// <param name="interfaceProgID">The interface name, preceeded by the library container name, e.g. "Excel._Worksheet"</param>
        /// <returns>bool indicating whether the class does inherit the specified interface</returns>
        bool DoesClassImplementInterface(ITypeLibWrapper projectTypeLib, string className, string interfaceProgID);

        /// <summary>
        /// Determines whether the specified VBA class implements a specific interface
        /// </summary>
        /// <param name="component">Safe-com wrapper representing the VBA component</param>
        /// <param name="interfaceProgID">The interface name, preceeded by the library container name, e.g. "Excel._Worksheet"</param>
        /// <returns>bool indicating whether the class does inherit the specified interface</returns>
        bool DoesClassImplementInterface(IVBComponent component, string interfaceProgID);

        /// <summary>
        /// Determines whether the specified VBA class implements a specific interface
        /// </summary>
        /// <param name="classTypeInfo">Low-level ITypeInfo wrapper representing the VBA project</param>
        /// <param name="interfaceProgID">The interface name, preceeded by the library container name, e.g. "Excel._Worksheet"</param>
        /// <returns>bool indicating whether the class does inherit the specified interface</returns>
        bool DoesClassImplementInterface(ITypeInfoWrapper classTypeInfo, string interfaceProgID);

        /// <summary>
        /// Determines whether the specified VBA class implements a specific interface
        /// </summary>
        /// <param name="ide">Safe-com wrapper representing the VBE</param>
        /// <param name="projectName">VBA Project name, as declared in the VBE</param>
        /// <param name="className">Document class name, as declared in the VBA project</param>
        /// <param name="interfaceProgIDs">An array of interface names, preceeded by the library container name, e.g. "Excel._Worksheet"</param>
        /// <param name="matchedIndex">on return indicates the index into interfaceProgIDs that matched, or -1 if no match</param>
        /// <returns>bool indicating whether the class does inherit one of the specified interfaces</returns>
        bool DoesClassImplementInterface(IVBE ide, string projectName, string className, string[] interfaceProgIDs, out int matchedIndex);

        /// <summary>
        /// Determines whether the specified VBA class implements a specific interface
        /// </summary>
        /// <param name="project">Safe-com wrapper representing the VBE project</param>
        /// <param name="className">Document class name, as declared in the VBA project</param>
        /// <param name="interfaceProgIDs">An array of interface names, preceeded by the library container name, e.g. "Excel._Worksheet"</param>
        /// <param name="matchedIndex">on return indicates the index into interfaceProgIDs that matched, or -1 if no match</param>
        /// <returns>bool indicating whether the class does inherit one of the specified interfaces</returns>
        bool DoesClassImplementInterface(IVBProject project, string className, string[] interfaceProgIDs, out int matchedIndex);

        /// <summary>
        /// Determines whether the specified VBA class implements a specific interface
        /// </summary>
        /// <param name="projectTypeLib">Low-level ITypeLib wrapper representing the VBA project</param>
        /// <param name="className">Document class name, as declared in the VBA project</param>
        /// <param name="interfaceProgIDs">An array of interface names, preceeded by the library container name, e.g. "Excel._Worksheet"</param>
        /// <param name="matchedIndex">on return indicates the index into interfaceProgIDs that matched, or -1 if no match</param>
        /// <returns>bool indicating whether the class does inherit one of the specified interfaces</returns>
        bool DoesClassImplementInterface(ITypeLibWrapper projectTypeLib, string className, string[] interfaceProgIDs, out int matchedIndex);

        /// <summary>
        /// Determines whether the specified VBA class implements a specific interface
        /// </summary>
        /// <param name="component">Safe-com wrapper representing the VBA component</param>
        /// <param name="interfaceProgIDs">An array of interface names, preceeded by the library container name, e.g. "Excel._Worksheet"</param>
        /// <param name="matchedIndex">on return indicates the index into interfaceProgIDs that matched, or -1 if no match</param>
        /// <returns>bool indicating whether the class does inherit one of the specified interfaces</returns>
        bool DoesClassImplementInterface(IVBComponent component, string[] interfaceProgIDs, out int matchedIndex);

        /// <summary>
        /// Determines whether the specified VBA class implements a specific interface
        /// </summary>
        /// <param name="classTypeInfo">Low-level ITypeInfo wrapper representing the VBA project</param>
        /// <param name="interfaceProgIDs">An array of interface names, preceeded by the library container name, e.g. "Excel._Worksheet"</param>
        /// <param name="matchedIndex">on return indicates the index into interfaceProgIDs that matched, or -1 if no match</param>
        /// <returns>bool indicating whether the class does inherit one of the specified interfaces</returns>
        bool DoesClassImplementInterface(ITypeInfoWrapper classTypeInfo, string[] interfaceProgIDs, out int matchedIndex);

        /// <summary>
        /// Determines whether the specified VBA class implements a specific interface
        /// </summary>
        /// <param name="ide">Safe-com wrapper representing the VBE</param>
        /// <param name="projectName">VBA Project name, as declared in the VBE</param>
        /// <param name="className">Document class name, as declared in the VBA project</param>
        /// <param name="interfaceIID">The interface IID</param>
        /// <returns>bool indicating whether the class does inherit the specified interface</returns>
        bool DoesClassImplementInterface(IVBE ide, string projectName, string className, Guid interfaceIID);

        /// <summary>
        /// Determines whether the specified VBA class implements a specific interface
        /// </summary>
        /// <param name="project">Safe-com wrapper representing the VBA project</param>
        /// <param name="className">Document class name, as declared in the VBA project</param>
        /// <param name="interfaceIID">The interface IID</param>
        /// <returns>bool indicating whether the class does inherit the specified interface</returns>
        bool DoesClassImplementInterface(IVBProject project, string className, Guid interfaceIID);

        /// <summary>
        /// Determines whether the specified VBA class implements a specific interface
        /// </summary>
        /// <param name="projectTypeLib">Low-level ITypeLib wrapper representing the VBA project</param>
        /// <param name="className">Document class name, as declared in the VBA project</param>
        /// <param name="interfaceIID">The interface IID</param>
        /// <returns>bool indicating whether the class does inherit the specified interface</returns>
        bool DoesClassImplementInterface(ITypeLibWrapper projectTypeLib, string className, Guid interfaceIID);

        /// <summary>
        /// Determines whether the specified VBA class implements a specific interface
        /// </summary>
        /// <param name="component">Safe-com wrapper representing the VBA component</param>
        /// <param name="interfaceIID">The interface IID</param>
        /// <returns>bool indicating whether the class does inherit the specified interface</returns>
        bool DoesClassImplementInterface(IVBComponent component, Guid interfaceIID);

        /// <summary>
        /// Determines whether the specified VBA class implements a specific interface
        /// </summary>
        /// <param name="classTypeInfo">Low-level ITypeInfo wrapper representing the VBA project</param>
        /// <param name="interfaceIID">The interface IID</param>
        /// <returns>bool indicating whether the class does inherit the specified interface</returns>
        bool DoesClassImplementInterface(ITypeInfoWrapper classTypeInfo, Guid interfaceIID);

        /// <summary>
        /// Determines whether the specified VBA class implements one of several possible interfaces
        /// </summary>
        /// <param name="ide">Safe-com wrapper representing the VBE</param>
        /// <param name="projectName">VBA Project name, as declared in the VBE</param>
        /// <param name="className">Document class name, as declared in the VBA project</param>
        /// <param name="interfaceIIDs">An array of interface IIDs to check against</param>
        /// <param name="matchedIndex">on return indicates the index into interfaceIIDs that matched, or -1 if no match</param>
        /// <returns>bool indicating whether the class does inherit one of the specified interfaces</returns>
        bool DoesClassImplementInterface(IVBE ide, string projectName, string className, Guid[] interfaceIIDs, out int matchedIndex);

        /// <summary>
        /// Determines whether the specified VBA class implements one of several possible interfaces
        /// </summary>
        /// <param name="project">Safe-com wrapper representing the VBA project</param>
        /// <param name="className">Document class name, as declared in the VBA project</param>
        /// <param name="interfaceIIDs">An array of interface IIDs to check against</param>
        /// <param name="matchedIndex">on return indicates the index into interfaceIIDs that matched, or -1 if no match</param>
        /// <returns>bool indicating whether the class does inherit one of the specified interfaces</returns>
        bool DoesClassImplementInterface(IVBProject project, string className, Guid[] interfaceIIDs, out int matchedIndex);

        /// <summary>
        /// Determines whether the specified VBA class implements one of several possible interfaces
        /// </summary>
        /// <param name="projectTypeLib">Low-level ITypeLib wrapper representing the VBA project</param>
        /// <param name="className">Document class name, as declared in the VBA project</param>
        /// <param name="interfaceIIDs">An array of interface IIDs to check against</param>
        /// <param name="matchedIndex">on return indicates the index into interfaceIIDs that matched, or -1 if no match</param>
        /// <returns>bool indicating whether the class does inherit one of the specified interfaces</returns>
        bool DoesClassImplementInterface(ITypeLibWrapper projectTypeLib, string className, Guid[] interfaceIIDs, out int matchedIndex);

        /// <summary>
        /// Determines whether the specified VBA class implements one of several possible interfaces
        /// </summary>
        /// <param name="component">Safe-com wrapper representing the VBA component</param>
        /// <param name="interfaceIIDs">An array of interface IIDs to check against</param>
        /// <param name="matchedIndex">on return indicates the index into interfaceIIDs that matched, or -1 if no match</param>
        /// <returns>bool indicating whether the class does inherit one of the specified interfaces</returns>
        bool DoesClassImplementInterface(IVBComponent component, Guid[] interfaceIIDs, out int matchedIndex);

        /// <summary>
        /// Determines whether the specified VBA class implements one of several possible interfaces
        /// </summary>
        /// <param name="classTypeInfo">Low-level ITypeInfo wrapper representing the VBA project</param>
        /// <param name="interfaceIIDs">An array of interface IIDs to check against</param>
        /// <param name="matchedIndex">on return indicates the index into interfaceIIDs that matched, or -1 if no match</param>
        /// <returns>bool indicating whether the class does inherit one of the specified interfaces</returns>
        bool DoesClassImplementInterface(ITypeInfoWrapper classTypeInfo, Guid[] interfaceIIDs, out int matchedIndex);

        /// <summary>
        /// Returns the class progID of a control on a UserForm
        /// </summary>
        /// <param name="ide">Safe-com wrapper representing the VBE</param>
        /// <param name="projectName">VBA Project name, as declared in the VBE</param>
        /// <param name="userFormName">UserForm class name, as declared in the VBA project</param>
        /// <param name="controlName">Control name, as declared on the UserForm</param>
        /// <returns>string class progID of the specified control on a UserForm, e.g. "MSForms.CommandButton"</returns>
        string GetUserFormControlType(IVBE ide, string projectName, string userFormName, string controlName);

        /// <summary>
        /// Returns the class progID of a control on a UserForm
        /// </summary>
        /// <param name="project">Safe-com wrapper representing the VBA project</param>
        /// <param name="userFormName">UserForm class name, as declared in the VBA project</param>
        /// <param name="controlName">Control name, as declared on the UserForm</param>
        /// <returns>string class progID of the specified control on a UserForm, e.g. "MSForms.CommandButton"</returns>
        string GetUserFormControlType(IVBProject project, string userFormName, string controlName);

        /// <summary>
        /// Returns the class progID of a control on a UserForm
        /// </summary>
        /// <param name="projectTypeLib">Low-level ITypeLib wrapper representing the VBA project</param>
        /// <param name="userFormName">UserForm class name, as declared in the VBA project</param>
        /// <param name="controlName">Control name, as declared on the UserForm</param>
        /// <returns>string class progID of the specified control on a UserForm, e.g. "MSForms.CommandButton"</returns>
        string GetUserFormControlType(ITypeLibWrapper projectTypeLib, string userFormName, string controlName);

        /// <summary>
        /// Returns the class progID of a control on a UserForm
        /// </summary>
        /// <param name="ide">Safe-com wrapper representing the UserForm VBA component</param>
        /// <param name="controlName">Control name, as declared on the UserForm</param>
        /// <returns>string class progID of the specified control on a UserForm, e.g. "MSForms.CommandButton"</returns>
        string GetUserFormControlType(IVBComponent component, string controlName);

        /// <summary>
        /// Returns the class progID of a control on a UserForm
        /// </summary>
        /// <param name="userFormTypeInfo">Low-level ITypeLib wrapper representing the UserForm VBA component</param>
        /// <param name="controlName">Control name, as declared on the UserForm</param>
        /// <returns>string class progID of the specified control on a UserForm, e.g. "MSForms.CommandButton"</returns>
        string GetUserFormControlType(ITypeInfoWrapper userFormTypeInfo, string controlName);

        /// <summary>
        /// Returns the class progID of a control on a UserForm
        /// </summary>
        /// <param name="ide">Safe-com wrapper representing the VBE</param>
        /// <param name="projectName">VBA Project name, as declared in the VBE</param>
        /// <param name="documentClassName">Document class name, as declared in the VBA project</param>
        /// <param name="controlName">Control name, as declared on the UserForm</param>
        /// <returns>string class progID of the specified control on a UserForm, e.g. "MSForms.CommandButton"</returns>
        string GetDocumentClassControlType(IVBE ide, string projectName, string documentClassName, string controlName);

        /// <summary>
        /// Returns the class progID of a control on a UserForm
        /// </summary>
        /// <param name="project">Safe-com wrapper representing the VBA project</param>
        /// <param name="documentClassName">Document class name, as declared in the VBA project</param>
        /// <param name="controlName">Control name, as declared on the UserForm</param>
        /// <returns>string class progID of the specified control on a UserForm, e.g. "MSForms.CommandButton"</returns>
        string GetDocumentClassControlType(IVBProject project, string documentClassName, string controlName);

        /// <summary>
        /// Returns the class progID of a control on a UserForm
        /// </summary>
        /// <param name="projectTypeLib">Low-level ITypeLib wrapper representing the VBA project</param>
        /// <param name="documentClassName">Document class name, as declared in the VBA project</param>
        /// <param name="controlName">Control name, as declared on the UserForm</param>
        /// <returns>string class progID of the specified control on a UserForm, e.g. "MSForms.CommandButton"</returns>
        string GetDocumentClassControlType(ITypeLibWrapper projectTypeLib, string documentClassName, string controlName);

        /// <summary>
        /// Returns the class progID of a control on a UserForm
        /// </summary>
        /// <param name="ide">Safe-com wrapper representing the UserForm VBA component</param>
        /// <param name="controlName">Control name, as declared on the UserForm</param>
        /// <returns>string class progID of the specified control on a UserForm, e.g. "MSForms.CommandButton"</returns>
        string GetDocumentClassControlType(IVBComponent component, string controlName);

        /// <summary>
        /// Returns the class progID of a control on a UserForm
        /// </summary>
        /// <param name="documentClass">Low-level ITypeLib wrapper representing the UserForm VBA component</param>
        /// <param name="controlName">Control name, as declared on the UserForm</param>
        /// <returns>string class progID of the specified control on a UserForm, e.g. "MSForms.CommandButton"</returns>
        string GetDocumentClassControlType(ITypeInfoWrapper documentClass, string controlName);

        /// <summary>
        /// Retrieves the TYPEFLAGS of a VBA component (e.g. module/class), providing flags like TYPEFLAG_FCANCREATE, TYPEFLAG_FPREDECLID
        /// </summary>
        /// <param name="ide">Safe-com wrapper representing the VBE</param>
        /// <param name="projectName">The VBA project name</param>
        /// <param name="componentName">The name of the component (module/class etc) to get flags for</param>
        /// <returns>ComTypes.TYPEFLAGS flags from the ITypeInfo</returns>
        System.Runtime.InteropServices.ComTypes.TYPEFLAGS GetComponentTypeFlags(IVBE ide, string projectName, string componentName);

        /// <summary>
        /// Retrieves the TYPEFLAGS of a VBA component (e.g. module/class), providing flags like TYPEFLAG_FCANCREATE, TYPEFLAG_FPREDECLID
        /// </summary>
        /// <param name="project">Safe-com wrapper representing the VBA project</param>
        /// <param name="componentName">The name of the component (module/class etc) to get flags for</param>
        /// <returns>ComTypes.TYPEFLAGS flags from the ITypeInfo</returns>
        System.Runtime.InteropServices.ComTypes.TYPEFLAGS GetComponentTypeFlags(IVBProject project, string componentName);

        /// <summary>
        /// Retrieves the TYPEFLAGS of a VBA component (e.g. module/class), providing flags like TYPEFLAG_FCANCREATE, TYPEFLAG_FPREDECLID
        /// </summary>
        /// <param name="projectTypeLib">Low-level ITypeLib wrapper representing the VBA project</param>
        /// <param name="componentName">The name of the component (module/class etc) to get flags for</param>
        /// <returns>ComTypes.TYPEFLAGS flags from the ITypeInfo</returns>
        System.Runtime.InteropServices.ComTypes.TYPEFLAGS GetComponentTypeFlags(ITypeLibWrapper projectTypeLib, string componentName);

        /// <summary>
        /// Retrieves the TYPEFLAGS of a VBA component (e.g. module/class), providing flags like TYPEFLAG_FCANCREATE, TYPEFLAG_FPREDECLID
        /// </summary>
        /// <param name="component">Safe-com wrapper representing the VBA component to get flags for</param>
        /// <returns>ComTypes.TYPEFLAGS flags from the ITypeInfo</returns>
        System.Runtime.InteropServices.ComTypes.TYPEFLAGS GetComponentTypeFlags(IVBComponent component);

        /// <summary>
        /// Retrieves the TYPEFLAGS of a VBA component (e.g. module/class), providing flags like TYPEFLAG_FCANCREATE, TYPEFLAG_FPREDECLID
        /// </summary>
        /// <param name="componentTypeInfo">Low-level ITypeInfo wrapper representing the VBA component to get flags for</param>
        /// <returns>ComTypes.TYPEFLAGS flags from the ITypeInfo</returns>
        System.Runtime.InteropServices.ComTypes.TYPEFLAGS GetComponentTypeFlags(ITypeInfoWrapper componentTypeInfo);

        /// <summary>
        /// Returns a TypeLibReference object containing information about the specified VBA project reference
        /// </summary>
        /// <param name="ide">Safe-com wrapper representing the VBE</param>
        /// <param name="projectName">VBA Project name, as declared in the VBE</param>
        /// <param name="referenceIdx">Index into the references collection</param>
        /// <returns>TypeLibReference containing information about the specified VBA project reference</returns>
        ITypeLibReference GetReferenceInfo(IVBE ide, string projectName, int referenceIdx);

        /// <summary>
        /// Returns a TypeLibReference object containing information about the specified VBA project reference
        /// </summary>
        /// <param name="project">Safe-com wrapper representing the VBA project</param>
        /// <param name="referenceIdx">Index into the references collection</param>
        /// <returns>TypeLibReference containing information about the specified VBA project reference</returns>
        ITypeLibReference GetReferenceInfo(IVBProject project, int referenceIdx);

        /// <summary>
        /// Returns a TypeLibReference object containing information about the specified VBA project reference
        /// </summary>
        /// <param name="projectTypeLib">Low-level ITypeLib wrapper representing the VBA project</param>
        /// <param name="referenceIdx">Index into the references collection</param>
        /// <returns>TypeLibReference containing information about the specified VBA project reference</returns>
        ITypeLibReference GetReferenceInfo(ITypeLibWrapper projectTypeLib, int referenceIdx);

        /// <summary>
        /// Returns a TypeLibReference object containing information about the specified VBA project reference
        /// </summary>
        /// <param name="project">Safe-com wrapper representing the VBA project</param>
        /// <param name="vbeReference">Safe-com wrapper representing the VBA project reference</param>
        /// <returns>TypeLibReference containing information about the specified VBA project reference</returns>
        ITypeLibReference GetReferenceInfo(IVBProject project, IReference vbeReference);

        /// <summary>
        /// Documents the type libaries of all loaded VBA projects
        /// </summary>
        /// <param name="ide">Safe-com wrapper representing the VBE</param>
        /// <returns>text document, in a non-standard format, useful for debugging purposes</returns>
        string DocumentAll(IVBE ide);

        /// <summary>
        /// Documents the type libary of a single VBA project
        /// </summary>
        /// <param name="project">Safe-com wrapper representing the VBA project</param>
        /// <returns>text document, in a non-standard format, useful for debugging purposes</returns>
        string DocumentAll(IVBProject project);

        /// <summary>
        /// Documents the type libary of a single VBA project
        /// </summary>
        /// <param name="projectTypeLib">Low-level ITypeLib wrapper representing the VBA project</param>
        /// <returns>text document, in a non-standard format, useful for debugging purposes</returns>
        string DocumentAll(ITypeLibWrapper projectTypeLib);
    }
}