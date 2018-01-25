using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using Rubberduck.VBEditor.ComManagement.TypeLibs;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Reflection = System.Reflection;
using System.Linq;

namespace Rubberduck.VBEditor.ComManagement.TypeLibsAPI
{
    // DEBUG TEMPORARY CLASS TO ALLOW ACCESS TO TypeLibsAPI from VBA
    [System.Runtime.InteropServices.ComVisible(true)]
    public class TypeLibsAPI_Object
    {
        IVBE _ide;
        public TypeLibsAPI_Object(IVBE ide) 
            => _ide = ide;

        public bool CompileProject(string projectName) 
            => VBETypeLibsAPI.CompileProject(_ide, projectName);
        public bool CompileComponent(string projectName, string componentName) 
            => VBETypeLibsAPI.CompileComponent(_ide, projectName, componentName);
        public object ExecuteCode(string projectName, string standardModuleName, string procName) 
            => VBETypeLibsAPI.ExecuteCode(_ide, projectName, standardModuleName, procName);
        public string GetProjectConditionalCompilationArgsRaw(string projectName)
            => VBETypeLibsAPI.GetProjectConditionalCompilationArgsRaw(_ide, projectName);
        public void SetProjectConditionalCompilationArgsRaw(string projectName, string newConditionalArgs)
            => VBETypeLibsAPI.SetProjectConditionalCompilationArgsRaw(_ide, projectName, newConditionalArgs);
        public bool DoesClassImplementInterface(string projectName, string className, string typeLibName, string interfaceName) 
            => VBETypeLibsAPI.DoesClassImplementInterface(_ide, projectName, className, typeLibName, interfaceName);
        public string GetUserFormControlType(string projectName, string userFormName, string controlName) 
            => VBETypeLibsAPI.GetUserFormControlType(_ide, projectName, userFormName, controlName);
        public string DocumentAll() 
            => VBETypeLibsAPI.DocumentAll(_ide);
    }

    public class VBETypeLibsAPI
    {
        // Compile the project, returning success (true)/failure (false)
        public static bool CompileProject(IVBE ide, string projectName)
        {
            using (var typeLibs = new VBETypeLibsAccessor(ide))
            {
                return CompileProject(typeLibs.Get(projectName));
            }
        }
        public static bool CompileProject(IVBProject project)
        {
            using (var typeLib = TypeLibWrapper.FromVBProject(project))
            {
                return CompileProject(typeLib);
            }
        }
        public static bool CompileProject(TypeLibWrapper projectTypeLib)
        {
            return projectTypeLib.CompileProject();
        }

        // Compile a module in a VBE project, returning success (true)/failure (false)
        // NOTE: This will only return success if ALL modules that this module depends on compile successfully
        public static bool CompileComponent(IVBE ide, string projectName, string componentName)
        {
            using (var typeLibs = new VBETypeLibsAccessor(ide))
            {
                return CompileComponent(typeLibs.Get(projectName), componentName);
            }
        }
        public static bool CompileComponent(IVBProject project, string componentName)
        {
            using (var typeLib = TypeLibWrapper.FromVBProject(project))
            {
                return CompileComponent(typeLib, componentName);
            }
        }
        public static bool CompileComponent(TypeLibWrapper projectTypeLib, string componentName)
        {
            return CompileComponent(projectTypeLib.TypeInfos.Get(componentName));
        }
        public static bool CompileComponent(IVBComponent component)
        {
            return CompileComponent(component.ParentProject, component.Name);
        }
        public static bool CompileComponent(TypeInfoWrapper componentTypeInfo)
        {
            return componentTypeInfo.CompileComponent();
        }
        
        public static object ExecuteCode(IVBE ide, string projectName, string standardModuleName, string procName, object[] args = null)
        {
            using (var typeLibs = new VBETypeLibsAccessor(ide))
            {
                return ExecuteCode(typeLibs.Get(projectName), standardModuleName, procName, args);
            }
        }
        public static object ExecuteCode(IVBProject project, string standardModuleName, string procName, object[] args = null)
        {
            using (var typeLib = TypeLibWrapper.FromVBProject(project))
            {
                return ExecuteCode(typeLib, standardModuleName, procName, args);
            }
        }
        public static object ExecuteCode(TypeLibWrapper projectTypeLib, string standardModuleName, string procName, object[] args = null)
        {
            return ExecuteCode(projectTypeLib.TypeInfos.Get(standardModuleName), procName, args);
        }
        public static object ExecuteCode(IVBComponent component, string procName, object[] args = null)
        {
            return ExecuteCode(component.ParentProject, component.Name, procName, args);
        }
        public static object ExecuteCode(TypeInfoWrapper standardModuleTypeInfo, string procName, object[] args = null)
        {
            return standardModuleTypeInfo.StdModExecute(procName, Reflection.BindingFlags.InvokeMethod, args);
        }
        
        // returns the raw conditional arguments string, e.g. "foo = 1 : bar = 2"
        public static string GetProjectConditionalCompilationArgsRaw(IVBE ide, string projectName)
        {
            using (var typeLibs = new VBETypeLibsAccessor(ide))
            {
                return GetProjectConditionalCompilationArgsRaw(typeLibs.Get(projectName));
            }
        }
        public static string GetProjectConditionalCompilationArgsRaw(IVBProject project)
        {
            using (var typeLib = TypeLibWrapper.FromVBProject(project))
            {
                return GetProjectConditionalCompilationArgsRaw(typeLib);
            }
        }
        public static string GetProjectConditionalCompilationArgsRaw(TypeLibWrapper projectTypeLib)
        {
            return projectTypeLib.ConditionalCompilationArguments;
        }

        // return the parsed conditional arguments string as a Dictionary<string, string>
        public static Dictionary<string, string> GetProjectConditionalCompilationArgs(IVBE ide, string projectName)
        {
            using (var typeLibs = new VBETypeLibsAccessor(ide))
            {
                return GetProjectConditionalCompilationArgs(typeLibs.Get(projectName));
            }
        }
        public static Dictionary<string, string> GetProjectConditionalCompilationArgs(IVBProject project)
        {
            using (var typeLib = TypeLibWrapper.FromVBProject(project))
            {
                return GetProjectConditionalCompilationArgs(typeLib);
            }
        }
        public static Dictionary<string, string> GetProjectConditionalCompilationArgs(TypeLibWrapper projectTypeLib)
        {
            // FIXME move dictionary stuff into the lower API here
            string args = GetProjectConditionalCompilationArgsRaw(projectTypeLib);

            if (args.Length > 0)
            {
                string[] argsArray = args.Split(new[] { ':' });
                return argsArray.Select(item => item.Split('=')).ToDictionary(s => s[0], s => s[1]);
            }
            else
            {
                return new Dictionary<string, string>();
            }
        }

        // sets the raw conditional arguments string, e.g. "foo = 1 : bar = 2"
        public static void SetProjectConditionalCompilationArgsRaw(IVBE ide, string projectName, string newConditionalArgs)
        {
            using (var typeLibs = new VBETypeLibsAccessor(ide))
            {
                SetProjectConditionalCompilationArgsRaw(typeLibs.Get(projectName), newConditionalArgs);
            }
        }
        public static void SetProjectConditionalCompilationArgsRaw(IVBProject project, string newConditionalArgs)
        {
            using (var typeLib = TypeLibWrapper.FromVBProject(project))
            {
                SetProjectConditionalCompilationArgsRaw(typeLib, newConditionalArgs);
            }
        }
        public static void SetProjectConditionalCompilationArgsRaw(TypeLibWrapper projectTypeLib, string newConditionalArgs)
        {
            projectTypeLib.ConditionalCompilationArguments = newConditionalArgs;
        }

        // sets the conditional arguments string via a Dictionary<string, string>
        public static void SetProjectConditionalCompilationArgs(IVBE ide, string projectName, Dictionary<string, string> newConditionalArgs)
        {
            using (var typeLibs = new VBETypeLibsAccessor(ide))
            {
                SetProjectConditionalCompilationArgs(typeLibs.Get(projectName), newConditionalArgs);
            }
        }
        public static void SetProjectConditionalCompilationArgs(IVBProject project, Dictionary<string, string> newConditionalArgs)
        {
            using (var typeLib = TypeLibWrapper.FromVBProject(project))
            {
                SetProjectConditionalCompilationArgs(typeLib, newConditionalArgs);
            }
        }
        public static void SetProjectConditionalCompilationArgs(TypeLibWrapper projectTypeLib, Dictionary<string, string> newConditionalArgs)
        {
            // FIXME move dictionary stuff into the lower API here
            var rawArgsString = string.Join(" : ", newConditionalArgs.Select(x => x.Key + " = " + x.Value));
            SetProjectConditionalCompilationArgsRaw(projectTypeLib, rawArgsString);
        }

        public static bool IsExcelWorkbook(IVBE ide, string projectName, string className)
        {
            using (var typeLibs = new VBETypeLibsAccessor(ide))
            {
                return IsExcelWorkbook(typeLibs.Get(projectName), className);
            }
        }
        public static bool IsExcelWorkbook(IVBProject project, string className)
        {
            using (var typeLib = TypeLibWrapper.FromVBProject(project))
            {
                return IsExcelWorkbook(typeLib, className);
            }
        }
        public static bool IsExcelWorkbook(IVBComponent component)
        {
            return IsExcelWorkbook(component.ParentProject, component.Name);
        }
        public static bool IsExcelWorkbook(TypeLibWrapper projectTypeLib, string className)
        {
            return DoesClassImplementInterface(projectTypeLib, className, "Excel", "_Workbook");
        }

        public static bool IsExcelWorksheet(IVBE ide, string projectName, string className)
        {
            using (var typeLibs = new VBETypeLibsAccessor(ide))
            {
                return IsExcelWorksheet(typeLibs.Get(projectName), className);
            }
        }
        public static bool IsExcelWorksheet(IVBProject project, string className)
        {
            using (var typeLib = TypeLibWrapper.FromVBProject(project))
            {
                return IsExcelWorksheet(typeLib, className);
            }
        }
        public static bool IsExcelWorksheet(IVBComponent component)
        {
            return IsExcelWorksheet(component.ParentProject, component.Name);
        }
        public static bool IsExcelWorksheet(TypeLibWrapper projectTypeLib, string className)
        {
            return DoesClassImplementInterface(projectTypeLib, className, "Excel", "_Worksheet");
        }
        
        public static bool DoesClassImplementInterface(IVBE ide, string projectName, string className, string typeLibName, string interfaceName)
        {
            using (var typeLibs = new VBETypeLibsAccessor(ide))
            {
                return DoesClassImplementInterface(typeLibs.Get(projectName), className, typeLibName, interfaceName);
            }
        }
        public static bool DoesClassImplementInterface(IVBProject project, string className, string typeLibName, string interfaceName)
        {
            using (var typeLib = TypeLibWrapper.FromVBProject(project))
            {
                return DoesClassImplementInterface(typeLib.TypeInfos.Get(className), typeLibName, interfaceName);
            }
        }
        public static bool DoesClassImplementInterface(TypeLibWrapper projectTypeLib, string className, string typeLibName, string interfaceName)
        {
            return DoesClassImplementInterface(projectTypeLib.TypeInfos.Get(className), typeLibName, interfaceName);
        }
        public static bool DoesClassImplementInterface(IVBComponent component, string typeLibName, string interfaceName)
        {
            return DoesClassImplementInterface(component.ParentProject, component.Name, typeLibName, interfaceName);
        }
        public static bool DoesClassImplementInterface(TypeInfoWrapper classTypeInfo, string typeLibName, string interfaceName)
        {
            return classTypeInfo.ImplementedInterfaces.DoesImplement(typeLibName, interfaceName);
        }        

        public static bool DoesClassImplementInterface(IVBE ide, string projectName, string className, Guid interfaceIID)
        {
            using (var typeLibs = new VBETypeLibsAccessor(ide))
            {
                return DoesClassImplementInterface(typeLibs.Get(projectName), className, interfaceIID);
            }
        }
        public static bool DoesClassImplementInterface(IVBProject project, string className, Guid interfaceIID)
        {
            using (var typeLib = TypeLibWrapper.FromVBProject(project))
            {
                return DoesClassImplementInterface(typeLib.TypeInfos.Get(className), interfaceIID);
            }
        }
        public static bool DoesClassImplementInterface(TypeLibWrapper projectTypeLib, string className, Guid interfaceIID)
        {
            return DoesClassImplementInterface(projectTypeLib.TypeInfos.Get(className), interfaceIID);
        }
        public static bool DoesClassImplementInterface(IVBComponent component, Guid interfaceIID)
        {
            return DoesClassImplementInterface(component.ParentProject, component.Name, interfaceIID);
        }
        public static bool DoesClassImplementInterface(TypeInfoWrapper classTypeInfo, Guid interfaceIID)
        {
            return classTypeInfo.ImplementedInterfaces.DoesImplement(interfaceIID);
        }

        public static string GetUserFormControlType(IVBE ide, string projectName, string userFormName, string controlName)
        {
            using (var typeLibs = new VBETypeLibsAccessor(ide))
            {
                return GetUserFormControlType(typeLibs.Get(projectName), userFormName, controlName);
            }
        }
        public static string GetUserFormControlType(IVBProject project, string userFormName, string controlName)
        {
            using (var typeLib = TypeLibWrapper.FromVBProject(project))
            {
                return GetUserFormControlType(typeLib, userFormName, controlName);
            }
        }
        public static string GetUserFormControlType(TypeLibWrapper projectTypeLib, string userFormName, string controlName)
        {
            return GetUserFormControlType(projectTypeLib.TypeInfos.Get(userFormName), controlName);
        }
        public static string GetUserFormControlType(IVBComponent component, string controlName)
        {
            return GetUserFormControlType(component.ParentProject, component.Name, controlName);
        }
        public static string GetUserFormControlType(TypeInfoWrapper userFormTypeInfo, string controlName)
        {
            return userFormTypeInfo.ImplementedInterfaces.Get("FormItf").GetControlType(controlName).Name;
        }

        public static string DocumentAll(IVBE ide)
        {
            using (var typeLibs = new VBETypeLibsAccessor(ide))
            {
                var output = new StringLineBuilder();

                foreach (var typeLib in typeLibs)
                {
                    typeLib.Document(output);
                }
                return output.ToString();
            }
        }
        public static string DocumentAll(IVBProject project)
        {
            using (var typeLib = TypeLibWrapper.FromVBProject(project))
            {
                return DocumentAll(typeLib);
            }
        }
        public static string DocumentAll(TypeLibWrapper projectTypeLib)
        {
            var output = new StringLineBuilder();
            projectTypeLib.Document(output);
            return output.ToString();
        }
    }
}
