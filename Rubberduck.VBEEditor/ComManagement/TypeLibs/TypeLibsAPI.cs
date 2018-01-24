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
                return typeLibs.FindTypeLib(projectName).CompileProject();
            }
        }
        public static bool CompileProject(IVBProject project)
        {
            using (var typeLib = TypeLibWrapper.FromVBProject(project))
            {
                return typeLib.CompileProject();
            }
        }

        // Compile a module in a VBE project, returning success (true)/failure (false)
        // NOTE: This will only return success if ALL modules that this module depends on compile successfully
        public static bool CompileComponent(IVBE ide, string projectName, string componentName)
        {
            using (var typeLibs = new VBETypeLibsAccessor(ide))
            {
                return typeLibs.FindTypeLib(projectName).FindTypeInfo(componentName).CompileComponent();
            }
        }
        public static bool CompileComponent(IVBProject project, string componentName)
        {
            using (var typeLib = TypeLibWrapper.FromVBProject(project))
            {
                return typeLib.FindTypeInfo(componentName).CompileComponent();
            }
        }

        public static object ExecuteCode(IVBE ide, string projectName, string standardModuleName, string procName, object[] args = null)
        {
            using (var typeLibs = new VBETypeLibsAccessor(ide))
            {
                return typeLibs.FindTypeLib(projectName).FindTypeInfo(standardModuleName)
                    .StdModExecute(procName, Reflection.BindingFlags.InvokeMethod, args);
            }
        }
        public static object ExecuteCode(IVBProject project, string standardModuleName, string procName, object[] args = null)
        {
            using (var typeLib = TypeLibWrapper.FromVBProject(project))
            {
                return typeLib.FindTypeInfo(standardModuleName)
                    .StdModExecute(procName, Reflection.BindingFlags.InvokeMethod, args);
            }
        }

        // returns the raw conditional arguments string, e.g. "foo = 1 : bar = 2"
        public static string GetProjectConditionalCompilationArgsRaw(IVBE ide, string projectName)
        {
            using (var typeLibs = new VBETypeLibsAccessor(ide))
            {
                return typeLibs.FindTypeLib(projectName).ConditionalCompilationArguments;
            }
        }
        public static string GetProjectConditionalCompilationArgsRaw(IVBProject project)
        {
            using (var typeLib = TypeLibWrapper.FromVBProject(project))
            {
                return typeLib.ConditionalCompilationArguments;
            }
        }

        // return the parsed conditional arguments string as a Dictionary<string, string>
        public static Dictionary<string, string> GetProjectConditionalCompilationArgs(IVBE ide, string projectName)
        {
            var args = GetProjectConditionalCompilationArgsRaw(ide, projectName);

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
        public static Dictionary<string, string> GetProjectConditionalCompilationArgs(IVBProject project)
        {
            string args = GetProjectConditionalCompilationArgsRaw(project);

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
                typeLibs.FindTypeLib(projectName).ConditionalCompilationArguments = newConditionalArgs;
            }
        }
        public static void SetProjectConditionalCompilationArgsRaw(IVBProject project, string newConditionalArgs)
        {
            using (var typeLib = TypeLibWrapper.FromVBProject(project))
            {
                typeLib.ConditionalCompilationArguments = newConditionalArgs;
            }
        }

        // sets the conditional arguments string via a Dictionary<string, string>
        public static void SetProjectConditionalCompilationArgs(IVBE ide, string projectName, Dictionary<string, string> newConditionalArgs)
        {
            var rawArgsString = string.Join(" : ", newConditionalArgs.Select(x => x.Key + " = " + x.Value));
            SetProjectConditionalCompilationArgsRaw(ide, projectName, rawArgsString);
        }
        public static void SetProjectConditionalCompilationArgs(IVBProject project, Dictionary<string, string> newConditionalArgs)
        {
            var rawArgsString = string.Join(" : ", newConditionalArgs.Select(x => x.Key + " = " + x.Value));
            SetProjectConditionalCompilationArgsRaw(project, rawArgsString);
        }

        public static bool IsExcelWorkbook(IVBE ide, string projectName, string className)
        {
            return DoesClassImplementInterface(ide, projectName, className, "Excel", "_Workbook");
        }
        public static bool IsExcelWorkbook(IVBProject project, string className)
        {
            return DoesClassImplementInterface(project, className, "Excel", "_Workbook");
        }

        public static bool IsExcelWorksheet(IVBE ide, string projectName, string className)
        {
            return DoesClassImplementInterface(ide, projectName, className, "Excel", "_Worksheet");
        }
        public static bool IsExcelWorksheet(IVBProject project, string className)
        {
            return DoesClassImplementInterface(project, className, "Excel", "_Worksheet");
        }

        public static bool DoesClassImplementInterface(IVBE ide, string projectName, string className, string typeLibName, string interfaceName)
        {
            using (var typeLibs = new VBETypeLibsAccessor(ide))
            {
                return typeLibs.FindTypeLib(projectName).FindTypeInfo(className).DoesImplement(typeLibName, interfaceName);
            }
        }
        public static bool DoesClassImplementInterface(IVBProject project, string className, string typeLibName, string interfaceName)
        {
            using (var typeLib = TypeLibWrapper.FromVBProject(project))
            {
                return typeLib.FindTypeInfo(className).DoesImplement(typeLibName, interfaceName);
            }
        }

        public static bool DoesClassImplementInterface(IVBE ide, string projectName, string className, Guid interfaceIID)
        {
            using (var typeLibs = new VBETypeLibsAccessor(ide))
            {
                return typeLibs.FindTypeLib(projectName).FindTypeInfo(className).DoesImplement(interfaceIID);
            }
        }
        public static bool DoesClassImplementInterface(IVBProject project, string className, Guid interfaceIID)
        {
            using (var typeLib = TypeLibWrapper.FromVBProject(project))
            {
                return typeLib.FindTypeInfo(className).DoesImplement(interfaceIID);
            }
        }

        public static string GetUserFormControlType(IVBE ide, string projectName, string userFormName, string controlName)
        {
            using (var typeLibs = new VBETypeLibsAccessor(ide))
            {
                return typeLibs.FindTypeLib(projectName).FindTypeInfo(userFormName)
                        .GetImplementedTypeInfo("FormItf").GetControlType(controlName).Name;
            }
        }
        public static string GetUserFormControlType(IVBProject project, string userFormName, string controlName)
        {
            using (var typeLib = TypeLibWrapper.FromVBProject(project))
            {
                return typeLib.FindTypeInfo(userFormName)
                        .GetImplementedTypeInfo("FormItf").GetControlType(controlName).Name;
            }
        }

        public static string DocumentAll(IVBE ide)
        {
            using (var typeLibs = new VBETypeLibsAccessor(ide))
            {
                var documenter = new TypeLibDocumenter();

                foreach (var typeLib in typeLibs)
                {
                    documenter.AddTypeLib(typeLib);
                }

                return documenter.ToString();
            }
        }
        public static string DocumentAll(IVBProject project)
        {
            using (var typeLib = TypeLibWrapper.FromVBProject(project))
            {
                var documenter = new TypeLibDocumenter();

                documenter.AddTypeLib(typeLib);
                
                return documenter.ToString();
            }
        }
    }
}
