using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using Org.Benf.OleWoo;
using Org.Benf.OleWoo.Typelib;

namespace Rubberduck.Deployment.Build.IdlGeneration
{
    public class IdlGenerator
    {
        public string GenerateIdl(string assemblyPath)
        {
            try
            {
                if (File.Exists(assemblyPath))
                {
                    var assembly = Assembly.LoadFrom(assemblyPath);
                    return GenerateIdl(assembly);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                throw;
            }

            return string.Empty;
        }

        public string GenerateIdl(Assembly assembly)
        {
            var converter = new TypeLibConverter();
            var sink = new TypeLibExporterNotifySink();
            var lib = (ITypeLib) converter.ConvertAssemblyToTypeLib(assembly, assembly.GetName().Name,
                TypeLibExporterFlags.None, sink);
            var formatter = new PlainIDLFormatter();
            var owLib = new OWTypeLib(lib);
            owLib.Listeners.Add(new IdlListener());

            formatter.AddString("midl_pragma warning( disable: 2400 2401 ) ");
            formatter.NewLine();

            owLib.BuildIDLInto(formatter);
            return formatter.ToString();
        }
    }

    public class TypeLibExporterNotifySink : ITypeLibExporterNotifySink
    {
        public void ReportEvent(ExporterEventKind eventKind, int eventCode, string eventMsg)
        {
            // no-op
        }

        public object ResolveRef(Assembly assembly)
        {
            var converter = new TypeLibConverter();
            var lib = converter.ConvertAssemblyToTypeLib(assembly, assembly.GetName().Name,
                TypeLibExporterFlags.None, this);
            return lib;
        }
    }
}
