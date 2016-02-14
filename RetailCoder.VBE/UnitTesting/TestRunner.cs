using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Vbe.Interop;

namespace Rubberduck.UnitTesting
{
    //[ComVisible(true)]
    //[ComDefaultInterface(typeof(ITestRunner))]
    //[Guid(ClassId)]
    //[ProgId(ProgId)]
    //public class TestRunner : ITestRunner
    //{
    //    private const string ClassId = "C46C141F-30C8-4A8A-B84B-EC04F1CC559B";
    //    private const string ProgId = "Rubberduck.TestRunner";

    //    private readonly TestEngine _engine = new TestEngine();

    //    public TestRunner()
    //    {
    //        _engine.MethodCleanup += TestEngineMethodCleanup;
    //        _engine.MethodInitialize += TestEngineMethodInitialize;
    //        _engine.ModuleCleanup += TestEngine_ModuleCleanup;
    //        _engine.ModuleInitialize += TestEngine_ModuleInitialize;
    //    }

    //    private void TestEngineMethodCleanup(object sender, TestModuleEventArgs e)
    //    {
    //        var module = e.QualifiedModuleName.Component.CodeModule;
    //        module.Parent.RunMethodsWithAttribute<TestCleanupAttribute>();
    //    }

    //    private void TestEngineMethodInitialize(object sender, TestModuleEventArgs e)
    //    {
    //        var module = e.QualifiedModuleName.Component.CodeModule;
    //        module.Parent.RunMethodsWithAttribute<TestInitializeAttribute>();
    //    }

    //    private void TestEngine_ModuleCleanup(object sender, TestModuleEventArgs e)
    //    {
    //        var module = e.QualifiedModuleName.Component.CodeModule;
    //        module.Parent.RunMethodsWithAttribute<ModuleCleanupAttribute>();
    //    }

    //    private void TestEngine_ModuleInitialize(object sender, TestModuleEventArgs e)
    //    {
    //        var module = e.QualifiedModuleName.Component.CodeModule;
    //        module.Parent.RunMethodsWithAttribute<ModuleInitializeAttribute>();
    //    }

    //    private void LoadAllTests(VBE vbe)
    //    {
    //        _engine.AllTests = vbe.VBProjects
    //                        .Cast<VBProject>().Where(project => project.Protection != vbext_ProjectProtection.vbext_pp_locked)
    //                        .SelectMany(project => project.TestMethods())
    //                        .ToDictionary(test => test, test => _engine.AllTests.ContainsKey(test) ? _engine.AllTests[test] : null);
    //    }

    //    public string RunAllTests(VBE vbe, string outputFilePath = null)
    //    {
    //        LoadAllTests(vbe);
    //        _engine.Run(_engine.AllTests.Keys);

    //        var results = OutputToString();

    //        if (outputFilePath != null)
    //        {
    //            OutputToFile(outputFilePath, results);
    //        }

    //        return results;
    //    }

    //    private string OutputToString()
    //    {
    //        var builder = new StringBuilder();
    //        builder.AppendLine("Rubberduck Unit Tests - " + string.Format("{0:G}", DateTime.Now));

    //        foreach (var result in _engine.AllTests)
    //        {
    //            var item = string.Format("{0}\t{1}ms\t{2} {3}", result.Key.QualifiedMemberName, result.Value.Duration, result.Value.Outcome, result.Value.Output);
    //            builder.AppendLine(item);
    //        }

    //        return builder.ToString();
    //    }

    //    private void OutputToFile(string path, string results)
    //    {
    //        File.AppendAllText(path, results);
    //    }
    //}
}
