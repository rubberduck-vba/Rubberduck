using System;
using System.IO;
using System.Linq;
using System.Text;
using Microsoft.Build.Framework;
using Microsoft.Build.Tasks;
using Microsoft.VisualStudio.Setup.Configuration;
using Rubberduck.Deployment.Build.Builders;
using Rubberduck.Deployment.Build.IdlGeneration;
using Rubberduck.Deployment.Build.Structs;
using Rubberduck.Deployment.Build.Writers;

namespace Rubberduck.Deployment.Build
{
    internal struct DllFileParameters
    {
        public string DllFile { get; }
        public string IdlFile { get; }
        public string Tlb32File { get; }
        public string Tlb64File { get; }
        public string SourceDll { get; }
        public string TargetDll { get; }
        public string SourceTlb32 { get; }
        public string TargetTlb32 { get; }
        public string SourceTlb64 { get; }
        public string TargetTlb64 { get; }
        public string DllXml { get; }
        public string TlbXml { get; }

        internal DllFileParameters(string dllFile, string sourceDir, string targetDir)
        {
            DllFile = dllFile;
            IdlFile = DllFile.Replace(".dll", ".idl");
            Tlb32File = DllFile.Replace(".dll", ".x32.tlb");
            Tlb64File = DllFile.Replace(".dll", ".x64.tlb");

            SourceDll = Path.Combine(sourceDir, DllFile);
            TargetDll = Path.Combine(targetDir, DllFile);
            SourceTlb32 = Path.Combine(sourceDir, Tlb32File);
            TargetTlb32 = Path.Combine(targetDir, Tlb32File);
            SourceTlb64 = Path.Combine(sourceDir, Tlb64File);
            TargetTlb64 = Path.Combine(targetDir, Tlb64File);
            DllXml = TargetDll + ".xml";
            TlbXml = TargetTlb32 + ".xml";
        }
    }

    public class PostBuildTask : ITask
    {
        public IBuildEngine BuildEngine { get; set; }
        public ITaskHost HostObject { get; set; }

        public string Config { get; set; }
        public string NetToolsDir { get; set; }
        public string WixToolsDir { get; set; }
        public string SourceDir { get; set; }
        public string TargetDir { get; set; }
        public string ProjectDir { get; set; }
        public string IncludeDir { get; set; }
        public string FilesToExtract { get; set; }

        private string RegFilePath =>
            Path.Combine(Path.Combine(ProjectDir, "LocalRegistryEntries"), "DebugRegistryEntries.reg");

        private string _rootPath;
        private string _batchPath;

        public bool Execute()
        {
            var result = true;
            try
            {
                this.LogCustom(FormatParameterList());

                CleanOldImports(ProjectDir);

                var dllFiles = FilesToExtract.Split(new[] {"|"}, StringSplitOptions.RemoveEmptyEntries);

                var message = SetVCToolsPath()
                    ? "No C++ build tools found; using tlbexp.exe to generate TLBs."
                    : "C++ build tools found; using midl.exe to generate TLBs.";
                this.LogCustom(message);
                
                foreach (var dllFile in dllFiles)
                {
                    ProcessDll(dllFile);
                }

                if (Config == "Debug")
                {
                    UpdateAddInRegistration();
                }
            }
            catch (Exception ex)
            {
                this.LogError(ex);
                result = false;
            }

            return result;
        }

        private string FormatParameterList()
        {
            var sb = new StringBuilder();
            sb.AppendLine("Parameters provided:");
            sb.AppendLine($"          Config: {Config}");
            sb.AppendLine($"     NetToolsDir: {NetToolsDir}");
            sb.AppendLine($"     WixToolsDir: {WixToolsDir}");
            sb.AppendLine($"       SourceDir: {SourceDir}");
            sb.AppendLine($"       TargetDir: {TargetDir}");
            sb.AppendLine($"      ProjectDir: {ProjectDir}");
            sb.AppendLine($"     IncludesDir: {IncludeDir}");
            sb.AppendLine($"  FilesToExtract: {FilesToExtract}");
            sb.AppendLine(string.Empty);

            return sb.ToString();
        }

        private bool SetVCToolsPath()
        {
            var configuration = new SetupConfiguration();
            var enumInstances = configuration.EnumInstances();
            int fetched;
            var instances = new ISetupInstance[1];
            do
            {
                enumInstances.Next(1, instances, out fetched);

                if (fetched <= 0)
                {
                    continue;
                }

                var instance = (ISetupInstance2) instances[0];
                var packages = instance.GetPackages();
                foreach (var package in packages)
                {
                    if (package.GetId() != "Microsoft.VisualStudio.Component.VC.Tools.x86.x64")
                    {
                        continue;
                    }

                    _rootPath = instance.ResolvePath();
                    _batchPath = Directory.GetFiles(_rootPath, "VsDevCmd.bat", SearchOption.AllDirectories)
                        .FirstOrDefault();

                    return true;
                }
            } while (fetched > 0);

            return false;
        }

        private void CleanOldImports(string dir)
        {
            this.LogCustom("Cleaing out old imports");

            var files = Directory.GetFiles(dir, "DebugRegistryEntries.reg.imported_*.txt").OrderByDescending(f => f);
            var i = 0;
            foreach (var file in files)
            {
                if (i > 10)
                {
                    this.LogCustom($"Deleting {file}");
                    File.Delete(file);
                }

                i++;
            }
        }

        private void UpdateAddInRegistration()
        {
            this.LogCustom("Updating addin registration...");
            var addInRegFile = Path.Combine(Path.GetDirectoryName(RegFilePath), "RubberduckAddinRegistry.reg");
            var command = $"reg.exe import \"{addInRegFile}";
            ExecuteTask(command);
        }

        private void ProcessDll(string file)
        {
            this.LogCustom($"Processing {file}...");

            var parameters = new DllFileParameters(file, SourceDir, TargetDir);
            
            if (_batchPath == string.Empty)
            {
                this.LogCustom("Compiling with tlbexp...");
                CompileWithTlbExp(parameters);
            }
            else
            {
                this.LogCustom("Compiling with midl...");
                CreateIdlFile(parameters);
                CompileWithMidl(parameters);
            }

            this.LogCustom("Extracting metadata using WiX...");
            HarvestMetadataWithWix(parameters);

            this.LogCustom("Building registry entries...");
            var entries = BuildRegistryEntriesFromMetadata(parameters);

            this.LogCustom("Creating InnoSetup registry entries...");
            CreateInnoSetupRegistryEntries(entries, parameters);

            if (Config != "Debug")
            {
                return;
            }

            RemovePreviousDebugRegistration();
            UpdateDebugRegistration(entries, parameters);
        }

        private void CreateIdlFile(DllFileParameters parameters)
        {
            var generator = new IdlGenerator();
            var idl = generator.GenerateIdl(parameters.SourceDll);
            // Encoding must be UTF8 BOM
            File.WriteAllText(parameters.IdlFile, idl, new UTF8Encoding(true));
        }

        private void CompileWithMidl(DllFileParameters parameters)
        {
            var targetPath = TargetDir.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            var command = _batchPath +
                          $" && midl.exe /win32 /tlb \"{parameters.Tlb32File}\" \"{parameters.IdlFile}\" /out \"{targetPath}\"" +
                          $" && midl.exe /amd64 /tlb \"{parameters.Tlb32File}\" \"{parameters.IdlFile}\" /out \"{targetPath}\"";
            ExecuteTask(command);
        }

        private void CompileWithTlbExp(DllFileParameters parameters)
        {
            var command = $"{NetToolsDir}tlbexp.exe \"{parameters.SourceDll}\" /win32 /out:\"{parameters.SourceTlb32}\"";
            ExecuteTask(command);

            command = $"{NetToolsDir}tlbexp.exe \"{parameters.SourceDll}\" /win64 /out:\"{parameters.SourceTlb64}\"";
            ExecuteTask(command);
        }

        private void HarvestMetadataWithWix(DllFileParameters parameters)
        {
            var command = $"{WixToolsDir}heat.exe file \"{parameters.SourceDll}\" -out \"{parameters.DllXml}\"";
            ExecuteTask(command);

            command = $"{WixToolsDir}heat.exe file \"{parameters.SourceTlb32}\" -out \"{parameters.TlbXml}\"";
            ExecuteTask(command);
        }

        private IOrderedEnumerable<RegistryEntry> BuildRegistryEntriesFromMetadata(DllFileParameters parameters)
        {
            var builder = new RegistryEntryBuilder();
            return builder.Parse(parameters.TlbXml, parameters.DllXml);
        }

        private void CreateInnoSetupRegistryEntries(IOrderedEnumerable<RegistryEntry> entries, DllFileParameters parameters)
        {
            var writer = new InnoSetupRegistryWriter();
            var content = writer.Write(entries, parameters.DllFile, parameters.Tlb32File, parameters.Tlb64File);
            var regFile = Path.Combine(IncludeDir, parameters.DllFile.Replace(".dll", ".reg.iss"));
            
            // Encoding must be UTF8 BOM
            File.WriteAllText(regFile, content, new UTF8Encoding(true));
        }

        private bool _previousRegistrationRemoved;
        private void RemovePreviousDebugRegistration()
        {
            if (_previousRegistrationRemoved)
            {
                return;
            }

            this.LogCustom("Removing previous debug build's registration...");

            // First see if there are registry script from the previous build
            // If so, execute them to delete previous build's keys (which may
            // no longer exist for the current build and thus won't be overwritten)
            var dir = Path.GetDirectoryName(RegFilePath);
            if (Directory.Exists(dir))
            {
                // The current reg file is now the previous build's reg file
                var lastRegFile = RegFilePath;
                if (File.Exists(lastRegFile))
                {
                    var now = DateTime.UtcNow;
                    if (Environment.Is64BitOperatingSystem)
                    {
                        var command = $"reg.exe import {lastRegFile} /reg:32";
                        ExecuteTask(command);
                        command = $"reg.exe import {lastRegFile} /reg:64";
                        ExecuteTask(command);
                    }
                    else
                    {
                        var command = $"reg.exe import {lastRegFile}";
                        ExecuteTask(command);
                    }
                    File.Move(lastRegFile, lastRegFile + ".imported_" + now.ToUniversalTime().ToString("yyyyMMddHHmmss") + ".txt");
                }
            }
            else
            {
                Directory.CreateDirectory(dir);
            }

            _previousRegistrationRemoved = true;
        }

        private void UpdateDebugRegistration(IOrderedEnumerable<RegistryEntry> entries, DllFileParameters parameters)
        {
            this.LogCustom($"Updating debug build reigstration for {parameters.DllFile}");

            // NOTE: The local writer will perform the actual registry changes; the return
            // is a registry script with deletion instructions for the keys to be deleted
            // in the next build.
            var writer = new LocalDebugRegistryWriter();
            var content = writer.Write(entries, parameters.DllFile, parameters.Tlb32File, parameters.Tlb64File);
            File.AppendAllText(RegFilePath, content, Encoding.ASCII);
        }

        private void ExecuteTask(string command, string workingDirectory = null)
        {
            var exec = new Exec
            {
                BuildEngine = BuildEngine,
                HostObject = HostObject,
                Command = command,
                ConsoleToMSBuild = true,
                EchoOff = false,
                IgnoreExitCode = false,
                WorkingDirectory = workingDirectory
            };

            if (!exec.Execute())
            {
                throw new InvalidOperationException("Execution of task failed");
            }
        }
    }
}

