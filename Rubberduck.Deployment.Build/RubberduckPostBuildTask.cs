using System;
using System.IO;
using System.Linq;
using System.Text;
using Microsoft.Build.Framework;
using Microsoft.Build.Tasks;
using Microsoft.Build.Utilities;
using Microsoft.VisualStudio.Setup.Configuration;
using Rubberduck.Deployment.Build.Builders;
using Rubberduck.Deployment.Build.IdlGeneration;
using Rubberduck.Deployment.Build.Structs;
using Rubberduck.Deployment.Build.Writers;

namespace Rubberduck.Deployment.Build
{
    internal readonly struct DllFileParameters
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

    public class RubberduckPostBuildTask : AppDomainIsolatedTask
    {
        /// <summary>
        /// Visual Studio Build Configuration (e.g. Debug or Release)
        /// </summary>
        [Required]
        public string Config { get; set; }

        /// <summary>
        /// Full path to NetFX SDK directory
        /// </summary>
        [Required]
        public string NetToolsDir { get; set; }

        /// <summary>
        /// Full path to WiX SDK directory
        /// </summary>
        [Required]
        public string WixToolsDir { get; set; }

        /// <summary>
        /// Full path to the directory containing all the source files we want to deploy.
        /// </summary>
        [Required]
        public string SourceDir { get; set; }

        /// <summary>
        /// Full path to the directory we want to write our modified files into.
        /// </summary>
        [Required]
        public string TargetDir { get; set; }

        /// <summary>
        /// Root path of the project executing the build task.
        /// </summary>
        [Required]
        public string ProjectDir { get; set; }

        /// <summary>
        /// Full path to the Inno Setup's include files.
        /// </summary>
        [Required]
        public string IncludeDir { get; set; }

        /// <summary>
        /// Pipe-delimited list of DLL to generate TLBs from. Should contain only name &amp; extension and exist in <see cref="SourceDir"/>.
        /// </summary> 
        [Required]
        public string FilesToExtract { get; set; }

        private const int MaxNumberOfPreviousRegistrationFiles = 10;

        private string RegFilePath =>
            Path.Combine(Path.Combine(ProjectDir, "LocalRegistryEntries"), "DebugRegistryEntries.reg");

        /// <remarks>
        /// Entry point for the build task. To use the build task in a csproj, the <c>UsingTask</c> element must be specified before defining the task. The task will
        /// have the same name as the class, followed by the parameters. In this case, it would be <see cref="RubberduckPreBuildTask"/> element. See <c>Rubberduck.Deployment.csproj</c>
        /// for usage example. The public properties are used as a parameter in the MSBuild task and are both settable and gettable. Thus, we must read from the properties when 
        /// we run the <c>Execute</c> which influences the behvaior of the task.
        /// </remarks>
        public override bool Execute()
        {
            var result = true;
            try
            {
                this.LogMessage(FormatParameterList());

                CleanOldImports(Path.Combine(ProjectDir, "LocalRegistryEntries"));

                var dllFiles = FilesToExtract.Split(new[] {"|"}, StringSplitOptions.RemoveEmptyEntries);
                var hasVCTools = TryGetVCToolsPath(out var batchPath);
                var message = hasVCTools
                    ? "C++ build tools found; using midl.exe to generate TLBs."
                    : "No C++ build tools found; using tlbexp.exe to generate TLBs.";
                this.LogMessage(message);
                
                foreach (var dllFile in dllFiles)
                {
                    ProcessDll(dllFile, batchPath);
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
            sb.AppendLine("Parameters provided:")
              .AppendLine($"          Config: {Config}")
              .AppendLine($"     NetToolsDir: {NetToolsDir}")
              .AppendLine($"     WixToolsDir: {WixToolsDir}")
              .AppendLine($"       SourceDir: {SourceDir}")
              .AppendLine($"       TargetDir: {TargetDir}")
              .AppendLine($"      ProjectDir: {ProjectDir}")
              .AppendLine($"     IncludesDir: {IncludeDir}")
              .AppendLine($"  FilesToExtract: {FilesToExtract}")
              .AppendLine(string.Empty);

            return sb.ToString();
        }

        private bool TryGetVCToolsPath(out string batchPath)
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

                if (!(instances[0] is ISetupInstance2 instance))
                {
                    continue;
                }

                var packages = instance.GetPackages();
                foreach (var package in packages)
                {
                    if (package.GetId() != "Microsoft.VisualStudio.Component.VC.Tools.x86.x64")
                    {
                        continue;
                    }

                    var rootPath = instance.ResolvePath();
                    batchPath = Directory.GetFiles(rootPath, "VsDevCmd.bat", SearchOption.AllDirectories)
                        .FirstOrDefault();

                    return true;
                }
            } while (fetched > 0);

            batchPath = null;
            return false;
        }

        private void CleanOldImports(string dir)
        {
            this.LogMessage("Cleaing out old imports");

            var files = Directory.GetFiles(dir, "DebugRegistryEntries.reg.imported_*.txt").OrderByDescending(f => f);
            var i = 0;
            foreach (var file in files)
            {
                if (i > MaxNumberOfPreviousRegistrationFiles)
                {
                    this.LogMessage($"Deleting {file}");
                    File.Delete(file);
                }

                i++;
            }
        }

        private void UpdateAddInRegistration()
        {
            this.LogMessage("Updating addin registration...");
            var addInRegFile = Path.Combine(Path.GetDirectoryName(RegFilePath), "RubberduckAddinRegistry.reg");
            var command = $"reg.exe import \"{addInRegFile}\"";
            ExecuteTask(command);
        }

        private void ProcessDll(string file, string batchPath)
        {
            this.LogMessage($"Processing {file}...");

            var parameters = new DllFileParameters(file, SourceDir, TargetDir);
            
            if (string.IsNullOrWhiteSpace(batchPath))
            {
                this.LogMessage("Compiling with tlbexp...");
                CompileWithTlbExp(parameters);
            }
            else
            {
                this.LogMessage("Compiling with midl...");
                CreateIdlFile(parameters);
                CompileWithMidl(parameters, batchPath);
            }

            this.LogMessage("Extracting metadata using WiX...");
            HarvestMetadataWithWixToFile(parameters);

            this.LogMessage("Building registry entries...");
            var entries = BuildRegistryEntriesFromMetadata(parameters);

            this.LogMessage("Creating InnoSetup registry entries...");
            CreateInnoSetupRegistryFile(entries, parameters);

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
            File.WriteAllText(Path.Combine(TargetDir, parameters.IdlFile), idl, new UTF8Encoding(true));
        }

        private void CompileWithMidl(DllFileParameters parameters, string batchPath)
        {
            var targetPath = Path.Combine(Path.GetTempPath(), "RubberduckMidl");
            if (Directory.Exists(targetPath))
            {
                Directory.Delete(targetPath, true);
            }
            Directory.CreateDirectory(targetPath);

            targetPath = targetPath.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);

            var command = $"call \"{batchPath}\"{Environment.NewLine}" +
                          $"midl.exe /win32 /tlb \"{parameters.Tlb32File}\" \"{parameters.IdlFile}\" /out \"{targetPath}\"{Environment.NewLine}" +
                          $"midl.exe /amd64 /tlb \"{parameters.Tlb64File}\" \"{parameters.IdlFile}\" /out \"{targetPath}\"";
            ExecuteTask(command, SourceDir);

            MoveFileWithOverwrite(Path.Combine(targetPath, parameters.Tlb32File), Path.Combine(TargetDir, parameters.Tlb32File));
            MoveFileWithOverwrite(Path.Combine(targetPath, parameters.Tlb64File), Path.Combine(TargetDir, parameters.Tlb64File));

            try
            {
                Directory.Delete(targetPath, true);
            }
            catch (Exception ex)
            {
                this.LogMessage($"Unable to delete temporary working directory: {targetPath}. Exception: {ex}");
            }
        }

        private static void MoveFileWithOverwrite(string sourceFile, string destFile)
        {
            if (File.Exists(destFile))
            {
                File.Delete(destFile);
            }
            File.Move(sourceFile, destFile);
        }

        private void CompileWithTlbExp(DllFileParameters parameters)
        {
            var command = $"\"{NetToolsDir}tlbexp.exe\" \"{parameters.SourceDll}\" /win32 /out:\"{parameters.SourceTlb32}\"";
            ExecuteTask(command);

            command = $"\"{NetToolsDir}tlbexp.exe\" \"{parameters.SourceDll}\" /win64 /out:\"{parameters.SourceTlb64}\"";
            ExecuteTask(command);
        }

        private void HarvestMetadataWithWixToFile(DllFileParameters parameters)
        {
            var command = $"\"{WixToolsDir}heat.exe\" file \"{parameters.SourceDll}\" -out \"{parameters.DllXml}\"";
            ExecuteTask(command);

            command = $"\"{WixToolsDir}heat.exe\" file \"{parameters.SourceTlb32}\" -out \"{parameters.TlbXml}\"";
            ExecuteTask(command);
        }

        private IOrderedEnumerable<RegistryEntry> BuildRegistryEntriesFromMetadata(DllFileParameters parameters)
        {
            var builder = new RegistryEntryBuilder();
            return builder.Parse(parameters.TlbXml, parameters.DllXml);
        }

        private void CreateInnoSetupRegistryFile(IOrderedEnumerable<RegistryEntry> entries, DllFileParameters parameters)
        {
            var writer = new InnoSetupRegistryWriter();
            var content = writer.Write(entries, parameters.DllFile, parameters.Tlb32File, parameters.Tlb64File);
            var regFile = Path.Combine(IncludeDir, parameters.DllFile.Replace(".dll", ".reg.iss"));
            
            // To use unicode with InnoSetup, encoding must be UTF8 BOM
            File.WriteAllText(regFile, content, new UTF8Encoding(true));
        }

        private bool _previousRegistrationRemoved;
        private void RemovePreviousDebugRegistration()
        {
            if (_previousRegistrationRemoved)
            {
                return;
            }

            this.LogMessage("Removing previous debug build's registration...");

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
                    //reg.exe should be present on all Windows since NT, so it's safe to assume
                    //that it's always there

                    var now = DateTime.UtcNow;
                    if (Environment.Is64BitOperatingSystem)
                    {
                        var command = $"reg.exe import \"{lastRegFile}\" /reg:32";
                        ExecuteTask(command);
                        command = $"reg.exe import \"{lastRegFile}\" /reg:64";
                        ExecuteTask(command);
                    }
                    else
                    {
                        var command = $"reg.exe import \"{lastRegFile}\"";
                        ExecuteTask(command);
                    }

                    var archivedRegFile = lastRegFile + ".imported_" +
                                          now.ToUniversalTime().ToString("yyyyMMddHHmmss") + ".txt";
                    File.Move(lastRegFile, archivedRegFile);
                    this.LogMessage($"Renamed the previous build reg file to '{archivedRegFile}'");
                }
            }
            else
            {
                Directory.CreateDirectory(dir);
                this.LogMessage($"Created the directory '{dir}'");
            }

            _previousRegistrationRemoved = true;
        }

        private void UpdateDebugRegistration(IOrderedEnumerable<RegistryEntry> entries, DllFileParameters parameters)
        {
            this.LogMessage($"Updating debug build reigstration for {parameters.DllFile}");

            // NOTE: The local writer will perform the actual registry changes; the return
            // is a registry script with deletion instructions for the keys to be deleted
            // in the next build.
            var writer = new LocalDebugRegistryWriter();
            writer.CurrentPath = TargetDir;
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
                LogStandardErrorAsError = false,
                IgnoreExitCode = false,
                UseCommandProcessor = false,
                WorkingDirectory = workingDirectory ?? ProjectDir,
                YieldDuringToolExecution = true
            };

            var log = new TaskLoggingHelper(exec);

            this.LogMessage(command);
            var result = exec.Execute();

            if (!result && !log.HasLoggedErrors)
            {
                throw new InvalidOperationException("Execution of task failed");
            }
        }
    }
}

