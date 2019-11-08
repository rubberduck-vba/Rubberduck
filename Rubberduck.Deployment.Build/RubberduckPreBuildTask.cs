using System;
using System.IO;
using Microsoft.Build.Framework;
using Microsoft.Build.Utilities;

namespace Rubberduck.Deployment.Build
{
    public class RubberduckPreBuildTask : AppDomainIsolatedTask
    {
        /// <summary>
        /// Full path to the directory containing the templates we need to modify.
        /// </summary>
        [Required]
        public string WorkingDir { get; set; }

        /// <summary>
        /// Full path to the directory where we want to place our modified files.
        /// </summary>
        [Required]
        public string OutputDir { get; set; }

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
                VerifyFileLocks();
                UpdateLicenseFile();
            }
            catch (Exception ex)
            {
                this.LogError(ex);
                result = false;
            }

            return result;
        }

        private void VerifyFileLocks()
        {
            try
            {
                var files = Directory.GetFiles(OutputDir);
                foreach (var file in files)
                {
                    var fileName = Path.GetFileName(file);
                    if (fileName.StartsWith("Rubberduck") && fileName.EndsWith(".dll"))
                    {
                        //If we can't delete it, then file is in use.
                        File.Delete(file);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException(
                    "Cannot build; files are in use. Please check for any processes that may be using the project.",
                    ex);
            }
        }

        private void UpdateLicenseFile()
        {
            var sourceFile = WorkingDir + @"\Licenses\License.rtf";
            var license = WorkingDir + @"\InnoSetup\Includes\license.rtf";
            var endYear = new DateTime().Year;

            var content = File.ReadAllText(sourceFile).Replace("$(YEAR$)", endYear.ToString());
            File.WriteAllText(license, content);
        }
    }
}
