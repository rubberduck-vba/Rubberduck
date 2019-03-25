using System;
using System.IO;
using Microsoft.Build.Framework;
using Microsoft.Build.Utilities;

namespace Rubberduck.Deployment.Build
{
    public class RubberduckPreBuildTask : AppDomainIsolatedTask
    {
        [Required]
        public string WorkingDir { get; set; }

        [Required]
        public string OutputDir { get; set; }

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
