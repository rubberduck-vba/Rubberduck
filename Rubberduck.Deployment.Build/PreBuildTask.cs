using System;
using System.IO;
using Microsoft.Build.Framework;

namespace Rubberduck.Deployment.Build
{
    public class PreBuildTask : ITask
    {
        public IBuildEngine BuildEngine { get; set; }

        public ITaskHost HostObject { get; set; }

        [Required]
        public string WorkingDir { get; set; }

        [Required]
        public string OutputDir { get; set; }

        public bool Execute()
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
                    if (file.StartsWith("Rubberduck") && file.EndsWith(".dll"))
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
            var sourceFile = WorkingDir + @"\Licenses\license.rtf";
            var license = WorkingDir + @"\InnoSetup\Includes\license.rtf";
            var endYear = new DateTime().Year;

            var content = File.ReadAllText(sourceFile).Replace("$(YEAR$)", endYear.ToString());
            File.WriteAllText(license, content);
        }
    }
}
