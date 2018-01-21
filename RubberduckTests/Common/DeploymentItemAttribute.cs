using System;
using System.IO;
using System.Reflection;
using NUnit.Framework;
using NUnit.Framework.Interfaces;

namespace RubberduckTests.Common
{
    // "borrowed" from https://stackoverflow.com/a/11569203
    [AttributeUsage(AttributeTargets.Method | AttributeTargets.Class | AttributeTargets.Struct,
        AllowMultiple = false,
        Inherited = false)]
    public class DeploymentItem : TestActionAttribute
    {
        private readonly string _fileProjectRelativePath;

        public override void BeforeTest(ITest test)
        {
            var filePath = _fileProjectRelativePath.Replace("/", @"\");

            var environmentDir = new DirectoryInfo(Environment.CurrentDirectory);
            var itemPathUri = new Uri(Path.Combine(environmentDir.FullName
                , filePath));

            var itemPath = itemPathUri.LocalPath;
            var binFolderPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

            var itemPathInBinUri = new Uri(Path.Combine(binFolderPath, filePath));
            var itemPathInBin = itemPathInBinUri.LocalPath;

            if (Directory.Exists(itemPath))
            {
                Directory.Delete(itemPath, true);
            }

            if (Directory.Exists(itemPathInBin))
            {
                foreach (var dirPath in Directory.GetDirectories(itemPathInBin, "*",
                    SearchOption.AllDirectories))
                {
                    Directory.CreateDirectory(dirPath.Replace(itemPathInBin, itemPath));
                }

                foreach (var newPath in Directory.GetFiles(itemPathInBin, "*.*",
                    SearchOption.AllDirectories))
                {
                    File.Copy(newPath, newPath.Replace(itemPathInBin, itemPath), true);
                }
            }
        }

        public DeploymentItem(string fileProjectRelativePath)
        {
            _fileProjectRelativePath = fileProjectRelativePath;
        }
    }
}