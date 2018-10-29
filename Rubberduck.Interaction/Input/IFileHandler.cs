using System.IO;

namespace Rubberduck.Interaction.Input
{
    public interface IFileHandler
    {
        string[] ReadAllLines(string path);
        void WriteToFile(string path, string contents);
        string RewrittenFilePath();
        bool Exists(string path);
        void Delete(string path);
    }
    public class FileHandler : IFileHandler
    {
        public string[] ReadAllLines(string path)
        {
            return File.ReadAllLines(path);
        }

        public void WriteToFile(string path, string contents)
        {
            var sw = File.CreateText(path);
            _updatedFilePath = path;
            sw.Write(contents);
            sw.Close();
        }

        private string _updatedFilePath;
        public string RewrittenFilePath() => _updatedFilePath;

        public bool Exists(string path)
        {
            return File.Exists(path);
        }

        public void Delete(string path)
        {
            if (Exists(path))
            {
                File.Delete(path);
            }
        }
    }
}
