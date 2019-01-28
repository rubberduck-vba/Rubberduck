namespace Rubberduck.Deployment.Build.Structs
{
    public struct FileMap
    {
        public string FileId { get; }
        public string FilePath { get; }
        
        public FileMap(string id, string filePath)
        {
            FileId = id;
            FilePath = filePath;
        }

        public string Replace(string source)
        {
            if (string.IsNullOrWhiteSpace(source) || !source.Contains(FileId))
            {
                return source;
            }
            
            return source.Replace(@"file:///[#" + FileId + @"]", FilePath)
                 .Replace(FileId, FilePath);
        }
    }
}