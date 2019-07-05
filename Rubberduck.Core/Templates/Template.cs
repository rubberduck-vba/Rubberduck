namespace Rubberduck.Templates
{
    /// <remarks>
    /// Template can be either built-in or user-defined. For a built-in template, the
    /// metadata should be stored in the <see cref="Rubberduck.Resources.Templates"/>
    /// resource, with specific entries, currently Name, Caption, Description and Code.
    /// Due to the fact that we cannot strong-type the reference to the resource entries
    /// the class has unit tests to validate that the crucial elements are present in the
    /// resource to guard against runtime errors/unexpected behavior due to missing/malformed
    /// entries in the resources. 
    /// </remarks>
    public class Template : ITemplate
    {
        private readonly ITemplateFileHandler _handler;
        public Template(string name, ITemplateFileHandler handler)
        {
            _handler = handler;

            Name = name;
            IsUserDefined = VerifyIfUserDefined(name);

            if (IsUserDefined)
            {
                //TODO: Devise a way for users to define their captions/descriptions simply
                Caption = Name;
                Description = Name;
            }
            else
            {
                VerifyFile(name, handler);
                (Caption, Description) = GetBuiltInMetaData(name);
            }
        }

        public string Name { get; }
        public bool IsUserDefined { get; }
        public string Caption { get; }
        public string Description { get; }

        public string Read() => _handler.Read();
        
        public void Write(string content) => _handler.Write(content);

        private static bool VerifyIfUserDefined(string name)
        {
            var builtInCode = Resources.Templates.ResourceManager.GetString(name + "_Code");
            return builtInCode == null;
        }

        private static void VerifyFile(string name, ITemplateFileHandler handler)
        {
            if (handler.Exists)
            {
                return;
            }

            var content = Resources.Templates.ResourceManager.GetString(name + "_Code");
            handler.Write(content);
        }

        private static (string caption, string description) GetBuiltInMetaData(string name)
        {
            var caption = Resources.Templates.ResourceManager.GetString(name + "_Caption");
            var description = Resources.Templates.ResourceManager.GetString(name + "_Description");

            return (caption, description);
        }
    }
}