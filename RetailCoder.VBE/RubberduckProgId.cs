namespace Rubberduck
{
    public static class RubberduckProgId
    {
        private const string BaseNamespace = "Rubberduck.";

        public const string ExtensionProgId = BaseNamespace + "Extension";
        public const string DockableWindowHostProgId = BaseNamespace + "UI.DockableWindowHost";

        public const string DeclarationProgId = BaseNamespace + "Declaration";
        public const string IdentifierReferenceProgId = BaseNamespace + "IdentifierReference";
        public const string ParserStateProgId = BaseNamespace + "ParserState";

        public const string AssertClassProgId = BaseNamespace + "AssertClass";
        public const string PermissiveAssertClassProgId = BaseNamespace + "PermissiveAssertClass";
        public const string FakesProviderProgId = BaseNamespace + "FakesProvider";
    }
}
