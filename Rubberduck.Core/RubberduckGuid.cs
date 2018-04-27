// ReSharper disable InconsistentNaming

namespace Rubberduck
{
    public static class RubberduckGuid
    {
        // TypeLib Guid:
        private const string TypeLibGuidspace = "e07c841c-14b4-4890-83e9-";
        public const string RubberduckTypeLibGuid = TypeLibGuidspace + "8c80b06dd59d";
        public const string RubberduckApiTypeLibGuid = TypeLibGuidspace + "8c80b06dd59e";

        // Addin Guids:
        private const string AddinGuidspace = "69E194DA-43F0-3B33-A105-";
        public const string ExtensionGuid = AddinGuidspace + "F697CA1F8A66"; 
        public const string DockableWindowHostGuid = AddinGuidspace + "F697CA1F8A67";

        // Unit testing Guids:
        private const string UnitTestingGuidspace = "69E194DA-43F0-3B33-B105-";
        public const string AssertClassGuid = UnitTestingGuidspace + "9B8188A6F040";
        public const string IAssertGuid = UnitTestingGuidspace + "9B8188A6F041";
        public const string PermissiveAssertClassGuid = UnitTestingGuidspace + "9B8188A6F042";
        public const string FakesProviderClassGuid = UnitTestingGuidspace + "9B8188A6F043";
        public const string IFakesProviderGuid = UnitTestingGuidspace + "9B8188A6F044";
        public const string IFakeGuid = UnitTestingGuidspace + "9B8188A6F045";
        public const string IVerifyGuid = UnitTestingGuidspace + "9B8188A6F046";
        public const string IStubGuid = UnitTestingGuidspace + "9B8188A6F047";

        // Rubberduck API Guids:
        private const string ApiGuidspace = "69E194DA-43F0-3B33-B106-";
        public const string IDeclarationGuid = ApiGuidspace + "9B8188A6F040";
        public const string DeclarationClassGuid = ApiGuidspace + "9B8188A6F041";
        public const string IIdentifierReferenceGuid = ApiGuidspace + "9B8188A6F042";
        public const string IdentifierReferenceClassGuid = ApiGuidspace + "9B8188A6F043";
        public const string IParserStateGuid = ApiGuidspace + "9B8188A6F044";
        public const string ParserStateClassGuid = ApiGuidspace + "9B8188A6F045";
        public const string IParserStateEventsGuid = ApiGuidspace + "9B8188A6F046";
        
        // Enum Guids:
        private const string RecordGuidspace = "69E194DA-43F0-3B33-C105-";
        public const string DeclarationTypeGuid = RecordGuidspace + "FEABE42C9725";
        public const string AccessibilityGuid = RecordGuidspace + "FEABE42C9726";
    }
}