namespace Rubberduck.Resources.Registration
{
    public static class RubberduckGuid
    {
        // Guid Suffix
        private const string GuidSuffix = "-43F0-3B33-B105-9B8188A6F040";

        // TypeLib Guid:
        private const string TypeLibGuidspace = "E07C84";
        public const string RubberduckTypeLibGuid = TypeLibGuidspace + "1C" + GuidSuffix;
        public const string RubberduckApiTypeLibGuid = TypeLibGuidspace + "1D" + GuidSuffix;
        
        // Addin Guids:
        private const string AddinGuidspace = "69E0F6";
        public const string ExtensionGuid = AddinGuidspace + "97" + GuidSuffix;
        public const string IDockableWindowHostGuid = AddinGuidspace + "98" + GuidSuffix;
        public const string DockableWindowHostGuid = AddinGuidspace + "99" + GuidSuffix;

        // Unit testing Guids:
        private const string UnitTestingGuidspace = "69E0F7";
        public const string AssertClassGuid = UnitTestingGuidspace + "DA" + GuidSuffix;
        public const string IAssertGuid = UnitTestingGuidspace + "DB" + GuidSuffix;
        public const string PermissiveAssertClassGuid = UnitTestingGuidspace + "DC" + GuidSuffix;
        public const string FakesProviderClassGuid = UnitTestingGuidspace + "DD" + GuidSuffix;
        public const string IFakesProviderGuid = UnitTestingGuidspace + "DE" + GuidSuffix;
        public const string IFakeGuid = UnitTestingGuidspace + "DF" + GuidSuffix;
        public const string IVerifyGuid = UnitTestingGuidspace + "E0" + GuidSuffix;
        public const string IStubGuid = UnitTestingGuidspace + "E1" + GuidSuffix;

        // Rubberduck API Guids:
        private const string ApiGuidspace = "69E0F7";
        public const string IDeclarationGuid = ApiGuidspace + "81" + GuidSuffix;
        public const string DeclarationClassGuid = ApiGuidspace + "82" + GuidSuffix;
        public const string IIdentifierReferenceGuid = ApiGuidspace + "83" + GuidSuffix;
        public const string IdentifierReferenceClassGuid = ApiGuidspace + "84" + GuidSuffix;
        public const string IParserGuid = ApiGuidspace + "85" + GuidSuffix;
        public const string ParserClassGuid = ApiGuidspace + "86" + GuidSuffix;
        public const string IParserEventsGuid = ApiGuidspace + "87" + GuidSuffix;
        public const string IDeclarationsGuid = ApiGuidspace + "88" + GuidSuffix;
        public const string DeclarationsClassGuid = ApiGuidspace + "89" + GuidSuffix;
        public const string IApiProviderGuid = ApiGuidspace + "8A" + GuidSuffix;
        public const string ApiProviderClassGuid = ApiGuidspace + "8B" + GuidSuffix;
        public const string IIdentifierReferencesGuid = ApiGuidspace + "8C" + GuidSuffix;
        public const string IdentifierReferencesClassGuid = ApiGuidspace + "8D" + GuidSuffix;

        // Enum Guids:
        private const string RecordGuidspace = "69E100";
        public const string DeclarationTypeGuid = RecordGuidspace + "23" + GuidSuffix;
        public const string AccessibilityGuid = RecordGuidspace + "24" + GuidSuffix;
        public const string ParserStateGuid = RecordGuidspace + "25" + GuidSuffix;

        // Debug Guids:
        private const string DebugGuidspace = "69E101";
        public const string DebugAddinObjectInterfaceGuid = DebugGuidspace + "23" + GuidSuffix;
        public const string DebugAddinObjectClassGuid = DebugGuidspace + "24" + GuidSuffix;
    }
}