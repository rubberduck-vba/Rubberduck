namespace Rubberduck.Resources.Registration
{
    public static class RubberduckGuid
    {
        public const string IID_IUnknown = "00000000-0000-0000-C000-000000000046";
        public const string IID_IDispatch = "00020400-0000-0000-C000-000000000046";

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
        public const string IMockProviderGuid = UnitTestingGuidspace + "E2" + GuidSuffix;
        public const string MockProviderGuid = UnitTestingGuidspace + "E3" + GuidSuffix;
        public const string IComMockGuid = UnitTestingGuidspace + "E4" + GuidSuffix;
        public const string ComMockGuid = UnitTestingGuidspace + "E5" + GuidSuffix;
        public const string ISetupArgumentDefinitionGuid = UnitTestingGuidspace + "E6" + GuidSuffix;
        public const string SetupArgumentDefinitionGuid = UnitTestingGuidspace + "E7" + GuidSuffix;
        public const string ISetupArgumentDefinitionsGuid = UnitTestingGuidspace + "E8" + GuidSuffix;
        public const string SetupArgumentDefinitionsGuid = UnitTestingGuidspace + "E9" + GuidSuffix;
        public const string ISetupArgumentCreatorGuid = UnitTestingGuidspace + "EA" + GuidSuffix;
        public const string SetupArgumentCreatorGuid = UnitTestingGuidspace + "EB" + GuidSuffix;
        public const string IComMockedGuid = UnitTestingGuidspace + "EC" + GuidSuffix;
        public const string ComMockedGuid = UnitTestingGuidspace + "ED" + GuidSuffix;

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
        public const string SetupArgumentRangeGuid = RecordGuidspace + "26" + GuidSuffix;
        public const string SetupArgumentTypeGuid = RecordGuidspace + "27" + GuidSuffix;

        // Debug Guids:
        private const string DebugGuidspace = "69E101";
        public const string DebugAddinObjectInterfaceGuid = DebugGuidspace + "23" + GuidSuffix;
        public const string DebugAddinObjectClassGuid = DebugGuidspace + "24" + GuidSuffix;
    }
}