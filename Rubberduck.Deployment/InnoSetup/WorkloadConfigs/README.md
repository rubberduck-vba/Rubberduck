## Workload Configurations

Workload configurations copy a `rubberduck.config` file under `%appdata%\Rubberduck` (under the current user's profile) where Rubberduck reads and writes its configuration settings.

If no workload option is specified on install, Rubberduck uses its default configuration and only writes a `rubberduck.config` file if the user changes the settings.

---

### Legacy

This workload aims to improve overall user experience by disabling features that can deteriorate performance in large "legacy" projects that generate lots of inspection results.

**Affected Settings**

- `AutoCompleteSettings` has its `IsEnabled` attribute set to `false`, in order to remove any in-editor interferences.
- The following inspections have their `Severity` set to `DoNotShow`:
  - EmptyStringLiteral
  - ObsoleteCallStatement
  - MultipleDeclarations
  - ParameterCanBeByVal
  - UseMeaningfulName
  - HungarianNotation
  - UnreachableCase
  - ImplicitPublicMember
  - ImplicitByRefModifier
  - ImplicitDefaultMemberAccess
  - ImplicitRecursiveDefaultMemberAccess
  - IndexedDefaultMemberAccess
  - RedundantByRefModifier
  - ShadowedDeclaration
  - UseOfBangNotation
  - UseOfRecursiveBangNotation

Disabled inspections can be re-enabled individually from the *Settings* dialog.