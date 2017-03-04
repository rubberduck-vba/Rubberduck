Attribute VB_Name = "modRegistry"
'
' Standard Registry handling module, written by Rob Bovey
'

Option Base 1
Option Explicit
Option Compare Text

Const MAX_STRING_LEN As Long = 128

'Register Value data types
Const REG_NONE = 0&
Const REG_SZ = 1&
Const REG_EXPAND_SZ = 2&
Const REG_BINARY = 3&
Const REG_DWORD = 4&
Const REG_DWORD_LITTLE_ENDIAN = 4&
Const REG_DWORD_BIG_ENDIAN = 5&
Const REG_LINK = 6&
Const REG_MULTI_SZ = 7&
Const REG_RESOURCE_LIST = 8&
Const REG_FULL_RESOURCE_DESCRIPTOR = 9&
Const REG_RESOURCE_REQUIREMENTS_LIST = 10&

'Registry error values
Const ERROR_SUCCESS = 0&
Const ERROR_BADDB = 1009&
Const ERROR_BADKEY = 2&
Const ERROR_CANTOPEN = 1011&
Const ERROR_CANTREAD = 1012&
Const ERROR_CANTWRITE = 1013&
Const ERROR_REGISTRY_RECOVERED = 1014&
Const ERROR_REGISTRY_CORRUPT = 1015&
Const ERROR_REGISTRY_IO_FAILED = 1016&
Const ERROR_NOT_REGISTRY_FILE = 1017&
Const ERROR_KEY_DELETED = 1018&
Const ERROR_NO_LOG_SPACE = 1019&
Const ERROR_KEY_HAS_CHILDREN = 1020&
Const ERROR_CHILD_MUST_BE_VOLATILE = 1021&

'Declare constants for standard registry keys
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_DYN_DATA = &H80000006

'Create a key in the registry
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal sKey As String, ByRef plKeyReturn As Long) As Long
Attribute RegCreateKey.VB_UserMemId = 1879048192

'Open a registry key
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal sKey As String, ByRef plKeyReturn As Long) As Long
Attribute RegOpenKey.VB_UserMemId = 1879048228

'Set a value in a registry key
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal sValueName As String, ByVal dwReserved As Long, ByVal dwType As Long, ByVal sBuffer As String, ByVal dwLen As Long) As Long
Attribute RegSetValueEx.VB_UserMemId = 1879048260

'Get the value from a registry
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal sValueName As String, ByVal dwReserved As Long, ByRef lValueType As Long, ByVal sValue As String, ByRef lResultLen As Long) As Long
Attribute RegQueryValueEx.VB_UserMemId = 1879048296

'Delete a value from a registry key
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal sValueName As String) As Long
Attribute RegDeleteValue.VB_UserMemId = 1879048336

'Close a registry key
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Attribute RegCloseKey.VB_UserMemId = 1879048372

'Delete a registry key
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Attribute RegDeleteKey.VB_UserMemId = 1879048404


'***************************************************************************
'*
'* PROCEDURE NAME:  SET REGISTRY VALUE
'*
'* DESCRIPTION:     Sets a value in the registry.
'*
'* PASS:
'*  (1) the key (e.g., HKEY_CURRENT_USER)
'*  (2) the subkey (e.g., "Software\Company\Product"
'*  (3) the value's name (e.g., "" [for default] or "whatever")
'*  (4) the value to be assigned to (3) (Pass binary values as strings)
'*  (5) the type of data (string, binary, Dword) (optional)
'*
'***************************************************************************

Sub procSetRegValue(ByVal iKey As Long, ByVal sSubKey As String, ByVal sValueName As Variant, _
                    ByVal vNameValue As Variant, Optional ByVal vValueType As Variant)
Attribute procSetRegValue.VB_UserMemId = 1610612736


    Dim iHkey As Long

    If Left(sSubKey, 1) = "\" Then sSubKey = Mid(sSubKey, 2)

    'Assume type of text if missing
    If IsMissing(vValueType) Then vValueType = REG_SZ

    'Convert binary or word values from strings to values
    Select Case vValueType
    Case REG_DWORD
        vNameValue = funFlipAndHexConvert(vNameValue)
    Case REG_BINARY
        vNameValue = funHexConvert(vNameValue)
    End Select

    'Create the correct iKey and return a handle to the iKey
    If RegCreateKey(iKey, sSubKey, iHkey) = ERROR_SUCCESS Then

        'Set the value in the iKey
        RegSetValueEx iHkey, sValueName, 0&, vValueType, vNameValue, Len(vNameValue)

        'Close the registry iKey
        RegCloseKey iHkey
    End If

End Sub


'***************************************************************************
'*
'* FUNCTION NAME:   GET REGISTRY VALUE
'*
'* DESCRIPTION:     Retrieves a value from the registry.
'*
'* PASS:
'*  (1) the key (e.g., HKEY_CURRENT_USER)
'*  (2) the subkey (e.g., "Software\Company\Product")
'*  (3) the value's name (e.g., "" [for default] or "whatever")
'*
'***************************************************************************

Function funGetRegValue(ByVal iKey As Long, ByVal sSubKey As String, ByVal sValueName As Variant, ByVal vDefault As Variant) As String
Attribute funGetRegValue.VB_UserMemId = 1610612737


    Dim sBuffer As String * MAX_STRING_LEN, sTempBuffer As String
    Dim iErrCode As Long, iHkey As Long, ValueType As Long, ValueLen As Long, i As Integer

    ValueLen = MAX_STRING_LEN

    If Left(sSubKey, 1) = "\" Then sSubKey = Mid(sSubKey, 2)

    'Open the registry key
    iErrCode = RegOpenKey(iKey, sSubKey, iHkey)

    'Continue if successful
    If iErrCode = ERROR_SUCCESS Then

        'Read the value from the registry into a buffer
        iErrCode = RegQueryValueEx(iHkey, sValueName, 0&, ValueType, sBuffer, ValueLen)

        'Close the registry key
        RegCloseKey iHkey

        If iErrCode = ERROR_SUCCESS Then

            'If successful ValueType contains data type of value and ValueLen its length
            Select Case ValueType
            Case REG_BINARY
                'Convert binary value to its string representation
                For i = 1 To ValueLen
                    sTempBuffer = sTempBuffer & funPadByte(Hex(Asc(Mid(sBuffer, i, 1)))) & " "
                Next

                funGetRegValue = sTempBuffer

            Case REG_DWORD
                'Convert long binary value to its string representation
                sTempBuffer = "0x"

                For i = 4 To 1 Step -1
                    sTempBuffer = sTempBuffer & funPadByte(Hex(Asc(Mid(sBuffer, i, 1))))
                Next

                funGetRegValue = sTempBuffer

            Case Else
                'If not binary, just return the string from the buffer
                funGetRegValue = Left(sBuffer, ValueLen - 1)

            End Select

            Exit Function
        End If
    End If

    'If an error occurred, return the default value
    funGetRegValue = vDefault

End Function


'***************************************************************************
'*
'* PROCEDURE NAME:  DELETE REGISTRY VALUE NAME
'*
'* DESCRIPTION:     Deletes a value from a registry key.
'*
'* Pass:
'*  (1) the iKey (e.g., HKEY_CURRENT_USER)
'*  (2) the sSubKey (e.g., "Software\Company\Product")
'*  (3) the value's name (e.g., "whatever")
'*
'***************************************************************************

Sub procDeleteRegValue(ByVal iKey As Long, ByVal sSubKey As String, ByVal sValueName As String)
Attribute procDeleteRegValue.VB_UserMemId = 1610612738

    Dim iErrCode As Long, iHkey As Long

    If Left(sSubKey, 1) = "\" Then sSubKey = Mid(sSubKey, 2)

    'Open the registry key
    iErrCode = RegOpenKey(iKey, sSubKey, iHkey)

    If iErrCode = ERROR_SUCCESS Then

        'Delete the registry value
        iErrCode = RegDeleteValue(iHkey, sValueName)

        'Close the registry key
        RegCloseKey iHkey
    End If

End Sub


'***************************************************************************
'*
'* PROCEDURE NAME:  DELETE REGISTRY KEY
'*
'* DESCRIPTION:     Deletes an entire key from the registry.
'*
'* Pass:
'*  (1) the iKey (e.g., HKEY_CURRENT_USER)
'*  (2) the sSubKey (e.g., "Software\Company\Product")
'*
'***************************************************************************

Sub procDeleteRegKey(ByVal iKey As Long, ByVal sSubKey As String)
Attribute procDeleteRegKey.VB_UserMemId = 1610612739

    Dim iErrCode As Long

    If Left(sSubKey, 1) = "\" Then sSubKey = Mid(sSubKey, 2)

    'Delete the key
    iErrCode = RegDeleteKey(iKey, sSubKey)

End Sub


'***************************************************************************
'*
'* FUNCTION NAME:   CONVERT HEX TEXT TO A NUMBER
'*
'* DESCRIPTION:     Converts text that looks like a hex number to a true
'*                  number.  For example, "41" is converted to the hex
'*                  value 41h = 65 ("A").  Returns the text string whose
'*                  Ascii codes are the number.
'*
'***************************************************************************

Function funHexConvert(ByVal ByteStr)
Attribute funHexConvert.VB_UserMemId = 1610612740


    Dim sTemp As String, i As Integer

    'Add a zero to make the string an even length
    If Len(ByteStr) Mod 2 = 1 Then ByteStr = "0" & ByteStr

    'Step through the hex string 2 characters at a time
    For i = 1 To Len(ByteStr) Step 2

        'Convert the hex characters to the text string
        sTemp = sTemp & Chr("&h" & Mid(ByteStr, i, 2))
    Next

    funHexConvert = sTemp

End Function


'***************************************************************************
'*
'* FUNCTION NAME:   CONVERT HEX TO FLIPPED VALUE
'*
'* DESCRIPTION:     Converts a hex-like string to a true number with the
'*                  least significant digit first, as DWORD values are
'*                  stored in the registry.
'*
'***************************************************************************

Function funFlipAndHexConvert(ByVal ByteStr)
Attribute funFlipAndHexConvert.VB_UserMemId = 1610612741


    Dim sTemp As String, i As Integer

    'Take the left-most 8 characters if longer
    ByteStr = Left(ByteStr, 8)

    'Pad out with zeros if shorter then 8 characters
    ByteStr = String(8 - Len(ByteStr), "0") & ByteStr

    'Loop backwards through the string, two at a time
    For i = 7 To 1 Step -2

        'Add the character equivalent to the string
        sTemp = sTemp & Chr("&h" & Mid(ByteStr, i, 2))
    Next

    funFlipAndHexConvert = sTemp

End Function


'***************************************************************************
'*
'* FUNCTION NAME:   PAD OUT A BINARY NUMBER
'*
'* DESCRIPTION:     Pads a leading zero if needed to make a byte string an
'*                  even number of characters in length
'*
'***************************************************************************

Function funPadByte(ByVal ByteStr As String) As String
Attribute funPadByte.VB_UserMemId = 1610612742

    'Add an extra zero if required
    Const sProc As String = "mod7_RegFuncs.funPadByte"    'ErrorHandler:$$N=mod7_RegFuncs.funPadByte

    If Len(ByteStr) = 1 Then ByteStr = "0" & ByteStr

    funPadByte = ByteStr

End Function


'***************************************************************************
'*
'* FUNCTION NAME:   GET REGISTRY ERROR TEXT
'*
'* DESCRIPTION:     Returns the string for a given registry error code.
'*
'***************************************************************************

Function funRegErrorText(ByVal iErrorCode As Long)
Attribute funRegErrorText.VB_UserMemId = 1610612743

    'Process the error code
    Select Case iErrorCode
    Case ERROR_SUCCESS
        funRegErrorText = "Success"

    Case ERROR_BADDB
        funRegErrorText = "Error - Corrupt Registry"

    Case ERROR_BADKEY
        funRegErrorText = "Error - Bad iKey"

    Case ERROR_CANTOPEN
        funRegErrorText = "Error - Can't Open"

    Case ERROR_CANTREAD
        funRegErrorText = "Error - Can't Read"

    Case ERROR_CANTWRITE
        funRegErrorText = "Error - Can't Write"

    Case ERROR_REGISTRY_RECOVERED
        funRegErrorText = "Error - Registry File Recovered"

    Case ERROR_REGISTRY_CORRUPT
        funRegErrorText = "Error - Corrupt Registry"

    Case ERROR_REGISTRY_IO_FAILED
        funRegErrorText = "Error - File I/O Failed"

    Case ERROR_NOT_REGISTRY_FILE
        funRegErrorText = "Error - File Not in Registry Format"

    Case ERROR_KEY_DELETED
        funRegErrorText = "Error - iKey Marked for Deletion"

    Case ERROR_NO_LOG_SPACE
        funRegErrorText = "Error - No Log Space"

    Case ERROR_KEY_HAS_CHILDREN
        funRegErrorText = "Error - iKey Has Children"

    Case ERROR_CHILD_MUST_BE_VOLATILE
        funRegErrorText = "Error - Child Must Be Volatile"

    Case Else
        funRegErrorText = "Error - Unknown Type"

    End Select

End Function



