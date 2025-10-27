'#Language "WWB-COM"

' POLYTEC CODE MODULE


Option Explicit

Declare Function RegCreateKeyEx 	Lib "advapi32.dll" Alias "RegCreateKeyExA"  (ByVal hKey As PortInt _
	, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As PortInt, ByVal dwOptions As Long _
	, ByVal samDesired As Long, ByVal lpSecurityAttributes As PortInt, ByRef phkResult As PortInt _
	, ByRef lpdwDisposition As Long) As Long

Declare Function RegOpenKeyEx 		Lib "advapi32.dll" Alias "RegOpenKeyExA" 	(ByVal hKey As PortInt _
	, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long _
	, ByRef phkResult As PortInt) As Long
	
Declare Function RegCloseKey 		Lib "advapi32.dll" 							(ByVal hKey As PortInt) As Long

Declare Function RegQueryValueEx 	Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As PortInt _
	, ByVal lpValueName As String, ByVal lpReserved As PortInt, ByRef lpType As Long, ByRef lpData As Any _
	, ByRef lpcbData As Long) As Long

Declare Function RegQueryValueStrEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As PortInt _
	, ByVal lpValueName As String, ByVal lpReserved As PortInt, ByRef lpType As Long, ByVal lpData As String _
	, ByRef lpcbData As Long) As Long

Declare Function RegSetValueLngEx 	Lib "advapi32.dll" Alias "RegSetValueExA"   (ByVal hKey As PortInt _
	, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByRef lpData As Long _
	, ByVal cbData As Long) As Long

Declare Function RegSetValueBoolEx 	Lib "advapi32.dll" Alias "RegSetValueExA"   (ByVal hKey As PortInt _
	, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByRef lpData As Boolean _
	, ByVal cbData As Long) As Long

Declare Function RegSetValueStrEx 	Lib "advapi32.dll" Alias "RegSetValueExA"   (ByVal hKey As PortInt _
	, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String _
	, ByVal cbData As Long) As Long

Public Const KEY_QUERY_VALUE 	As Long = &H1
Public Const KEY_SET_VALUE 		As Long = &H2
Public Const KEY_CREATE_SUB_KEY As Long = &H4

Public Const HKEY_LOCAL_MACHINE As PortInt = &H80000002
Public Const HKEY_CURRENT_USER 	As PortInt = &H80000001

Public Const ERROR_SUCCESS 		As Long = 0&

Public Const REG_SZ 			As Long = 1
Public Const REG_BINARY 		As Long = 3
Public Const REG_DWORD 			As Long = 4

Declare Function FormatMessage      Lib "kernel32.dll" Alias "FormatMessageA"   (ByVal dwFlags As Long _
	, lpSource As PortInt, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String _
	, ByVal nSize As Long, Arguments As PortInt) As Long
	
Declare Function GetTempFileName	Lib "kernel32.dll" Alias "GetTempFileNameA" (ByVal lpszPath As String _
	, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long

Declare Function GetTempPath 		Lib "kernel32.dll" Alias "GetTempPathA"     (ByVal nBufferLength As Long _
	, ByVal lpBuffer As String) As Long

Public Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Public Const FORMAT_MESSAGE_ARGUMENT_ARRAY 	= &H2000
Public Const FORMAT_MESSAGE_FROM_HMODULE 	= &H800
Public Const FORMAT_MESSAGE_FROM_STRING 	= &H400
Public Const FORMAT_MESSAGE_FROM_SYSTEM 	= &H1000
Public Const FORMAT_MESSAGE_IGNORE_INSERTS 	= &H200
Public Const FORMAT_MESSAGE_MAX_WIDTH_MASK 	= &HFF

Public Const LANG_USER_DEFAULT As Long 		= &H400&
Public Const REG_OPTION_NON_VOLATILE As Long = &H0&

Function APIErrorDescription(ByVal ErrLastDllError As Long) As String

  Dim sBuffer As String
  Dim lBufferLen As Long

  'Allocate memory for return value:
  sBuffer = Space$(1024)

  'Transform error code to error string:
  lBufferLen = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_MAX_WIDTH_MASK Or FORMAT_MESSAGE_IGNORE_INSERTS, _
                     0&, ErrLastDllError, LANG_USER_DEFAULT, sBuffer, Len(sBuffer), 0)

  If lBufferLen > 0 Then
    'Error string was found
    APIErrorDescription = Left$(sBuffer, lBufferLen)
  Else
    'Erros string was not found. return a default error string
    APIErrorDescription = "unknown error: &H" & Hex$(ErrLastDllError)
  End If
End Function
