Attribute VB_Name = "Module4"
'required?
Option Explicit

  Private m_lngRetVal As Long
  Private Const REG_NONE As Long = 0                  ' No value type
  Private Const REG_SZ As Long = 1                    ' nul terminated string
  Private Const REG_EXPAND_SZ As Long = 2             ' nul terminated string w/enviornment var
  Private Const REG_BINARY As Long = 3                ' Free form binary
  Private Const REG_DWORD As Long = 4                 ' 32-bit number
  Private Const REG_DWORD_LITTLE_ENDIAN As Long = 4   ' 32-bit number (same as REG_DWORD)
  Private Const REG_DWORD_BIG_ENDIAN As Long = 5      ' 32-bit number
  Private Const REG_LINK As Long = 6                  ' Symbolic Link (unicode)
  Private Const REG_MULTI_SZ As Long = 7              ' Multiple Unicode strings
  Private Const REG_RESOURCE_LIST As Long = 8         ' Resource list in the resource map
  Private Const REG_FULL_RESOURCE_DESCRIPTOR As Long = 9 ' Resource list in the hardware description
  Private Const REG_RESOURCE_REQUIREMENTS_LIST As Long = 10
  Private Const KEY_QUERY_VALUE As Long = &H1
  Private Const KEY_SET_VALUE As Long = &H2
  Private Const KEY_CREATE_SUB_KEY As Long = &H4
  Private Const KEY_ENUMERATE_SUB_KEYS As Long = &H8
  Private Const KEY_NOTIFY As Long = &H10
  Private Const KEY_CREATE_LINK As Long = &H20
  Private Const KEY_ALL_ACCESS As Long = &H3F
  Public Const HKEY_CLASSES_ROOT As Long = &H80000000
  Public Const HKEY_CURRENT_USER As Long = &H80000001
  Public Const HKEY_LOCAL_MACHINE As Long = &H80000002
  Public Const HKEY_USERS As Long = &H80000003
  Public Const HKEY_PERFORMANCE_DATA As Long = &H80000004
  Public Const HKEY_CURRENT_CONFIG As Long = &H80000005
  Public Const HKEY_DYN_DATA As Long = &H80000006
  Private Const ERROR_SUCCESS As Long = 0
  Private Const ERROR_ACCESS_DENIED As Long = 5
  Private Const ERROR_NO_MORE_ITEMS As Long = 259
  Private Const REG_OPTION_NON_VOLATILE As Long = 0
  Private Const REG_OPTION_VOLATILE As Long = &H1
  Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal lngRootKey As Long) As Long
  
  Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" _
            (ByVal lngRootKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
  
  Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" _
            (ByVal lngRootKey As Long, ByVal lpSubKey As String) As Long
  
  Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" _
            (ByVal lngRootKey As Long, ByVal lpValueName As String) As Long
  
  Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" _
            (ByVal lngRootKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
  
  Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" _
            (ByVal lngRootKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
             lpType As Long, lpData As Any, lpcbData As Long) As Long
  
  Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" _
            (ByVal lngRootKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, _
             ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Public Function regDelSubKey(ByVal lngRootKey As Long, _
                             ByVal strRegKeyPath As String, _
                             ByVal strRegSubKey As String)
    
'    regDelete_Sub_Key HKEY_CURRENT_USER, "Software\AAA-Registry Test\Products", "StringTestData"
  Dim lngKeyHandle As Long
  If regIsKey(lngRootKey, strRegKeyPath) Then
      m_lngRetVal = RegOpenKey(lngRootKey, strRegKeyPath, lngKeyHandle)
      m_lngRetVal = RegDeleteValue(lngKeyHandle, strRegSubKey)
      m_lngRetVal = RegCloseKey(lngKeyHandle)
  End If
End Function

Public Function regIsKey(ByVal lngRootKey As Long, _
                                  ByVal strRegKeyPath As String) As Boolean
    
'    strKeyQuery = regIsKey(HKEY_CURRENT_USER, "Software\AAA-Registry Test\Products")
  Dim lngKeyHandle As Long
  lngKeyHandle = 0
  m_lngRetVal = RegOpenKey(lngRootKey, strRegKeyPath, lngKeyHandle)
  If lngKeyHandle = 0 Then
      regIsKey = False
  Else
      regIsKey = True
  End If
  m_lngRetVal = RegCloseKey(lngKeyHandle)
End Function

Public Function regReadKey(ByVal lngRootKey As Long, _
                           ByVal strRegKeyPath As String, _
                           ByVal strRegSubKey As String) As Variant
'    strKeyQuery = regReadKey(HKEY_CURRENT_USER, "Software\AAA-Registry Test\Products", "StringTestData")
  Dim intPosition As Integer
  Dim lngKeyHandle As Long
  Dim lngDataType As Long
  Dim lngBufferSize As Long
  Dim lngBuffer As Long
  Dim strBuffer As String
  Dim strTemp As String
  lngKeyHandle = 0
  lngBufferSize = 0
  m_lngRetVal = RegOpenKey(lngRootKey, strRegKeyPath, lngKeyHandle)
  If lngKeyHandle = 0 Then
      regReadKey = ""
      m_lngRetVal = RegCloseKey(lngKeyHandle)   ' always close the handle
      Exit Function
  End If
  m_lngRetVal = RegQueryValueEx(lngKeyHandle, strRegSubKey, 0&, _
                         lngDataType, ByVal 0&, lngBufferSize)
  If lngKeyHandle = 0 Then
      regReadKey = ""
      m_lngRetVal = RegCloseKey(lngKeyHandle)   ' always close the handle
      Exit Function
  End If
  Select Case lngDataType
         Case REG_SZ:       ' String data (most common)
              strBuffer = Space(lngBufferSize)
      
              m_lngRetVal = RegQueryValueEx(lngKeyHandle, strRegSubKey, 0&, 0&, _
                                     ByVal strBuffer, lngBufferSize)
              
              If m_lngRetVal <> ERROR_SUCCESS Then
                  regReadKey = ""
              Else
                  intPosition = InStr(1, strBuffer, Chr(0))  ' look for the first null char
                  If intPosition > 0 Then
                      strTemp = Mid$(strBuffer, 1, intPosition - 1)
                      regReadKey = strTemp
                  Else
                      strTemp = strBuffer
                      regReadKey = strTemp
                  End If
              End If
              
         Case REG_DWORD:    ' Numeric data (Integer)
              m_lngRetVal = RegQueryValueEx(lngKeyHandle, strRegSubKey, 0&, lngDataType, _
                                     lngBuffer, 4&)  ' 4& = 4-byte word (long integer)
              If m_lngRetVal <> ERROR_SUCCESS Then
                  regReadKey = ""
              Else
                  regReadKey = lngBuffer
              End If
         
         Case Else:    ' unknown
              regReadKey = ""
  End Select
  m_lngRetVal = RegCloseKey(lngKeyHandle)
End Function

Public Sub regWriteSubKey(ByVal lngRootKey As Long, ByVal strRegKeyPath As String, _
                               ByVal strRegSubKey As String, varRegData As Variant)
'    regWriteSubKey HKEY_CURRENT_USER, _
'                      "Software\AAA-Registry Test\Products", _
'                      "StringTestData", "22 Jun 1999"
  Dim lngKeyHandle As Long
  Dim lngDataType As Long
  Dim lngKeyValue As Long
  Dim strKeyValue As String
  If IsNumeric(varRegData) Then
      lngDataType = REG_DWORD
  Else
      lngDataType = REG_BINARY
  End If
  m_lngRetVal = RegCreateKey(lngRootKey, strRegKeyPath, lngKeyHandle)
  Select Case lngDataType
         Case REG_BINARY:       ' String data
              strKeyValue = Trim(varRegData) & Chr(0)     ' null terminated
              m_lngRetVal = RegSetValueEx(lngKeyHandle, strRegSubKey, 0&, lngDataType, _
                                          ByVal strKeyValue, Len(strKeyValue))
                                   
         Case REG_DWORD:    ' numeric data
              lngKeyValue = CLng(varRegData)
              m_lngRetVal = RegSetValueEx(lngKeyHandle, strRegSubKey, 0&, lngDataType, _
                                          lngKeyValue, 4&)  ' 4& = 4-byte word (long integer)
                                   
  End Select
  m_lngRetVal = RegCloseKey(lngKeyHandle)
End Sub

Public Function regWriteKey(ByVal lngRootKey As Long, ByVal strRegKeyPath As String)
'   regWriteKey HKEY_CURRENT_USER, "Software\AAA-Registry Test"
'   regWriteKey HKEY_CURRENT_USER, "Software\AAA-Registry Test\Products"
  Dim lngKeyHandle As Long
  m_lngRetVal = RegCreateKey(lngRootKey, strRegKeyPath, lngKeyHandle)
  m_lngRetVal = RegCloseKey(lngKeyHandle)
End Function

Public Function regDelKey(ByVal lngRootKey As Long, _
                                ByVal strRegKeyPath As String, _
                                ByVal strRegKeyName As String) As Boolean
'    regDelKey HKEY_CURRENT_USER, "Software", "AAA-Registry Test"
  Dim lngKeyHandle As Long
  regDelKey = False
  If regIsKey(lngRootKey, strRegKeyPath) Then
      m_lngRetVal = RegOpenKey(lngRootKey, strRegKeyPath, lngKeyHandle)
      m_lngRetVal = RegDeleteKey(lngKeyHandle, strRegKeyName)
      If m_lngRetVal = 0 Then regDelKey = True
      m_lngRetVal = RegCloseKey(lngKeyHandle)
  End If
End Function





