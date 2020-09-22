VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAddressCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'////  Simple class for checking the validity of an e-mail address.
'////  Joe Bourne 2001
'////  zhap018@aol.com
'////  Feel free to use this, but please Mail me and let me know
'////  what you use this for - Cheers!!


'////  Constant declerations

'Valid punctuation allowed in an e-mail addresses - sure i've missed some!
Private Const VALID_PUNC As String = "_-"  'Add any extra valid punctuation marks here if i forgot any!
Private Const AT_SYMBOL As String = "@"
Private Const DOT_SYMBOL As String = "."
Private Const UNDERSCORE_SYMBOL As String = "_" 'Not sure if we are allowed this after the @ symbol - anyone know??


'////  Class house keeping
Private Sub Class_Initialize()
    Debug.Print "clsAdressCheck: Initialize"
End Sub
Private Sub Class_Terminate()
    Debug.Print "clsAdressCheck: Terminate"
End Sub



'////  function to check an e-mail address (as string)
'////  Will return true if the address is valid, Fasle if not
Public Function CheckEmailAddress(ByVal strAddress As String) As Boolean
    Dim i As Integer             'current character index
    Dim intLen As Integer        'Address length
    Dim strChar As String        'a single character from the string
    Dim intPos As Integer        'Position of a character within a string
    Dim blnPunc As Boolean       'Set to true when we find a punctuation mark
    Dim intAtPos As Integer      'Postiton of the @ character
    Dim intLastDotPos As Integer 'Position of last Dot Found in string


    On Local Error GoTo errHandler    'Simple error handling

    strAddress = Trim$(strAddress)   'Remove any whitespace characters

    If strAddress = vbNullString Then
        Exit Function   'Zero length address - Bail out
    End If

    strAddress = LCase$(strAddress)  'make address lowercase for comparison purposes
    intLen = Len(strAddress)         'Save the length of the address for later use

    '////  loop through each character
    For i = 1 To intLen        'Loop for each character in the address
        strChar = Mid(strAddress, i, 1) 'Take a character from the address

        If (Not (strChar Like "[a-z]")) And (Not (strChar Like "[0-9]")) Then
            'Not an alpha numeric character - check to see if it's punctuation
            If InStr(1, VALID_PUNC & DOT_SYMBOL & AT_SYMBOL, strChar) > 0 Then
                If blnPunc = True Then
                    'Second punctuation mark in succession - bail!
                    'Not allowed two punctuation characters in a row
                    Debug.Print "Not allowed two punctuation marks in a row!: " & strChar
                    Exit Function   'Bail out!
                Else
                    blnPunc = True  'Set flag to true to show we've found punctuation
                    '////  Check the first and last char aren't punctuation
                    If i = 1 Or i = intLen Then
                        Debug.Print "Not allowed punctuation as first or last character!: " & strChar
                        Exit Function   'Bail out!
                    End If
                    '////  Now lets look for the @ symbol - not allowed at begining or end,
                    '////  and must have at least one dot following it (Not directly)
                    If strChar = AT_SYMBOL Then
                        If intAtPos = 0 Then
                            intAtPos = i    'Save position of the '@' Character
                        Else
                            Debug.Print "Not allowed two @ symbols!"
                            Exit Function   'Bail out!
                        End If
                    End If
                    If strChar = DOT_SYMBOL Then    'If its a dot character
                        intLastDotPos = i   'Save position of the last dot found
                    End If
                    'Add any extra punctuation checks here
                    'eg, not sure if allowed an underscore '_' after the '@' symbol anyone know??
                    'If strChar = UNDERSCORE_SYMBOL And intAtPos > 0 Then
                    '    Debug.Print "Not allowed _ after @ symbol!"
                    '    Exit Function   'Bail out!
                    'End If
                End If
            Else
                Debug.Print "BAD CHARACTER IN ADDRESS!: " & strChar
                Exit Function   'Bail out!
            End If
        Else
            blnPunc = False 'Reset the punctuation flag - we've got a character
        End If
    Next i  'Repeat for next character

    '////  There must be an @ in the address!
    If intAtPos = 0 Or intLastDotPos = 0 Then
        Debug.Print "NO @ or . in address!"
        Exit Function   'Bail out!
    End If
    
    
    'Check there is a Dot after the @
    If intLastDotPos < intAtPos Then
        Debug.Print "NO DOT FOLLOWING THE @!"
        Exit Function   'Bail out!
    End If

    '////  if we get here we've passes all the tests - its a good address!
    CheckEmailAddress = True

    Exit Function
errHandler: 'Simple error handling
    On Local Error Resume Next  'Prevent error handler raising an error
    Debug.Print "ERROR: " & Err.Description
    CheckEmailAddress = False
End Function






