Attribute VB_Name = "modCryptData"
'@Lang VBA
                                                                                                                                                                                            ' _
    |||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||                                                                                     ' _
    ||||||||||||||||||||||||||           modCryptData (v1.0)         ||||||||||||||||||||||||||||||||||                                                                                     ' _
    |||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
                                                                                                                                                                                            ' _
    AUTHOR:   Kallun Willock | Eduardo A. Morcillo                                                                                                                                          ' _
              (Many thanks to Extreminador @ Discord for helping me with the testing)                                                                                                       ' _
    URL:      Reddit - https://www.reddit.com/r/vba/comments/13s2hk2/encryptingauthenticating_in_vba/                                                                                       ' _
    SEE:      https://www.mvps.org/emorcillo/en/code/vb6/protect.shtml                                                                                                                      ' _
              CryptProtectData - https://learn.microsoft.com/en-us/windows/win32/api/dpapi/nf-dpapi-cryptprotectdata                                                                        ' _
              CryptUnprotectData - https://learn.microsoft.com/en-us/windows/win32/api/dpapi/nf-dpapi-cryptunprotectdata
                                                                                                                                                                                            ' _
    NOTES:    Based on code published by Eduardo A. Morcillo - the published code referenced external functions and                                                                         ' _
              so was incomplete, and it needed several corrections and the API declarations needed to be added/updated                                                                      ' _
              in order to be 64-bit compatible.
                                                                                                                                                                                            ' _
    LICENSE:  MIT                                                                                                                                                                           ' _
    VERSION:  1.0        08/06/2023         Published v1.0 on Github.                                                                                                                       ' _
                                                                                                                                                                                            ' _                                                                                                                                                   ' _
    TODO:     [ ] Error Handling                                                                                                                                                            ' _

Option Explicit

' See dpapi.h for enumerations, structs and API declarations
' https://github.com/Alexpux/mingw-w64/blob/master/mingw-w64-headers/include/dpapi.h

Enum ProtectDataPromptFlags
   PromptOnUnprotect = &H1      ' This flag can be combined with PromptOnProtect to enforce the UI (user interface) policy of the caller.
                                ' When CryptUnprotectData is called, the dwPromptFlags specified in the CryptProtectData call are enforced.
   PromptOnProtect = &H2        ' This flag is used to provide the prompt for the protect phase.
   Strong = &H8
   RequireStrong = &H10
End Enum

Enum ProtectDataFlags
   UIForbidden = &H1            ' This flag is used for remote situations where the user interface (UI) is not an option.
                                ' When this flag is set and UI is specified for either the protect or unprotect operation,
                                ' the operation fails and GetLastError returns the ERROR_PASSWORD_RESTRICTION code.
   LocalMachine = &H4           ' Default. When this flag is set, it associates the data encrypted with the current computer
                                ' instead of with an individual user. Any user on the computer on which CryptProtectData
                                ' is called can use CryptUnprotectData to decrypt the data.
   CredSync = &H8
   Audit = &H10                 ' This flag generates an audit on protect and unprotect operations.
                                ' Audit log entries are recorded only if szDataDescr is not NULL and not empty.
   NoRecovery = &H20
   VerifyProtection = &H40      ' This flag verifies the protection of a protected BLOB.
                                ' If the default protection level configured of the host is higher
                                ' than the current protection level for the BLOB, the function
                                ' returns CRYPT_I_NEW_PROTECTION_REQUIRED to advise the caller to again
                                ' protect the plaintext contained in the BLOB.
   CredRegenerate = &H80
End Enum

#If VBA7 Then
    Private Declare PtrSafe Function CryptProtectData Lib "crypt32.dll" (pDataIn As Any, ByVal szDataDescr As LongPtr, pOptionalEntropy As Any, ByVal pvReserved As Long, pPromptStruct As Any, ByVal dwFlags As Long, pDataOut As Any) As Long
    Private Declare PtrSafe Function CryptUnprotectData Lib "crypt32.dll" (pDataIn As Any, ppszDataDescr As LongPtr, pOptionalEntropy As Any, ByVal pvReserved As Long, pPromptStruct As Any, ByVal dwFlags As Long, pDataOut As Any) As Long
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal length As LongPtr)
    Private Declare PtrSafe Function LocalFree Lib "kernel32.dll" (ByVal Ptr As LongPtr) As Long
    Private Declare PtrSafe Function lstrlenW Lib "kernel32.dll" (ByVal lpString As LongPtr) As Long
#Else
    Private Enum LongPtr
        [_]
    End Enum
    Private Declare Function CryptProtectData Lib "crypt32.dll" (pDataIn As Any, ByVal szDataDescr As LongPtr, pOptionalEntropy As Any, ByVal pvReserved As Long, pPromptStruct As Any, ByVal dwFlags As long, pDataOut As Any) As Long
    Private Declare Function CryptUnprotectData Lib "crypt32.dll" (pDataIn As Any, ppszDataDescr As LongPtr, pOptionalEntropy As Any, ByVal pvReserved As Long, pPromptStruct As Any, ByVal dwFlags As Long, pDataOut As Any) As Long
    Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As LongPtr)
    Private Declare Function LocalFree Lib "kernel32.dll" (ByVal Ptr As LongPtr) As Long
    Private Declare Function lstrlenW Lib "kernel32.dll" (ByVal lpString As LongPtr) As Long
#End If

' The CRYPTPROTECT_PROMPTSTRUCT structure provides the text of a prompt and information
' about when and where that prompt is to be displayed when using the CryptProtectData and
' CryptUnprotectData functions.

Private Type CRYPTPROTECT_PROMPTSTRUCT
   cbSize           As Long
   dwPromptFlags    As ProtectDataPromptFlags
   hwndApp          As LongPtr
   szPrompt         As LongPtr
End Type

Private Type CRYPTOAPI_BLOB
   cbData           As Long
   pbData           As LongPtr
End Type

Public Function ProtectData(Data() As Byte, _
                            Optional ByVal DataDescription As String, _
                            Optional ByVal ParentWnd As LongPtr, _
                            Optional ByVal DialogTitle As String, _
                            Optional ByVal Flags As ProtectDataFlags = LocalMachine, _
                            Optional ByVal PromptFlags As ProtectDataPromptFlags) As Byte()
    
    ' The CryptProtectData API requires the target data to be a byte array.
    ' This API was designed for use in specific scenarios - namely, for use by  a user with the same
    ' logon credential as the user who encrypted the data can decrypt the data.
    ' The encryption and decryption 'usually' must be done on the same computer.
    
    Dim DataBlobIn          As CRYPTOAPI_BLOB
    Dim DataBlobOut         As CRYPTOAPI_BLOB
    Dim Prompt              As CRYPTPROTECT_PROMPTSTRUCT
    Dim EncryptedData()     As Byte
    Dim Result              As Long

    With DataBlobIn
       .cbData = UBound(Data) - LBound(Data) + 1
       .pbData = VarPtr(Data(0))
    End With
    
    With Prompt
       .cbSize = Len(Prompt)
       .hwndApp = ParentWnd
       .dwPromptFlags = PromptFlags
       If Len(DialogTitle) Then .szPrompt = StrPtr(DialogTitle)
    End With
    
    Result = CryptProtectData(DataBlobIn, StrPtr(DataDescription), ByVal 0&, 0&, Prompt, Flags, DataBlobOut)
    If Result = 0 Then Err.Raise &H80070000 Or Err.LastDllError
       
    ReDim EncryptedData(0 To DataBlobOut.cbData - 1)
    CopyMemory EncryptedData(0), ByVal DataBlobOut.pbData, DataBlobOut.cbData      ' Copy the encrypted data to a byte array
    
    ProtectData = EncryptedData                                                    ' Return the encrypted data
    LocalFree DataBlobOut.pbData                                                   ' Release the returned data
   
End Function

Public Function UnProtectData(Data() As Byte, _
                            Optional ByRef DataDescription As Variant, _
                            Optional ByVal ParentWnd As LongPtr, _
                            Optional ByVal DialogTitle As String, _
                            Optional ByVal Flags As ProtectDataFlags = LocalMachine, _
                            Optional ByVal PromptFlags As ProtectDataPromptFlags) As Byte()
    
    Dim DataBlobIn         As CRYPTOAPI_BLOB
    Dim DataBlobOut        As CRYPTOAPI_BLOB
    Dim Prompt             As CRYPTPROTECT_PROMPTSTRUCT
    Dim DecryptedData()    As Byte
    Dim DescriptionPtr     As LongPtr
    Dim Result             As Long

    With DataBlobIn
       .cbData = UBound(Data) - LBound(Data) + 1
       .pbData = VarPtr(Data(0))
    End With
    
    With Prompt
       .cbSize = Len(Prompt)
       .hwndApp = ParentWnd                                                         ' This works even if left blank
       .dwPromptFlags = PromptFlags
       If Len(DialogTitle) Then .szPrompt = StrPtr(DialogTitle)
    End With
    
    Result = CryptUnprotectData(DataBlobIn, DescriptionPtr, ByVal 0&, 0&, Prompt, Flags, DataBlobOut)
    If Result = 0 Then Err.Raise &H80070000 Or Err.LastDllError
       
    ReDim DecryptedData(0 To DataBlobOut.cbData - 1)
    CopyMemory DecryptedData(0), ByVal DataBlobOut.pbData, DataBlobOut.cbData       ' Copy the data to a byte array
    
    If Not IsMissing(DataDescription) Then
        DataDescription = Pointer2String(DescriptionPtr)                            ' Get the description
    End If
    
    UnProtectData = DecryptedData
    
    LocalFree DataBlobOut.pbData                                                    ' Release the returned data pointer
    LocalFree DescriptionPtr
   
End Function

Public Function Pointer2String(ByVal StringPointer As LongPtr) As String

    If StringPointer Then
        Dim StringLength As Long
        StringLength = lstrlenW(StringPointer)
        If StringLength Then
            Pointer2String = Space$(StringLength)
            CopyMemory ByVal StrPtr(Pointer2String), ByVal StringPointer, StringLength * 2
        Else
            ' It is not necessarily an error if the length of the string is 0, but I set out the
            ' code to raise an error below for completeness.
            ' Err.Raise Number:=vbObjectError + 514, Description:="Pointer returns a 0 length string"
        End If
    Else
        Err.Raise Number:=vbObjectError + 513, DESCRIPTION:="No Pointer provided"
    End If

End Function

Sub TestRoutine()

    Dim SecretAPIKey() As Byte, DummyDescription As String
    Dim EncryptedData() As Byte, UnEncryptedData() As Byte
    
    ' The CryptProtectData API requires the target data to be a byte array. In this test routine,
    ' I demonstrate a quick and easy way of converting a string to a byte array ... by simply assigning
    ' the string literal to the declared byte array variable.
    
    SecretAPIKey = "ThisIsMySuperSecretRedditAPIKey"
    
    ' The CryptProtectData API allows for you to assign a description to the encrypted data.
    ' This is optional. It needs to be in a string data type.
    
    DummyDescription = "Super Secret Reddit API Key"
    
    Dim TempString As String
    TempString = SecretAPIKey
    
    Debug.Print "1. Original:   " & TempString & vbNewLine
    EncryptedData = ProtectData(SecretAPIKey, DummyDescription)
    Debug.Print "2. Encrypted:  " & StrConv(EncryptedData, vbUnicode) & vbNewLine
    
    ' In order to retrieve the Description information, call the UnProtectData function
    ' with an empty string variable
    
    DummyDescription = ""
    TempString = ""
    
    UnEncryptedData = UnProtectData(EncryptedData, DummyDescription)
    TempString = UnEncryptedData
    
    Debug.Print "3. UnEncrypted:  " & TempString
    Debug.Print "4. Description:  " & DummyDescription
    
End Sub


Sub TestRoutine2()

    Dim SecretAPIKey() As Byte
    Dim EncryptedData() As Byte, UnEncryptedData() As Byte
    
    SecretAPIKey = "ThisIsMySuperSecretRedditAPIKey"
    
    Dim TempString As String
    TempString = SecretAPIKey
    
    Debug.Print "1. Original:   " & TempString & vbNewLine
    EncryptedData = ProtectData(SecretAPIKey, , , "This is the dialog title", , PromptOnProtect)
    Debug.Print "2. Encrypted:  " & StrConv(EncryptedData, vbUnicode) & vbNewLine
    
    TempString = ""
    
    UnEncryptedData = UnProtectData(EncryptedData)
    TempString = UnEncryptedData
    
    Debug.Print "3. UnEncrypted:  " & TempString
    
End Sub


