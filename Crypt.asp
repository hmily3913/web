<% 
Dim g_Key

Const g_CryptThis = "Now is the time for all good men to come to the aid of their country."
Const g_KeyLocation = "c:\key.txt"

g_Key = mid(ReadKeyFromFile(g_KeyLocation),1,Len(g_CryptThis))

Response.Write "<p>ORIGINAL STRING: " & g_CryptThis & "<p>"
Response.Write "<p>KEY VALUE: " & g_Key & "<p>"
Response.Write "<p>ENCRYPTED CYPHERTEXT: " & EnCrypt(g_CryptThis) & "<p>"
Response.Write "<p>DECRYPTED CYPHERTEXT: " & DeCrypt(EnCrypt(g_CryptThis)) & "<p>"

Function EnCrypt(strCryptThis)
Dim strChar, iKeyChar, iStringChar, I
for I = 1 to Len(strCryptThis)
iKeyChar = Asc(mid(g_Key,I,1))
iStringChar = Asc(mid(strCryptThis,I,1))
' *** uncomment below to encrypt with addition,
' iCryptChar = iStringChar + iKeyChar
iCryptChar = iKeyChar Xor iStringChar
strEncrypted = strEncrypted & Chr(iCryptChar)
next
EnCrypt = strEncrypted
End Function

Function DeCrypt(strEncrypted)
Dim strChar, iKeyChar, iStringChar, I
for I = 1 to Len(strEncrypted)
iKeyChar = (Asc(mid(g_Key,I,1)))
iStringChar = Asc(mid(strEncrypted,I,1))
' *** uncomment below to decrypt with subtraction 
' iDeCryptChar = iStringChar - iKeyChar 
iDeCryptChar = iKeyChar Xor iStringChar
strDecrypted = strDecrypted & Chr(iDeCryptChar)
next
DeCrypt = strDecrypted
End Function

Function ReadKeyFromFile(strFileName)
Dim keyFile, fso, f
set fso = Server.CreateObject("Scripting.FileSystemObject") 
set f = fso.GetFile(strFileName) 
set ts = f.OpenAsTextStream(1, -2)

Do While not ts.AtEndOfStream
keyFile = keyFile & ts.ReadLine
Loop 

ReadKeyFromFile = keyFile
End Function

%> 