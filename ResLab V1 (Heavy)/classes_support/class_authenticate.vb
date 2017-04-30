Imports System.Security.Cryptography
Imports ResLab_V1_Heavy.cDatabaseInfo

Public Class class_authenticate

    Private _saltLength As Integer = 4

    Public Function authenticate_login(username As String, password As String) As Boolean

        'Get password for username from database
        Dim pw_stored As String = Me.get_password(username)

        If Not IsNothing(pw_stored) Then
            Return Me.compare_Passwords(pw_stored, password)
        Else
            'Username not found
            Return False
        End If

    End Function

    Public Function encrypt_password(password As String) As String
        'Takes a password as plain text and return hashed and salted version as a byte array converted to a string 

        Dim pw_bytes As Byte() = Me.create_DbPassword(System.Text.Encoding.Unicode.GetBytes(password), Me.create_salt)
        Dim pw_string As String = Convert.ToBase64String(pw_bytes)

        Return pw_string

    End Function

    Private Function create_salt() As Byte()

        Dim salt As Byte()
        ReDim salt(_saltLength - 1)

        'Create a salt value.
        Dim rng = New RNGCryptoServiceProvider()
        rng.GetBytes(salt)

        salt(0) = 197
        salt(1) = 176
        salt(2) = 104
        salt(3) = 213

        Return salt

    End Function

    Private Function create_DbPassword(ByVal Password As Byte(), saltValue As Byte()) As Byte()
        ' Create salted password, return as byte array


        ' Add the salt to the hash.
        Dim rawSalted(Password.Length + saltValue.Length - 1) As Byte
        Password.CopyTo(rawSalted, 0)
        saltValue.CopyTo(rawSalted, Password.Length)

        'Create the salted hash.         
        Dim sha1 As SHA1 = sha1.Create()
        Dim saltedPassword As Byte() = sha1.ComputeHash(rawSalted)

        ' Add the salt value to the salted hash.
        Dim dbPassword(saltedPassword.Length + saltValue.Length - 1) As Byte
        saltedPassword.CopyTo(dbPassword, 0)
        saltValue.CopyTo(dbPassword, saltedPassword.Length)

        Return dbPassword

    End Function

    Private Function compare_Passwords(ByVal storedPassword As String, enteredPassword As String) As Boolean
        ' Compare the entered password against the stored password.

        Dim i As Integer = 0

        'Convert entered plain text to a byte array
        Dim pw_entered_b As Byte() = System.Text.Encoding.Unicode.GetBytes(enteredPassword)

        'Convert stored base64 string to byte array
        Dim pw_stored_b As Byte() = Convert.FromBase64String(storedPassword)

        'Get the saved saltValue.
        Dim saltValue(_saltLength - 1) As Byte
        Dim saltOffset As Integer = pw_stored_b.Length - _saltLength
        For i = 0 To _saltLength - 1
            saltValue(i) = pw_stored_b(saltOffset + i)
        Next

        'Now hash entered password and salt with the stored salt value
        Dim pw_entered_b_h As Byte() = Me.create_DbPassword(pw_entered_b, saltValue)

        ' Compare the values.
        Return CompareByteArray(pw_stored_b, pw_entered_b_h)

    End Function

    Private Function CompareByteArray(ByVal array1 As Byte(), ByVal array2 As Byte()) As Boolean
        ' Compare the contents of two byte arrays.

        If (array1.Length <> array2.Length) Then
            Return False
        End If

        Dim i As Integer
        For i = 0 To array1.Length - 1
            If (array1(i) <> array2(i)) Then
                Return False
            End If
        Next

        Return True

    End Function

    Public Function update_password(personID As Long, encryptedPassword As String) As Boolean

        Dim d As Dictionary(Of String, String) = Nothing
        d.Add("personid", personID)
        d.Add("user_password", "'" & encryptedPassword & "'")
        Dim sql As String = cDAL.Build_UpdateQuery(eTables.persons, d)
        Return cDAL.Update_Record(sql)

    End Function

    Public Function get_password(user_name As String) As String

        Dim sql As String = "SELECT user_password FROM persons WHERE user_name='" & user_name & "'"
        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)

        If Not IsNothing(Ds) Then
            If Ds.Tables(0).Rows.Count = 0 Then
                'user_name not found
                Return Nothing
            ElseIf Ds.Tables(0).Rows.Count = 1 Then
                Return Ds.Tables(0).Rows(0)("user_password")
            Else
                '>1 user_name Should never happen
                Return Nothing
            End If
        Else
            Return Nothing
        End If

        Ds = Nothing

    End Function

End Class
