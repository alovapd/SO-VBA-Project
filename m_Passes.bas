Attribute VB_Name = "m_Passes"
Function Passes() As Boolean
    
    Dim arrUsers As Variant
    Dim assPass As Variant
    Dim strUsers As String
    Dim strPass As String
    Dim i As Long
    Dim x As Long
                '  1     2       3        4       5      6      7         8        9       10         11      12     13      14
    strUsers = "Snaith|Biddle|Brackett|Grissom|Penkava|Scow|Kennedy, R|Whipple|Adams, D|Ainsworth|Henderson|Kocina|Waldon|Sickler"
                        '     15       16     17      18     19      20      21    22    23    24
    strUsers = strUsers & "|Adams, C|Avery|Bartlett|Cam, B|Cam, T|Campbell|Cicero|Cote|Crisp|Osborne|"
                '  1     2       3        4       5      6      7      8      9       10       11       12     13     14
     strPass = "Snaith|Biddle|Brackett|Grissom|Penkava|Scow|Kennedy|Whipple|Adams|Ainsworth|Henderson|Kocina|Waldon|Sickler"
                    '         15    16    17      18   19   20        21   22    23     24
     strPass = strPass & "|AdamsC|Avery|Bartlett|Cam|CamT|Campbell|Cicero|Cote|Crisp|Osborne|"
                
    arrUsers = Split(strUsers, "|")
    arrpass = Split(strPass, "|")
    
    For i = 0 To UBound(arrUsers)
        If fm_Password.cb_User.Value = arrUsers(i) Then
            If fm_Password.txt_Pass.Value = arrpass(i) Then
                Passes = True
                Exit For
            Else
                Passes = False
                Exit For
            End If
        End If
    Next i
 
End Function

'Denton
'Duke
'Ettinger
'French
'Friend
'Grieve
'Groom
'Hagert
'Hermant
'Hohl
'Johnson
'Josephson
'Kennedy , T
'Lance
'Lenhardt
'Lewis
'Mcgrath
'Mckay
'Murillo
'Osborne
'Prins
'Richmond
'Shuart
'Strohmeyer
'Tuff
'Winston

