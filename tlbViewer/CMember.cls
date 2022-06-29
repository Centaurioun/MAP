VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMember"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'License:   GPL
'Copyright: 2005 iDefense a Verisign Company
'Site:      http://labs.idefense.com
'
'Author:  David Zimmer <david@idefense.com, dzzie@yahoo.com>
'
'         This program is free software; you can redistribute it and/or modify it
'         under the terms of the GNU General Public License as published by the Free
'         Software Foundation; either version 2 of the License, or (at your option)
'         any later version.
'
'         This program is distributed in the hope that it will be useful, but WITHOUT
'         ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or
'         FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for
'         more details.
'
'         You should have received a copy of the GNU General Public License along with
'         this program; if not, write to the Free Software Foundation, Inc., 59 Temple
'         Place, Suite 330, Boston, MA 02111-1307 USA

Option Explicit

'this class encapsulates one function as defined in the typelib
'this class is owned by an CInterface class and held as part of its mMembers col
'this class in turn holds a direct ref to the member tli obj as well as a
'simplified and parsed collection of CArgument classes used by .Invoke to help
'call it


Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long

Public Enum FType
    ftsub = 0
    ftFunc = 1
    ftPropLet = 2
    ftPropGet = 3
    ftPropSet = 4
    ftUnknown = 5
    ftEvent = 6
End Enum

Public ClassGUID As String
Public CallType As FType
Public ProtoString As String
Public mMemberInfo As MemberInfo
Public ArgCount As Integer
Public SupportsFuzzing As Boolean 'entire method cant be fuzzed
Public Args As New Collection 'of CArgument
Public name As String
Public vTableOffset As Long

Private Declare Function GetTickCount Lib "kernel32" () As Long


Sub LoadMember(m As MemberInfo, ownerClassGuid As String, isEvent As Boolean)
    
    'If m.Name = "insertBefore" Then Stop
    
    Set mMemberInfo = m
    name = m.name
    'If Name = "OpenTextFile" Then Stop
    
    If isEvent Then CallType = ftEvent 'used in buildfuncname
    
    ClassGUID = ownerClassGuid
    ArgCount = m.Parameters.Count
    ProtoString = GetProtoString()
    vTableOffset = m.vTableOffset

    Dim a As CArgument
    Dim x As Long
    
    For Each a In Args
        If Not a.isWorthFuzzing Then x = x + 1
    Next
    
    If x = Args.Count Then SupportsFuzzing = False
    
End Sub





Private Function GetProtoString() As String
    
Dim pi As ParameterInfo
Dim TIResolved As TypeInfo
Dim arg As CArgument
Dim tmp() As String
Dim isFirst As Boolean
       
    SupportsFuzzing = True
    isFirst = True
     
    BuildFuncName tmp 'sets some props and builds prototype string head
       
    With mMemberInfo.Parameters
        If .Count Then
              glue tmp, "(" & vbCrLf
              
              For Each pi In mMemberInfo.Parameters
              
                    Set arg = New CArgument
                    If isFirst Then isFirst = False Else glue tmp, ", " & vbCrLf 'not first parameter
                    
                    arg.isDefault = pi.default
                    arg.isOptional = pi.Optional
                    
                    If pi.default Or pi.Optional Then
                          If .OptionalCount = -1 Then  'paramarray
                              arg.dataType = atParamAry
                              push tmp, "[ParamArray " 'This will be the only optional parameter
                              SupportsFuzzing = False
                          End If
                    End If
                    
                    With pi.VarTypeInfo
                          If (.VarType And Not (VT_ARRAY Or VT_VECTOR)) = 0 Then
                              ParseObjArg tmp, pi, arg, TIResolved
                          Else
                              ParseStdArg tmp, pi, arg
                          End If
                      
                          If pi.Optional Then
                              If pi.default Then
                                  glue tmp, ProduceDefaultValue(pi.DefaultValue, TIResolved)
                              End If
                              glue tmp, "]"
                          End If
                    End With
                    
                    arg.name = pi.name 'added late in game!
                    Args.Add arg
              Next
              
              glue tmp, vbCrLf & ")"
        Else
            glue tmp, "()", False
        End If
    End With

    GetReturnType tmp, False

    GetProtoString = Join(tmp, vbCrLf)
  
End Function

Private Sub ParseStdArg(tmp, pi As ParameterInfo, arg As CArgument)

    Dim strTypeName As String

    With pi.VarTypeInfo
    
        arg.isByVal = IIf(.PointerLevel = 0, True, False)
        glue tmp, vbTab & IIf(arg.isOptional, "[ ", "") & IIf(arg.isByVal, "ByVal ", "ByRef ") & pi.name
                                    
        If .VarType <> vbVariant Then
              strTypeName = TypeName(.TypedVariant)
              If .VarType And (VT_ARRAY Or VT_VECTOR) Then
                    glue tmp, " As " & Left$(strTypeName, Len(strTypeName) - 2) & "()"
                    If Not pi.Optional Then SupportsFuzzing = False
                    arg.ManSetType atarray, False, .VarType
              Else
                    arg.SetStdType strTypeName
                    glue tmp, " As " & strTypeName
              End If
        Else
              arg.ManSetType atVariant, True, VT_VARIANT
              glue tmp, " As Variant"
        End If
        
    End With
                                
End Sub

Private Sub ParseObjArg(tmp, pi As ParameterInfo, arg As CArgument, TIResolved As TypeInfo)
    Dim TIType As TypeInfo
    Dim TKind As TypeKinds
    Dim overRideAsLong As Boolean
    
    On Error Resume Next
    
    TKind = TKIND_MAX
    
    arg.ManSetType atObject, False ' True 'default guess might change
    
    With pi.VarTypeInfo
                   
            Set TIType = .TypeInfo
            
            If Not TIType Is Nothing Then
                
                Set TIResolved = TIType
                TKind = TIResolved.TypeKind
                
                Do While TKind = TKIND_ALIAS
                    TKind = TKIND_MAX
                    Set TIResolved = TIResolved.ResolvedType
                    If Err Then Err.Clear Else TKind = TIResolved.TypeKind
                Loop
                
            End If
              
            Select Case TKind
                  Case TKIND_INTERFACE, TKIND_COCLASS, TKIND_DISPATCH
                          arg.isByVal = .PointerLevel = 1
                  Case TKIND_RECORD 'Records not passed ByVal in VB
                          arg.isByVal = False
                          arg.ManSetType atOther, True  'switched from long: 6-28-22
                          overRideAsLong = False
                  Case TKIND_ENUM
                          arg.ManSetType atLong, True
                          arg.isByVal = .PointerLevel = 0
                          overRideAsLong = True
                  Case Else
                          arg.isByVal = .PointerLevel = 0
            End Select
            
            glue tmp, vbTab & IIf(arg.isOptional, "[ ", "") & IIf(arg.isByVal, "ByVal ", "ByRef ") & pi.name
            
            If TIType Is Nothing Then 'Error
                  glue tmp, "As ?"
            Else
                  If .IsExternalType Then
                          glue tmp, "As " & .TypeLibInfoExternal.name & "." & TIType.name
                          arg.ObjectType = .TypeLibInfoExternal.name & "." & TIType.name
                  Else
                          glue tmp, "As " & IIf(overRideAsLong, "Long", TIType.name)
                          arg.ObjectType = IIf(overRideAsLong, "Long", TIType.name)
                  End If
            End If
            
            If .VarType And (VT_ARRAY Or VT_VECTOR) Then
                glue tmp, "()"
                arg.ManSetType atarray, False
            End If

    End With
    
   
                                
End Sub


Private Sub BuildFuncName(tmp)
    
    With mMemberInfo
    
        If .InvokeKind = INVOKE_FUNC Or .InvokeKind = INVOKE_EVENTFUNC Then
            Select Case .ReturnType.VarType
                  Case VT_VOID, VT_HRESULT
                      If CallType = ftEvent Then
                            push tmp, "Event"
                            SupportsFuzzing = False
                      Else
                            push tmp, "Sub"
                            CallType = ftsub
                            If .Parameters.Count = 0 Then SupportsFuzzing = False
                       End If
                  Case Else
                      push tmp, "Function"
                      CallType = ftFunc
                      If .Parameters.Count = 0 Then SupportsFuzzing = False
            End Select
        Else
        
            Select Case .InvokeKind
                Case INVOKE_PROPERTYGET
                        CallType = ftPropGet
                        SupportsFuzzing = False
                        push tmp, "Property Get"
                        
                Case INVOKE_PROPERTYPUT
                        CallType = ftPropLet
                        SupportsFuzzing = True
                        push tmp, "Property Let"
                        
                Case INVOKE_PROPERTYPUTREF
                        CallType = ftPropSet
                        SupportsFuzzing = False
                        push tmp, "Property Set"
                
                Case INVOKE_UNKNOWN
                        CallType = ftPropLet 'experimental treating these as PROPPUT
                        SupportsFuzzing = True
                        'CallType = ftUnknown
                        'SupportsFuzzing = False
                        push tmp, "Invoke_Unknown "
                        
            End Select
            
        End If
        
        glue tmp, .name
        'If .Name = "GetSpecialFolder" Then Stop
    End With
    
End Sub

Private Function GetReturnType(tmp, fIsConstant As Boolean)
    Dim strTypeName As String
    Dim VarTypeCur As Integer
    Dim ConstVal
    Dim arg As CArgument
   
    With mMemberInfo
    
        If fIsConstant Then
               ConstVal = .value
               glue tmp, " = " & ConstVal
               Select Case VarType(ConstVal)
                 Case vbInteger, vbLong
                   If ConstVal < 0 Or ConstVal > 15 Then
                       glue tmp, " (&H" & Hex$(ConstVal) & ")"
                   End If
               End Select
         Else
               With .ReturnType
                 VarTypeCur = .VarType
                 If VarTypeCur = 0 Or (VarTypeCur And Not (VT_ARRAY Or VT_VECTOR)) = 0 Then
                       
                       On Error Resume Next
                       
                       'we cant fuzz prop puts who want any of these
                       If mMemberInfo.InvokeKind = INVOKE_PROPERTYPUT Then SupportsFuzzing = False
                       
                       If Not .TypeInfo Is Nothing Then
                           
                           If Err Then 'Information not available
                               glue tmp, " As ?"
                           Else
                               If .IsExternalType Then
                                       glue tmp, " As " & .TypeLibInfoExternal.name & "." & .TypeInfo.name
                               Else
                                       glue tmp, " As " & .TypeInfo.name
                               End If
                           End If
                       End If
                       
                       If VarTypeCur And (VT_ARRAY Or VT_VECTOR) Then glue tmp, "()"
                       On Error GoTo 0
                 Else
                       Select Case VarTypeCur
                         Case VT_VARIANT, VT_VOID, VT_HRESULT
                         Case Else
                           strTypeName = TypeName(.TypedVariant)
                           If VarTypeCur And (VT_ARRAY Or VT_VECTOR) Then
                               If mMemberInfo.InvokeKind = INVOKE_PROPERTYPUT Then
                                    SupportsFuzzing = False
                               End If
                               glue tmp, " As " & Left$(strTypeName, Len(strTypeName) - 2) & "()"
                           Else
                               glue tmp, " As " & strTypeName
                               
                               'If mMemberInfo.InvokeKind = INVOKE_PROPERTYPUT Then
                               If CallType = ftPropLet Then
                                    Set arg = New CArgument      'only chance to add it special case logic
                                    arg.isByVal = True
                                    arg.SetStdType strTypeName   'will handle this special case
                                    arg.isDefault = False
                                    arg.isOptional = False
                                    arg.name = mMemberInfo.name
                                    Args.Add arg
                                    ArgCount = 1
                               End If
                                
                           End If
                       End Select
                 End If
               End With
         End If
    End With

End Function

Private Function ProduceDefaultValue(DefVal As Variant, ByVal ti As TypeInfo) As String
Dim lTrackVal As Long
Dim mi As MemberInfo
Dim TKind As TypeKinds
    If ti Is Nothing Then
        Select Case VarType(DefVal)
            Case vbString
                If Len(DefVal) Then
                    ProduceDefaultValue = """" & DefVal & """"
                End If
            Case vbBoolean 'Always show for Boolean
                ProduceDefaultValue = DefVal
            Case vbDate
                If DefVal Then
                    ProduceDefaultValue = "#" & DefVal & "#"
                End If
            Case Else 'Numeric Values
                If DefVal <> 0 Then
                    ProduceDefaultValue = DefVal
                End If
        End Select
    Else
        'See if we have an enum and track the matching member
        'If the type is an object, then there will never be a
        'default value other than Nothing
        TKind = ti.TypeKind
        Do While TKind = TKIND_ALIAS
            TKind = TKIND_MAX
            On Error Resume Next
            Set ti = ti.ResolvedType
            If Err = 0 Then TKind = ti.TypeKind
            On Error GoTo 0
        Loop
        If ti.TypeKind = TKIND_ENUM Then
            lTrackVal = DefVal
            For Each mi In ti.members
                If mi.value = lTrackVal Then
                    ProduceDefaultValue = mi.name
                    Exit For
                End If
            Next
        End If
    End If
End Function

