VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CClass"
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

'Private Type strucGUID
'    Data1 As Long
'    Data2 As Integer
'    Data3 As Integer
'    Data4(7) As Byte
'End Type
'
'Private Declare Function CLSIDFromProgID Lib "ole32.dll" (ByVal lpszProgID As Long, pCLSID As strucGUID) As Long
'Private Declare Function CLSIDFromString Lib "ole32.dll" (ByVal lpszProgID As Long, pCLSID As strucGUID) As Long
'Private Declare Function ProgIDFromCLSID Lib "ole32.dll" (pCLSID As strucGUID, lpszProgID As Long) As Long
'Private Declare Function StringFromCLSID Lib "ole32.dll" (pCLSID As strucGUID, lpszProgID As Long) As Long
'Private Declare Sub CoCreateInstance Lib "ole32.dll" (ByRef rclsid As strucGUID, ByVal pUnkOuter As Long, ByVal dwClsContext As Long, ByRef riid As strucGUID, ppv As Any)
'
'Private IID_IUnknown As strucGUID
'Private Const IIDSTR_IUnknown = "{00000000-0000-0000-C000-000000000046}"
'Private Const CLSCTX_INPROC_SERVER As Long = 1&

Public name As String
Public GUID As String
Public ProgID As String
Public BeenAudited As Boolean
Public mInterfaces  As New Collection 'as CInterface
Public SafeForScripting As Boolean
Public SafeForInitilization As Boolean
Public KillBitSet As Boolean
Public DefaultInterface As String
Public ObjectSafetyReport As String

Public isRegisteredOnSystem As Boolean
Public isControl As Boolean
Public isInsertable As Boolean
Public isDesignTime As Boolean
Public isDotNet As Boolean
Public isProgrammable As Boolean

'Function CreateInstance() As IUnknown
'
'    Dim udtGUID As strucGUID
'    Dim iUnk As IUnknown
'
'    CLSIDFromString StrPtr(IIDSTR_IUnknown), IID_IUnknown
'    CLSIDFromString StrPtr(GUID), udtGUID
'    CoCreateInstance udtGUID, 0, CLSCTX_INPROC_SERVER, IID_IUnknown, iUnk
'
'    Set CreateInstance = iUnk
'
'End Function


Friend Sub LoadInterfaces(ii As Interfaces, defInterface As String)
        
        Dim iface As InterfaceInfo
        Dim mi As CInterface
        Dim cnt As Long
          
        reg.hive = HKEY_CLASSES_ROOT
'        If reg.keyExists(clsid & GUID & "\Programmable") Then Programmable = True

        isRegisteredOnSystem = reg.keyExists("\CLSID\" & GUID)
        
        Const tail = "\Implemented Categories\{7DD95801-9882-11CF-9FA9-00AA006C42C4}"
        Const tail2 = "\Implemented Categories\{7DD95802-9882-11CF-9FA9-00AA006C42C4}"
        Const tail7 = "\Implemented Categories\{40FC6ED5-2438-11CF-A3DB-080036F12502}" ' CATID_Programmable /Embedable
        Const tail3 = "\Implemented Categories\{40FC6ED4-2438-11cf-A3DB-080036F12502}"
        Const tail4 = "\Implemented Categories\{40FC6ED3-2438-11cf-A3DB-080036F12502}"
        Const tail5 = "\Implemented Categories\{F2BB56D1-DB07-11D1-AA6B-006097DB9539}"
        Const tail6 = "\Implemented Categories\{62C8FE65-4EBB-45e7-B440-6E39B2CDBF29}"
       
        
        
        'HKEY_CLASSES_ROOT\Component Categories\{F2BB56D1-DB07-11D1-AA6B-006097DB9539} DesignTimeActivatable
        'HKEY_CLASSES_ROOT\Component Categories\{62C8FE65-4EBB-45e7-B440-6E39B2CDBF29} .Net Catagory
        
        If isRegisteredOnSystem Then
            If reg.keyExists("\CLSID\" & GUID & tail) Then SafeForScripting = True
            If reg.keyExists("\CLSID\" & GUID & tail2) Then SafeForInitilization = True
            If reg.keyExists("\CLSID\" & GUID & tail3) Then isControl = True
            If reg.keyExists("\CLSID\" & GUID & tail4) Then isInsertable = True
            If reg.keyExists("\CLSID\" & GUID & tail5) Then isDesignTime = True
            If reg.keyExists("\CLSID\" & GUID & tail6) Then isDotNet = True
            If reg.keyExists("\CLSID\" & GUID & tail7) Then isProgrammable = True
        End If
        
        If KeyExistsInCollection(killbitted, GUID) Then KillBitSet = True
        
        DefaultInterface = defInterface
        
        For Each iface In ii
            AddInterface iface
        Next
        
End Sub

Sub AddInterface(iface As InterfaceInfo)
        
        On Error Resume Next
        Dim mi As New CInterface
        
        mi.DerivedString = GetDerivedInterfaces(iface, mi)
        dbg "Interface: " & iface.name & " " & mi.DerivedString & IIf(mi.isPublic, " : Public", " : Private")
            
        'If InStr(1, iface.Name, "command", vbTextCompare) > 0 Then Stop
        
        With iface
            mi.GUID = .GUID
            mi.name = .name
            If .name = DefaultInterface Then mi.isDefault = True
            mi.LoadMembers .members, Me.GUID, iface
            mInterfaces.Add mi
        End With
        
End Sub

Private Function GetDerivedInterfaces(iface As InterfaceInfo, mi As CInterface) As String
    Dim ii As InterfaceInfo
    Dim tmp As String
    Dim aas() As String
    
    'Debug.Print ii.AttributeStrings(aas)
    
    
    mi.AttributeMask = iface.AttributeMask
    
    For Each ii In iface.ImpliedInterfaces
        If ii.name = "IDispatch" Then mi.hasIDispatch = True
        tmp = tmp & " : " & ii.name
    Next
    
    GetDerivedInterfaces = tmp
    
End Function
