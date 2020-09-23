Attribute VB_Name = "Module4"
      
      
      
      
'please visit my site
'I DONT NEED TO BE VOTE
'JUST TAG IN MY SITE
'HTTP://XYREN.USA.GS


'THANKS FOR DOWNLOADING MY PROGRAM
      
      
      
      
      Public Const MMSYSERR_NOERROR = 0
      Public Const MAXPNAMELEN = 32
      Public Const MIXER_LONG_NAME_CHARS = 64
      Public Const MIXER_SHORT_NAME_CHARS = 16
      Public Const MIXER_GETLINEINFOF_COMPONENTTYPE = &H3&
      Public Const MIXER_GETCONTROLDETAILSF_VALUE = &H0&
      Public Const MIXER_GETLINECONTROLSF_ONEBYTYPE = &H2&
      Public Const MIXERLINE_COMPONENTTYPE_DST_FIRST = &H0&
      Public Const MIXERLINE_COMPONENTTYPE_SRC_FIRST = &H1000&
      
      Public Const MIXERLINE_COMPONENTTYPE_DST_SPEAKERS = _
                     (MIXERLINE_COMPONENTTYPE_DST_FIRST + 4)
                     
      Public Const MIXERLINE_COMPONENTTYPE_SRC_MICROPHONE = _
                     (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 3)
      
      Public Const MIXERLINE_COMPONENTTYPE_SRC_LINE = _
                     (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 2)
      
      Public Const MIXERCONTROL_CT_CLASS_FADER = &H50000000
      Public Const MIXERCONTROL_CT_UNITS_UNSIGNED = &H30000
      
      Public Const MIXERCONTROL_CONTROLTYPE_FADER = _
                     (MIXERCONTROL_CT_CLASS_FADER Or _
                     MIXERCONTROL_CT_UNITS_UNSIGNED)
      
      Public Const MIXERCONTROL_CONTROLTYPE_VOLUME = _
                     (MIXERCONTROL_CONTROLTYPE_FADER + 1)
      
      Declare Function mixerClose Lib "winmm.dll" _
                     (ByVal hmx As Long) As Long
         
      Declare Function mixerGetControlDetails Lib "winmm.dll" _
                     Alias "mixerGetControlDetailsA" _
                     (ByVal hmxobj As Long, _
                     pmxcd As MIXERCONTROLDETAILS, _
                     ByVal fdwDetails As Long) As Long
         
      Declare Function mixerGetDevCaps Lib "winmm.dll" _
                     Alias "mixerGetDevCapsA" _
                     (ByVal uMxId As Long, _
                     ByVal pmxcaps As MIXERCAPS, _
                     ByVal cbmxcaps As Long) As Long
         
      Declare Function mixerGetID Lib "winmm.dll" _
                     (ByVal hmxobj As Long, _
                     pumxID As Long, _
                     ByVal fdwId As Long) As Long
                     
      Declare Function mixerGetLineControls Lib "winmm.dll" _
                     Alias "mixerGetLineControlsA" _
                     (ByVal hmxobj As Long, _
                     pmxlc As MIXERLINECONTROLS, _
                     ByVal fdwControls As Long) As Long
                     
      Declare Function mixerGetLineInfo Lib "winmm.dll" _
                     Alias "mixerGetLineInfoA" _
                     (ByVal hmxobj As Long, _
                     pmxl As MIXERLINE, _
                     ByVal fdwInfo As Long) As Long
                     
      Declare Function mixerGetNumDevs Lib "winmm.dll" () As Long
      
      Declare Function mixerMessage Lib "winmm.dll" _
                     (ByVal hmx As Long, _
                     ByVal uMsg As Long, _
                     ByVal dwParam1 As Long, _
                     ByVal dwParam2 As Long) As Long
                     
      Declare Function mixerOpen Lib "winmm.dll" _
                     (phmx As Long, _
                     ByVal uMxId As Long, _
                     ByVal dwCallback As Long, _
                     ByVal dwInstance As Long, _
                     ByVal fdwOpen As Long) As Long
                     
      Declare Function mixerSetControlDetails Lib "winmm.dll" _
                     (ByVal hmxobj As Long, _
                     pmxcd As MIXERCONTROLDETAILS, _
                     ByVal fdwDetails As Long) As Long
                     
      Declare Sub CopyStructFromPtr Lib "kernel32" _
                     Alias "RtlMoveMemory" _
                     (struct As Any, _
                     ByVal ptr As Long, ByVal cb As Long)
                     
      Declare Sub CopyPtrFromStruct Lib "kernel32" _
                     Alias "RtlMoveMemory" _
                     (ByVal ptr As Long, _
                     struct As Any, _
                     ByVal cb As Long)
                     
      Declare Function GlobalAlloc Lib "kernel32" _
                     (ByVal wFlags As Long, _
                     ByVal dwBytes As Long) As Long
                     
      Declare Function GlobalLock Lib "kernel32" _
                     (ByVal hmem As Long) As Long
                     
      Declare Function GlobalFree Lib "kernel32" _
                     (ByVal hmem As Long) As Long
      
      Type MIXERCAPS
         wMid As Integer
         wPid As Integer
         vDriverVersion As Long
         szPname As String * MAXPNAMELEN
         fdwSupport As Long
         cDestinations As Long
      End Type
      
      Type MIXERCONTROL
         cbStruct As Long
         dwControlID As Long
         dwControlType As Long
         fdwControl As Long
         cMultipleItems As Long
         szShortName As String * MIXER_SHORT_NAME_CHARS
         szName As String * MIXER_LONG_NAME_CHARS
         lMinimum As Long
         lMaximum As Long
         reserved(10) As Long
         End Type
      
      Type MIXERCONTROLDETAILS
         cbStruct As Long
         dwControlID As Long
         cChannels As Long
         item As Long
         cbDetails As Long
         paDetails As Long
      End Type
      
      Type MIXERCONTROLDETAILS_UNSIGNED
         dwValue As Long
      End Type
      
      Type MIXERLINE
         cbStruct As Long
         dwDestination As Long
         dwSource As Long
         dwLineID As Long
         fdwLine As Long
         dwUser As Long
         dwComponentType As Long
         cChannels As Long
         cConnections As Long
         cControls As Long
         szShortName As String * MIXER_SHORT_NAME_CHARS
         szName As String * MIXER_LONG_NAME_CHARS
         dwType As Long
         dwDeviceID As Long
         wMid  As Integer
         wPid As Integer
         vDriverVersion As Long
         szPname As String * MAXPNAMELEN
      End Type
      
      Type MIXERLINECONTROLS
         cbStruct As Long
         dwLineID As Long
                                
         dwControl As Long
         cControls As Long
         cbmxctrl As Long
         pamxctrl As Long
      End Type
      
      Function GetVolumeControl(ByVal hmixer As Long, _
                              ByVal componentType As Long, _
                              ByVal ctrlType As Long, _
                              ByRef mxc As MIXERCONTROL) As Boolean
                              
         Dim mxlc As MIXERLINECONTROLS
         Dim mxl As MIXERLINE
         Dim hmem As Long
         Dim rc As Long
             
         mxl.cbStruct = Len(mxl)
         mxl.dwComponentType = componentType
      
         rc = mixerGetLineInfo(hmixer, mxl, MIXER_GETLINEINFOF_COMPONENTTYPE)
         
         If (MMSYSERR_NOERROR = rc) Then
             mxlc.cbStruct = Len(mxlc)
             mxlc.dwLineID = mxl.dwLineID
             mxlc.dwControl = ctrlType
             mxlc.cControls = 1
             mxlc.cbmxctrl = Len(mxc)
             
             hmem = GlobalAlloc(&H40, Len(mxc))
             mxlc.pamxctrl = GlobalLock(hmem)
             mxc.cbStruct = Len(mxc)
             
             rc = mixerGetLineControls(hmixer, _
                                       mxlc, _
                                       MIXER_GETLINECONTROLSF_ONEBYTYPE)
                  
             If (MMSYSERR_NOERROR = rc) Then
                 GetVolumeControl = True
                 
                 CopyStructFromPtr mxc, mxlc.pamxctrl, Len(mxc)
             Else
                 GetVolumeControl = False
             End If
             GlobalFree (hmem)
             Exit Function
         End If
      
         GetVolumeControl = False
      End Function
      
      Function SetVolumeControl(ByVal hmixer As Long, _
                              mxc As MIXERCONTROL, _
                              ByVal volume As Long) As Boolean
         Dim mxcd As MIXERCONTROLDETAILS
         Dim vol As MIXERCONTROLDETAILS_UNSIGNED
      
         mxcd.item = 0
         mxcd.dwControlID = mxc.dwControlID
         mxcd.cbStruct = Len(mxcd)
         mxcd.cbDetails = Len(vol)
         
         hmem = GlobalAlloc(&H40, Len(vol))
         mxcd.paDetails = GlobalLock(hmem)
         mxcd.cChannels = 1
         vol.dwValue = volume
         
         CopyPtrFromStruct mxcd.paDetails, vol, Len(vol)
         
         rc = mixerSetControlDetails(hmixer, _
                                    mxcd, _
                                    MIXER_SETCONTROLDETAILSF_VALUE)
         
         GlobalFree (hmem)
         If (MMSYSERR_NOERROR = rc) Then
             SetVolumeControl = True
         Else
             SetVolumeControl = False
         End If
      End Function


