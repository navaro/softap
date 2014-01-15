Imports System
Imports System.Runtime.InteropServices
Imports System.ComponentModel

Module wlanapi

    Public Enum WLAN_HOSTED_NETWORK_REASON
        wlan_hosted_network_reason_success
        wlan_hosted_network_reason_unspecified
        wlan_hosted_network_reason_bad_parameters
        wlan_hosted_network_reason_service_shutting_down
        wlan_hosted_network_reason_insufficient_resources
        wlan_hosted_network_reason_elevation_required
        wlan_hosted_network_reason_read_only
        wlan_hosted_network_reason_persistence_failed
        wlan_hosted_network_reason_crypt_error
        wlan_hosted_network_reason_impersonation
        wlan_hosted_network_reason_stop_before_start
        wlan_hosted_network_reason_interface_available
        wlan_hosted_network_reason_interface_unavailable
        wlan_hosted_network_reason_miniport_stopped
        wlan_hosted_network_reason_miniport_started
        wlan_hosted_network_reason_incompatible_connection_started
        wlan_hosted_network_reason_incompatible_connection_stopped
        wlan_hosted_network_reason_user_action
        wlan_hosted_network_reason_client_abort
        wlan_hosted_network_reason_ap_start_failed
        wlan_hosted_network_reason_peer_arrived
        wlan_hosted_network_reason_peer_departed
        wlan_hosted_network_reason_peer_timeout
        wlan_hosted_network_reason_gp_denied
        wlan_hosted_network_reason_service_unavailable
        wlan_hosted_network_reason_device_change
        wlan_hosted_network_reason_properties_change
        wlan_hosted_network_reason_virtual_station_blocking_use
        wlan_hosted_network_reason_service_available_on_virtual_station
    End Enum

    Public Enum WLAN_HOSTED_NETWORK_STATE
        wlan_hosted_network_unavailable
        wlan_hosted_network_idle
        wlan_hosted_network_active
    End Enum

    Public Enum DOT11_PHY_TYPE
        dot11_phy_type_unknown = 0
        dot11_phy_type_any = 0
        dot11_phy_type_fhss = 1
        dot11_phy_type_dsss = 2
        dot11_phy_type_irbaseband = 3
        dot11_phy_type_ofdm = 4
        dot11_phy_type_hrdsss = 5
        dot11_phy_type_erp = 6
        dot11_phy_type_ht = 7
        dot11_phy_type_IHV_start = &H80000000
        dot11_phy_type_IHV_end = &HFFFFFFFF
    End Enum

    Public Enum WLAN_HOSTED_NETWORK_PEER_AUTH_STATE
        wlan_hosted_network_peer_state_invalid
        wlan_hosted_network_peer_state_authenticated
    End Enum

    Public Enum WLAN_HOSTED_NETWORK_OPCODE
        wlan_hosted_network_opcode_connection_settings
        wlan_hosted_network_opcode_security_settings
        wlan_hosted_network_opcode_station_profile
        wlan_hosted_network_opcode_enable
    End Enum

    Public Enum WLAN_OPCODE_VALUE_TYPE
        wlan_opcode_value_type_query_only = 0
        wlan_opcode_value_type_set_by_group_policy = 1
        wlan_opcode_value_type_set_by_user = 2
        wlan_opcode_value_type_invalid = 3
    End Enum

    <StructLayout(LayoutKind.Sequential)> _
    Public Structure CGUID
        Public a As UInt32
        Public b As UInt16
        Public c As UInt16
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=8)> _
        Public d() As Byte
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure DOT11_MAC_ADDRESS
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=6)> _
        Public mac() As Byte
    End Structure

    <StructLayout(LayoutKind.Sequential)> _
    Public Structure WLAN_HOSTED_NETWORK_PEER_STATE
        Public PeerMacAddress As DOT11_MAC_ADDRESS
        Public PeerAuthState As WLAN_HOSTED_NETWORK_PEER_AUTH_STATE
    End Structure

    Const DOT11_SSID_MAX_LENGTH = 32
    <StructLayout(LayoutKind.Sequential)> _
    Public Structure DOT11_SSID
        Public uSSIDLength As UInt32
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=DOT11_SSID_MAX_LENGTH)> _
        Public ucSSID() As Byte

        Public Sub New(ssid As String)
            If ssid.Length > DOT11_SSID_MAX_LENGTH Then
                ssid = ssid.Substring(0, DOT11_SSID_MAX_LENGTH)
            End If
            ucSSID = System.Text.Encoding.ASCII.GetBytes(ssid)
            uSSIDLength = ucSSID.Length
            ReDim Preserve ucSSID(DOT11_SSID_MAX_LENGTH)
        End Sub
    End Structure

    <StructLayout(LayoutKind.Sequential)> _
    Public Structure WLAN_HOSTED_NETWORK_CONNECTION_SETTINGS
        Public hostedNetworkSSID As DOT11_SSID
        Public dwMaxNumberOfPeers As UInt32

        Public Sub New(p As IntPtr)
            hostedNetworkSSID = Marshal.PtrToStructure(p, GetType(DOT11_SSID))
            dwMaxNumberOfPeers = Marshal.ReadInt32(p + Marshal.SizeOf(hostedNetworkSSID))
            WlanFreeMemory(p)
        End Sub
    End Structure


    <StructLayout(LayoutKind.Sequential)>
    Public Structure WLAN_HOSTED_NETWORK_STATUS
        Public HostedNetworkState As WLAN_HOSTED_NETWORK_STATE
        Public IPDeviceID As CGUID
        Public wlanHostedNetworkBSSID As DOT11_MAC_ADDRESS
        Public dot11PhyType As DOT11_PHY_TYPE
        Public ulChannelFrequency As ULong
        Public dwNumberOfPeers As UInt32
        '<MarshalAs(UnmanagedType.Struct)> _
        Public PeerList() As WLAN_HOSTED_NETWORK_PEER_STATE

        Public Sub New(pList As IntPtr)

            Dim i As Integer

            HostedNetworkState = Marshal.ReadInt32(pList, 0)
            If HostedNetworkState <> WLAN_HOSTED_NETWORK_STATE.wlan_hosted_network_unavailable Then
                IPDeviceID = Marshal.PtrToStructure(pList + 4, GetType(CGUID))
                wlanHostedNetworkBSSID = Marshal.PtrToStructure(pList + 20, GetType(DOT11_MAC_ADDRESS))
                dot11PhyType = Marshal.ReadInt32(pList, 28)
                ulChannelFrequency = Marshal.ReadInt32(pList, 32)
                dwNumberOfPeers = Marshal.ReadInt32(pList, 36)

                ReDim PeerList(dwNumberOfPeers - 1)
                For i = 0 To dwNumberOfPeers - 1
                    Dim peer As WLAN_HOSTED_NETWORK_PEER_STATE = New WLAN_HOSTED_NETWORK_PEER_STATE
                    peer = Marshal.PtrToStructure(pList + 40 + Marshal.SizeOf(peer) * i, GetType(WLAN_HOSTED_NETWORK_PEER_STATE))
                    PeerList(i) = peer
                Next
            End If

            WlanFreeMemory(pList)

        End Sub
    End Structure


    <DllImport("Wlanapi", EntryPoint:="WlanOpenHandle")> _
    Public Function WlanOpenHandle(ByVal dwClientVersion As UInteger, ByVal pReserved As IntPtr, <Out> ByRef pdwNegotiatedVersion As UInteger, ByRef phClientHandle As IntPtr) As UInteger
    End Function

    <DllImport("Wlanapi", EntryPoint:="WlanCloseHandle")> _
    Public Function WlanCloseHandle(<[In]> ByVal hClientHandle As IntPtr, ByVal pReserved As IntPtr) As UInteger
    End Function

    <DllImport("Wlanapi", EntryPoint:="WlanHostedNetworkStartUsing")> _
    Public Function WlanHostedNetworkStartUsing(<[In]> ByVal hClientHandle As IntPtr, <Out> ByRef pReason As WLAN_HOSTED_NETWORK_REASON, ByVal pReserved As IntPtr) As UInteger
    End Function

    <DllImport("Wlanapi", EntryPoint:="WlanHostedNetworkStopUsing")> _
    Public Function WlanHostedNetworkStopUsing(<[In]> ByVal hClientHandle As IntPtr, <Out> ByRef pReason As WLAN_HOSTED_NETWORK_REASON, ByVal pReserved As IntPtr) As UInteger
    End Function

    <DllImport("Wlanapi", EntryPoint:="WlanHostedNetworkForceStop")> _
    Public Function WlanHostedNetworkForceStop(<[In]> ByVal hClientHandle As IntPtr, <Out> ByRef pReason As WLAN_HOSTED_NETWORK_REASON, ByVal pReserved As IntPtr) As UInteger
    End Function

    <DllImport("Wlanapi", EntryPoint:="WlanHostedNetworkQueryStatus")> _
    Public Function WlanHostedNetworkQueryStatus(<[In]> ByVal hClientHandle As IntPtr, <Out> ByRef pStatus As IntPtr, ByVal pReserved As IntPtr) As UInteger
    End Function

    <DllImport("Wlanapi", EntryPoint:="WlanHostedNetworkQueryProperty")> _
    Public Function WlanHostedNetworkQueryProperty(<[In]> ByVal hClientHandle As IntPtr, <[In]> ByVal OpCode As WLAN_HOSTED_NETWORK_OPCODE, <Out> ByRef pdwDataSize As UInt32, <Out> ByRef pvData As IntPtr, <Out> ByRef pWlanOpcodeValueType As WLAN_OPCODE_VALUE_TYPE, ByVal pReserved As IntPtr) As UInteger
    End Function

    <DllImport("Wlanapi", EntryPoint:="WlanHostedNetworkSetProperty")> _
    Public Function WlanHostedNetworkSetProperty(<[In]> ByVal hClientHandle As IntPtr, <[In]> ByVal OpCode As WLAN_HOSTED_NETWORK_OPCODE, <[In]> ByVal dwDataSize As UInt32, <[In]> ByVal pvData As IntPtr, <Out> ByRef pReason As WLAN_HOSTED_NETWORK_REASON, ByVal pReserved As IntPtr) As UInteger
    End Function

    <DllImport("Wlanapi", EntryPoint:="WlanHostedNetworkQuerySecondaryKey")> _
    Public Function WlanHostedNetworkQuerySecondaryKey(<[In]> ByVal hClientHandle As IntPtr, <Out> ByRef pdwKeyLength As UInt32, <Out> ByRef ppucKeyData As IntPtr, <Out> ByRef pbIsPassPhrase As Int32, <Out> ByRef pbPersistent As Int32, <Out> ByRef pReason As WLAN_HOSTED_NETWORK_REASON, ByVal pReserved As IntPtr) As UInteger
    End Function

    <DllImport("Wlanapi", EntryPoint:="WlanHostedNetworkSetSecondaryKey")> _
    Public Function WlanHostedNetworkSetSecondaryKey(<[In]> ByVal hClientHandle As IntPtr, <[In]> ByVal pdwKeyLength As UInt32, <[In]> ByVal pucKeyData As IntPtr, <[In]> ByVal bIsPassPhrase As Int32, <[In]> ByVal bPersistent As Int32, <Out> ByRef pReason As WLAN_HOSTED_NETWORK_REASON, ByVal pReserved As IntPtr) As UInteger
    End Function


    <DllImport("Wlanapi", EntryPoint:="WlanFreeMemory")> _
    Public Sub WlanFreeMemory(<[In]> ByVal ptr As IntPtr)
    End Sub




    <DllImport("Kernel32.dll", EntryPoint:="FormatMessageW", SetLastError:=True, CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.StdCall)> _
    Public Function FormatMessage(ByVal dwFlags As Integer, ByRef lpSource As IntPtr, ByVal dwMessageId As Integer, ByVal dwLanguageId As Integer, ByRef lpBuffer As [String], ByVal nSize As Integer, ByRef Arguments As IntPtr) As Integer
    End Function

    Public Function WlanapiFormatMessage(ByVal dwMessageId As Integer) As String

        Dim myEx As New Win32Exception(dwMessageId)
        Return myEx.Message

    End Function

    Public Function WlanapiFormatReasonMessage(ByVal reason As WLAN_HOSTED_NETWORK_REASON) As String

        Select Case reason


            Case WLAN_HOSTED_NETWORK_REASON.wlan_hosted_network_reason_success
                Return "Success"
            Case WLAN_HOSTED_NETWORK_REASON.wlan_hosted_network_reason_unspecified
                Return "Unspecified"
            Case WLAN_HOSTED_NETWORK_REASON.wlan_hosted_network_reason_bad_parameters
                Return "Bad Parameters"
            Case WLAN_HOSTED_NETWORK_REASON.wlan_hosted_network_reason_service_shutting_down
                Return "Service Shutting Down"
            Case WLAN_HOSTED_NETWORK_REASON.wlan_hosted_network_reason_insufficient_resources
                Return "Insufficient Resources"
            Case WLAN_HOSTED_NETWORK_REASON.wlan_hosted_network_reason_elevation_required
                Return "Elevation Required"
            Case WLAN_HOSTED_NETWORK_REASON.wlan_hosted_network_reason_read_only
                Return "Read Only"
            Case WLAN_HOSTED_NETWORK_REASON.wlan_hosted_network_reason_persistence_failed
                Return "Persistence Failed"
            Case WLAN_HOSTED_NETWORK_REASON.wlan_hosted_network_reason_crypt_error
                Return "Crypt Error"
            Case WLAN_HOSTED_NETWORK_REASON.wlan_hosted_network_reason_impersonation
                Return "Impersonation"
            Case WLAN_HOSTED_NETWORK_REASON.wlan_hosted_network_reason_stop_before_start
                Return "Stop Before Start"
            Case WLAN_HOSTED_NETWORK_REASON.wlan_hosted_network_reason_interface_available
                Return "Interface Available"
            Case WLAN_HOSTED_NETWORK_REASON.wlan_hosted_network_reason_interface_unavailable
                Return "Interface Unavailable"
            Case WLAN_HOSTED_NETWORK_REASON.wlan_hosted_network_reason_miniport_stopped
                Return "Miniport Stopped"
            Case WLAN_HOSTED_NETWORK_REASON.wlan_hosted_network_reason_miniport_started
                Return "Miniport Started"
            Case WLAN_HOSTED_NETWORK_REASON.wlan_hosted_network_reason_incompatible_connection_started
                Return "Connection Started"
            Case WLAN_HOSTED_NETWORK_REASON.wlan_hosted_network_reason_incompatible_connection_stopped
                Return "Connection Stopped"
            Case WLAN_HOSTED_NETWORK_REASON.wlan_hosted_network_reason_user_action
                Return "User Action"
            Case WLAN_HOSTED_NETWORK_REASON.wlan_hosted_network_reason_client_abort
                Return "Client Abort"
            Case WLAN_HOSTED_NETWORK_REASON.wlan_hosted_network_reason_ap_start_failed
                Return "Start Failed"
            Case WLAN_HOSTED_NETWORK_REASON.wlan_hosted_network_reason_peer_arrived
                Return "Peer Arrived"
            Case WLAN_HOSTED_NETWORK_REASON.wlan_hosted_network_reason_peer_departed
                Return "Peer Departed"
            Case WLAN_HOSTED_NETWORK_REASON.wlan_hosted_network_reason_peer_timeout
                Return "Peer Timeout"
            Case WLAN_HOSTED_NETWORK_REASON.wlan_hosted_network_reason_gp_denied
                Return "GP Denied"
            Case WLAN_HOSTED_NETWORK_REASON.wlan_hosted_network_reason_service_unavailable
                Return "Service Unavailable"
            Case WLAN_HOSTED_NETWORK_REASON.wlan_hosted_network_reason_device_change
                Return "Device Change"
            Case WLAN_HOSTED_NETWORK_REASON.wlan_hosted_network_reason_properties_change
                Return "Properties Change"
            Case WLAN_HOSTED_NETWORK_REASON.wlan_hosted_network_reason_virtual_station_blocking_use
                Return "Blocking Use"
            Case WLAN_HOSTED_NETWORK_REASON.wlan_hosted_network_reason_service_available_on_virtual_station
                Return "Service Available on Virtual Station"
        End Select

        Return "Unknnown"

    End Function



End Module
