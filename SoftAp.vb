Imports System
Imports System.Runtime.InteropServices
Imports System.ComponentModel

Public Class SoftAp

    Private hClientHandle As IntPtr
    Private dwNegotiatedVersion As UInteger

    Public Sub New()

        Dim result As UInt32 = WlanOpenHandle(2, Nothing, dwNegotiatedVersion, hClientHandle)
        If result <> 0 Then
            hClientHandle = IntPtr.Zero
            Throw New Exception(WlanapiFormatMessage(result))
        End If

    End Sub

    Public Property Enabled As Boolean
        Get
            Dim size As UInt32
            Dim settings As IntPtr
            Dim valueType As WLAN_OPCODE_VALUE_TYPE
            Dim result As UInt32 = WlanHostedNetworkQueryProperty(
                            hClientHandle,
                            WLAN_HOSTED_NETWORK_OPCODE.wlan_hosted_network_opcode_enable,
                            size,
                            settings,
                            valueType,
                            IntPtr.Zero)
            If result <> 0 Then
                Throw New Exception(WlanapiFormatMessage(result))
            End If

            Return Marshal.ReadInt32(settings)

        End Get
        Set(value As Boolean)

            Dim bEnabled(0) As Int32
            Dim reason As WLAN_HOSTED_NETWORK_REASON
            Dim pnt As IntPtr
            Dim result As UInt32

            bEnabled(0) = value
            pnt = Marshal.AllocHGlobal(4)
            Marshal.Copy(bEnabled, 0, pnt, 1)

            result = WlanHostedNetworkSetProperty(
                        hClientHandle,
                        WLAN_HOSTED_NETWORK_OPCODE.wlan_hosted_network_opcode_enable,
                        4,
                        pnt,
                        reason,
                        IntPtr.Zero)

            If result <> 0 Then
                Throw New Exception(WlanapiFormatMessage(result))
            End If
            If reason <> 0 Then
                Throw New Exception(WlanapiFormatReasonMessage(reason))
            End If

        End Set
    End Property

    Public Property SSID As String
        Get
            Dim size As UInt32
            Dim settings As IntPtr
            Dim connectionSettings As WLAN_HOSTED_NETWORK_CONNECTION_SETTINGS
            Dim valueType As WLAN_OPCODE_VALUE_TYPE
            Dim result As UInt32 = WlanHostedNetworkQueryProperty(
                            hClientHandle,
                            WLAN_HOSTED_NETWORK_OPCODE.wlan_hosted_network_opcode_connection_settings,
                            size,
                            settings,
                            valueType,
                            IntPtr.Zero)
            If result <> 0 Then
                hClientHandle = IntPtr.Zero
                Throw New Exception(WlanapiFormatMessage(result))
            End If
            connectionSettings = New WLAN_HOSTED_NETWORK_CONNECTION_SETTINGS(settings)

            Return System.Text.Encoding.UTF8.GetString(connectionSettings.hostedNetworkSSID.ucSSID).Substring(0, connectionSettings.hostedNetworkSSID.uSSIDLength)
        End Get
        Set(value As String)
            Dim connectionSettings As WLAN_HOSTED_NETWORK_CONNECTION_SETTINGS
            Dim reason As WLAN_HOSTED_NETWORK_REASON
            Dim pnt As IntPtr
            Dim result As UInt32

            connectionSettings.hostedNetworkSSID = New DOT11_SSID("aervida")
            connectionSettings.dwMaxNumberOfPeers = 101

            pnt = Marshal.AllocHGlobal(Marshal.SizeOf(connectionSettings))
            Marshal.StructureToPtr(connectionSettings, pnt, False)

            result = WlanHostedNetworkSetProperty(
                        hClientHandle,
                        WLAN_HOSTED_NETWORK_OPCODE.wlan_hosted_network_opcode_connection_settings,
                        Marshal.SizeOf(connectionSettings),
                        pnt,
                        reason,
                        IntPtr.Zero)

            If result <> 0 Then
                Throw New Exception(WlanapiFormatMessage(result))
            End If
            If reason <> 0 Then
                Throw New Exception(WlanapiFormatReasonMessage(reason))
            End If

        End Set
    End Property

    Public Property PSK As String
        Get
            Dim reason As WLAN_HOSTED_NETWORK_REASON
            Dim pnt As IntPtr
            Dim result As UInt32
            Dim len As UInt32
            Dim bIsPassPhrase As Int32
            Dim bPersistent As Int32

            result = WlanHostedNetworkQuerySecondaryKey(
                hClientHandle,
                len,
                pnt,
                bIsPassPhrase,
                bPersistent,
                reason,
                IntPtr.Zero)

            If result <> 0 Then
                Throw New Exception(WlanapiFormatMessage(result))
            End If
            If reason <> 0 Then
                Throw New Exception(WlanapiFormatReasonMessage(reason))
            End If

            PSK = Marshal.PtrToStringAnsi(pnt)
            WlanFreeMemory(pnt)

        End Get
        Set(value As String)
            Dim reason As WLAN_HOSTED_NETWORK_REASON
            Dim result As UInt32
            Dim ar() As Byte
            Dim pnt As IntPtr

            value &= ControlChars.NullChar
            ar = System.Text.Encoding.ASCII.GetBytes(value)
            pnt = Marshal.AllocHGlobal(ar.Length)
            Marshal.Copy(ar, 0, pnt, ar.Length)

            result = WlanHostedNetworkSetSecondaryKey(
                hClientHandle,
                ar.Length,
                pnt,
                True,
                True,
                reason,
                IntPtr.Zero)

            Marshal.FreeHGlobal(pnt)

            If result <> 0 Then
                Throw New Exception(WlanapiFormatMessage(result))
            End If
            If reason <> 0 Then
                Throw New Exception(WlanapiFormatReasonMessage(reason))
            End If

        End Set
    End Property

    Friend ReadOnly Property ConnectionSettings As WLAN_HOSTED_NETWORK_CONNECTION_SETTINGS
        Get
            Dim result As UInt32
            Dim size As UInt32
            Dim settings As IntPtr
            Dim valueType As WLAN_OPCODE_VALUE_TYPE

            result = WlanHostedNetworkQueryProperty(
                            hClientHandle,
                            WLAN_HOSTED_NETWORK_OPCODE.wlan_hosted_network_opcode_connection_settings,
                            size,
                            settings,
                            valueType,
                            IntPtr.Zero)
            If result <> 0 Then
                Throw New Exception(WlanapiFormatMessage(result))
            End If

            Return New WLAN_HOSTED_NETWORK_CONNECTION_SETTINGS(settings)
        End Get
    End Property

    Friend ReadOnly Property NetworkStatus As WLAN_HOSTED_NETWORK_STATUS
        Get
            Dim status As IntPtr
            Dim result As UInt32
            result = WlanHostedNetworkQueryStatus(hClientHandle, status, IntPtr.Zero)
            If result <> 0 Then
                Throw New Exception(WlanapiFormatMessage(result))
            End If
            NetworkStatus = New WLAN_HOSTED_NETWORK_STATUS(status)

        End Get
    End Property

    Public ReadOnly Property IsStarted As Boolean
        Get
            Dim status As WLAN_HOSTED_NETWORK_STATUS = NetworkStatus
            Return status.HostedNetworkState = WLAN_HOSTED_NETWORK_STATE.wlan_hosted_network_active
        End Get
    End Property

    Public Sub StartUsing()
        Dim reason As WLAN_HOSTED_NETWORK_REASON
        Dim result As UInt32
        result = WlanHostedNetworkStartUsing(hClientHandle, reason, IntPtr.Zero)
        If result <> 0 Then
            Throw New Exception(WlanapiFormatMessage(result))
        End If
        If reason <> 0 Then
            Throw New Exception(WlanapiFormatReasonMessage(reason))
        End If
    End Sub

    Public Sub StopUsing()
        Dim reason As WLAN_HOSTED_NETWORK_REASON
        Dim result As UInt32
        result = WlanHostedNetworkStopUsing(hClientHandle, reason, IntPtr.Zero)
        If result <> 0 Then
            result = WlanHostedNetworkForceStop(hClientHandle, reason, IntPtr.Zero)
            If result <> 0 Then
                Throw New Exception(WlanapiFormatMessage(result))
            End If
            If reason <> 0 Then
                Throw New Exception(WlanapiFormatReasonMessage(reason))
            End If
        End If
        If reason <> 0 Then
            Throw New Exception(WlanapiFormatReasonMessage(reason))
        End If
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
        If hClientHandle <> IntPtr.Zero Then

            Try

                If IsStarted Then
                    StopUsing()
                End If
                WlanCloseHandle(hClientHandle, IntPtr.Zero)

            Catch ex As Exception


            End Try

        End If
    End Sub
End Class
