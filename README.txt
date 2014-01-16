Example:

    Private mAP As SoftAp = New SoftAp

...
...

        If mAP.Enabled = False Then
            Try
                mAP.Enabled = True
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Enable SoftAP Error!", MessageBoxButton.OK, MessageBoxImage.Error)
            End Try
        End If

        Try
            If mAP.IsStarted Then
                mAP.StopUsing()
                grdSoftAp.Background = mImgStopped
            Else
                mAP.SSID = mSSID
                mAP.PSK = mPSK
                mAP.StartUsing()
                grdSoftAp.Background = mImgStarted
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Enable SoftAp Error!", MessageBoxButton.OK, MessageBoxImage.Error)
            Return
        End Try

