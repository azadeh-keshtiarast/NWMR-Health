Imports System
Imports System.Drawing
Imports System.Text
Imports System.Windows.Forms
Imports System.Runtime.InteropServices
Imports ESRI.ArcGIS.Controls
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Geoprocessor
Imports ESRI.ArcGIS.AnalysisTools
Imports Microsoft.VisualBasic
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.Display

Public Class Health
    Inherits ESRI.ArcGIS.Desktop.AddIns.Button
    Private m_hookHelper As IHookHelper = Nothing

    Public Sub New()
        If Hook Is Nothing Then
            Return
        End If

        If m_hookHelper Is Nothing Then
            m_hookHelper = New HookHelperClass()
        End If

        m_hookHelper.Hook = Hook
        pApp = My.ArcMap.Application
    End Sub

    Protected Overrides Sub OnClick()
        '
        '  TODO: Sample code showing how to access button host
        '
        My.ArcMap.Application.CurrentTool = Nothing
        pApp = My.ArcMap.Application

        If Nothing Is m_hookHelper Then
            Return
        End If
        If m_hookHelper.FocusMap.LayerCount > 0 Then
            'Dim bufferDlg As BufferDlg = New BufferDlg(m_hookHelper)
            'bufferDlg.Show()
            If ffrmMain Is Nothing Then
                ffrmMain = New frmMain(m_hookHelper)
                ffrmMain.Show()
            ElseIf ffrmMain.IsDisposed = True Then
                ffrmMain = New frmMain(m_hookHelper)
                ffrmMain.Show()
            Else
                ffrmMain.Focus()
            End If
        End If
    End Sub

    Protected Overrides Sub OnUpdate()
        Enabled = My.ArcMap.Application IsNot Nothing
    End Sub
End Class
