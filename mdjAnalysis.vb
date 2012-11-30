Imports System.Collections
Imports System.Text
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Runtime.InteropServices
Imports Microsoft.VisualBasic
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.Controls
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Catalog
Imports ESRI.ArcGIS.CatalogUI
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Geoprocessor
Imports ESRI.ArcGIS.Geoprocessing
Imports ESRI.ArcGIS.AnalysisTools
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Display


Module mdjAnalysis
    Public m_hookHelper As IHookHelper = Nothing

    Public Function GetFeatureClassPath(ByVal LayerName As String) As String
        On Error GoTo err_h

        Dim pMxDoc As IMxDocument
        Dim pMap As IMap
        Dim pActiveView As IActiveView
        Dim pFeatureLayer As IFeatureLayer
        Dim player As ILayer
        Dim pLayers As IEnumLayer
        Dim Path As String = ""
        pMap = m_hookHelper.FocusMap
        pActiveView = pMap


        Dim pUID As New UID
        pUID.Value = "{E156D7E5-22AF-11D3-9F99-00C04F6BC78E}" ' CLSID for GeoFeatureLayer

        pLayers = pMap.Layers(pUID, True)
        player = pLayers.Next
        Do While (Not player Is Nothing)
            pFeatureLayer = player
            If Not pFeatureLayer Is Nothing Then
                If pFeatureLayer.FeatureClass.AliasName = LayerName Then
                    Dim pDataset As IDataset
                    pDataset = pFeatureLayer
                    Dim pWorkspace As IWorkspace
                    pWorkspace = pDataset.Workspace
                    Path = pWorkspace.PathName
                    Return Path
                    Exit Function
                End If
            End If
            player = pLayers.Next
        Loop
        MsgBox("Layer Name '" + LayerName + "' does not exist!", MsgBoxStyle.Critical, LayerName)
err_h:
        MsgBox(Err.Number & " = " & Err.Description)
        Err.Clear()
    End Function
    Public Sub Fun_CreateGeodatabase(ByVal Path As String, ByVal FileName As String)
        On Error GoTo err_h

        Dim factoryType As Type = Type.GetTypeFromProgID("esriDataSourcesGDB.AccessWorkspaceFactory")
        Dim workspaceFactory As IWorkspaceFactory = CType(Activator.CreateInstance(factoryType), IWorkspaceFactory)

        ' Create a personal geodatabase.
        Dim workspaceName As IWorkspaceName = workspaceFactory.Create(Path, FileName, Nothing, 0)

        Exit Sub
err_h:
        ErrorExist = True
        ErrorDescription = "Error CreateGeodatabase :" + Err.Description
    End Sub

    Public Sub FindFeatures(ByVal LayerName As String, ByVal Str_Query As String)
        On Error GoTo err_h

        Dim pMxDoc As IMxDocument
        Dim pMap As IMap
        Dim pActiveView As IActiveView
        Dim pFeatureLayer As IFeatureLayer
        Dim pFeatureSelection As IFeatureSelection
        Dim pQueryFilter As IQueryFilter
        Dim player As ILayer
        Dim pLayers As IEnumLayer
        '   Unselect()
        pMap = m_hookHelper.FocusMap
        pActiveView = pMap

        'QI
        pQueryFilter = New QueryFilter
        'Dim Str_Query As String
        Dim pFeatureCursor As IFeatureCursor

        Dim pUID As New UID
        pUID.Value = "{E156D7E5-22AF-11D3-9F99-00C04F6BC78E}" ' CLSID for GeoFeatureLayer
        FeatureLayerName = ""

        pLayers = pMap.Layers(pUID, True)
        player = pLayers.Next
        Do While (Not player Is Nothing)

            pFeatureLayer = player
            pFeatureSelection = pFeatureLayer
            If Not pFeatureLayer Is Nothing Then
                If pFeatureLayer.FeatureClass.AliasName = LayerName Then
                    FeatureLayerName = pFeatureLayer.Name
                    pQueryFilter.WhereClause = Str_Query
                    pFeatureSelection.SelectFeatures(pQueryFilter, 1, False)
                    If (pFeatureSelection.SelectionSet.Count > 0) Then
                        pFeatureSelection.SelectionSet.Search(Nothing, True, pFeatureCursor)
                    End If
                    pActiveView.PartialRefresh(esriViewDrawPhase.esriViewGeoSelection, Nothing, Nothing)

                    Exit Sub
                End If
            End If
            player = pLayers.Next
        Loop

        MsgBox("Layer Name '" + LayerName + "' does not exist!", MsgBoxStyle.Critical, LayerName)
err_h:
        MsgBox(Err.Number & " = " & Err.Description)
        Err.Clear()
    End Sub
    Public Function FindGetLayerName(ByVal LayerName As String) As String
        On Error GoTo err_h

        Dim pMxDoc As IMxDocument
        Dim pMap As IMap
        Dim pActiveView As IActiveView
        Dim pFeatureLayer As IFeatureLayer
        Dim pFeatureSelection As IFeatureSelection
        Dim pQueryFilter As IQueryFilter
        Dim player As ILayer
        Dim pLayers As IEnumLayer
        '   Unselect()
        pMap = m_hookHelper.FocusMap
        pActiveView = pMap

        'QI
        pQueryFilter = New QueryFilter
        Dim pFeatureCursor As IFeatureCursor

        Dim pUID As New UID
        pUID.Value = "{E156D7E5-22AF-11D3-9F99-00C04F6BC78E}" ' CLSID for GeoFeatureLayer
        FeatureLayerName = ""

        pLayers = pMap.Layers(pUID, True)
        player = pLayers.Next
        Do While (Not player Is Nothing)
            pFeatureLayer = player
            pFeatureSelection = pFeatureLayer
            If Not pFeatureLayer Is Nothing Then
                If pFeatureLayer.FeatureClass.AliasName = LayerName Then
                    FeatureLayerName = pFeatureLayer.Name
                    Return FeatureLayerName
                    Exit Function
                End If
            End If
            player = pLayers.Next
        Loop
        Return FeatureLayerName
        Exit Function
        MsgBox("Layer Name '" + LayerName + "' does not exist!", MsgBoxStyle.Critical, LayerName)
err_h:
        MsgBox(Err.Number & " = " & Err.Description)
        Err.Clear()
    End Function
    Public Sub Fun_SelectByAttribute(ByVal inFeatureClass As String, ByVal where_clause As String, ByVal selection_type As String)

        If m_hookHelper.FocusMap.LayerCount = 0 Then
            Return
        End If

        'get an instance of the geoprocessor
        Dim gp As ESRI.ArcGIS.Geoprocessor.Geoprocessor = New ESRI.ArcGIS.Geoprocessor.Geoprocessor()
        gp.OverwriteOutput = True
        gp.AddOutputsToMap = False

        Dim SelectLayerByAttribute As ESRI.ArcGIS.DataManagementTools.SelectLayerByAttribute = New ESRI.ArcGIS.DataManagementTools.SelectLayerByAttribute(inFeatureClass)
        SelectLayerByAttribute.selection_type = selection_type
        If where_clause <> "" Then
            SelectLayerByAttribute.where_clause = where_clause
        End If

        Dim results As IGeoProcessorResult
        Try
            results = CType(gp.Execute(SelectLayerByAttribute, Nothing), IGeoProcessorResult)
        Catch ex As Exception
            GoTo err_h
        End Try

        Exit Sub
err_h:
        ErrorExist = True
        ErrorDescription = "Error SelectLayerByAttribute :" + (gp.GetMessage(2))

    End Sub
    Public Sub Fun_CopyFeatures(ByVal inFeatureClass As String, ByVal PathOutput As String, ByVal AddOutputsToMap As Boolean)

        If m_hookHelper.FocusMap.LayerCount = 0 Then
            Return
        End If

        'get an instance of the geoprocessor
        Dim gp As ESRI.ArcGIS.Geoprocessor.Geoprocessor = New ESRI.ArcGIS.Geoprocessor.Geoprocessor()
        gp.OverwriteOutput = True
        gp.AddOutputsToMap = AddOutputsToMap
        gp.SetEnvironmentValue("workspace", PathOutput)
        gp.SetEnvironmentValue("scratchWorkspace", PathOutput)
        'create a new instance of a copy tool
        Dim Copy As ESRI.ArcGIS.DataManagementTools.CopyFeatures = New ESRI.ArcGIS.DataManagementTools.CopyFeatures(inFeatureClass, PathOutput)
        Dim results As IGeoProcessorResult
        Try
            results = CType(gp.Execute(Copy, Nothing), IGeoProcessorResult)
        Catch ex As Exception
            GoTo err_h
        End Try

        Exit Sub
err_h:
        ErrorExist = True
        ErrorDescription = "Error CopyFeatures :" + (gp.GetMessage(2))

    End Sub
  
    Public Sub Unselect()
        Dim pSelectTool As ICommandItem
        Dim pCommandBars As ICommandBars
        Dim u As New UID
        u.Value = "{37C833F3-DBFD-11D1-AA7E-00C04FA37860}"
        pCommandBars = pApp.Document.CommandBars
        pSelectTool = pCommandBars.Find(u)
        pSelectTool.Execute()
    End Sub

    Public Sub Fun_IntersectLayers(ByVal InputLayers As String, ByVal OutputPath As String, ByVal JoinAttribute As String, ByVal AddOutputsToMap As Boolean)


        Dim txtMessages As String = ""
        If m_hookHelper.FocusMap.LayerCount = 0 Then
            Return
        End If

        'add message to the messages box
        txtMessages &= "Intersect layers: " & InputLayers & Constants.vbCrLf
        txtMessages += Constants.vbCrLf & "Get the geoprocessor. This might take a few seconds..." & Constants.vbCrLf

        'get an instance of the geoprocessor
        Dim gp As ESRI.ArcGIS.Geoprocessor.Geoprocessor = New ESRI.ArcGIS.Geoprocessor.Geoprocessor()
        gp.OverwriteOutput = True
        gp.AddOutputsToMap = AddOutputsToMap

        'create a new instance of a Union tool
        Dim Intersect As ESRI.ArcGIS.AnalysisTools.Intersect = New ESRI.ArcGIS.AnalysisTools.Intersect(InputLayers, OutputPath)
        If JoinAttribute <> "" Then
            Intersect.join_attributes = JoinAttribute
        End If

        Dim results As IGeoProcessorResult
        Try
            results = CType(gp.Execute(Intersect, Nothing), IGeoProcessorResult)
            Dim comob As Object = results
            System.Runtime.InteropServices.Marshal.ReleaseComObject(comob)
        Catch ex As Exception
            GoTo err_h
        End Try

        Exit Sub

err_h:
        ErrorExist = True
        ErrorDescription = "Error IntersectLayers :" + gp.GetMessage(gp.MessageCount - 1)
    End Sub
    Public Sub Fun_EraseLayers(ByVal InputLayers As String, ByVal OutputPath As String, ByVal EraseLayer As String, ByVal AddOutputsToMap As Boolean)


        Dim txtMessages As String = ""

        If m_hookHelper.FocusMap.LayerCount = 0 Then
            Return
        End If

        'add message to the messages box
        txtMessages &= "Erase layers: " & InputLayers & Constants.vbCrLf
        txtMessages += Constants.vbCrLf & "Get the geoprocessor. This might take a few seconds..." & Constants.vbCrLf

        'get an instance of the geoprocessor
        Dim gp As ESRI.ArcGIS.Geoprocessor.Geoprocessor = New ESRI.ArcGIS.Geoprocessor.Geoprocessor()
        gp.OverwriteOutput = True
        gp.AddOutputsToMap = AddOutputsToMap

        'create a new instance of a Union tool
        Dim Erase1 As ESRI.ArcGIS.AnalysisTools.Erase = New ESRI.ArcGIS.AnalysisTools.Erase(InputLayers, EraseLayer, OutputPath)

        Dim results As IGeoProcessorResult
        Try
            results = CType(gp.Execute(Erase1, Nothing), IGeoProcessorResult)
            Dim comob As Object = results
            System.Runtime.InteropServices.Marshal.ReleaseComObject(comob)
        Catch ex As Exception
            GoTo err_h
        End Try


        Exit Sub

err_h:
        ErrorExist = True
        ErrorDescription = "Error EraseLayers :" + gp.GetMessage(gp.MessageCount - 1)
    End Sub

    Public Sub Fun_SelectByLocation(ByVal InputLayers As String, ByVal SelectFeature As String, ByVal Overlap_type As String, ByVal selection_type As String, Optional ByVal Distance As Integer = 0)
        'make sure that all parameters are okay
        Dim txtMessages As String = ""
        If m_hookHelper.FocusMap.LayerCount = 0 Then
            Return
        End If


        'add message to the messages box
        txtMessages &= "Fun_SelectByLocation: " & InputLayers & Constants.vbCrLf
        txtMessages += Constants.vbCrLf & "Get the geoprocessor. This might take a few seconds..." & Constants.vbCrLf

        'get an instance of the geoprocessor
        Dim gp As ESRI.ArcGIS.Geoprocessor.Geoprocessor = New ESRI.ArcGIS.Geoprocessor.Geoprocessor()

        gp.OverwriteOutput = True
        gp.AddOutputsToMap = False
        Dim results As IGeoProcessorResult
        Try
            'create a new instance of a Union tool
            Dim SelectByLocation As ESRI.ArcGIS.DataManagementTools.SelectLayerByLocation = New ESRI.ArcGIS.DataManagementTools.SelectLayerByLocation(InputLayers)
            SelectByLocation.select_features = SelectFeature
            SelectByLocation.overlap_type = Overlap_type
            SelectByLocation.selection_type = selection_type
            If Distance > 0 Then
                SelectByLocation.search_distance = Distance
            End If
            results = CType(gp.Execute(SelectByLocation, Nothing), IGeoProcessorResult)
        Catch ex As Exception
            GoTo err_h
        End Try


        Exit Sub
err_h:
        ErrorExist = True
        ErrorDescription = "Error Fun_SelectByLocation :" + gp.GetMessage(gp.MessageCount - 1)
    End Sub

    Public Sub Fun_DeleteLayerfromMap(ByVal LayerName As String)
        On Error GoTo err_h

        Dim pMap As IMap
        Dim pFeatureLayer As IFeatureLayer
        Dim player As ILayer
        Dim pLayers As IEnumLayer
        Dim pUID As New UID

        pMap = m_hookHelper.FocusMap
        pUID.Value = "{E156D7E5-22AF-11D3-9F99-00C04F6BC78E}" ' CLSID for GeoFeatureLayer
        'Unselect()
        pLayers = pMap.Layers(pUID, True)
        player = pLayers.Next

        Do While (Not player Is Nothing)
            pFeatureLayer = player
            If Not pFeatureLayer Is Nothing Then
                If pFeatureLayer.Name = LayerName Then
                    pMap.DeleteLayer(player)
                    Exit Sub
                End If
            End If
            player = pLayers.Next
        Loop
        Exit Sub
err_h:
        ErrorExist = True
        ErrorDescription = "Error Delete Layer From Map "

    End Sub

    Public Sub Display_Symbology_Layer(ByVal LayerName As String, ByVal SymLayerPath As String)
        On Error GoTo err_h

        Dim pGxLayer As IGxLayer
        Dim pGxFile As IGxFile
        Dim pSymLayer As ILayer
        Dim pLyr As IGeoFeatureLayer
        Dim pGeoSymLyr As IGeoFeatureLayer
        Dim pMap As IMap
        Dim pFLayer As IFeatureLayer
        Dim player As ILayer
        Dim pLayers As IEnumLayer
        Dim StrSum As String = ""
        Dim pUID As New UID

        pMxDoc = pApp.Document
        pMap = m_hookHelper.FocusMap
        pActiveView = pMap

        pUID.Value = "{E156D7E5-22AF-11D3-9F99-00C04F6BC78E}" ' CLSID for GeoFeatureLayer

        'Get the symbology layer
        pGxLayer = New ESRI.ArcGIS.Catalog.GxLayer

        pGxFile = pGxLayer
        pGxFile.Path = SymLayerPath


        ' Unselect()
        pLayers = pMap.Layers(pUID, True)
        player = pLayers.Next

        Do While (Not player Is Nothing)
            pFLayer = player
            If Not pFLayer Is Nothing Then
                If pFLayer.Name = LayerName Then

                    pSymLayer = pGxLayer.Layer
                    pLyr = pFLayer
                    pGeoSymLyr = pSymLayer

                    pLyr.Renderer = pGeoSymLyr.Renderer

                    Exit Sub

                End If

            End If
            player = pLayers.Next
        Loop

        ''''''''''''''''''
        Exit Sub

err_h:
        MsgBox("Layer Name '" + LayerName + "' does not exist!", MsgBoxStyle.Critical, LayerName)
        MsgBox(Err.Number & " = " & Err.Description)
        Err.Clear()

    End Sub
    Sub TurnOFFONLayer(ByVal LayerName As String)

        Dim pMap As IMap
        Dim player As ILayer
        Dim pLayers As IEnumLayer
        Dim StrSum As String = ""
        Dim pUID As New UID
        pMxDoc = pApp.Document
        pMap = m_hookHelper.FocusMap
        pActiveView = pMap

        pLayers = pMap.Layers
        player = pLayers.Next

        Do While (Not player Is Nothing)
            If player.Name = LayerName Then
                player.Visible = Not (player.Visible)
            End If
            player = pLayers.Next
        Loop


    End Sub
End Module
