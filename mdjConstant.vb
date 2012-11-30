Imports System
Imports System.Configuration
Imports System.Reflection
Imports Microsoft.VisualBasic
Imports System.Drawing
Imports System.Text
Imports System.Windows.Forms
Imports System.IO
Imports System.Collections.Specialized
Imports System.Windows.Forms.Cursors
Imports System.Runtime.InteropServices
Imports ESRI.ArcGIS.Controls
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Framework

Module mdjConstant

    Public TempPath As String
    Public TempGDBName As String
    Public ConnectionString As String
    Public PathGDB As String
    Public TempOutputPath As String
    Public fcSEIFA As String
    Public fcDepression As String
    Public fcDiabetes As String
    Public fcGPs As String


    ''''''''''''''''''''''''''''''''''''''''''''''''''
    ''Public Variable
   
    Public ffrmMain As frmMain
    Public ErrorExist As Boolean
    Public ErrorDescription As String
    Public dsMain As DataSet
    Public daMain As OleDb.OleDbDataAdapter
    Public Con As OleDb.OleDbConnection
    Public pApp As IApplication
    Public pMxDoc As IMxDocument
    Public pMap As IMap
    Public pActiveView As IActiveView
    Public FeatureLayerName As String
    Public IntersectPolyStr As String

    Public Sub Fun_SetConstraint()
        On Error GoTo L1

        Dim asm As Assembly = Assembly.GetAssembly(ffrmMain.[GetType]())
        Dim map As New ExeConfigurationFileMap()
        map.ExeConfigFilename = ffrmMain.[GetType]().Assembly.Location & ".config"
        Dim config As Configuration = ConfigurationManager.OpenMappedExeConfiguration(map, ConfigurationUserLevel.None)

        ''''''''''''''''''
        TempPath = "C:\tmp"
        TempGDBName = "Output.mdb"

        fcSEIFA = "CCD2006_ML_INWM"
        fcDepression = "MoodProblems_ML_NWIM"
        fcDiabetes = "Type2Diabetes_ML_NWIM"
        fcGPs = "GPs_ML_INWM"

        '---------------------------------------------------------------------------------------------------------
        PathGDB = GetFeatureClassPath(fcSEIFA)

        Exit Sub
L1:
        MsgBox(Err.Description)
    End Sub
    Public Sub Fun_ConnectDB()
        Con = New OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + PathGDB)
    End Sub


End Module
