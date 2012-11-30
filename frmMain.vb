Imports Microsoft.VisualBasic
Imports System
Imports System.Drawing
Imports System.Text
Imports System.Windows.Forms
Imports System.Runtime.InteropServices
Imports ESRI.ArcGIS.Controls
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Geoprocessor
Imports ESRI.ArcGIS.Geoprocessing
Imports ESRI.ArcGIS.AnalysisTools
Imports System.IO
Imports System.Configuration
Imports System.Collections.Specialized
Imports System.Reflection

Public Class frmMain
    Private OutputFileName As String
    Private chkSetSymbol As Boolean
    Private ProgressVal1 As Integer
    Private ProgressVal2 As Integer

    Public Sub New(ByVal hookHelper As IHookHelper)
        InitializeComponent()
        m_hookHelper = hookHelper
    End Sub



    Private Sub chkSEIFA_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkSEIFA.CheckedChanged
        If chkSEIFA.Checked = True Then
            cmbSEIFA.Enabled = True
            trkb_SEIFA.Enabled = True
            lblSEIFA.Enabled = True
        Else
            cmbSEIFA.Enabled = False
            trkb_SEIFA.Enabled = False
            lblSEIFA.Enabled = False

            trkb_SEIFA.Value = trkb_SEIFA.Minimum
            cmbSEIFA.Text = "="

        End If
    End Sub

    Private Sub chkDiabets_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkDiabets.CheckedChanged
        If chkDiabets.Checked = True Then
            cmbDiabets.Enabled = True
            trkb_Diabets.Enabled = True
            lblDiabets.Enabled = True
        Else
            cmbDiabets.Enabled = False
            trkb_Diabets.Enabled = False
            lblDiabets.Enabled = False

            trkb_Diabets.Value = trkb_Diabets.Minimum
            cmbDiabets.Text = "="
        End If

    End Sub

    Private Sub chkDepression_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkDepression.CheckedChanged
        If chkDepression.Checked = True Then
            cmbDepression.Enabled = True
            trkb_Depression.Enabled = True
            lblDepression.Enabled = True
        Else
            cmbDepression.Enabled = False
            trkb_Depression.Enabled = False
            lblDepression.Enabled = False

            trkb_Depression.Value = trkb_Depression.Minimum
            cmbDepression.Text = "="
        End If

    End Sub

    Private Sub trkbSEIFA_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles trkb_SEIFA.Scroll, trkb_SEIFA.ValueChanged
        Dim trkb_Name As String = DirectCast(sender, System.Windows.Forms.TrackBar).Name
        Dim lbl1 As Control
        Dim lblName As String = "lbl" + Microsoft.VisualBasic.Right(trkb_Name, trkb_Name.Length - InStr(trkb_Name, "_"))
        If Me.Controls.Find(lblName, True).Length > 0 Then
            lbl1 = Me.Controls.Find(lblName, True)(0)
            lbl1.Text = sender.Value
        End If
    End Sub
    Private Sub trkb_Diabets_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles trkb_Diabets.Scroll, trkb_Depression.Scroll, trkb_Depression.ValueChanged, trkb_Diabets.ValueChanged
        Dim trkb_Name As String = DirectCast(sender, System.Windows.Forms.TrackBar).Name
        Dim lbl1 As Control
        Dim lblName As String = "lbl" + Microsoft.VisualBasic.Right(trkb_Name, trkb_Name.Length - InStr(trkb_Name, "_"))
        If Me.Controls.Find(lblName, True).Length > 0 Then
            lbl1 = Me.Controls.Find(lblName, True)(0)
            lbl1.Text = sender.Value / 2
        End If
    End Sub
    Private Sub trkb_GP_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles trkb_GP.Scroll
        Dim trkb_Name As String = DirectCast(sender, System.Windows.Forms.TrackBar).Name
        Dim lbl1 As Control
        Dim lblName As String = "lbl" + Microsoft.VisualBasic.Right(trkb_Name, trkb_Name.Length - InStr(trkb_Name, "_"))
        If Me.Controls.Find(lblName, True).Length > 0 Then
            lbl1 = Me.Controls.Find(lblName, True)(0)
            lbl1.Text = Int(sender.Value / 100) * 100
        End If
    End Sub

    Private Sub frmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        On Error GoTo L1

        ErrorExist = False
        Fun_SetConstraint()
        TempOutputPath = TempPath + "\" + TempGDBName
       

        trkb_SEIFA.Value = trkb_SEIFA.Maximum
        trkb_Depression.Value = trkb_Depression.Maximum
        trkb_Diabets.Value = trkb_Diabets.Maximum

        trkb_SEIFA.Value = trkb_SEIFA.Minimum
        trkb_Depression.Value = trkb_Depression.Minimum
        trkb_Diabets.Value = trkb_Diabets.Minimum
        Exit Sub
L1:
        MsgBox(Err.Description)
    End Sub

    Private Sub btnBrowseAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBrowseAdd.Click
        Dim pFCFilter As New ESRI.ArcGIS.Catalog.GxFilterFeatureClasses
        Dim pGxDialog As New ESRI.ArcGIS.CatalogUI.GxDialog

        pGxDialog.Title = "Save feature class"
        pGxDialog.Name = ""
        pGxDialog.ObjectFilter = pFCFilter

        If Not pGxDialog.DoModalSave(My.ArcMap.Application.hWnd) Then Exit Sub
        If InStr(pGxDialog.FinalLocation.FullName, ".mdb") > 0 Then
            txtOutput.Text = pGxDialog.FinalLocation.FullName + "\" + pGxDialog.Name
            OutputFileName = pGxDialog.Name
        Else
            If Microsoft.VisualBasic.Right(pGxDialog.Name, 4) <> ".shp" Then
                txtOutput.Text = pGxDialog.FinalLocation.FullName + "\" + pGxDialog.Name + ".shp"
                OutputFileName = pGxDialog.Name
            Else
                txtOutput.Text = pGxDialog.FinalLocation.FullName + "\" + pGxDialog.Name
                OutputFileName = pGxDialog.Name
            End If

        End If

        If File.Exists(txtOutput.Text) = True Then
            pb_ExistAdd.Visible = True
        ElseIf pGxDialog.ReplacingObject = True Then
            pb_ExistAdd.Visible = True
        End If
        ffrmMain.Focus()
    End Sub

    Private Sub chkGPBulkBilling_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkGPBulkBilling.CheckedChanged
        If chkGPBulkBilling.Checked = True Then
            chkGPFee.Checked = False
            chkGPBulkbillingOnly.Checked = False
        End If
    End Sub

    Private Sub chkGPFee_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkGPFee.CheckedChanged
        If chkGPFee.Checked = True Then
            chkGPBulkBilling.Checked = False
            chkGPBulkbillingOnly.Checked = False
        End If
    End Sub
    Private Sub chkGPBulkbillingOnly_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkGPBulkbillingOnly.CheckedChanged
        If chkGPBulkbillingOnly.Checked = True Then
            chkGPFee.Checked = False
            chkGPBulkBilling.Checked = False
        End If

    End Sub

    Private Sub btnOutput_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOutput.Click
        On Error GoTo err_h

        ErrorExist = False
        chkSetSymbol = False
        Dim QueryProperty As String
        IntersectPolyStr = ""
        '--------Ouput Address Validation--------------------------------------
        '----------------------------------------------------------------------
        If txtOutput.Text = "" Then
            MsgBox("Please Specify Output Address ", MsgBoxStyle.Critical)
            Exit Sub
        ElseIf Microsoft.VisualBasic.Right(txtOutput.Text, 4) <> ".shp" And Microsoft.VisualBasic.InStr(txtOutput.Text, ".mdb") = 0 Then
            MsgBox("Output Address is not valid", MsgBoxStyle.Critical)
            Exit Sub
        End If
        If pb_ExistAdd.Visible = True Then
            MsgBox("Output Address '" + txtOutput.Text + "' already exists ", MsgBoxStyle.Critical)
            txtOutput.Text = ""
            Exit Sub
        End If

        If chkGP.Checked = True Then
            If lblGP.Text = 0 Then
                MsgBox(" Please Determine the Distance from GPs by Slider bar", MsgBoxStyle.Critical)
                Exit Sub
            End If

        End If
        '--------------------------------------------------------------------
        '--------------------------------------------------------------------
        '--Start-------------------------------------------------------------
        Me.Height = 415
        ProgressBar1.Value = 0
        btnClear.Enabled = False
        btnOutput.Enabled = False


        If (Not System.IO.Directory.Exists(TempPath)) Then
            System.IO.Directory.CreateDirectory(TempPath)
        End If
        If (Not System.IO.File.Exists(TempOutputPath)) Then
            Fun_CreateGeodatabase(TempPath, TempGDBName)
        End If
        '--------------------------------------------------------------------
        '--------------------------------------------------------------------
        ProgressVal1 = 100
        ProgressVal2 = 0

        Fun_output_Health()

        '--Final--------------------------------------------------------------------------
Final:

        Unselect()
        ProgressBar1.Value = 100
        Me.Height = 390
        txtOutput.Text = ""
        btnClear.Enabled = True
        btnOutput.Enabled = True
        If ErrorExist = False Then
            MsgBox("The analysis has done succesfully.You can find the result on the map", MsgBoxStyle.OkOnly, "DONE")
        End If

        Exit Sub
err_h:

    End Sub
    Private Sub Fun_output_Health()
        On Error GoTo err_h
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim NumCriteria As Integer
        Dim QueryAfterHour As String = ""

        If chkSEIFA.Checked = True Then
            NumCriteria = NumCriteria + 1
        End If
        If chkDepression.Checked = True Then
            NumCriteria = NumCriteria + 1
        End If
        If chkDiabets.Checked = True Then
            NumCriteria = NumCriteria + 1
        End If

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '------Analysis-------------------------------------------------------------
        '---------------------------------------------------------------------------
        ''----SEIFA-----------------------------------------------------------------
        Dim NameOutlayer As String = ""
        ProgressBar1.Value = 10
        Dim Querystr As String = ""
        If ErrorExist = False Then
            If chkSEIFA.Checked = True Then
                If cmbSEIFA.Text <> "" Then
                    If cmbSEIFA.Text = "<" Or cmbSEIFA.Text = "<=" Then
                        Querystr = "IRSD_Decil" + cmbSEIFA.Text + lblSEIFA.Text + " and IRSD_Decil>0"
                    Else
                        Querystr = "IRSD_Decil" + cmbSEIFA.Text + lblSEIFA.Text
                    End If

                End If

                If ErrorExist = False Then
                    If Querystr <> "" Then
                        FindFeatures(fcSEIFA, Querystr)
                        If FeatureLayerName <> "" Then
                            NameOutlayer = "Selected_SEIFA"
                            Fun_CopyFeatures(FeatureLayerName, TempOutputPath + "\" + "Selected_SEIFA", True)
                            If IntersectPolyStr = "" Then
                                IntersectPolyStr = TempOutputPath + "\" + "Selected_SEIFA"
                            Else
                                IntersectPolyStr = IntersectPolyStr + ";" + TempOutputPath + "\" + "Selected_SEIFA"
                            End If


                        End If
                        ProgressBar1.Value = ProgressBar1.Value + 20
                    Else
                        GoTo err_h
                    End If
                End If
            End If
        End If
        ''----Depression-----------------------------------------------------------------
        If ErrorExist = False Then
            Querystr = ""
            If chkDepression.Checked = True Then
                If cmbDepression.Text <> "" Then
                    Querystr = "RatePer100" + cmbDepression.Text + lblDepression.Text
                End If

                If ErrorExist = False Then
                    If Querystr <> "" Then
                        FindFeatures(fcDepression, Querystr)
                        If FeatureLayerName <> "" Then
                            NameOutlayer = "Selected_Depression"
                            Fun_CopyFeatures(FeatureLayerName, TempOutputPath + "\" + "Selected_Depression", True)

                            If IntersectPolyStr = "" Then
                                IntersectPolyStr = TempOutputPath + "\" + "Selected_Depression"
                            Else
                                IntersectPolyStr = IntersectPolyStr + ";" + TempOutputPath + "\" + "Selected_Depression"
                            End If
                        End If

                        ProgressBar1.Value = ProgressBar1.Value + 20
                    Else
                        GoTo err_h
                    End If
                End If
            End If
        End If
        ''----Diabets-----------------------------------------------------------------
        If ErrorExist = False Then
            Querystr = ""
            If chkDiabets.Checked = True Then
                If cmbDiabets.Text <> "" Then
                    Querystr = "RatePer100" + cmbDiabets.Text + lblDiabets.Text
                End If

                If ErrorExist = False Then
                    If Querystr <> "" Then
                        FindFeatures(fcDiabetes, Querystr)
                        If FeatureLayerName <> "" Then
                            NameOutlayer = "Selected_Diabetes"
                            Fun_CopyFeatures(FeatureLayerName, TempOutputPath + "\" + "Selected_Diabetes", True)

                            If IntersectPolyStr = "" Then
                                IntersectPolyStr = TempOutputPath + "\" + "Selected_Diabetes"
                            Else
                                IntersectPolyStr = IntersectPolyStr + ";" + TempOutputPath + "\" + "Selected_Diabetes"
                            End If

                        End If
                        ProgressBar1.Value = ProgressBar1.Value + 20
                    Else
                        GoTo err_h
                    End If
                End If
            End If
        End If

        '-GPs-------------------------------------------------------------------------
        Dim SelectedGPLayer As String = ""
        Dim chkGPs As Boolean = False
        If ErrorExist = False Then
            Querystr = ""
            If chkGP.Checked = True Then
                If lblGP.Text <> "0" Then
                    'fcGPs = "_" + lblGP.Text + "m"
                    fcGPs = lblGP.Text + "m"
                End If
                If chkGPBulkBilling.Checked = True Or chkGPFee.Checked = True Or chkGPBulkbillingOnly.Checked = True Or chkAfterHSSat.Checked = True Or chkAfterHSSun.Checked = True Or chkAfterHSW5_8.Checked = True Or chkAfterHSW8.Checked = True Then
                    chkGPs = True

                    If chkGPBulkBilling.Checked = True Then
                        Querystr = "FreeProvis = 'Fees & Bulk Billing'"
                    ElseIf chkGPFee.Checked = True Then
                        Querystr = "FreeProvis = 'Fees Apply'"
                    ElseIf chkGPBulkbillingOnly.Checked = True Then
                        Querystr = "FreeProvis = 'Bulkbilling only'"
                    End If

                    QueryAfterHour = Fun_CreateAfterHourQuery()

                    If QueryAfterHour <> "" Then
                        If Querystr = "" Then
                            Querystr = QueryAfterHour
                        Else
                            Querystr = Querystr + " AND " + QueryAfterHour
                        End If
                    End If

                    If Querystr <> "" Then

                        FindFeatures(fcGPs, Querystr)
                        If FeatureLayerName <> "" Then
                            Fun_CopyFeatures(FeatureLayerName, TempOutputPath + "\" + "Selected_GPs", True)
                            SelectedGPLayer = "Selected_GPs"
                        End If
                        ProgressBar1.Value = ProgressBar1.Value + 10
                    End If

                Else
                    SelectedGPLayer = FindGetLayerName(fcGPs)
                End If
            End If
        End If
        ''------------------Intersection Polygon------------------------------------
        If chkGP.Checked = True Then
            If NumCriteria > 1 Then
                If IntersectPolyStr <> "" Then
                    If ErrorExist = False Then
                        Fun_IntersectLayers(IntersectPolyStr, TempOutputPath + "\Intersect_HealthIndex", "ONLY_FID", True)
                        ProgressBar1.Value = ProgressBar1.Value + 2
                        NameOutlayer = "Intersect_HealthIndex"
                    Else
                        GoTo err_h
                    End If
                End If
            End If

            '=---------------------------------------------------------------------------
            If ErrorExist = False Then
                'Fun_SelectByLocation(NameOutlayer, SelectedGPLayer, "WITHIN_A_DISTANCE", "NEW_SELECTION", lblGP.Text)
                'Fun_SelectByAttribute(NameOutlayer, "", "SWITCH_SELECTION")
                Fun_EraseLayers(NameOutlayer, txtOutput.Text, fcGPs, True)
                'Fun_CopyFeatures(NameOutlayer, txtOutput.Text, True)
                If ErrorExist = False Then
                    Unselect()
                Else
                    GoTo err_h
                End If


            Else
                GoTo err_h
            End If
        Else
            If IntersectPolyStr <> "" Then
                If ErrorExist = False Then
                    If NumCriteria > 1 Then
                        Fun_IntersectLayers(IntersectPolyStr, txtOutput.Text, "ONLY_FID", True)
                        ProgressBar1.Value = ProgressBar1.Value + 10
                    Else
                        Fun_CopyFeatures(NameOutlayer, txtOutput.Text, True)
                    End If
                Else
                    GoTo err_h
                End If
            End If

        End If


        Display_Symbology_Layer(OutputFileName, Microsoft.VisualBasic.Left(PathGDB, InStrRev(PathGDB, "\")) & "OutputSymbol.lyr")

        Fun_DeleteLayerfromMap("Selected_GPs")
        Fun_DeleteLayerfromMap("Intersect_HealthIndex")
        Fun_DeleteLayerfromMap("Selected_SEIFA")
        Fun_DeleteLayerfromMap("Selected_Depression")
        Fun_DeleteLayerfromMap("Selected_Diabetes")
       
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Exit Sub
        '----------------------------------------------------------------------------
err_h:
        If ErrorExist = True Then
            MsgBox(ErrorDescription, MsgBoxStyle.OkOnly, "Error")
        Else
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Error")
        End If
        ProgressBar1.Value = 100
        Me.Height = 390
        txtOutput.Text = ""
        btnClear.Enabled = True
        btnOutput.Enabled = True
    End Sub
    Private Function Fun_CreateAfterHourQuery() As String
        Dim QueryAfterHour As String = ""

        If chkAfterHSW5_8.Checked = True Then
            QueryAfterHour = "(Monday LIKE '*-5:30PM*' OR Tuesday LIKE '*-5:30PM*' OR Wednesday LIKE '*-5:30PM*' OR Thursday LIKE '*-5:30PM*' OR Friday LIKE '*-5:30PM*' OR" +
                             " Monday LIKE '*-6:00PM*' OR Tuesday LIKE '*-6:00PM*' OR Wednesday LIKE '*-6:00PM*' OR Thursday LIKE '*-6:00PM*' OR Friday LIKE '*-6:00PM*' OR" +
                             " Monday LIKE '*-6:30PM*' OR Tuesday LIKE '*-6:30PM*' OR Wednesday LIKE '*-6:30PM*' OR Thursday LIKE '*-6:30PM*' OR Friday LIKE '*-6:30PM*' OR" +
                             " Monday LIKE '*-7:00PM*' OR Tuesday LIKE '*-7:00PM*' OR Wednesday LIKE '*-7:00PM*' OR Thursday LIKE '*-7:00PM*' OR Friday LIKE '*-7:00PM*' OR" +
                             " Monday LIKE '*-7:30PM*' OR Tuesday LIKE '*-7:30PM*' OR Wednesday LIKE '*-7:30PM*' OR Thursday LIKE '*-7:30PM*' OR Friday LIKE '*-7:30PM*' OR" +
                             " Monday LIKE '*-8:00PM*' OR Tuesday LIKE '*-8:00PM*' OR Wednesday LIKE '*-8:00PM*' OR Thursday LIKE '*-8:00PM*' OR Friday LIKE '*-8:00PM*' )"
        End If
        If chkAfterHSW8.Checked = True Then
            QueryAfterHour = "(Monday LIKE '*-8:30PM*' OR Tuesday LIKE '*-8:30PM*' OR Wednesday LIKE '*-8:30PM*' OR Thursday LIKE '*-8:30PM*' OR Friday LIKE '*-8:30PM*' OR" +
                            " Monday LIKE '*-9:00PM*' OR Tuesday LIKE '*-9:00PM*' OR Wednesday LIKE '*-9:00PM*' OR Thursday LIKE '*-9:00PM*' OR Friday LIKE '*-9:00PM*' OR" +
                            " Monday LIKE '*-9:30PM*' OR Tuesday LIKE '*-9:30PM*' OR Wednesday LIKE '*-9:30PM*' OR Thursday LIKE '*-9:30PM*' OR Friday LIKE '*-9:30PM*' OR" +
                            " Monday LIKE '*-10:00PM*' OR Tuesday LIKE '*-10:00PM*' OR Wednesday LIKE '*-10:00PM*' OR Thursday LIKE '*-10:00PM*' OR Friday LIKE '*-10:00PM*' OR" +
                            " Monday LIKE '*-10:30PM*' OR Tuesday LIKE '*-10:30PM*' OR Wednesday LIKE '*-10:30PM*' OR Thursday LIKE '*-10:30PM*' OR Friday LIKE '*-10:30PM*' OR" +
                            " Monday LIKE '*-11:00PM*' OR Tuesday LIKE '*-11:00PM*' OR Wednesday LIKE '*-11:00PM*' OR Thursday LIKE '*-11:00PM*' OR Friday LIKE '*-11:00PM*' OR" +
                            " Monday LIKE '*-11:30PM*' OR Tuesday LIKE '*-11:30PM*' OR Wednesday LIKE '*-11:30PM*' OR Thursday LIKE '*-11:30PM*' OR Friday LIKE '*-11:30PM*' )"
        End If
        If chkAfterHSSat.Checked = True Then
            If QueryAfterHour <> "" Then
                QueryAfterHour = QueryAfterHour + "and (Saturday LIKE '*PM*')"
            Else
                QueryAfterHour = "(Saturday LIKE '*PM*')"
            End If
        End If

        If chkAfterHSSun.Checked = True Then
            If QueryAfterHour <> "" Then
                QueryAfterHour = QueryAfterHour + "and (Sunday  LIKE '*PM*')"
            Else
                QueryAfterHour = "(Sunday LIKE '*AM*' or Sunday LIKE '*PM*')"
            End If
        End If
        If QueryAfterHour <> "" Then
            QueryAfterHour = "(" + QueryAfterHour + ")"
        End If
        Return QueryAfterHour
    End Function
 
    Private Sub chkGP_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkGP.CheckedChanged
        If chkGP.Checked = True Then
            trkb_GP.Enabled = True
            chkGPAfterHours.Enabled = True

            chkGPBulkBilling.Enabled = True
            chkGPFee.Enabled = True
            chkGPBulkbillingOnly.Enabled = True

            chkAfterHSSat.Enabled = True
            chkAfterHSSun.Enabled = True
            chkAfterHSW5_8.Enabled = True
            chkAfterHSW8.Enabled = True
        Else

            trkb_GP.Enabled = False
            chkGPAfterHours.Enabled = False

            chkGPBulkBilling.Enabled = False
            chkGPFee.Enabled = False
            chkGPBulkbillingOnly.Enabled = False

            chkAfterHSSat.Enabled = False
            chkAfterHSSun.Enabled = False
            chkAfterHSW5_8.Enabled = False
            chkAfterHSW8.Enabled = False

            chkGPAfterHours.Checked = False

            chkGPBulkBilling.Checked = False
            chkGPFee.Checked = False
            chkGPBulkbillingOnly.Checked = False

            chkAfterHSSat.Checked = False
            chkAfterHSSun.Checked = False
            chkAfterHSW5_8.Checked = False
            chkAfterHSW8.Checked = False
        End If


    End Sub

   
    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        ErrorExist = True
        Me.Close()
    End Sub

    Private Sub txtOutput_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOutput.TextChanged
        If File.Exists(txtOutput.Text) = True Then
            pb_ExistAdd.Visible = True
        Else
            pb_ExistAdd.Visible = False
        End If
    End Sub
End Class