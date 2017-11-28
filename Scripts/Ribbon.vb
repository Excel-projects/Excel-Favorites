Option Strict Off
Option Explicit On

Imports System.Windows.Forms

Namespace Scripts

    <Runtime.InteropServices.ComVisible(True)>
    Public Class Ribbon
        Implements Office.IRibbonExtensibility
        Private ribbon As Office.IRibbonUI

        Private mySettings As TaskPane.Settings
        Private myTaskPaneSettings As Microsoft.Office.Tools.CustomTaskPane

#Region "| Ribbon Events |"

        Public Sub New()
        End Sub

        Public Function GetCustomUI(ByVal ribbonID As String) As String Implements Office.IRibbonExtensibility.GetCustomUI
            Return GetResourceText("Favorites.Ribbon.xml")
        End Function

        Private Shared Function GetResourceText(ByVal resourceName As String) As String
            Dim asm As Reflection.Assembly = Reflection.Assembly.GetExecutingAssembly()
            Dim resourceNames() As String = asm.GetManifestResourceNames()
            For i As Integer = 0 To resourceNames.Length - 1
                If String.Compare(resourceName, resourceNames(i), StringComparison.OrdinalIgnoreCase) = 0 Then
                    Using resourceReader As IO.StreamReader = New IO.StreamReader(asm.GetManifestResourceStream(resourceNames(i)))
                        If resourceReader IsNot Nothing Then
                            Return resourceReader.ReadToEnd()
                        End If
                    End Using
                End If
            Next
            Return Nothing
        End Function

        Public Sub Ribbon_Load(ByVal ribbonUI As Office.IRibbonUI)
            Me.ribbon = ribbonUI
        End Sub

        Public Function GetLabelText(ByVal control As Office.IRibbonControl) As String
            Try
                Select Case control.Id.ToString
                    Case Is = "tabFavorites"
                        If Application.ProductVersion.Substring(0, 2) = "15" Then
                            Return My.Application.Info.Title.ToUpper()
                        Else
                            Return My.Application.Info.Title
                        End If
                    Case Is = "txtCopyright"
                        Return "© " & My.Application.Info.Copyright.ToString
                    Case Is = "txtDescription"
                        Dim version As String = My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor & "." & My.Application.Info.Version.Build & "." & My.Application.Info.Version.Revision
                        Return My.Application.Info.Title.ToString.Replace("&", "&&") & Space(1) & version
                    Case Is = "txtReleaseDate"
                        Return My.Settings.App_ReleaseDate.ToString("dd-MMM-yyyy hh:mm tt")
                    Case Else
                        Return String.Empty
                End Select

            Catch ex As Exception
                Call ErrorHandler.DisplayMessage(ex)
                Return String.Empty

            End Try

        End Function

        Public Sub OnAction(ByVal Control As Office.IRibbonControl)
            Try
                Select Case Control.Id
                    Case "btnCut"
                        CutSelection()
                    Case "btnOpenReadMe"
                        OpenReadMe()
                    Case "btnOpenNewIssue"
                        OpenNewIssue()
                    Case "btnSettings"
                        OpenSettings()
                End Select

            Catch ex As Exception
                Call ErrorHandler.DisplayMessage(ex)

            End Try

        End Sub

#End Region

#Region "| Ribbon Buttons |"

        Public Sub CutSelection()
            Try
                Globals.ThisAddIn.Application.Selection.Cut()

            Catch ex As Exception
                Call ErrorHandler.DisplayMessage(ex)

            End Try

        End Sub

        Public Sub OpenReadMe()
            System.Diagnostics.Process.Start(My.Settings.App_PathReadMe)
        End Sub

        Public Sub OpenNewIssue()
            System.Diagnostics.Process.Start(My.Settings.App_PathReportIssue)

        End Sub

        Public Sub OpenSettings()
            Try
                If myTaskPaneSettings IsNot Nothing Then
                    If myTaskPaneSettings.Visible = True Then
                        myTaskPaneSettings.Visible = False
                    Else
                        myTaskPaneSettings.Visible = True
                    End If
                Else
                    mySettings = New Favorites.TaskPane.Settings()
                    myTaskPaneSettings = Globals.ThisAddIn.CustomTaskPanes.Add(mySettings, "Settings For " + My.Application.Info.Title)
                    myTaskPaneSettings.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight
                    myTaskPaneSettings.DockPositionRestrict = Office.MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange
                    myTaskPaneSettings.Width = 675
                    myTaskPaneSettings.Visible = True

                End If

            Catch ex As Exception
                Call ErrorHandler.DisplayMessage(ex)

            End Try

        End Sub

#End Region

#Region "| Subroutines |"

        Public Sub OpenFile(ByVal fileName As String)
            Dim pStart As New System.Diagnostics.Process
            Try
                If fileName = String.Empty Then Exit Try
                pStart.StartInfo.FileName = fileName
                pStart.Start()

            Catch ex As System.ComponentModel.Win32Exception
                'MessageBox.Show("No application Is assicated To this file type." & vbCrLf & vbCrLf & pstrFile, "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Try

            Catch ex As Exception
                Call ErrorHandler.DisplayMessage(ex)
                Exit Try

            Finally
                pStart.Dispose()

            End Try

        End Sub

#End Region

    End Class

End Namespace