Option Strict On
Option Explicit On

Imports System.Windows.Forms
Imports Favorites.Code
Imports System.Reflection

Namespace Forms

    ''' <summary>
    ''' The settings used for the addin
    ''' </summary>
    ''' <remarks></remarks>
    Public Class Settings

        ''' <summary>
        ''' Procedures to run during before the form opens
        ''' </summary>
        ''' <param name="sender">contains the sender of the event, so if you had one method bound to multiple controls, you can distinguish them.</param>
        ''' <param name="e">refers to the event arguments for the used event, they usually come in the form of properties/functions/methods that get to be available on it.</param>
        ''' <remarks></remarks>
        Private Sub FrmSettingsLoad(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
            Try
                Me.pgdSettings.SelectedObject = My.Settings
                Call SetLabelColumnWidth(Me.pgdSettings, 200)
                Call SetFormIcon(Me, My.Resources.Settings)
                Dim strVersion As String = My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor & "." & My.Application.Info.Version.Build & "." & My.Application.Info.Version.Revision
                Me.Text = "Settings for " & My.Application.Info.Title.ToString.Replace("&", "&&") & Space(1) & strVersion

                ''Only show "user" settings
                'Dim userAttr As New System.Configuration.UserScopedSettingAttribute
                'Dim attrs As New System.ComponentModel.AttributeCollection(userAttr)
                'pgdSettings.BrowsableAttributes = attrs

            Catch ex As Exception
                Call ErrorMsg(ex)
                Exit Try

            End Try

        End Sub

        ''' <summary>
        ''' Procedures to run when the forms closes
        ''' </summary>
        ''' <param name="sender">contains the sender of the event, so if you had one method bound to multiple controls, you can distinguish them.</param>
        ''' <param name="e">refers to the event arguments for the used event, they usually come in the form of properties/functions/methods that get to be available on it.</param>
        ''' <remarks></remarks>
        Private Sub FrmSettingsFormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
            Try
                My.Settings.Save()

            Catch ex As Exception
                Call ErrorMsg(ex)
                Exit Try

            End Try

        End Sub

        ''' <summary> 
        ''' Set the width of the property grid columns
        ''' </summary>
        ''' <param name="grid">the property grid control</param>
        ''' <param name="width">the width for resizing the control</param>
        Public Sub SetLabelColumnWidth(grid As PropertyGrid, width As Integer)
            Try

                If grid Is Nothing Then
                    Return
                End If

                Dim fi As FieldInfo = grid.[GetType]().GetField("gridView", BindingFlags.Instance Or BindingFlags.NonPublic)
                If fi Is Nothing Then
                    Return
                End If

                Dim view As Control = TryCast(fi.GetValue(grid), Control)
                If view Is Nothing Then
                    Return
                End If

                Dim mi As MethodInfo = view.[GetType]().GetMethod("MoveSplitterTo", BindingFlags.Instance Or BindingFlags.NonPublic)
                If mi Is Nothing Then
                    Return
                End If
                mi.Invoke(view, New Object() {width})

            Catch ex As Exception
                Call ErrorMsg(ex)
                Exit Try

            End Try
        End Sub

    End Class

End Namespace