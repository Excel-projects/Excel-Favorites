Option Strict Off
Option Explicit On

Namespace Scripts

	''' <summary>
	''' The ribbon code used for the addin
	''' </summary>
	''' <remarks></remarks>
	<Runtime.InteropServices.ComVisible(True)>
	Public Class Ribbon
		Implements Office.IRibbonExtensibility
		<Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")>
		Private ribbon As Office.IRibbonUI

		Private mySettings As TaskPane.Settings
		Private myTaskPaneSettings As Microsoft.Office.Tools.CustomTaskPane

#Region "| Ribbon Events |"

		''' <summary>
		''' 
		''' </summary>
		Public Sub New()
		End Sub

		''' <summary>
		''' Loads the XML markup, either from an XML customization file or from XML markup embedded in the procedure, that customizes the Ribbon user interface.
		''' </summary>
		''' <param name="ribbonID">Represents the XML customization file</param>
		''' <returns>A method that returns a bitmap image for the control id.</returns>
		''' <remarks></remarks>
		<Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId:="0#")>
		Public Function GetCustomUI(ByVal ribbonID As String) As String Implements Office.IRibbonExtensibility.GetCustomUI
			Return GetResourceText("Favorites.Ribbon.xml")
		End Function

		''' <summary>
		''' 
		''' </summary>
		''' <param name="resourceName"></param>
		''' <returns></returns>
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

		''' <summary>
		''' Load the ribbon
		''' </summary>
		''' <param name="ribbonUI">Represents the IRibbonUI instance that is provided by the Microsoft Office application to the Ribbon extensibility code.</param>
		''' <remarks></remarks>
		<Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Naming", "CA1707:IdentifiersShouldNotContainUnderscores")>
		Public Sub Ribbon_Load(ByVal ribbonUI As Office.IRibbonUI)
			Me.ribbon = ribbonUI
		End Sub

		''' <summary>
		'''To assign a images to the controls on the ribbon in the xml file
		''' </summary>
		''' <param name="control">Represents the object passed into the callback procedure of a control in a ribbon or another user interface that can be customized by using Office Fluent ribbon extensibility.</param>
		''' <returns>A method that returns a bitmap image for the control id.</returns>
		''' <remarks></remarks>
		<Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes")>
		<Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1062:Validate arguments of public methods", MessageId:="0")>
		Public Function GetButtonImage(ByVal control As Office.IRibbonControl) As System.Drawing.Bitmap
			Try
				Select Case control.Id.ToString
					Case Is = "btnSettings"
						Return My.Resources.Resources.Settings
					Case Is = "btnCut"
						Return My.Resources.Resources.Cut
					Case Else
						Return Nothing
				End Select

			Catch ex As Exception
				Call DisplayMessage(ex)
				Return Nothing

			End Try

		End Function

		''' <summary>
		''' To assign text to controls on the ribbon from the xml file
		''' </summary>
		''' <param name="control">Represents the object passed into the callback procedure of a control in a ribbon or another user interface that can be customized by using Office Fluent ribbon extensibility.</param>
		''' <returns>A method that returns a string for a label. </returns>
		''' <remarks></remarks>
		<Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Globalization", "CA1305:SpecifyIFormatProvider", MessageId:="System.DateTime.ToString(System.String)")>
		<Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes")>
		<Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1062:Validate arguments of public methods", MessageId:="0")>
		Public Function GetLabelText(ByVal control As Office.IRibbonControl) As String
			Try
				Select Case control.Id.ToString
					Case Is = "tabFavorites"
						Return My.Application.Info.Title
					Case Is = "txtCopyright"
						Return "© " & My.Application.Info.Copyright.ToString
					Case Is = "txtDescription"
						Dim strVersion As String = My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor & "." & My.Application.Info.Version.Build & "." & My.Application.Info.Version.Revision
						Return My.Application.Info.Title.ToString.Replace("&", "&&") & Space(1) & strVersion
					Case Is = "txtReleaseDate"
						Dim dteCreateDate As DateTime = System.IO.File.GetLastWriteTime(My.Application.Info.DirectoryPath.ToString & "\" & My.Application.Info.AssemblyName.ToString & ".dll") 'get creation date 
						Return dteCreateDate.ToString("dd-MMM-yyyy hh:mm tt")
					Case Else
						Return String.Empty
				End Select

			Catch ex As Exception
				Call DisplayMessage(ex)
				'Console.WriteLine(ex.Message.ToString)
				Return String.Empty

			End Try

		End Function

#End Region

#Region "| Ribbon Buttons |"

		''' <summary>
		''' Using the application defined "Cut" so I can show a different icon file
		''' </summary>
		''' <param name="control">Represents the object passed into the callback procedure of a control in a ribbon or another user interface that can be customized by using Office Fluent ribbon extensibility.</param>
		''' <remarks></remarks>
		<Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA1801:ReviewUnusedParameters", MessageId:="control")>
		<Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes")>
		Public Sub CutSelection(ByVal control As Office.IRibbonControl)
			Try
				Globals.ThisAddIn.Application.Selection.Cut()

			Catch ex As Exception
				Call DisplayMessage(ex)

			End Try

		End Sub

		''' <summary>
		''' show the settings form
		''' </summary>
		''' <param name="control">Represents the object passed into the callback procedure of a control in a ribbon or another user interface that can be customized by using Office Fluent ribbon extensibility.</param>
		''' <remarks></remarks>
		<Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes")>
		<Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA1801:ReviewUnusedParameters", MessageId:="control")>
		Public Sub OpenSettingsForm(ByVal control As Office.IRibbonControl)
			Try
				If myTaskPaneSettings IsNot Nothing Then
					If myTaskPaneSettings.Visible = True Then
						myTaskPaneSettings.Visible = False
					Else
						myTaskPaneSettings.Visible = True
					End If
				Else
					mySettings = New Favorites.TaskPane.Settings()
					myTaskPaneSettings = Globals.ThisAddIn.CustomTaskPanes.Add(mySettings, "Settings for " + My.Application.Info.Title)
					myTaskPaneSettings.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight
					myTaskPaneSettings.DockPositionRestrict = Office.MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange
					myTaskPaneSettings.Width = 675
					myTaskPaneSettings.Visible = True

				End If

			Catch ex As Exception
				Call DisplayMessage(ex)

			End Try

		End Sub

		''' <summary>
		''' show the read me file
		''' </summary>
		''' <param name="control">Represents the object passed into the callback procedure of a control in a ribbon or another user interface that can be customized by using Office Fluent ribbon extensibility.</param>
		''' <remarks></remarks>
		<Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes")>
		<Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA1801:ReviewUnusedParameters", MessageId:="control")>
		Public Sub OpenHelpAsBuiltFile(ByVal control As Office.IRibbonControl)
			Try
				Call OpenFile(My.Settings.App_PathReadMe)

			Catch ex As Exception
				Call DisplayMessage(ex)

			End Try

		End Sub

#End Region

#Region "| Subroutines |"

		''' <summary>
		''' open a file from the source list
		''' </summary>
		''' <param name="fileName">The selected file to open</param>
		''' <remarks></remarks>
		<Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes")>
		Public Sub OpenFile(ByVal fileName As String)
			Dim pStart As New System.Diagnostics.Process
			Try
				If fileName = String.Empty Then Exit Try
				pStart.StartInfo.FileName = fileName
				pStart.Start()

			Catch ex As System.ComponentModel.Win32Exception
				'MessageBox.Show("No application is assicated to this file type." & vbCrLf & vbCrLf & pstrFile, "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Information)
				Exit Try

			Catch ex As Exception
				Call DisplayMessage(ex)
				Exit Try

			Finally
				pStart.Dispose()

			End Try

		End Sub

#End Region

	End Class

End Namespace