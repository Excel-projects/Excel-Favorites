Option Strict On
Option Explicit On

Imports System.Windows.Forms
Imports System.Reflection

Namespace Favorites.TaskPane
	''' <summary>
	''' Settings TaskPane
	''' </summary>
	Partial Public Class Settings
		Inherits UserControl
		''' <summary>
		''' Initialize the controls in the object
		''' </summary>
		Public Sub New()
			InitializeComponent()
			Me.pgdSettings.SelectedObject = My.Settings
		End Sub

		Friend WithEvents pgdSettings As PropertyGrid

		Private Sub InitializeComponent()
		Me.pgdSettings = New System.Windows.Forms.PropertyGrid()
		Me.SuspendLayout
		'
		'pgdSettings
		'
		Me.pgdSettings.Dock = System.Windows.Forms.DockStyle.Fill
		Me.pgdSettings.Location = New System.Drawing.Point(0, 0)
		Me.pgdSettings.Name = "pgdSettings"
		Me.pgdSettings.Size = New System.Drawing.Size(650, 750)
		Me.pgdSettings.TabIndex = 0
		'
		'Settings
		'
		Me.Controls.Add(Me.pgdSettings)
		Me.Name = "Settings"
		Me.Size = New System.Drawing.Size(650, 750)
		Me.ResumeLayout(false)

End Sub

		''' <summary> 
		''' Sets the column width of a property grid 
		''' </summary>
		''' <param name="grid">Represents the property grid object. </param>
		''' <param name="width">Represents the width of the column. </param>
		''' <remarks></remarks>
		Public Shared Sub SetLabelColumnWidth(grid As PropertyGrid, width As Integer)
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
		End Sub

		''' <summary>
		''' Update the ribbon with the changed settings values
		''' </summary>
		''' <param name="s">contains the sender of the event, so if you had one method bound to multiple controls, you can distinguish them.</param>
		''' <param name="e">refers to the event arguments for the used event, they usually come in the form of properties/functions/methods that get to be available on it.</param>
		Private Sub pgdSettings_PropertyValueChanged(s As Object, e As PropertyValueChangedEventArgs)
			'Scripts.Ribbon.ribbonref.InvalidateRibbon()
		End Sub

	End Class

End Namespace