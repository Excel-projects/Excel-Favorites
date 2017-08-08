Imports Favorites.Scripts
 
Public Class ThisAddIn

    Protected Overrides Function CreateRibbonExtensibilityObject() As Microsoft.Office.Core.IRibbonExtensibility
        Return New Ribbon()
    End Function

	<Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1822:MarkMembersAsStatic")>
	Private Sub ThisAddIn_Startup() Handles Me.Startup

	End Sub

	<Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1822:MarkMembersAsStatic")>
	Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

	End Sub

End Class
