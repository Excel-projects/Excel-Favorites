Option Strict On
Option Explicit On

Imports System.Environment
Imports System.Windows.Forms

Namespace Scripts

	Module ErrorHandler

		''' <summary> 
		''' Global error message for all procedures
		''' </summary>
		''' <param name="ex">the handled exception</param>
		''' <param name="silent">show or hide the message</param>
		Public Sub DisplayMessage(ByRef ex As Exception, Optional ByVal silent As Boolean = False)
			Dim sf As New System.Diagnostics.StackFrame(1)
			Dim caller As System.Reflection.MethodBase = sf.GetMethod()
			Dim procedure As String = (caller.Name).Trim
			Dim msg As String = "Contact your system administrator."
			msg += NewLine & "Procedure: " & procedure
			msg += NewLine & "Description: " & ex.ToString
			Console.WriteLine(msg)
			If silent = False Then
				MessageBox.Show(msg, "Unexpected Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
			End If

		End Sub

	End Module

End Namespace
