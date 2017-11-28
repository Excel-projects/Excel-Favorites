Option Strict On
Option Explicit On

Imports System.Environment
Imports System.Windows.Forms

Namespace Scripts

        Public Class ErrorHandler

            Public Shared Sub DisplayMessage(ex As Exception)
                Dim sf As New System.Diagnostics.StackFrame(1)
                Dim caller As System.Reflection.MethodBase = sf.GetMethod()
                Dim currentProcedure As String = (caller.Name).Trim()
                Dim errorMessageDescription As String = ex.ToString()
                errorMessageDescription = System.Text.RegularExpressions.Regex.Replace(errorMessageDescription, "\r\n+", " ")
                Dim msg As String = "Contact your system administrator. A record has been created in the log file." + Environment.NewLine
                msg += (Convert.ToString("Procedure: ") & currentProcedure) + Environment.NewLine
                msg += "Description: " + ex.ToString() + Environment.NewLine
                MessageBox.Show(msg, "Unexpected Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
            End Sub

        End Class

    End Namespace
