Option Strict On
Option Explicit On

Imports System.Windows.Forms
Imports Excel = Microsoft.Office.Interop.Excel

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

        Public Shared Function IsActiveDocument(Optional showMsg As Boolean = False) As Boolean
            Try
                If Globals.ThisAddIn.Application.ActiveWorkbook Is Nothing Then
                    If showMsg = True Then
                        MessageBox.Show("The command could not be completed.  Please open a document and select a range.", My.Application.Info.Description, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    End If
                    Return False
                Else
                    Return True
                End If

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)
                Return False
            End Try

        End Function

        Public Shared Function IsActiveSelection(Optional showMsg As Boolean = False) As Boolean
            Dim checkRange As Excel.Range = Nothing
            Try
                checkRange = TryCast(Globals.ThisAddIn.Application.Selection, Excel.Range)
                'must cast the selection as range or errors
                If checkRange Is Nothing Then
                    If showMsg = True Then
                        MessageBox.Show("The command could not be completed by using the range specified.  Select a single cell within the range and try the command again. [Range]", My.Application.Info.Description, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    End If
                    Return False
                Else
                    Return True
                End If

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)
                Return False

            Finally
                If checkRange IsNot Nothing Then
                    'Marshal.ReleaseComObject(checkRange)
                End If
            End Try

        End Function

        Private Shared Function IsInCellEditingMode(Optional showMsg As Boolean = False) As Boolean
            Dim flag As Boolean = False
            Try
                'This will throw an Exception if Excel is in Cell Editing Mode
                Globals.ThisAddIn.Application.DisplayAlerts = False

            Catch generatedExceptionName As Exception
                If showMsg = True Then
                    MessageBox.Show("The procedure can not run while you are editing a cell.", "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
                flag = True
            End Try
            Return flag

        End Function

        Public Shared Function IsEnabled(Optional showMsg As Boolean = False) As Boolean
            Try
                If IsActiveDocument(showMsg) = False Then
                    Return False
                Else
                    If IsActiveSelection(showMsg) = False Then
                        Return False
                    Else
                        If IsInCellEditingMode(showMsg) = True Then
                            Return False
                        Else
                            Return True
                        End If
                    End If
                End If

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)
                Return False
            End Try

        End Function

    End Class

End Namespace
