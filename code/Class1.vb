Imports System.Diagnostics.Eventing.Reader
Imports System.Runtime.InteropServices
Imports System.Text

Public Class LCDSmartie


    Private Delegate Function EnumWindowsDelegate(ByVal hWnd As IntPtr, ByVal lParam As IntPtr) As Boolean

    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function EnumWindows(ByVal lpEnumFunc As EnumWindowsDelegate, ByVal lParam As IntPtr) As Boolean
    End Function

    Public Shared Function CompareWindowTextToAll(ByVal textToCompare As String) As Integer
        Dim found As Boolean = False

        EnumWindows(Function(hWnd, lParam)
                        Dim windowTitle As New Text.StringBuilder(256)
                        GetWindowText(hWnd, windowTitle, 256)

                        If windowTitle.ToString().Contains(textToCompare) Then
                            found = True
                            Return False ' Stop enumeration
                        End If

                        Return True
                    End Function, IntPtr.Zero)

        If found Then
            Return 1 ' Match found
        Else
            Return 0 ' No match found
        End If

    End Function




    Public Declare Function GetForegroundWindow Lib "user32" Alias "GetForegroundWindow" () As IntPtr
    Public Declare Auto Function GetWindowText Lib "user32" (ByVal hWnd As System.IntPtr, ByVal lpString As System.Text.StringBuilder, ByVal cch As Integer) As Integer
    Function GetCaption() As String
        Dim Caption As New System.Text.StringBuilder(256)
        Dim hWnd As IntPtr = GetForegroundWindow()
        GetWindowText(hWnd, Caption, Caption.Capacity)
        Return Caption.ToString()
    End Function
    Public Function function1(param1 As String, param2 As String)
        Dim windowName As String = GetCaption()
        '       Dim comparisionResult As Integer
        If Trim(param1) <> "" Then

            If windowName.Contains(param1) Then
                Return "1"
            Else
                Return "0"
            End If


        Else
            Return "param1 should be the title to compare - partial"
        End If
    End Function


    Public Function function2(param1 As String, param2 As String)

        Dim windowName As String = GetCaption()
        '  Dim comparisionResult As Integer
        If Trim(param1) <> "" Then
            If windowName.Equals(param1) Then
                Return "1"
            Else
                Return "0"
            End If


        Else
            Return "param1 should be the title to compare - exact"
        End If
    End Function



    Public Function function3(param1 As String, param2 As String)
        Dim windowName As String = GetCaption()
        '  Dim comparisionResult As Integer
        Dim compType As String
        If param2 = "partial" Or param2 = "p" Or param2 = "par" Or param2 = "2" Then
            If windowName.Contains(param1) Then
                Return param1
            Else
                Return ""
            End If
        ElseIf param2 = "exact" Or param2 = "e" Or param2 = "exa" Or param2 = "1" Then
            If windowName.Equals(param1) Then
                Return param1
            Else
                Return ""
            End If
        Else
            If windowName.Contains(param1) Then
                Return param1
            Else
                Return ""
            End If
        End If


        If Trim(param1) = "" Then

            Return "param1 should be the title to compare -  param2 the comparision type exact or partial"
        End If
    End Function
    Public Function function4(param1 As String, param2 As String)

        Dim windowName As String = GetCaption()
        Return windowName
    End Function


    Public Function function19(param1 As String, param2 As String)
        Return CompareWindowTextToAll(param1).ToString
        If Trim(param1) = "" Or param1 = Nothing Then
            Return "param1 should be the title to search for in all running windows"
        End If
    End Function

    Public Function function20(param1 As String, param2 As String)

        Return "Probe v 1.0"



    End Function



    Public Function SmartieDemo()
        Dim demolist As New StringBuilder()

        demolist.AppendLine("probe plugin for LCD Smartie")
        demolist.AppendLine("This plugin returns info about the window which have focus on windows")
        demolist.AppendLine("It is created to automate screen and theme switching depending on user activity")
        demolist.AppendLine("------ Function1 ------")
        demolist.AppendLine("Returns 1 or 0 if the param1 is a part of the title of the current focused window")
        demolist.AppendLine("Name of the current window contains 'pad++': $dll(probe,1,pad ++,)")
        demolist.AppendLine("Name of the current window contains 'YouTube': $dll(probe,1,YouTube,)")
        demolist.AppendLine("Name of the current window contains 'Chrome': $dll(probe,1,,Chrome,)")
        demolist.AppendLine("")
        demolist.AppendLine("------ Function2 ------")
        demolist.AppendLine("Returns 1 or 0 if the param1 is a exact match with the title of the current focused window")
        demolist.AppendLine("Name of the current window is 'This PC': $dll(probe,2,This PC,)")
        demolist.AppendLine("Name of the current window is: $dll(probe,2,Spotify Premium,)")
        demolist.AppendLine("")
        demolist.AppendLine("------ Function3 ------")
        demolist.AppendLine("Returns window name (param1) if match otherwise returns nothing")
        demolist.AppendLine("param1 should be the term to search for (it is also the return value on match)")
        demolist.AppendLine("param2 is the comparision type partial,p,par,2 or exact,e,exa,1.")
        demolist.AppendLine("")

        demolist.AppendLine("------ Function4 ------")
        demolist.AppendLine("This function is created to give diagnostic info")
        demolist.AppendLine("displays the current window name ")
        demolist.AppendLine("")

        demolist.AppendLine("------ Function19 ------")
        demolist.AppendLine("Searches for the param1 to all running programs titles (partial match and returns 1 or 0")
        demolist.AppendLine("")

        demolist.AppendLine("------ Function20 ------")
        demolist.AppendLine(">>> Credits <<<")

        demolist.AppendLine("---------------------------------------------------------------------------------------------------------")
        demolist.AppendLine(" *** Visit ***")
        demolist.AppendLine("> Old home page")
        demolist.AppendLine("https://lcdsmartie.sourceforge.net")
        demolist.AppendLine("> Forums")
        demolist.AppendLine("https://www.lcdsmartie.org")
        demolist.AppendLine("> New official development branch (latest version)")
        demolist.AppendLine("https://github.com/LCD-Smartie/LCDSmartie")
        demolist.AppendLine("")

        Dim result As String = demolist.ToString()
        Return result
    End Function

    Public Function SmartieInfo()
        Return "Developer: Nikos Georgousis (limbo)" & vbNewLine & "Version: 1.0 "
    End Function

    Public Function GetMinRefreshInterval() As Integer

        Return 100 ' ms 

    End Function

End Class
