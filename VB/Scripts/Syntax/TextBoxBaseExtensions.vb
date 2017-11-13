
Imports System.Runtime.InteropServices
Imports System.Windows.Forms

Namespace Scripts
    Namespace Syntax

        Public NotInheritable Class TextBoxBaseExtensions
            Private Sub New()
            End Sub

            Public Shared Sub DisableThenDoThenEnable(textBox As TextBoxBase, action As Action)
                Dim stateLocked As IntPtr = IntPtr.Zero

                Lock(textBox, stateLocked)

                Dim hscroll As Integer = GetHScrollPos(textBox)
                Dim vscroll As Integer = GetVScrollPos(textBox)

                Dim selstart As Integer = textBox.SelectionStart

                action()

                textBox.[Select](selstart, 0)

                SetHScrollPos(textBox, hscroll)
                SetVScrollPos(textBox, vscroll)

                Unlock(textBox, stateLocked)
            End Sub

            Public Shared Function GetHScrollPos(textBox As TextBoxBase) As Integer
                Return GetScrollPos(CInt(textBox.Handle), SB_HORZ)
            End Function

            Public Shared Sub SetHScrollPos(textBox As TextBoxBase, value As Integer)
                SetScrollPos(textBox.Handle, SB_HORZ, value, True)
                PostMessageA(textBox.Handle, WM_HSCROLL, SB_THUMBPOSITION + &H10000 * value, 0)
            End Sub

            Public Shared Function GetVScrollPos(textBox As TextBoxBase) As Integer
                Return GetScrollPos(CInt(textBox.Handle), SB_VERT)
            End Function

            Public Shared Sub SetVScrollPos(textBox As TextBoxBase, value As Integer)
                SetScrollPos(textBox.Handle, SB_VERT, value, True)
                PostMessageA(textBox.Handle, WM_VSCROLL, SB_THUMBPOSITION + &H10000 * value, 0)
            End Sub

            Private Shared Sub Lock(textBox As TextBoxBase, ByRef stateLocked As IntPtr)
                ' Stop redrawing:  
                SendMessage(textBox.Handle, WM_SETREDRAW, 0, IntPtr.Zero)
                ' Stop sending of events:  
                stateLocked = SendMessage(textBox.Handle, EM_GETEVENTMASK, 0, IntPtr.Zero)
                ' change colors and stuff in the RichTextBox  
            End Sub

            Private Shared Sub Unlock(textBox As TextBoxBase, ByRef stateLocked As IntPtr)
                ' turn on events  
                SendMessage(textBox.Handle, EM_SETEVENTMASK, 0, stateLocked)
                ' turn on redrawing  
                SendMessage(textBox.Handle, WM_SETREDRAW, 1, IntPtr.Zero)

                stateLocked = IntPtr.Zero
                textBox.Invalidate()
            End Sub

#Region "Win API Stuff"

            ' Windows APIs
            <DllImport("user32", CharSet:=CharSet.Auto)>
            Private Shared Function SendMessage(hWnd As IntPtr, msg As Integer, wParam As Integer, lParam As IntPtr) As IntPtr
            End Function

            <DllImport("user32.dll")>
            Private Shared Function PostMessageA(hWnd As IntPtr, nBar As Integer, wParam As Integer, lParam As Integer) As Boolean
            End Function

            <DllImport("user32.dll", CharSet:=CharSet.Auto)>
            Private Shared Function GetScrollPos(hWnd As Integer, nBar As Integer) As Integer
            End Function

            <DllImport("user32.dll")>
            Private Shared Function SetScrollPos(hWnd As IntPtr, nBar As Integer, nPos As Integer, bRedraw As Boolean) As Integer
            End Function

            Private Const WM_SETREDRAW As Integer = &HB
            Private Const WM_USER As Integer = &H400
            Private Const EM_GETEVENTMASK As Integer = (WM_USER + 59)
            Private Const EM_SETEVENTMASK As Integer = (WM_USER + 69)
            Private Const SB_HORZ As Integer = &H0
            Private Const SB_VERT As Integer = &H1
            Private Const WM_HSCROLL As Integer = &H114
            Private Const WM_VSCROLL As Integer = &H115
            Private Const SB_THUMBPOSITION As Integer = 4
            'private const int UNDO_BUFFER = 100;

#End Region
        End Class

    End Namespace
End Namespace