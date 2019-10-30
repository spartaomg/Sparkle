Imports System.Windows.Forms

Namespace SubClassCtrl
    'From https://www.codeproject.com/Articles/3234/Subclassing-in-NET-The-pure-NET-way

    Public Class SubClassing
        Inherits System.Windows.Forms.NativeWindow

        Public Event CallBackProc(ByRef m As Message)

        Public Sub New(ByVal handle As IntPtr)
            AssignHandle(handle)
        End Sub

        Public Property SubClass() As Boolean = False

        Protected Overrides Sub WndProc(ByRef m As Message)
            If SubClass Then
                RaiseEvent CallBackProc(m)
            End If
            MyBase.WndProc(m)
        End Sub

        Protected Overrides Sub Finalize()
            MyBase.Finalize()
        End Sub
    End Class

End Namespace
