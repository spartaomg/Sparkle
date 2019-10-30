Imports System.Windows.Forms

Namespace SubClassCtrl
    'To Subclass any Control, You need to Inherit From NativeWindow Class

    Public Class SubClassing
        Inherits System.Windows.Forms.NativeWindow

        'Event Declaration. This event will be raised when any message will be posted to the Control
        Public Event CallBackProc(ByRef m As Message)

        'Flag which indicates that either Event should be raised or not
        Private m_Subclassed As Boolean = False

        'During Creation of Object of this class, Pass the Handle of Control which you want to SubClass
        Public Sub New(ByVal handle As IntPtr)
            MyBase.AssignHandle(handle)
        End Sub

        'Terminate The SubClassing
        'There is no need to Create this Method. Cuz, when you will create the Object of this class, You will have the Method Named ReleaseHandle.
        'Just call that as you can see in this Sub

        'Public Sub RemoveHandle()
        '    MyBase.ReleaseHandle()
        'End Sub

        'To Enable or Disable Receiving Messages
        Public Property SubClass() As Boolean
            Get
                Return m_Subclassed
            End Get
            Set(ByVal Value As Boolean)
                m_Subclassed = Value
            End Set
        End Property

        Protected Overrides Sub WndProc(ByRef m As Message)
            If m_Subclassed Then 'If Subclassing is Enabled 
                RaiseEvent CallBackProc(m) 'then RaiseEvent
            End If
            MyBase.WndProc(m)
        End Sub

        Protected Overrides Sub Finalize()
            MyBase.Finalize()
        End Sub
    End Class

End Namespace
