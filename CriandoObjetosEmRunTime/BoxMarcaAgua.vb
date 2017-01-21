Imports System.Runtime.InteropServices
Imports System.Windows.Forms

Public Class BoxMarcaAgua
    Inherits TextBox
    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)> _
    Private Shared Function SendMessage(ByVal hWnd As IntPtr, ByVal Msg As UInteger, ByVal wParam As IntPtr, ByVal lParam As String) As IntPtr
    End Function
    Shadows Property Multiline() As Boolean
        Get
            Return False
        End Get
        Set(ByVal value As Boolean)
            value = False
        End Set
    End Property

    Const _MarcaAguaDefault As String = "Escreva algo"
    Private _MarcaAguaText As String = _MarcaAguaDefault

    Public Property MarcaAguaText() As String
        Get
            Return _MarcaAguaText
        End Get
        Set(ByVal value As String)
            _MarcaAguaText = value
            Invalidate()
            Marca_Modificada()
        End Set
    End Property

    Const _ShowOnFocus As Boolean = True
    Private _show As Boolean = _ShowOnFocus
    Public Property TextOnFocus() As Boolean
        Get
            Return _show
        End Get
        Set(ByVal value As Boolean)
            _show = value
        End Set
    End Property
    Sub New()
        Me.Multiline = False
    End Sub

    Private Sub BoxMarcaAgua_HandleCreated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.HandleCreated
        Marca_Modificada()
    End Sub

    Private Sub Marca_Modificada()
        SendMessage(Me.Handle, &H1501, TextOnFocus, MarcaAguaText)
    End Sub
End Class