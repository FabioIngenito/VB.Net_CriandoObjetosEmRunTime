Imports System.Windows.Forms

Module mdlPrincipal
    'Não dá para usar "WithEvents" junto com uma array!
    'You can't use "WithEvents" together with an array.
    'Public WithEvents cbnMolde As New ComboBox
    Public m_ComboBoxes() As ComboBox = {}

    Sub main()
        'Não use assim:
        'Do not use like this:
        'frmTESTE.Show()

        'Use assim:
        'Use like this:
        Application.Run(frmTESTE)
    End Sub

End Module
