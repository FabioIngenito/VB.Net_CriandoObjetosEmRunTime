Imports System.Windows.Forms
Imports System.Drawing
Imports Project2

''' <summary>
''' Criando Componentes no código fonte sem usar o "Project / Components" (sem referenciar no projeto). 
''' Exemplos de como usar "Licenses.Add", "VBControlExtender" e o "ObjectEvent". Cria em 'run-time' uma Label e uma ImageCombo.
''' 
''' Creating components in source code without using the "Project/Components" (without reference to the project). 
''' Examples of how to use "Licenses. Add", "VBControlExtender" and "ObjectEvent. Creates in 'run-time' a Label and an ImageCombo.
''' </summary>
Public Class frmTeste

    'Objetos da "frmTeste.vb"
    'Objects of "frmTeste.vb"
    Private lblCriar As New Label
    Private cBoxCriar As New BoxMarcaAgua
    Private lblColunas As New Label
    Private WithEvents cBoxColunas As New BoxMarcaAgua
    Private lblExplicacao As New Label
    Private WithEvents btnCriar As New Button
    Private WithEvents btnRetirarTudo As New Button

    'Objeto da "frmP3"
    'Object of "frmP3"
    Private txtTeste As New TextBox

    Dim clsC As New clsBox
    Dim frmP2 As New Form
    Dim frmP3 As New Form

    'http://msdn.microsoft.com/en-us/library/5185a713(v=vs.90).aspx
    ' Upgrade Notes
    ' --------------------------------------------------------------------------------
    ' When a Visual Basic 6.0 project is upgraded to Visual Basic 2008, any instances of the VBControlExtender object are ignored.
    '     A COM interop wrapper is created for each ActiveX control; property, method, and events are mapped to their equivalents.
    '     Where there are no equivalents, upgrade warnings are added to the code.
    ' Private WithEvents cboTeste2 As VBControlExtender

    'Notas de atualização
    '------------------------------------------------- -------------------------------
    'Quando um projeto Visual Basic 6.0 é atualizado para Visual Basic 2008, todas as instâncias do objeto VBControlExtender são ignorados.
    '     A invólucro de interoperabilidade COM é criada para cada controle ActiveX, propriedade, método e eventos são mapeados para seus equivalentes.
    '     Onde não há equivalentes, avisos de atualização são adicionados ao código.
    'Private WithEvents cboTeste2 Como VBControlExtender

    Private Sub frmTESTE_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Cria um objetos em tempo de execução neste projeto.
        'Creates a runtime objects in this project.
        MontaForm1()
        'Cria um único objeto ComboBox para servir de base para criação dos outros objetos em forma de array
        'Creates a single ComboBox object to serve as a basis for creation of other objects in the form of array
        extra2()
    End Sub

    ''' <summary>
    ''' Monta o formulário principal que gerencia o formulário secundário
    ''' Assemble the main form that manages the secondary form
    ''' 
    ''' Cria os OCXs no "waBoxMarcaAgua" (neste form)
    ''' Creates the OCXs in the "waBoxMarcaAgua" (in this form)
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub MontaForm1()

        Try
            'Esta forma abaixo era como se fazia no VB6:
            'This form below was like they did in VB6:
            'lblCriar = Me.Controls.Add(System.Windows.Forms.Label)
            lblCriar.AutoSize = False
            lblCriar.Text = "Criar mais:"
            lblCriar.Width = 55
            lblCriar.Top = 10
            lblCriar.Left = 10
            lblCriar.Visible = True
            lblCriar.UseMnemonic = True
            lblCriar.ImageAlign = ContentAlignment.TopLeft
            lblCriar.BorderStyle = System.Windows.Forms.BorderStyle.None
            lblCriar.ImageList = ImageList1
            lblCriar.ImageIndex = 1
            lblCriar.Size = New Size(lblCriar.PreferredWidth, lblCriar.PreferredHeight)
            ' "Add" é para adicionar o projeto ao form
            ' "Add" is to add the project to the form ... Dããã ... This comment is ridiculous in English...
            Me.Controls.Add(lblCriar)

            cBoxCriar.Width = 70
            cBoxCriar.Top = 10
            cBoxCriar.Left = 70
            cBoxCriar.MaxLength = 3
            'Coloca "Handler" no Controle cBox:
            'Puts "Handler" in the cBox Control:
            AddHandler cBoxCriar.KeyPress, AddressOf cBox_AtualizaTexto
            Me.Controls.Add(cBoxCriar)

            lblColunas.Text = "Colunas:"
            lblColunas.Width = 50
            lblColunas.Top = 10
            lblColunas.Left = 150
            Me.Controls.Add(lblColunas)

            cBoxColunas.Width = 70
            cBoxColunas.Top = 10
            cBoxColunas.Left = 200
            cBoxColunas.MaxLength = 2
            Me.Controls.Add(cBoxColunas)

            btnCriar.Text = "&Criar"
            btnCriar.Width = 50
            btnCriar.Top = 10
            btnCriar.Left = 280
            btnCriar.Visible = True
            Me.Controls.Add(btnCriar)

            btnRetirarTudo.Text = "&Retirar Tudo"
            btnRetirarTudo.Width = 100
            btnRetirarTudo.Top = 10
            btnRetirarTudo.Left = 340
            btnRetirarTudo.Visible = False
            Me.Controls.Add(btnRetirarTudo)

            lblExplicacao.AutoSize = False
            lblExplicacao.Text = "Ao retirar eu mando um 'm_ComboBoxes(0).Visible = False', pois esta é a COMBO que serve de modelo para criação das outras combos."
            lblExplicacao.Width = 400
            lblExplicacao.Top = 10
            lblExplicacao.Left = 450
            lblExplicacao.Height = 26
            lblExplicacao.Visible = True
            Me.Controls.Add(lblExplicacao)

            Me.FormBorderStyle = Windows.Forms.FormBorderStyle.FixedSingle
            Me.WindowState = FormWindowState.Normal
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.TopMost = True
            Me.Width = 850
            Me.Height = 370
            Me.StartPosition = FormStartPosition.CenterScreen
            Me.BringToFront()

        Catch ex As Exception
            MessageBox.Show(ex.Message)

        Finally

        End Try

    End Sub

    ''' <summary>
    ''' Cria a ComboBox Principal (número 0) no form secundário (frmteste)
    ''' Creates the Main ComboBox (number 0) in the form (frmteste)
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub extra2()

        Try
            frmP2 = New frmFormVazio

            'Não precisa mais de "LICENÇA" ... o VB6 precisava
            'You don't need more than "license" ... the VB6 needed
            'Licenses.Add("MSCOMCTLLIB.ImageComboCtl.2")
            'cboTeste = Me.Controls.Add("MSCOMCTLLIB.ImageComboCtl.2", "cboTeste")
            ReDim Preserve m_ComboBoxes(0)
            m_ComboBoxes(0) = New ComboBox

            m_ComboBoxes(0).Name = "cboTeste"
            m_ComboBoxes(0).Width = 100
            m_ComboBoxes(0).Top = 10
            m_ComboBoxes(0).Left = 10
            m_ComboBoxes(0).Visible = True
            m_ComboBoxes(0).DropDownStyle = ComboBoxStyle.DropDownList

            frmP2.Controls.Add(m_ComboBoxes(0))

            frmP2.MaximizeBox = False
            frmP2.MinimizeBox = False
            frmP2.ShowIcon = False
            frmP2.ShowInTaskbar = False
            frmP2.FormBorderStyle = Windows.Forms.FormBorderStyle.FixedSingle
            frmP2.Width = 850
            frmP2.Height = 800
            frmP2.WindowState = FormWindowState.Maximized
            frmP2.Visible = True
            frmP2.Show()

            'Código abaixo somente para mostrar que não precisamos de um Form "físico" para criar um form.
            'Code below only to show that we do not need a Form "physical" to create a form.
            txtTeste = New TextBox
            txtTeste.Text = "EXEMPLO"
            txtTeste.Width = 500
            txtTeste.Top = 10
            txtTeste.Left = 10

            frmP3 = New Form
            frmP3.Controls.Add(txtTeste)
            frmP3.Location = New Point(100, 300)
            frmP3.MaximizeBox = False
            frmP3.MinimizeBox = False
            frmP3.ShowIcon = False
            frmP3.ShowInTaskbar = False
            frmP3.FormBorderStyle = Windows.Forms.FormBorderStyle.FixedSingle
            frmP3.Width = 525
            frmP3.Height = 200
            frmP3.Opacity = 900
            frmP3.StartPosition = FormStartPosition.Manual
            frmP3.WindowState = FormWindowState.Normal
            frmP3.Visible = True
            frmP3.Text = "Form Criado do Nada!"
            frmP3.Show()

        Catch ex As Exception
            MessageBox.Show(ex.Message)

        Finally

        End Try

    End Sub

    ''' <summary>
    ''' Cria e gerencia as ComboBox
    ''' Creates and manages the ComboBox
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnCriar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCriar.Click

        If Not cBoxCriar.Text = "" AndAlso Not cBoxColunas.Text = "" Then
            btnCriar.Enabled = False
            btnRetirarTudo.Visible = True
            m_ComboBoxes(0).Visible = True
            txtTeste.Text = ""

            clsC.CriaCombos(frmP2, m_ComboBoxes(0), Val(cBoxCriar.Text), Val(cBoxColunas.Text))
            clsC.PreencheCombos(frmP2)
        ElseIf cBoxCriar.Text = "" Then
            cBoxCriar.Text = "10"
            txtTeste.Text = "Por favor, digite o número de Combos que deseja Criar."
            cBoxCriar.SelectAll()

            'ATENÇÃO: Código abaixo NÃO funciona direito, pois perde o foco do formulário:
            'WARNING: code below won't work right because it loses focus of the form:
            'MessageBox.Show("Por favor, digite o número de Combos que deseja Criar.", "Qtde Combos", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            cBoxCriar.Focus()
        ElseIf cBoxColunas.Text = "" Then
            cBoxColunas.Text = "5"
            txtTeste.Text = "Por favor, digite o número de Colunas que deseja Criar."
            cBoxColunas.SelectAll()

            'ATENÇÃO: Código abaixo NÃO funciona direito, pois perde o foco do formulário:
            'WARNING: code below won't work right because it loses focus of the form:
            'MessageBox.Show("Por favor, digite o número de Colunas que deseja Criar.", "Qtde Colunas", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            cBoxColunas.Focus()
        End If

    End Sub

    ''' <summary>
    ''' Retira TODAS as ComboBox, menos a Número ZERO
    ''' Removes all ComboBox, less the number ZERO
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnRetirarTudo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRetirarTudo.Click
        btnCriar.Enabled = True
        btnRetirarTudo.Visible = False
        clsC.RetiraTodasCombos(frmP2)
    End Sub

    ''' <summary>
    ''' Handler de Atulização de Texto
    ''' Text update handler
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>Serve para inibir a digitação de letras, permite a digitação somente de números.
    ''' Serves to inhibit typing letters, allows to type numbers only.</remarks>
    Private Sub cBox_AtualizaTexto(ByVal sender As Object, ByVal e As KeyPressEventArgs)
        If Not Char.IsNumber(e.KeyChar) And Not Convert.ToInt32(e.KeyChar) = Keys.Back Then e.Handled = True
    End Sub

    ''' <summary>
    ''' Handler NORMAL
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>Serve para inibir a digitação de letras, permite a digitação somente de números.
    ''' Serves to inhibit typing letters, allows to type numbers only.</remarks>
    Private Sub cBoxColunas_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cBoxColunas.KeyPress
        If Not Char.IsNumber(e.KeyChar) And Not Convert.ToInt32(e.KeyChar) = Keys.Back Then e.Handled = True
    End Sub

End Class