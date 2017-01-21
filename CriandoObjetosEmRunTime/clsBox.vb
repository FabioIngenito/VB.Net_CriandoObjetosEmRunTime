Imports System.Windows.Forms

Public Class clsBox

    ''' <summary>
    ''' Cria Combos em sequência de acordo com os parâmetros passados
    ''' Create Combos in sequence according to the parameters passed
    ''' </summary>
    ''' <param name="frm">Form onde serão criadas - Form where you will be created</param>
    ''' <param name="cbn">Molde de ComboBox para copiar nas outras - ComboBox template to copy other</param>
    ''' <param name="NroCombos">Número de ComboBox a serem criadas - Number of ComboBox to be created</param>
    ''' <param name="NroColunas">Número de Colunas onde serão criadas as Combobox (somente pra controle) - Number of columns where the Combobox will be created (only to control)</param>
    ''' <remarks>Esta Sub serve para criar as ComboBox de acordo com o usuário desta Sub necessitar - This Sub serves to make ComboBox according to the user of this Sub you need</remarks>
    Public Sub CriaCombos(ByVal frm As Form, ByVal cbn As ComboBox, ByVal NroCombos As Long, ByVal NroColunas As Long)
        Dim lngX As Long
        Dim lngY As Long
        Dim lngAcrescenta As Long
        Dim intColuna As Integer

        lngAcrescenta = cbn.Height
        intColuna = NroColunas

        ReDim Preserve m_ComboBoxes(NroCombos)

        For lngX = 1 To NroCombos
            'Parte da "física" das Combos (desenho das Combos no form)
            'Part of the "Physics" of Combos (design of Combos in the form)
            m_ComboBoxes(lngX) = New ComboBox

            m_ComboBoxes(lngX).Tag = lngX
            m_ComboBoxes(lngX).Visible = True
            m_ComboBoxes(lngX).Width = m_ComboBoxes(0).Width

            PosicionaCombo(lngX, intColuna)

            lngY = lngY + lngAcrescenta

            If lngX >= intColuna Then intColuna = intColuna + NroColunas

            'Coloca "Handler" no Controle:
            'Puts "Handler" in control:
            AddHandler m_ComboBoxes(lngX).TextUpdate, AddressOf m_ComboBoxes_AtualizaTexto

            'Adicionando a Combo no formulário:
            'Adding the Combo in the form:
            frm.Controls.Add(m_ComboBoxes(lngX))
        Next

    End Sub

    ''' <summary>
    ''' Faz o controle das ComboBox na tela
    ''' Makes the control of the ComboBox on the screen
    ''' </summary>
    ''' <param name="X">Número de ComboBox - Number of ComboBox</param>
    ''' <param name="cols">Número de Colunas - Number of columns</param>
    ''' <remarks>Acerta a posição das ComboBox na tela controlando as Colunas
    ''' Sets the position of the ComboBox on the screen controlling the columns</remarks>
    Private Sub PosicionaCombo(ByVal X As Long, ByVal cols As Integer)

        m_ComboBoxes(X).Top = m_ComboBoxes(X - 1).Top
        m_ComboBoxes(X).Left = (m_ComboBoxes(0).Width + m_ComboBoxes(X - 1).Left) + 10

        If X >= cols Then
            m_ComboBoxes(X).Top = (m_ComboBoxes(0).Height + m_ComboBoxes(X - 1).Top) + 10
            m_ComboBoxes(X).Left = m_ComboBoxes(0).Left
        End If

    End Sub

    ''' <summary>
    ''' Retira as Combos de uma determinada Form
    ''' Removes the Combos of a particular Form
    ''' </summary>
    ''' <param name="frm">Form onde serão retiradas as ComboBox - Form where you will be taken the ComboBox</param>
    ''' <remarks>Retira TODAS as ComboBox da Tela - Removes all the ComboBox of the screen</remarks>
    Public Sub RetiraTodasCombos(ByVal frm As Form)
        Dim X As Integer

        If m_ComboBoxes(0).Items.Count > 0 Then m_ComboBoxes(0).Items.Clear()
        m_ComboBoxes(0).Visible = False

        For X = 1 To m_ComboBoxes.Count - 1
            frm.Controls.Remove(m_ComboBoxes(X))
        Next

        'Não funciona direito o código abaixo?!?
        'Not working right the code below?!?
        'Dim MyControl As Control

        'For Each MyControl In frm.Controls
        '    If TypeOf MyControl Is ComboBox Then frm.Controls.Remove(MyControl)
        'Next
    End Sub

    ''' <summary>
    ''' Preenche as ComboBox
    ''' Fills the ComboBox
    ''' </summary>
    ''' <param name="frm">Formulário onde serão preenchidas as ComboBox - Form where the ComboBox</param>
    ''' <remarks>Para que não fique vazias, preenche com duas linhas cada ComboBox - To empty, fills with two lines each ComboBox</remarks>
    Public Sub PreencheCombos(ByVal frm As Form)
        Dim lngConta As Long

        For lngConta = 0 To m_ComboBoxes.Count - 1
            m_ComboBoxes(lngConta).Items.Add("Item: " & lngConta)
            m_ComboBoxes(lngConta).Items.Add("Ou outra opção!")
            m_ComboBoxes(lngConta).SelectedIndex = 0
        Next

    End Sub

    ''' <summary>
    ''' Exemplo de Handler na ComboBox
    ''' Handler example in the ComboBox
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>Mostra como funciona um Evento na Array de ComboBox - Shows how to work an event in the Array of ComboBox</remarks>
    Private Sub m_ComboBoxes_AtualizaTexto(ByVal sender As Object, ByVal e As EventArgs)

        MessageBox.Show("O texto foi alterado em: " & sender.GetType.FullName.ToString())

    End Sub

End Class
