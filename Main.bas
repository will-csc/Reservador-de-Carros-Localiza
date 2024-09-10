Attribute VB_Name = "Main"
Dim Navegador As New Selenium.ChromeDriver
Sub ReservaLocaliza()
    Navegador.Start
    Navegador.Get "https://empresas-prd.localiza.com/passo1?empresas=Empresas"
    
    On Error GoTo Erro
    '-------------- LOGANDO NO SITE -----------------------
    Set Usuario = Navegador.FindElementById("login")
    Set Password = Navegador.FindElementById("senha")
    Set Button = Navegador.FindElementById("btnEfetuarLogin")
    
    Usuario.SendKeys "[Digite seu usuario]"
    Password.SendKeys "[Digite sua senha]"
    Button.Click
    Navegador.Wait 1500
    
    '-------------- DADOS DA RESERVA 1° (LOCAL,DIA,HORARIO) -----------------------
    Set Agência1 = Navegador.FindElementById("mat-input-2")
    Set Data1 = Navegador.FindElementById("dataRetirada")
    Set Horario1 = Navegador.FindElementById("mat-input-1")
    
    Agência1.SendKeys Range("D4").Value2
    Navegador.Wait 1000
    Set Agência1 = Navegador.FindElementByCss(".places-list").FindElementByCss("[data-testid='place-item']")
    Agência1.Click
    Data1.SendKeys Range("A4").Value
    Navegador.SendKeys Navegador.Keys.Escape
    Horario1.Click
    Horario1.SendKeys Format(Range("B4").Value, "hhmm")
    
    '-------------- DADOS DA RESERVA 2° (LOCAL,DIA,HORARIO) -----------------------
    Set Data2 = Navegador.FindElementById("dataDevolucao")
    Set Horario2 = Navegador.FindElementById("mat-input-5")
    Set Button = Navegador.FindElementById("btnContinuar")
    
    Data2.SendKeys Range("A7").Value
    Navegador.SendKeys Navegador.Keys.Escape
    Horario2.Click
    Horario2.Clear
    Navegador.Wait 5000
    Horario2.SendKeys Format(Range("B7").Value, "hhmm")
    
    Button.Click
    
    '------------ ESCOLHA DO CARRO ------------------------
    Navegador.Wait 6500
    Set Carros = Navegador.FindElementsByCss(".new-group-car-content")
    
    i = 0
    For Each Carro In Carros
        Set Esgostado = Carro.FindElementByTag("button")
        If Esgostado.Text = Empty Then
            Set GRUPO = Carro.FindElementByTag("h3")
            Set Motor = Carro.FindElementByTag("p")
            Set Valor = Carro.FindElementsByTag("span").Item(2)
            Valores = Replace(Replace(Valor.Text, "R$ ", ""), ".", ",")
            If InStr(1, GRUPO.Text, "GRUPO FS") > 0 Or InStr(1, GRUPO.Text, "GRUPO FX") > 0 Then
                Set Button = Carro.FindElementByTag("ds-new-button")
                Button.Click
                If Range("D7").Value = "C6" Or Range("D7").Value = "C8" Then
                    Set InputCPF = Navegador.FindElementById("cpf")
                    InputCPF.SendKeys Range("C15").Value
                End If
                Exit For
            ElseIf InStr(1, Motor.Text, "1.6", vbTextCompare) > 0 And Valores < 270 Then
                Set Button = Carro.FindElementByTag("ds-new-button")
                Button.Click
                If Range("D7").Value = "C6" Or Range("D7").Value = "C8" Then
                    Navegador.Wait 1500
                    Set InputCPF = Navegador.FindElementById("cpf")
                    InputCPF.SendKeys Range("C15").Value
                End If
                Exit For
            End If
        End If
        i = i + 1
    Next Carro
    
    '------------ DEFINIÇÃO VALORES ------------------------
    Set Ajuste = Navegador.FindElementById("idOferta-0")
    Ajuste.Click
    Navegador.Wait 3500
    
    Set ProteçãoVidros = Navegador.FindElementsById("labelReservaDescricaoOpcional-0").Item(2)
    ProteçãoVidros.Click
    Navegador.Wait 15000
    
    Set LimpezaG = Navegador.FindElementsById("labelReservaDescricaoOpcional-1").Item(2)
    LimpezaG.Click
    Navegador.Wait 15000
    
    Set CompensaçãoC = Navegador.FindElementsById("labelReservaValorOpcional-2").Item(2)
    CompensaçãoC.Click
    Navegador.Wait 15000
    
    Set Button = Navegador.FindElementById("btn-continuar-grupo-carro")
    Button.Click
    
    '------------ DADOS DO CONDUTOR PART I ------------------------
    Set Nacionalidade = Navegador.FindElementById("mat-select-value-3")
    Set DocumentoTipo = Navegador.FindElementById("select")
    Set N°Documento = Navegador.FindElementById("numeroDocumento")
    Set NomeCond = Navegador.FindElementById("nomeCondutor")
    Set Email = Navegador.FindElementById("emailCondutor")
    Set ddd = Navegador.FindElementById("ddd")
    Set Telefone = Navegador.FindElementById("telefone")
    Set CentrodeCusto = Navegador.FindElementById("centroCusto")
    Set Setor = Navegador.FindElementById("setorEmpresa")
    
    Nacionalidade.Click
    Navegador.Wait 1500
    Application.SendKeys "BRASIL"
    Navegador.Wait 1500
    Navegador.SendKeys Navegador.Keys.Enter
    Navegador.Wait 5000
    
    DocumentoTipo.Click
    Navegador.Wait 1000
    Set DocumentoTipo = Navegador.FindElementByTag("mat-option")
    DocumentoTipo.Click
    
    N°Documento.Click
    N°Documento.SendKeys Range("C15").Value
    
    NomeCond.Click
    NomeCond.SendKeys Range("A17").Value
    
    Email.Click
    Email.SendKeys Range("D22").Value
    
    ddd.Click
    ddd.SendKeys "11"
    
    Telefone.Click
    Telefone.SendKeys Range("C22").Value
    
    CentrodeCusto.SendKeys "Comercial"
    Setor.SendKeys "Bufalo Ind. Com. Produtos Quimico LTDA"
    
    '------------ DADOS DO CONDUTOR PART II ------------------------
    Set DadosCondutor = Navegador.FindElementByCss(".requisitante__linha")
    Set ddd = DadosCondutor.FindElementById("ddd")
    Set Telefone = DadosCondutor.FindElementById("telefone")
    Set Email = DadosCondutor.FindElementById("emailRequisitante")
    
    ddd.Click
    ddd.SendKeys "11"
    
    Telefone.Click
    Telefone.SendKeys "[Numero para apoio]"
    
    Email.Click
    Email.SendKeys "[email para apoio]"
    
    '------------ DADOS DO CONDUTOR PART III ------------------------
    Set DadosCondutor = Navegador.FindElementByCss(".forma-pagamento")
    Set Pagamento = DadosCondutor.FindElementById("select")
    
    Pagamento.Click
    Navegador.Wait 1500
    Set Pagamento = Navegador.FindElementById("select-panel").FindElementByTag("mat-option")
    Pagamento.Click
    
    Exit Sub
    
Erro:
    Navegador.Wait 1000
    Limite = Limite + 1
    If Limite > 30 Then Stop
    Resume
End Sub
