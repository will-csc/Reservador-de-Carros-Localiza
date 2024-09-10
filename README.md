<h2><b>CRIADOR DE RESERVAS LOCALIZA | VBA</b></h2>

Olá! 👋
Ao longo da minha carreira precisei adaptar muito dos meus trabalhos para ganhar tempo e consequentemente agilidade em outros trabalhos; com isso, reserva de carros, exigiam
um intervalo muito grande. Ao longo dos meus estudos, essa tarefa foi automatizada graças aos conhecimentos em VBA,Selenium e HTML.<br>
Nesse repositório, os senhores(a) vão encontrar um script VBA que automatiza o processo de fazer uma reserva no site da Localiza Empresas. 
Ele utiliza o Selenium Chrome Driver para interagir com o site, realizando o login, inserindo os dados da reserva, selecionando o veículo e preenchendo as informações do condutor.

<h1><b>Pré-requisitos</b></h1>
-Selenium Basic instalado para VBA.<br>
-ChromeDriver compatível com a versão do Chrome instalada no seu sistema.<br>
-Microsoft Excel para armazenar os dados utilizados pelo script.<br>
-Como Funciona<br>

<h1><b><b>1. Login no Site da Localiza<b></b></h1>
O script faz o login no site da Localiza Empresas utilizando as credenciais inseridas no código. Substitua os textos de marcador [Digite seu usuario] e [Digite sua senha] pelas suas credenciais reais.

<h1><b>2. Detalhes da Reserva</b></h1>
O script coleta os dados da reserva, incluindo:

<b>Data de Retirada (Data1):</b> Data em que o veículo será retirado.<br>
<b>Horário de Retirada (Horario1):</b> Horário da retirada.<br>
<b>Data de Devolução (Data2):</b> Data em que o veículo será devolvido.<br>
<b>Horário de Devolução (Horario2):</b> Horário da devolução.<br>
<b>Agência (Agência1):</b> Local de retirada do veículo.<br><br>

Esses valores são preenchidos dinamicamente usando dados de uma planilha Excel. As células relevantes são:

A4 para a Data de Retirada
B4 para o Horário de Retirada
A7 para a Data de Devolução
B7 para o Horário de Devolução
D4 para a Agência

<h1><b>3. Seleção do Veículo<b></b></h1>
O script seleciona um veículo com base nos seguintes critérios:

Grupo de Veículo: Verifica se o veículo pertence ao Grupo FS ou FX.
Tipo de Motor: Verifica se o veículo possui motor 1.6.
Preço: Caso o preço do veículo seja inferior a R$270, o carro será selecionado.

<h1><b><b>4. Informações do Condutor<b></b></h1>
O script preenche os dados do condutor, incluindo:

- Nacionalidade<br>
- Tipo de Documento e Número do Documento<br>
- Nome do Condutor<br>
- Email<br>
- Número de Telefone<br>
- Centro de Custo<br>
- Setor/Empresa<br><br>

Os dados do condutor são preenchidos utilizando as seguintes células da planilha Excel:

C15 para o CPF (número do documento)
A17 para o nome do condutor
D22 para o email do condutor
C22 para o telefone do condutor

<h1><b>5. Etapas Finais</b></h1>
Após selecionar o veículo e preencher os dados do condutor, o script prossegue para a seleção do método de pagamento.

Tratamento de Erros
O script inclui um tratamento básico de erros com um mecanismo de repetição. Se ocorrer um erro, o processo tentará novamente até 30 vezes antes de parar.

Como Usar o Script
Instale o Selenium Basic para VBA e o ChromeDriver correspondente à versão do Chrome instalada no seu computador.
Abra o arquivo Excel onde os detalhes da reserva estão armazenados.
Modifique o script com suas credenciais de login e certifique-se de que os intervalos de células correspondem à sua planilha.
Execute a macro ReservaLocaliza para automatizar o processo de reserva.
Melhorias Futuras
Adicionar mensagens de erro dinâmicas para facilitar a solução de problemas.
Habilitar a leitura de credenciais de login de uma fonte segura, como um arquivo criptografado.
Implementar seleção dinâmica de métodos de pagamento com base nas opções disponíveis.

## Principal Linguagem Utilizada
<img align="center" alt="Will-VBA" height="30" width="40" src="https://raw.githubusercontent.com/vscode-icons/vscode-icons/master/icons/file_type_vba.svg">
