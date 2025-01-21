Descrição do Software de Automação de Análise de Contas

Visão Geral
Este software foi desenvolvido para automatizar a análise de contas do Google, otimizando um processo que, manualmente, levaria um tempo significativamente maior e estaria sujeito a erros humanos. A automação utiliza Python, juntamente com as bibliotecas Selenium e openpyxl, para interagir com o sistema, extrair informações e atualizar uma base de dados em tempo real.
O objetivo principal é analisar aproximadamente 4 mil contas antes de 2025, garantindo precisão e economia de tempo.

Funcionalidades

Leitura de Base de Dados:
O software utiliza uma planilha de controle chamada gmails.xlsx, localizada no diretório C:\Analisar.
A planilha contém os e-mails das contas que precisam ser analisadas.
Análise Automática de Contas:
Para cada e-mail, o software realiza login no sistema, onde diversas informações da conta estão disponíveis.
O HTML da página é analisado para identificar os elementos necessários à verificação.

Atualização de Status:

Após a análise, o software atualiza a planilha com o status de cada conta.
Caso a conta atenda a todos os critérios estabelecidos, o sistema a marca automaticamente como Data Reviewed (DR), indicando que está pronta para ser deletada.
Relatórios e Registro:
O resultado da análise é registrado diretamente na planilha gmails.xlsx, permitindo rastreabilidade e auditoria.
Interface Simples:
Um executável (.exe) foi criado para facilitar o uso do software, permitindo que qualquer membro da equipe o utilize sem necessidade de configurar o ambiente Python.
Além disso, um manual foi elaborado para orientar os analistas sobre como operar o software de forma eficiente.

Impacto e Benefícios


Produtividade:

A automação consegue analisar 1 conta por minuto, enquanto o processo manual levaria entre 10 e 15 minutos por conta.
Com a automação, cada analista pode processar até 480 contas por dia em um turno de 8 horas.
Considerando uma equipe de 10 analistas, é possível analisar cerca de 3 mil contas por dia, mesmo considerando eventuais erros.
Economia de Tempo:
O processo manual levaria meses para ser concluído, enquanto a automação permite finalizar o trabalho em apenas alguns dias.
Precisão:

A lógica programada reduz significativamente a chance de erros humanos, garantindo que os critérios sejam aplicados de forma consistente.

Tecnologias Utilizadas


Linguagem: Python
Bibliotecas Principais:

Selenium: Para interação com o sistema e análise de elementos HTML.
openpyxl: Para manipulação da planilha de controle gmails.xlsx.

Como Usar (Somente funciona no ambiente interno da empresa)

1 - Certifique-se de que a planilha gmails.xlsx esteja localizada no diretório C:\Analisar.
2- Coloque seu email no arquivo config.json
3- Execute o arquivo .exe (disponível no repositório).
4- O software realizará automaticamente a análise das contas e atualizará o status na planilha.

Observações

Este software foi projetado para funcionar somente no ambiente da empresa, mas pode ser ajustado para atender a novos critérios ou alterações em outros ambientes. Ele é uma solução prática para grandes volumes de dados que exigem análise criteriosa e rápida.

Autor: Ivan Bragante
Linguagem: Python
Bibliotecas: Selenium, openpyxl
