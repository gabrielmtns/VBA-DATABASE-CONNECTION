# Classes para facilitar a conexão e a interação de formulários com o banco de dados.

Diferenças entre o arquivo clsConnection e clsConnectionComplete:
    clsConnectionComplete fornece suporte diretamente as listViews, textBox e comboBox. Implementação com listBox ainda será lançada.
    Se no seu projeto não for necessário o uso de formulários utilize a clsConnection.

Funcionalidades básicas:

    • Conexão com o banco de dados.
    • Criação rapida de recordsets.
    • Execução simples de comandos SQL no geral.
    • Fornece suporte para a execução de comandos SQL com ADODB.Parameters.
    • Metodo para pegar os cabeçalhos da consulta (fieldToArray).
    • Metodo para gerar arquivo .xlsx diretamente do recordset.
    • Preenche formulários com informações do banco. 

Como utilizar:

    Importe o arquivo mFunctionToConnect,  clsConnection ou clsConnectionComplete para o seu projeto.
    
Exemplo de consumo da classe:
    
![instructionBasicConnection](https://user-images.githubusercontent.com/35817813/56043337-6dd5ca00-5d13-11e9-9909-62b42c38a64c.PNG)
