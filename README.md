# ManipuladorXML

Este projeto tem como objetivo manipular arquivos XML de notas fiscais eletrônicas (NFe), realizando edições automáticas em campos específicos conforme configurações definidas em um arquivo JSON de constantes.

## Estrutura do Projeto

- `manipuladorXML.py`: Script principal responsável pela leitura, edição e salvamento dos arquivos XML.
- `constantes.json`: Arquivo de configuração contendo os parâmetros de edição, caminhos de pastas, dados do emitente, produto, impostos e flags de controle.

## Como Funciona

1. **Configuração**: O arquivo `constantes.json` define:

   - Caminhos das pastas de origem e edição dos XMLs.
   - Dados do emitente, produto e impostos a serem inseridos ou alterados.
   - Flags para ativar/desativar cada tipo de alteração (emitente, produtos, impostos, data, referência de NFe).
   - Nova data para ser aplicada nos XMLs.

2. **Execução**:
   - O script lê as configurações do JSON.
   - Percorre os arquivos XML na pasta de origem.
   - Realiza as alterações conforme as flags e dados definidos.
   - Salva os arquivos editados na pasta de edição (pode ser a mesma de origem).
   - Pode renomear os arquivos processados, se configurado.

## Como Usar

1. Edite o arquivo `constantes.json` conforme necessário.
2. Execute o script `manipuladorXML.py`:

```bash
python manipuladorXML.py
```

3. Os arquivos XML serão processados conforme as configurações.

## Observações

- Certifique-se de ter permissão de leitura e escrita nas pastas configuradas.
- O script foi desenvolvido para facilitar ajustes em lote em arquivos XML de NFe, útil para empresas que precisam padronizar ou corrigir informações em grande quantidade de documentos.

## Autor

Desenvolvido por [Seu Nome].
