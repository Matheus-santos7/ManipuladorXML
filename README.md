# ManipuladorXML

> Ferramenta para automação de manipulação, edição e renomeação de arquivos XML de Notas Fiscais Eletrônicas (NFe), com base em regras de negócio e configurações customizáveis.

---

## Sumário

- [Visão Geral](#visão-geral)
- [Estrutura do Projeto](#estrutura-do-projeto)
- [Configuração](#configuração)
- [Fluxo de Processamento e Lógica de Renomeação](#fluxo-de-processamento-e-lógica-de-renomeação)
- [Exemplo de Fluxo Completo](#exemplo-de-fluxo-completo)
- [Como Usar](#como-usar)
- [Observações e Recomendações](#observações-e-recomendações)

---

## Visão Geral

O ManipuladorXML foi desenvolvido para automatizar tarefas recorrentes de ajuste, padronização e correção em arquivos XML de NFe. Ele permite:

- Alterar campos do emitente, produtos, impostos, datas e referências de NFe.
- Renomear arquivos XML conforme regras de negócio baseadas em CFOP, natureza da operação e relações entre notas.
- Processar grandes volumes de arquivos de forma rápida e segura.

## Estrutura do Projeto

- `manipuladorXML.py`: Script principal com toda a lógica de processamento, edição e renomeação dos XMLs.
- `constantes.json`: Arquivo de configuração central, onde são definidos caminhos, dados de alteração e flags de controle.

## Configuração

O arquivo `constantes.json` permite customizar:

- **Caminhos**: Defina as pastas de origem e edição dos XMLs.
- **Dados do emitente, produto e impostos**: Informações que podem ser inseridas ou alteradas nos XMLs.
- **Flags de alteração**: Ative ou desative alterações específicas (`emitente`, `produtos`, `impostos`, `data`, `refNFe`).
- **Nova data**: Data a ser aplicada nos campos de emissão e saída dos XMLs.

Exemplo de configuração:

```json
{
   "caminhos": {
      "pasta_origem": "/caminho/para/origem",
      "pasta_edicao": "/caminho/para/edicao"
   },
   "configuracao_execucao": {
      "processar_e_renomear": true,
      "editar_arquivos": true
   },
   ...
}
```

## Fluxo de Processamento e Lógica de Renomeação

O script executa duas etapas principais:

1. **Organização e Renomeação** (`processar_arquivos`):

   - Analisa cada XML, identifica o tipo de operação (Remessa, Retorno, Venda, Devolução, etc.) com base no CFOP, natureza da operação e referências.
   - Renomeia os arquivos conforme regras de negócio, facilitando o rastreio do fluxo de mercadorias.

2. **Edição dos Arquivos** (`editar_arquivos`):
   - Altera campos internos dos XMLs conforme as flags e dados definidos no `constantes.json`.
   - Garante que as alterações sejam consistentes e rastreáveis.

## Exemplo de Fluxo Completo

| Etapa do Fluxo       | Exemplo de Arquivo | Lógica no Script (`processar_arquivos`)                                                                           | Resultado Esperado                         |
| -------------------- | ------------------ | ----------------------------------------------------------------------------------------------------------------- | ------------------------------------------ |
| 1. Remessa           | 4163.xml           | Identifica o CFOP (5949) e a ausência de uma refNFe para nomear como "Remessa".                                   | 4163 - Remessa.xml                         |
| 2. Retorno Simbólico | 4252.xml           | Identifica o CFOP (1949), a natOp "Retorno Simbolico..." e a refNFe para a remessa 4163.                          | 4252 - Retorno da remessa 4163.xml         |
| 3. Venda             | 4253.xml           | Identifica o CFOP de venda (5105) e nomeia como "Venda", mantendo a referência à nota de retorno 4252.            | 4253 - Venda.xml                           |
| 4. Devolução         | 4297.xml           | Identifica o CFOP de devolução (1201), a refNFe para a venda 4253 e o texto DEVOLUTION_devolution.                | 4297 - Devolucao da venda 4253.xml         |
| 5. Remessa Simbólica | 4298.xml           | Identifica o CFOP (5949) e a presença de uma refNFe (para a devolução 4297) para nomear como "Remessa simbólica". | 4298 - Remessa simbólica da venda 4297.xml |
| 6. Retorno Simbólico | 4429.xml           | A mesma lógica da etapa 2 é aplicada, identificando a refNFe para a remessa simbólica 4298.                       | 4429 - Retorno da remessa 4298.xml         |
| 7. Venda             | 4430.xml           | A mesma lógica da etapa 3 é aplicada, identificando a refNFe para o retorno 4429.                                 | 4430 - Venda.xml                           |

---

## Como Usar

1. Ajuste o arquivo `constantes.json` conforme sua necessidade.
2. Coloque os arquivos XML na pasta de origem definida.
3. Execute o script:

```bash
python manipuladorXML.py
```

4. Os arquivos serão processados, editados e renomeados conforme as regras e configurações.

## Observações e Recomendações

- Certifique-se de ter permissão de leitura e escrita nas pastas configuradas.
- O script não insere quebras de linha entre as tags dos XMLs processados, garantindo compatibilidade com sistemas que exigem arquivos "em linha única".
- Ideal para empresas que precisam padronizar, corrigir ou rastrear grandes volumes de documentos fiscais eletrônicos.
- Faça sempre um backup dos arquivos antes de processar em lote.

---

## Autores

Desenvolvido por [Claudio Santos, Matheus Rodrigues, Matheus Gonçalves, Rodrigo Siqueira].
