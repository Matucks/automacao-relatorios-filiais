# Automação de Relatórios para Filiais

Este repositório contém um script em Python para automatizar o processamento e envio de relatórios de inventário e faturamento para diferentes filiais. Ele foi desenvolvido para garantir eficiência, organização e padronização no gerenciamento de dados.

## Funcionalidades

- **Processamento de Planilhas**:

  - Identifica automaticamente os arquivos na pasta de entrada.
  - Remove colunas desnecessárias para faturados.
  - Gera arquivos Excel organizados por filial.
  - Cria um arquivo consolidado para filiais agrupadas.

- **Envio Automatizado de E-mails**:

  - Integração com SMTP para envio dos relatórios processados.
  - Suporte a listas de destinatários e cópias configuráveis por filial.

## Estrutura do Projeto

- **Pastas de Entrada e Saída**:

  - `C:/relatorios/entrada_faturados`: Contém os arquivos de faturados.
  - `C:/relatorios/entrada_inventario`: Contém os arquivos de inventário.
  - `C:/relatorios/saida_faturados`: Contém os arquivos processados de faturados.
  - `C:/relatorios/saida_inventario`: Contém os arquivos processados de inventário.

- **Mapeamento de Filiais**:

  - As filiais são identificadas por códigos genéricos, como `Filial_01`, `Filial_02`, etc.

- **Configuração de Colunas**:

  - Algumas colunas, como "Localidade", "Cliente Final", e outras, são removidas durante o processamento para faturados.

## Como Usar

### Requisitos

- Python 3.8 ou superior.
- Bibliotecas necessárias:
  - `pandas`
  - `xlsxwriter`
  - `openpyxl`

### Instalação

1. Clone este repositório:
   ```bash
   git clone https://github.com/seu-usuario/automacao-relatorios-filiais.git
   ```
2. Navegue até o diretório do projeto:
   ```bash
   cd automacao-relatorios-filiais
   ```
3. Instale as dependências:
   ```bash
   pip install pandas openpyxl
   ```

### Configuração

1. **Caminhos de Entrada e Saída**:
   Certifique-se de que as pastas configuradas no script existem:

   ```python
   PASTA_ENTRADA_FATURADOS = 'C:/relatorios/entrada_faturados'
   PASTA_ENTRADA_INVENTARIO = 'C:/relatorios/entrada_inventario'
   PASTA_SAIDA_FATURADOS = 'C:/relatorios/saida_faturados/'
   PASTA_SAIDA_INVENTARIO = 'C:/relatorios/saida_inventario/'
   ```

2. **Servidor SMTP**:
   Configure o servidor de e-mail para envio dos relatórios:

   ```python
   SMTP_SERVER = 'smtp.seuservidor.com'
   SMTP_PORT = 587
   EMAIL_USER = 'seu.email@dominio.com'
   EMAIL_PASS = 'sua_senha'
   ```

3. **Mapeamento de Filiais**:
   Ajuste o dicionário `email_map` para incluir os destinatários das filiais:

   ```python
   email_map = {
       "Filial_01": {"to": ["email1@dominio.com"], "cc": ["email2@dominio.com"]},
       "Filial_02": {"to": ["email3@dominio.com"], "cc": []}
   }
   ```

### Execução

1. Certifique-se de que os arquivos de entrada estejam nas pastas configuradas.
2. Execute o script:
   ```bash
   python script.py
   ```
3. Os arquivos processados serão salvos nas pastas de saída e os e-mails serão enviados automaticamente (se configurados).

## Estrutura dos Relatórios

- **Relatórios por Filial**:

  - Arquivos separados para cada filial com os dados processados.
  - Incluem informações relevantes ao inventário ou faturamento.

- **Relatório Consolidado**:

  - Para filiais agrupadas (exemplo: `Filial_01` e `Filial_02`), gera-se um arquivo único.

## Contribuições

Contribuições são bem-vindas! Caso encontre problemas ou tenha sugestões, sinta-se à vontade para abrir uma issue ou criar um pull request.

## Licença

Este projeto está licenciado sob a Licença MIT. Consulte o arquivo `LICENSE` para mais informações.

