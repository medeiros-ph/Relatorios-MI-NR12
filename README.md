# Relatórios Automatizados para Extração de Dados de Planilhas Excel

Este projeto visa a automação do processo de extração, consolidação e validação de dados de múltiplos arquivos Excel. Ele gera relatórios em formato CSV, com foco na confiabilidade e resiliência a erros durante a execução. A solução foi construída para facilitar a análise de informações críticas, reduzindo o tempo e o esforço manual.

---

## Tecnologias Utilizadas

- **Python**: Linguagem principal para manipulação de dados.
- **Pandas**: Para manipulação e criação de DataFrames (estruturas tabulares) e exportação para CSV.
- **OpenPyXL**: Para leitura de arquivos Excel (`.xlsx` e `.xlsm`), extraindo dados e validando conteúdo.
- **OS e Glob**: Navegação e varredura de diretórios e subdiretórios.
- **CSV**: Exportação dos relatórios gerados.

---

## Instalação

1. Certifique-se de ter o Python instalado (versão >= 3.8).
2. Instale as dependências do projeto com o comando:

   ```bash
   pip install pandas openpyxl
   ```

---

## Relatórios Gerados

### **1. TabelaMaquina.csv**
- **Objetivo:** Consolidar informações sobre máquinas, incluindo parâmetros como nome, modelo, localização, e outros atributos relevantes.
- **Planilhas Analisadas:**
  - `"1- Capa"`
- **Colunas Extraídas:**  
  - Exemplos: `AnoFabricacao`, `Nome`, `Fabricante`, `Peso`, entre outros.
- **Diferenciais:**  
  - Incremental: Salva progressivamente os dados no CSV.
  - Resiliente: Registra erros de processamento por arquivo.

---

### **2. TabelaCheckList.csv**
- **Objetivo:** Consolidar informações sobre checklists de conformidade e status.
- **Planilhas Analisadas:**
  - `"2- Check List 01"`
  - `"3- Check List 02"`
- **Colunas Extraídas:**  
  - `Referência NR-12`, `Status`
- **Diferenciais:**
  - Validação de ranges na coluna `D` e associação com status na coluna `CN`.
  - Respeita múltiplos ranges em uma única célula.

---

### **3. TabelaAcoes.csv**
- **Objetivo:** Registrar descrições de ações e informações complementares.
- **Planilhas Analisadas:**
  - `"4A- Não Conform."`
- **Colunas Extraídas:**  
  - `descricao`, `descricao_en`
- **Diferenciais:**  
  - Estrutura flexível para adicionar colunas como `maquina`, `nc`, `Projeto`, entre outras no futuro.

---

### **4. RelatorioAnaliseDeRisco.csv**
- **Objetivo:** Mapear análises de risco com informações de nomes, locais, e FEs.
- **Planilhas Analisadas:**
  - `"1- Capa"`
- **Colunas Extraídas:**  
  - `Nome`, `Local`, `FE`
- **Diferenciais:**  
  - Ignora células vazias para evitar dados inconsistentes.
  - Relatório consolidado com o nome do arquivo fonte.

---

## Recursos Adicionais

1. **Incremental Processing:**
   - Cada relatório é salvo progressivamente para evitar perda de dados em caso de interrupções.

2. **Relatórios de Erros:**
   - Um CSV separado para registrar arquivos que apresentaram problemas durante o processamento.

3. **Varredura de Subdiretórios:**
   - O código navega por todos os diretórios e subdiretórios de uma pasta raiz para encontrar os arquivos Excel.

4. **Tratamento de Erros:**
   - Arquivos bloqueados ou com permissões inadequadas são ajustados automaticamente.
   - Arquivos com falhas de processamento são ignorados, mas documentados.

---

## Como Utilizar

1. Ajuste o caminho da pasta principal no código:
   ```python
   pasta_principal = r"C:\Users\Raphael Medeiros\Desktop\07 - Relatórios"
   ```
2. Configure o diretório de saída para os relatórios CSV:
   ```python
   caminho_saida_csv = r"C:\Users\Raphael Medeiros\Desktop\07 - Relatórios\outputs"
   ```
3. Execute os scripts individualmente para testar cada relatório.
4. Integre os scripts em um pipeline unificado após validação.

---

## Arquivos Gerados

1. **Relatórios CSV:**
   - `TabelaMaquina.csv`
   - `TabelaCheckList.csv`
   - `TabelaAcoes.csv`
   - `RelatorioAnaliseDeRisco.csv`

2. **Relatórios de Erros:**
   - `ErrosTabelaMaquina.csv`
   - `ErrosTabelaCheckList.csv`
   - `ErrosTabelaAcoes.csv`
   - `ErrosAnaliseDeRisco.csv`

---

## Roadmap

- **Integração:** Unificar todos os scripts em um único fluxo automatizado.
- **Validação:** Adicionar testes para verificar inconsistências nos dados.
- **Documentação:** Expandir explicações e exemplos de uso.

---

Esta solução foi projetada para ser escalável e confiável, atendendo às necessidades de processamento massivo e automatizado de dados em ambientes corporativos.

