# Automação de Extração e Organização de Dados

### Problema Resolvido

O processo de extração, organização e armazenamento de dados de uma base consolidada era realizado manualmente, demandando tempo e sendo propenso a erros. Com a automação implementada, esse fluxo foi otimizado, garantindo eficiência, padronização e menor intervenção manual.

### Solução Implementada

O código desenvolvido executa as seguintes etapas:

### 1️⃣ Identificação e Leitura do Arquivo Mais Recente

O script busca automaticamente o arquivo Excel mais recente em um diretório específico.

Isso assegura que sempre os dados mais atualizados sejam usados no processamento.

### 2️⃣ Filtragem de Dados

Os dados são filtrados conforme critérios específicos (exemplo: regionais, organização de vendas).

Essa segmentação permite que apenas as informações relevantes sejam processadas e armazenadas.

### 3️⃣ Criação de Estrutura de Pastas

O código gera automaticamente uma hierarquia de pastas organizada da seguinte forma:

DIRETÓRIO BASE / [Regional] / [Mês Atual] / [ID Regional - Nome Regional]

Isso garante uma organização padronizada e facilita a consulta futura dos arquivos.

### 4️⃣ Inserção de Dados em Arquivos Excel Preservando Formatação

Os dados filtrados são inseridos em um modelo de arquivo Excel pré-formatado.

O novo arquivo atualizado é salvo na pasta correspondente sem alterar a estrutura original do modelo.

### 5️⃣ Execução Automática do Fluxo Completo

A função executarAutomacao() dispara todo o processo automaticamente.

O tempo de execução é monitorado para avaliação de performance.

### 🚀 Benefícios da Automação

- ✅ Redução de Trabalho Manual → O processo de extração e organização ocorre automaticamente.
- ✅ Minimização de Erros → Reduz riscos de falhas humanas na manipulação dos arquivos.
- ✅ Padronização → Garante que os arquivos estejam sempre organizados da mesma maneira.
- ✅ Otimização de Tempo → Automação pode ser executada em segundo plano, permitindo que outras tarefas possam ser executadas ao mesmo tempo.
