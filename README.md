# Dashboard Turma B — Painel Operacional

Dashboard alimentado diretamente pelo Google Drive via Google Apps Script.

## Arquivos do projeto
- `index.html` — Dashboard principal (sem dados embutidos)
- `Code.gs` — Script Google Apps a implementar
- `logo_OPE.png`, `logo_MNT.png`, `logo_report.png` — Logos

## Como configurar

### 1. Google Apps Script
1. Acesse [script.google.com](https://script.google.com)
2. Crie um novo projeto e cole o conteúdo de `Code.gs`
3. Ative o serviço "Drive API": clique em **Serviços → +** → selecione **Drive API**
4. Salve e clique em **Implantar → Nova implantação**
   - Tipo: **Aplicativo da web**
   - Executar como: **Usuário que instalou**
   - Quem tem acesso: **Qualquer pessoa**
5. Copie a URL gerada

### 2. index.html
- A URL do Apps Script já está configurada com a URL fornecida
- Se mudar o script, atualize a constante `APPS_SCRIPT_URL` no `index.html`

## Funcionalidades
- ✅ Alimentação 100% via Google Drive (sem dados offline)
- ✅ Nomes e meses dinâmicos conforme planilha
- ✅ Tooltip com OS reais por tipo (detalhes de cada OS)
- ✅ Botão copiar OS no formato `,=` para uso no Maximo
- ✅ Aderência com fórmulas diferenciadas: todos = (dias×7)/HH_mnt | individual = (dias×2,34)/HH_mnt
- ✅ Resumo mensal dinâmico conforme evolução da planilha
- ✅ Layout responsivo otimizado para celular
- ✅ Link "Origem dos dados" aponta para pasta do Drive

## IDs das planilhas configurados
- Horas: `1Sog9TjfIbq17VJyZKlD5Wj53JW4ZfZjm`
- Status: `1vsJdRzwO7XvAuQLTkbuQ2VSFUTzU0PeG`
