# 🖥️ Inventário de Máquinas

Ferramenta plug-and-play para levantar configuração e modelo de computadores Windows e consolidar tudo num painel central.

## O que ele coleta

- **Identificação**: fabricante, modelo, número de série, patrimônio, hostname, usuário logado, domínio/grupo de trabalho
- **Hardware**: CPU, memória RAM (com módulos), discos (modelo, tamanho, tipo SSD/HDD), placa de vídeo
- **Sistema**: versão do Windows, build, arquitetura
- **Rede**: IP, MAC, gateway de cada adaptador
- **Softwares instalados**: nome, versão e fabricante de cada programa

## Como funciona

1. **`coletor.bat`** — o usuário só dá dois cliques (não precisa instalar nada, não precisa ser administrador).
2. O PowerShell (`coletor.ps1`) roda por trás, coleta tudo e:
   - Salva um relatório local em `Desktop\Inventario\` (HTML bonito + JSON) e já abre o relatório no navegador.
   - Envia os dados automaticamente para uma Planilha Google central (se configurado).
3. **`dashboard.html`** — painel web que lista todas as máquinas que já enviaram dados, com busca, filtro por fabricante, detalhes de cada máquina e exportação para CSV.

Se uma máquina rodar o coletor mais de uma vez, a linha dela na planilha é **atualizada**, não duplicada (identificada pelo hostname).

## Configuração (passo a passo)

### 1. Crie a planilha central

1. Crie uma nova planilha no Google Sheets.
2. Copie o **ID da planilha** (a parte da URL entre `/d/` e `/edit`).

### 2. Configure o Apps Script

1. Acesse [script.google.com](https://script.google.com) → **Novo projeto**.
2. Cole o conteúdo de `apps-script-inventario.gs`.
3. Substitua `COLE_AQUI_O_ID_DA_SUA_PLANILHA` pelo ID copiado no passo 1.
4. Rode a função `inicializarPlanilha` uma vez (menu Executar) para criar o cabeçalho — na primeira execução o Google vai pedir para autorizar o script.
5. Implante como **App da Web**:
   - Executar como: **Eu**
   - Quem pode acessar: **Qualquer pessoa**
6. Copie a URL gerada (termina em `/exec`).

### 3. Configure o coletor

Abra `coletor.ps1` e cole a URL na primeira linha de configuração:

```powershell
$APPS_SCRIPT_URL = 'https://script.google.com/macros/s/SEU_ID/exec'
```

### 4. Configure o painel

Abra `dashboard.html` e cole a mesma URL:

```js
const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/SEU_ID/exec';
```

Publique `dashboard.html` onde quiser (GitHub Pages, uma pasta compartilhada, etc.) para acompanhar o inventário em tempo real.

## Como distribuir para as máquinas

Depois de configurado, basta levar a pasta `inventario` (ou só os arquivos `coletor.bat` + `coletor.ps1`) para cada máquina e executar `coletor.bat`. Formas comuns de distribuir:

- Pen drive
- Pasta de rede compartilhada
- Enviar por e-mail/Teams e pedir para cada usuário rodar
- Script de logon de domínio (GPO), se quiser automatizar 100%

Não precisa instalar nada nem ter permissão de administrador — o PowerShell já vem com o Windows.

## Sem planilha central?

Sem problema — se você deixar `APPS_SCRIPT_URL` como está (`COLE_AQUI_A_URL_DO_SEU_APPS_SCRIPT`), o coletor só gera o relatório local em `Desktop\Inventario\` e não tenta enviar nada.
