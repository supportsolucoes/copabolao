# 🖥️ Inventário de Máquinas

Ferramenta plug-and-play para levantar configuração e modelo de computadores Windows e consolidar tudo num painel central, usando **Firebase Realtime Database** — sem precisar de nenhum script de servidor.

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
   - Envia os dados automaticamente para o Firebase (se configurado).
3. **`dashboard.html`** — painel web que lista todas as máquinas que já enviaram dados, com busca, filtro por fabricante, detalhes de cada máquina e exportação para CSV. Lê os dados direto do Firebase, sem precisar de nenhum backend.

Se uma máquina rodar o coletor mais de uma vez, a entrada dela é **atualizada**, não duplicada (identificada pelo hostname).

## Configuração (passo a passo)

### 1. Crie o projeto e o Realtime Database no Firebase

1. Acesse [console.firebase.google.com](https://console.firebase.google.com) → **Criar projeto** (pode usar um projeto já existente também).
2. No menu lateral, vá em **Compilação → Realtime Database → Criar banco de dados**.
3. Escolha a localização e inicie em **modo bloqueado** (vamos colar as regras corretas no próximo passo).
4. Copie a **Database URL** mostrada no topo da página (algo como `https://seu-projeto-default-rtdb.firebaseio.com`).

### 2. Configure as regras de segurança

Na aba **Regras** do Realtime Database, substitua o conteúdo por:

```json
{
  "rules": {
    "chaves": {
      ".read": false,
      "$hostname": {
        ".write": "newData.val() === 'TROQUE_POR_UM_TOKEN_SECRETO'"
      }
    },
    "maquinas": {
      ".read": true,
      "$hostname": {
        ".write": "root.child('chaves').child($hostname).val() === 'TROQUE_POR_UM_TOKEN_SECRETO'"
      }
    }
  }
}
```

Troque `TROQUE_POR_UM_TOKEN_SECRETO` (nos **dois** lugares) por uma senha só sua, por exemplo `inv-8k2m9x`. Clique em **Publicar**.

**Por que assim:** o nó `/maquinas` fica com leitura aberta pro painel funcionar sem login, mas a escrita só é aceita se o token secreto já tiver sido gravado antes no nó `/chaves` — que ninguém consegue ler de volta (`.read: false`). Assim o token nunca fica exposto, mesmo com o banco público para leitura.

### 3. Configure o coletor

Abra `coletor.ps1` e cole a Database URL e o token nas primeiras linhas:

```powershell
$FIREBASE_DB_URL = 'https://seu-projeto-default-rtdb.firebaseio.com'
$FIREBASE_TOKEN  = 'inv-8k2m9x'
```

### 4. Configure o painel

Abra `dashboard.html` e cole a mesma Database URL (o token **não** entra aqui — leitura é aberta):

```js
const FIREBASE_DB_URL = 'https://seu-projeto-default-rtdb.firebaseio.com';
```

## Como distribuir para as máquinas

Depois de configurado, basta levar a pasta `inventario` (ou só os arquivos `coletor.bat` + `coletor.ps1`, com a URL e o token já preenchidos) para cada máquina e executar `coletor.bat`. Formas comuns de distribuir:

- Pen drive
- Pasta de rede compartilhada
- Enviar por e-mail/Teams e pedir para cada usuário rodar
- Script de logon de domínio (GPO), se quiser automatizar 100%

Não precisa instalar nada nem ter permissão de administrador — o PowerShell já vem com o Windows.

Para ver o inventário, é só abrir `dashboard.html` (2 cliques) — ele busca os dados do Firebase automaticamente e atualiza a cada clique em "Atualizar".

## Sem Firebase configurado?

Sem problema — se você deixar `FIREBASE_DB_URL` como está (`COLE_AQUI_A_URL_DO_SEU_FIREBASE_REALTIME_DATABASE`), o coletor só gera o relatório local em `Desktop\Inventario\` e não tenta enviar nada.
