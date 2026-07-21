# ⚽ Bolão Copa 2026 🇧🇷

Aplicação web para o bolão da Copa do Mundo 2026 — Seleção Brasileira, Grupo C.

## Como funciona

- Participante escolhe o jogo do Brasil
- Adiciona quantos palpites quiser (carrinho)
- Paga **R$ 5,00 por palpite** via Pix
- Envia o comprovante no grupo do WhatsApp com os códigos gerados

## Jogos

| Rodada | Jogo | Data | Hora |
|--------|------|------|------|
| 1 | 🇧🇷 Brasil × Marrocos 🇲🇦 | 13/06/2026 | 19:00 |
| 2 | 🇧🇷 Brasil × Haiti 🇭🇹 | 19/06/2026 | 21:30 |
| 3 | 🏴󠁧󠁢󠁳󠁣󠁴󠁿 Escócia × Brasil 🇧🇷 | 24/06/2026 | 19:00 |

## Configuração

Antes de publicar, edite as constantes no início do `index.html`:

```js
const APPS_SCRIPT_URL = 'COLE_AQUI_A_URL_DO_SEU_APPS_SCRIPT';
const PIX_CHAVE       = '(11) 99999-9999'; // sua chave Pix real
```

## Google Apps Script

O arquivo `codigo-apps-script.gs` contém o script que salva os palpites no Google Sheets.

**Passos:**
1. Acesse [script.google.com](https://script.google.com) → Novo projeto
2. Cole o conteúdo de `codigo-apps-script.gs`
3. Substitua `COLE_AQUI_O_ID_DA_SUA_PLANILHA` pelo ID da sua planilha
4. Implante como **App da Web** (acesso: Qualquer pessoa)
5. Cole a URL gerada no `index.html`

## Hospedagem com GitHub Pages

1. Vá em **Settings → Pages**
2. Source: `main` branch → `/root`
3. Acesse em: `https://supportsolucoes.github.io/copabolao`

---

# 🖥️ Inventário de Máquinas

Este repositório também contém, na pasta [`inventario/`](inventario/), uma ferramenta plug-and-play separada para levantar configuração e modelo de computadores Windows (hardware, sistema, rede e softwares instalados) e consolidar tudo num painel central. Veja [`inventario/README.md`](inventario/README.md) para o passo a passo completo.
