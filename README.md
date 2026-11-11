# Livro de Ocorrências de Turno

Site para tablet que substitui a planilha do livro de ocorrências de turno da manutenção.
Roda em Google Planilhas + Apps Script: a planilha é a base de dados e o Apps Script serve o site e envia o relatório por e-mail aos líderes ao fim de cada turno.

## Arquivos

- `Codigo.gs` — servidor (escala, lançamentos, pendências, envio de e-mail, área do administrador)
- `docs/index.html` — página do GitHub Pages que só redireciona para o link do sistema no Google (se o link da implantação mudar, atualize aqui)
- `Index.html` — tela do tablet (aberta fora do Google, roda em modo demonstração; senha do administrador na demonstração: `demo123`)

## Instalação

1. Crie uma Planilha Google em branco e abra **Extensões → Apps Script**.
2. Cole o `Codigo.gs` no arquivo `Código.gs`. Crie um arquivo HTML chamado `Index` e cole o `Index.html`.
3. Execute a função `setup` (ou recarregue a planilha e use o menu **Livro de Ocorrências → 1. Configuração inicial**). Ela cria as abas, liga o envio automático e pede a senha do administrador.
4. Preencha a aba **Cadastro** e os e-mails dos líderes em **Config**.
5. **Implantar → Nova implantação → App da Web** (Executar como: Eu; Acesso: Qualquer pessoa).

## Atualização

Substitua o conteúdo dos dois arquivos, execute `setup` de novo e publique em
**Implantar → Gerenciar implantações → lápis → Nova versão**.

## Turnos

| Turno | Horário | Escala |
|---|---|---|
| A | 07:00–15:35 | 5x1 |
| B | 15:20–23:40 | 5x1 |
| C | 23:40–07:00 | 5x1 |
| 6x1 (código `D`) | 07:00–16:48 | 6x1 |
| Folguista | 2 dias B, 2 dias C, 2 folgas | — |

Os horários de cada turno e do envio do e-mail são ajustados em **Administrador → Configurações**.

Dados de funcionários, crachás e lançamentos ficam só na planilha e não fazem parte deste repositório.
