# Consultório

## Gerador de Plano Alimentar (Dietbox → Word → PDF)

O programa lê o PDF **"Plano alimentar ativo"** exportado do Dietbox e gera
automaticamente o plano no **modelo Word personalizado**, já com a capa, e
uma **cópia em PDF**.

### Regras aplicadas

| Dietbox               | Word        |
|-----------------------|-------------|
| Refeição principal    | OPÇÃO 1     |
| Substituição 1        | OPÇÃO 2     |
| Substituição 2        | OPÇÃO 3     |
| Substituição 3        | OPÇÃO 4     |

- Cada refeição aparece como **NOME DA REFEIÇÃO – HHhMM** (ex.: `ALMOÇO – 12h30`).
- A **Obs** da refeição principal (ex.: *Composição do Almoço…*) vira uma
  **célula mesclada** abaixo de todas as opções.
- A refeição com **uma única opção** (ex.: Ceia) ocupa as 4 colunas.
- A palavra **OU** é destacada em negrito.
- O cabeçalho (PACIENTE / Data / REFEIÇÕES / OPÇÕES) se repete em cada página,
  e nenhuma refeição é cortada entre duas páginas.

### Instalação (uma única vez)

1. Instale o Python 3.10 ou mais recente em https://www.python.org
   (marque **"Add Python to PATH"** durante a instalação).
2. Abra o terminal na pasta do projeto e rode:

```bash
pip install -r requirements.txt
```

O PDF é gerado pelo **Microsoft Word**. Se o Word não estiver instalado,
o programa usa o **LibreOffice**, se ele estiver disponível.

### Como usar

- **Duplo clique** em `Gerar Plano Alimentar.bat`. Uma janela abre para
  você escolher um ou mais PDFs do Dietbox.
- Ou pelo terminal:

```bash
python plano_alimentar/gerar_plano.py "Plano alimentar ativo de Fulana.pdf"
```

Opções:

- `--saida PASTA`: salva os arquivos em outra pasta.
- `--nao-abrir`: não abre a pasta de saída no final.
- `--sem-pdf`: gera somente o Word.
- `--modelo ARQ.docx`: usa outro modelo de Word.
- É possível informar uma **pasta**. Nesse caso, o programa processa todos os PDFs
  da pasta que tenham "ativo" no nome.

Os arquivos são salvos na pasta **`plano_alimentar/output`**, que é aberta
automaticamente no final, com o nome
`Plano Alimentar - <Paciente> - <dd-mm-aaaa>.docx` e `.pdf`.

### Modelo

O modelo fica em `plano_alimentar/modelo/Modelo_Plano_Alimentar.docx` e contém
somente a capa, o cabeçalho e o rodapé. Para mudar a capa ou o logotipo, basta
editar esse arquivo no Word. As tabelas são criadas pelo programa.
