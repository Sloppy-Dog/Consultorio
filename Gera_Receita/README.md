# Receitas

Gera um único arquivo Word, no papel timbrado, a partir das receitas em PDF (impressas pelo Dietbox).

## Pastas

| Pasta | Conteúdo |
|---|---|
| `input/` | Coloque aqui os PDFs das receitas (um PDF por receita). |
| `output/` | Saída: `Receitas.docx` e `classificacao.csv` (conferência doce/salgada). |
| `modelo/` | `papel_timbrado.docx` — o papel timbrado usado como base. |

## Como usar

1. Instale o Python 3.10 ou mais novo (no Windows, marque "Add Python to PATH").
2. Copie os PDFs das receitas para `input/`.
3. Rode:
   - **Windows:** dê dois cliques em `Gerar Receitas.bat`.
   - **Terminal:**
     ```bash
     pip install -r requirements.txt
     python gerar_receitas.py
     ```
4. Abra `output/Receitas.docx`.

## O que o documento contém

1. **Capa** com o título "Receitas".
2. **Índice** com as receitas salgadas e depois as doces.
3. **Receitas Salgadas**, em ordem alfabética, uma por página.
4. **Receitas Doces**, em ordem alfabética, uma por página.

Cada receita traz Ingredientes, Modo de preparo e Observações (quando houver).
Texto em Calibri 12, títulos em Calibri 14 (negrito). O cabeçalho do Dietbox
(nome, e-mail, CRN, data) e o aviso de confidencialidade são removidos, já que o
papel timbrado traz essas informações.

## Doce ou salgada?

A separação é automática, por palavras-chave no título e nos ingredientes
(ex.: *bolo, bala, açúcar, cacau* → doce; *sal, alho, cebola, carne* → salgada).
Depois de rodar, confira `output/classificacao.csv`. Se alguma receita cair no
grupo errado, adicione uma linha em `categorias.csv` e rode de novo:

```
Nome_do_arquivo.pdf;Doce
Outra_receita.pdf;Salgada
```

Receitas com o mesmo título aparecem uma vez só (o programa avisa no terminal).
