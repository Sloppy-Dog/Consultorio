#!/usr/bin/env python3
"""
Gerador de livro de receitas em papel timbrado.

Lê todas as receitas em PDF da pasta ``input/`` (formato de impressão do
Dietbox), extrai título, ingredientes, modo de preparo e observações, separa
as receitas em SALGADAS e DOCES, ordena cada grupo alfabeticamente e gera um
único arquivo Word em ``output/`` usando o papel timbrado de
``modelo/papel_timbrado.docx``.

Uso:
    python gerar_receitas.py
    python gerar_receitas.py --input input --output output/Receitas.docx

A classificação doce/salgada é automática (por palavras-chave). Para corrigir
alguma receita, edite ``categorias.csv`` (veja o README).
"""

from __future__ import annotations

import argparse
import copy
import csv
import re
import statistics
import sys
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path

import pdfplumber
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor

BASE = Path(__file__).resolve().parent

FONTE = "Calibri"
TAM_TEXTO = 12
TAM_TITULO = 14
TAM_CAPA = 40
COR_TITULO = RGBColor(0x9E, 0x5A, 0x63)   # rosé escuro, combina com o timbrado
COR_SUBTITULO = RGBColor(0x40, 0x40, 0x40)

SALGADAS = "Salgadas"
DOCES = "Doces"


# --------------------------------------------------------------------------
# Utilidades de texto
# --------------------------------------------------------------------------

def sem_acentos(texto: str) -> str:
    nfkd = unicodedata.normalize("NFKD", texto)
    return "".join(c for c in nfkd if not unicodedata.combining(c))


def chave(texto: str) -> str:
    """Chave para comparação/ordenação: minúsculas e sem acentos."""
    return re.sub(r"\s+", " ", sem_acentos(texto).casefold()).strip()


def limpar(texto: str) -> str:
    # Algumas fontes do Dietbox perdem a ligadura "ti" e viram \x00.
    texto = texto.replace("\x00", "ti")
    texto = texto.replace(" ", " ")
    return re.sub(r"[ \t]+", " ", texto).strip()


# --------------------------------------------------------------------------
# Leitura do PDF
# --------------------------------------------------------------------------

SECOES = {
    "ingredientes": "ingredientes",
    "modo de preparo": "preparo",
    "observacoes": "observacoes",
    "observacao": "observacoes",
    "dicas": "observacoes",
    "dica": "observacoes",
}

RE_DATA = re.compile(r"^data:\s*\d{1,2}/\d{1,2}/\d{2,4}$", re.I)
RE_RODAPE = re.compile(r"^este documento e de uso exclusivo", re.I)
RE_LIXO = [
    re.compile(r"^\d{1,2}/\d{1,2}/\d{2,4},?\s+\d{1,2}:\d{2}"),     # cabeçalho do navegador
    re.compile(r"^https?://", re.I),                               # URL do navegador
    re.compile(r"^\d+\s*/\s*\d+$"),                                # "1/1"
    re.compile(r"@"),                                              # e-mail
    re.compile(r"^crn\b", re.I),
    re.compile(r"^telefone", re.I),
    re.compile(r"^proibido\.?$", re.I),
]


@dataclass
class Linha:
    texto: str
    topo: float
    x0: float


@dataclass
class Receita:
    titulo: str
    arquivo: str
    ingredientes: list[str] = field(default_factory=list)
    preparo: list[str] = field(default_factory=list)
    observacoes: list[str] = field(default_factory=list)
    categoria: str = ""
    motivo: str = ""


def ler_linhas(pdf_path: Path) -> tuple[str | None, list[Linha]]:
    """Retorna (título do cabeçalho, linhas do corpo sem cabeçalho/rodapé)."""
    titulo = None
    corpo: list[Linha] = []
    deslocamento = 0.0  # para manter as posições crescentes entre páginas
    with pdfplumber.open(pdf_path) as pdf:
        for pagina in pdf.pages:
            linhas = [
                Linha(limpar(l["text"]), l["top"] + deslocamento, l["x0"])
                for l in pagina.extract_text_lines()
            ]
            linhas = [l for l in linhas if l.texto]
            deslocamento += pagina.height

            # Cabeçalho: tudo até a linha "Data: dd/mm/aaaa".
            idx_data = next(
                (i for i, l in enumerate(linhas) if RE_DATA.match(chave(l.texto))), None
            )
            if idx_data is not None:
                if titulo is None:
                    # O título fica entre "Receita" e "Data:".
                    idx_rec = max(
                        (i for i in range(idx_data) if chave(linhas[i].texto) == "receita"),
                        default=None,
                    )
                    if idx_rec is not None:
                        titulo = " ".join(l.texto for l in linhas[idx_rec + 1 : idx_data])
                    elif idx_data > 0:
                        titulo = linhas[idx_data - 1].texto
                linhas = linhas[idx_data + 1 :]

            # Rodapé: a partir do aviso de confidencialidade.
            idx_rod = next(
                (i for i, l in enumerate(linhas) if RE_RODAPE.match(chave(l.texto))), None
            )
            if idx_rod is not None:
                linhas = linhas[:idx_rod]

            corpo.extend(
                l for l in linhas if not any(r.search(chave(l.texto)) for r in RE_LIXO)
            )
    return titulo, corpo


def montar_paragrafos(linhas: list[Linha]) -> list[str]:
    """Junta linhas quebradas em parágrafos usando o espaço vertical entre elas."""
    if not linhas:
        return []
    gaps = [b.topo - a.topo for a, b in zip(linhas, linhas[1:]) if b.topo > a.topo]
    passo = statistics.median(gaps) if gaps else 12
    passo = min(passo, 14)  # entrelinha "normal" nunca é maior que isso
    inicio_novo = re.compile(r"^(\d+\s*[.)-]|[-•*]\s|[A-ZÀ-Ý][A-ZÀ-Ý ]{2,}:)")

    paragrafos: list[str] = []
    anterior: Linha | None = None
    for l in linhas:
        novo = (
            anterior is None
            or (l.topo - anterior.topo) > passo * 1.45
            or bool(inicio_novo.match(l.texto))
        )
        if novo:
            paragrafos.append(l.texto)
        else:
            paragrafos[-1] = juntar(paragrafos[-1], l.texto)
        anterior = l
    return paragrafos


def juntar(a: str, b: str) -> str:
    if a.endswith("-") and not a.endswith(" -") and b[:1].islower():
        return a[:-1] + b  # hifenização
    return f"{a} {b}"


def montar_ingredientes(linhas: list[Linha]) -> list[str]:
    itens: list[str] = []
    for l in linhas:
        # Linha que começa com minúscula é continuação do ingrediente anterior.
        if itens and l.texto[:1].islower():
            itens[-1] = juntar(itens[-1], l.texto)
        else:
            itens.append(l.texto)
    return itens


def ler_receita(pdf_path: Path) -> Receita:
    titulo, linhas = ler_linhas(pdf_path)
    if not titulo:
        titulo = re.sub(r"[_\s]+", " ", pdf_path.stem).strip()
        titulo = re.sub(r"\s\d+$", "", titulo)

    # Remove a repetição do título logo no início do corpo.
    if linhas and chave(linhas[0].texto) == chave(titulo):
        linhas = linhas[1:]

    blocos: dict[str, list[Linha]] = {"ingredientes": [], "preparo": [], "observacoes": []}
    atual = "preparo"
    for l in linhas:
        sec = SECOES.get(chave(l.texto).rstrip(":"))
        if sec:
            atual = sec
            continue
        blocos[atual].append(l)

    return Receita(
        titulo=titulo,
        arquivo=pdf_path.name,
        ingredientes=montar_ingredientes(blocos["ingredientes"]),
        preparo=montar_paragrafos(blocos["preparo"]),
        observacoes=montar_paragrafos(blocos["observacoes"]),
    )


# --------------------------------------------------------------------------
# Classificação doce x salgada
# --------------------------------------------------------------------------

# Palavras no TÍTULO têm peso maior que nos ingredientes.
DOCE_TITULO = """
bala balas bolo bolos brigadeiro brigadeiros beijinho brownie cookie cookies
biscoito doce doces docinho pudim mousse sorvete picole geleia compota torta-doce
cocada quindim pave manjar gelatina chocolate cacau mingau panqueca-doce
creme-de-avela sobremesa trufa trufas bombom bombons muffin cupcake churros
crepe-doce granola barrinha barra-de-cereal smoothie vitamina milkshake
""".split()

SALGADO_TITULO = """
salgado salgada salgadinho torta-salgada quiche empada empadao pastel coxinha
esfiha pizza lasanha escondidinho omelete frango carne peixe file bife hamburguer
almondega almondegas strogonoff estrogonofe risoto arroz feijao sopa caldo
creme-de-legumes salada farofa pure batata batatas mandioca aipim inhame
chuchu abobrinha berinjela legumes quibe kibe pao-de-queijo patê pate molho
tapioca crepioca wrap sanduiche hamburguer nhoque macarrao massa
""".split()

DOCE_ING = """
acucar mel melado adocante stevia estevia xilitol eritritol xylitol
leite-condensado chocolate cacau achocolatado doce-de-leite canela baunilha
essencia-de-baunilha banana morango uva suco-de-uva abacaxi manga maca
tamara uva-passa uvas-passas frutas-vermelhas coco leite-de-coco geleia
""".split()

SALGADO_ING = """
sal alho cebola cebolinha salsa salsinha carne frango peixe atum sardinha
presunto bacon linguica calabresa queijo-parmesao parmesao mussarela muçarela
oregano pimenta pimentao tomate cominho paprica curcuma acafrao tomilho
alecrim louro caldo shoyu molho-de-tomate
""".split()


def _tem(texto: str, termo: str) -> bool:
    termo = termo.replace("-", " ")
    return re.search(rf"\b{re.escape(termo)}\b", texto) is not None


def classificar(rec: Receita, manual: dict[str, str]) -> None:
    for k in (chave(rec.arquivo), chave(Path(rec.arquivo).stem), chave(rec.titulo)):
        if k in manual:
            rec.categoria, rec.motivo = manual[k], "categorias.csv"
            return

    titulo = chave(rec.titulo).replace("-", " ")
    ingred = chave(" | ".join(rec.ingredientes)).replace("-", " ")

    pontos_doce = 3 * sum(_tem(titulo, t) for t in DOCE_TITULO)
    pontos_sal = 3 * sum(_tem(titulo, t) for t in SALGADO_TITULO)
    pontos_doce += sum(_tem(ingred, t) for t in DOCE_ING)
    pontos_sal += sum(_tem(ingred, t) for t in SALGADO_ING)

    if pontos_doce > pontos_sal:
        rec.categoria = DOCES
    else:
        rec.categoria = SALGADAS  # empate vai para salgadas (revisar no relatório)
    rec.motivo = f"automático (doce={pontos_doce}, salgado={pontos_sal})"


def carregar_categorias(caminho: Path) -> dict[str, str]:
    manual: dict[str, str] = {}
    if not caminho.exists():
        return manual
    with caminho.open(encoding="utf-8-sig", newline="") as f:
        amostra = f.read(2048)
        f.seek(0)
        delim = ";" if amostra.count(";") >= amostra.count(",") else ","
        for linha in csv.reader(f, delimiter=delim):
            if len(linha) < 2 or not linha[0].strip() or linha[0].lstrip().startswith("#"):
                continue
            nome, cat = linha[0].strip(), chave(linha[1])
            if nome.casefold() in ("arquivo", "receita"):
                continue  # cabeçalho
            if cat.startswith("doce"):
                manual[chave(nome)] = DOCES
            elif cat.startswith("salgad"):
                manual[chave(nome)] = SALGADAS
    return manual


# --------------------------------------------------------------------------
# Geração do Word
# --------------------------------------------------------------------------

def fonte_run(run, tamanho: int, negrito=False, cor: RGBColor | None = None):
    run.font.name = FONTE
    run.font.size = Pt(tamanho)
    run.font.bold = negrito
    if cor is not None:
        run.font.color.rgb = cor
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = rpr.makeelement(qn("w:rFonts"), {})
        rpr.insert(0, rfonts)
    for att in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rfonts.set(qn(att), FONTE)


def configurar_estilos(doc: Document) -> None:
    normal = doc.styles["Normal"]
    normal.font.name = FONTE
    normal.font.size = Pt(TAM_TEXTO)
    rpr = normal.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = rpr.makeelement(qn("w:rFonts"), {})
        rpr.insert(0, rfonts)
    for att in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rfonts.set(qn(att), FONTE)
    for att in ("w:asciiTheme", "w:hAnsiTheme", "w:cstheme", "w:eastAsiaTheme"):
        if rfonts.get(qn(att)) is not None:
            del rfonts.attrib[qn(att)]
    pf = normal.paragraph_format
    pf.space_after = Pt(6)
    pf.line_spacing = 1.15

    # Estilos de título com nível de estrutura (aparecem no Painel de Navegação).
    for nome, nivel, cor, alinhamento in (
        ("Receitas - Seção", 0, COR_TITULO, WD_ALIGN_PARAGRAPH.CENTER),
        ("Receitas - Título", 1, COR_TITULO, WD_ALIGN_PARAGRAPH.CENTER),
        ("Receitas - Subtítulo", 2, COR_SUBTITULO, WD_ALIGN_PARAGRAPH.LEFT),
    ):
        try:
            st = doc.styles[nome]
        except KeyError:
            st = doc.styles.add_style(nome, WD_STYLE_TYPE.PARAGRAPH)
        st.base_style = normal
        st.font.name = FONTE
        st.font.size = Pt(TAM_TITULO)
        st.font.bold = True
        st.font.color.rgb = cor
        st.paragraph_format.alignment = alinhamento
        st.paragraph_format.keep_with_next = True
        st.paragraph_format.space_before = Pt(12 if nivel == 2 else 0)
        st.paragraph_format.space_after = Pt(6 if nivel == 2 else 12)
        ppr = st.element.get_or_add_pPr()
        ol = ppr.find(qn("w:outlineLvl"))
        if ol is None:
            ol = ppr.makeelement(qn("w:outlineLvl"), {})
            ppr.append(ol)
        ol.set(qn("w:val"), str(nivel))


def abrir_modelo(modelo: Path) -> Document:
    """Abre o papel timbrado e apaga o conteúdo do corpo (mantém cabeçalho/margens)."""
    doc = Document(str(modelo))
    body = doc.element.body
    for el in list(body):
        if el.tag != qn("w:sectPr"):
            body.remove(el)
    return doc


def paragrafo(doc, texto="", estilo=None, tamanho=TAM_TEXTO, negrito=False,
              cor=None, alinhamento=None, espaco_depois=None):
    p = doc.add_paragraph(style=estilo)
    if texto:
        r = p.add_run(texto)
        fonte_run(r, tamanho, negrito, cor)
    if alinhamento is not None:
        p.alignment = alinhamento
    if espaco_depois is not None:
        p.paragraph_format.space_after = Pt(espaco_depois)
    return p


def quebra_pagina(doc):
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)


def titulo(doc, texto, estilo):
    p = doc.add_paragraph(style=estilo)
    fonte_run(p.add_run(texto), TAM_TITULO, True,
              COR_SUBTITULO if estilo.endswith("Subtítulo") else COR_TITULO)
    return p


def item_lista(doc, texto):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.left_indent = Pt(18)
    pf.first_line_indent = Pt(-12)
    pf.space_after = Pt(2)
    fonte_run(p.add_run("•  " + texto), TAM_TEXTO)


def gerar_word(grupos: dict[str, list[Receita]], modelo: Path, saida: Path) -> None:
    doc = abrir_modelo(modelo)
    configurar_estilos(doc)

    # ---- Capa
    for _ in range(8):
        paragrafo(doc)
    paragrafo(doc, "Receitas", tamanho=TAM_CAPA, negrito=True, cor=COR_TITULO,
              alinhamento=WD_ALIGN_PARAGRAPH.CENTER)
    total = sum(len(v) for v in grupos.values())
    paragrafo(doc, f"{total} receitas", tamanho=TAM_TITULO, cor=COR_SUBTITULO,
              alinhamento=WD_ALIGN_PARAGRAPH.CENTER)

    # ---- Índice
    quebra_pagina(doc)
    titulo(doc, "Índice", "Receitas - Seção")
    for cat in (SALGADAS, DOCES):
        if not grupos.get(cat):
            continue
        titulo(doc, f"Receitas {cat}", "Receitas - Subtítulo")
        for r in grupos[cat]:
            item_lista(doc, r.titulo)

    # ---- Receitas
    for cat in (SALGADAS, DOCES):
        for i, r in enumerate(grupos.get(cat, [])):
            quebra_pagina(doc)
            if i == 0:
                titulo(doc, f"RECEITAS {cat.upper()}", "Receitas - Seção")
            titulo(doc, r.titulo, "Receitas - Título")
            if r.ingredientes:
                titulo(doc, "Ingredientes", "Receitas - Subtítulo")
                for ing in r.ingredientes:
                    item_lista(doc, ing)
            if r.preparo:
                titulo(doc, "Modo de preparo", "Receitas - Subtítulo")
                for par in r.preparo:
                    paragrafo(doc, par, alinhamento=WD_ALIGN_PARAGRAPH.JUSTIFY)
            if r.observacoes:
                titulo(doc, "Observações", "Receitas - Subtítulo")
                for par in r.observacoes:
                    paragrafo(doc, par, alinhamento=WD_ALIGN_PARAGRAPH.JUSTIFY)

    # Remove o parágrafo vazio que sobra antes da 1ª quebra de página da capa? Não
    # é necessário; apenas garante que o sectPr continue sendo o último elemento.
    body = doc.element.body
    sect = body.find(qn("w:sectPr"))
    body.remove(sect)
    body.append(sect)

    saida.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(saida))


# --------------------------------------------------------------------------
# Programa principal
# --------------------------------------------------------------------------

def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="Gera o livro de receitas em Word.")
    ap.add_argument("--input", default=BASE / "input", type=Path)
    ap.add_argument("--output", default=BASE / "output" / "Receitas.docx", type=Path)
    ap.add_argument("--modelo", default=BASE / "modelo" / "papel_timbrado.docx", type=Path)
    ap.add_argument("--categorias", default=BASE / "categorias.csv", type=Path)
    args = ap.parse_args(argv)

    pdfs = sorted(p for p in args.input.iterdir() if p.suffix.lower() == ".pdf")
    if not pdfs:
        print(f"Nenhum PDF encontrado em {args.input}")
        return 1

    manual = carregar_categorias(args.categorias)
    receitas: list[Receita] = []
    vistos: set[str] = set()
    for pdf in pdfs:
        try:
            rec = ler_receita(pdf)
        except Exception as e:  # noqa: BLE001
            print(f"  [ERRO] {pdf.name}: {e}")
            continue
        if chave(rec.titulo) in vistos:
            print(f"  [AVISO] receita repetida ignorada: {rec.titulo} ({pdf.name})")
            continue
        vistos.add(chave(rec.titulo))
        if not rec.ingredientes or not rec.preparo:
            print(f"  [AVISO] {pdf.name}: ingredientes ou modo de preparo não encontrados")
        classificar(rec, manual)
        receitas.append(rec)

    grupos = {SALGADAS: [], DOCES: []}
    for r in receitas:
        grupos[r.categoria].append(r)
    for lista in grupos.values():
        lista.sort(key=lambda r: chave(r.titulo))

    gerar_word(grupos, args.modelo, args.output)

    # Relatório para conferência da classificação.
    relatorio = args.output.with_name("classificacao.csv")
    with relatorio.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["arquivo", "categoria", "titulo", "criterio"])
        for cat in (SALGADAS, DOCES):
            for r in grupos[cat]:
                w.writerow([r.arquivo, cat, r.titulo, r.motivo])

    for cat in (SALGADAS, DOCES):
        print(f"\nReceitas {cat} ({len(grupos[cat])}):")
        for r in grupos[cat]:
            print(f"  - {r.titulo}   [{r.motivo}]")
    print(f"\nArquivo gerado: {args.output}")
    print(f"Conferência da classificação: {relatorio}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
