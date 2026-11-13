# -*- coding: utf-8 -*-
"""
Gerador de Plano Alimentar personalizado
=========================================

Lê o PDF "Plano alimentar ativo" exportado do Dietbox e gera:
  1) o plano no modelo Word personalizado (capa + tabelas por refeição);
  2) a cópia em PDF desse Word.

Regras de conversão:
  - Refeição principal  -> OPÇÃO 1
  - Substituição 1      -> OPÇÃO 2
  - Substituição 2      -> OPÇÃO 3
  - Substituição 3      -> OPÇÃO 4
  - "Obs:" da refeição principal (ex.: "Composição do Almoço...") vira uma
    célula mesclada logo abaixo de todas as opções da refeição.
  - Refeição com apenas uma opção (ex.: Ceia) ocupa as 4 colunas mescladas.
  - A palavra "OU" (maiúscula) é destacada em negrito.

Uso:
  python gerar_plano.py                          -> abre janela para escolher o(s) PDF(s)
  python gerar_plano.py "Plano alimentar ativo de Fulana.pdf"
  python gerar_plano.py PASTA_COM_PDFS           -> processa todos os PDFs da pasta
  python gerar_plano.py ARQ.pdf --saida PASTA    -> define a pasta de saída
  python gerar_plano.py ARQ.pdf --sem-pdf        -> gera só o Word

Os arquivos são salvos na pasta plano_alimentar/output (ou em --saida), e a
pasta é aberta automaticamente no final.
"""
from __future__ import annotations

import argparse
import os
import re
import shutil
import subprocess
import sys
import tempfile
from dataclasses import dataclass, field
from pathlib import Path
from xml.sax.saxutils import escape

import pdfplumber
from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn

BASE_DIR = Path(__file__).resolve().parent
MODELO_PADRAO = BASE_DIR / "modelo" / "Modelo_Plano_Alimentar.docx"
SAIDA_PADRAO = BASE_DIR / "output"  # plano_alimentar/output

NUM_OPCOES = 4
# Larguras das colunas (em twips), iguais às do modelo original
LARG_COLUNAS = [1919, 3184, 3215, 3697, 3129]
COR_CABECALHO = "FFEBEB"
TAM_FONTE = 28  # meio-pontos (28 = 14 pt), como no modelo


# --------------------------------------------------------------------------- #
# Estrutura de dados
# --------------------------------------------------------------------------- #
@dataclass
class Opcao:
    itens: list[str] = field(default_factory=list)
    obs: str = ""


@dataclass
class Refeicao:
    horario: str  # "07:30"
    nome: str     # "Café da manhã"
    opcoes: list[Opcao] = field(default_factory=list)

    @property
    def titulo(self) -> str:
        h, m = self.horario.split(":")
        return f"{self.nome.upper()} – {int(h):02d}h{m}"

    @property
    def obs_principal(self) -> str:
        return self.opcoes[0].obs if self.opcoes else ""


@dataclass
class Plano:
    paciente: str
    data: str  # "11/11/2026"
    refeicoes: list[Refeicao]
    avisos: list[str] = field(default_factory=list)


# --------------------------------------------------------------------------- #
# Leitura do PDF do Dietbox
# --------------------------------------------------------------------------- #
RE_REFEICAO = re.compile(r"^(\d{1,2}:\d{2})\s*-\s*(.+)$")
RE_SUBST = re.compile(r"^Substitui[çc][ãa]o\s*\d*\s*$", re.I)
RE_OBS = re.compile(r"^Obs\.?\s*:\s*(.*)$", re.I)
RE_DATA = re.compile(r"^Data:\s*(\d{2}/\d{2}/\d{4})")
RE_RODAPE = re.compile(r"^Este documento é de uso exclusivo", re.I)


def ler_linhas_pdf(caminho: Path) -> list[str]:
    linhas: list[str] = []
    with pdfplumber.open(caminho) as pdf:
        for pagina in pdf.pages:
            texto = pagina.extract_text() or ""
            linhas.extend(l.strip() for l in texto.splitlines())
    # remove o aviso de confidencialidade do rodapé (2 linhas)
    limpas, pular = [], False
    for l in linhas:
        if RE_RODAPE.match(l):
            pular = True
            continue
        if pular and l.lower().rstrip(".") == "proibido":
            pular = False
            continue
        pular = False
        if l:
            limpas.append(l)
    return limpas


def ler_plano_dietbox(caminho: Path) -> Plano:
    linhas = ler_linhas_pdf(caminho)
    avisos: list[str] = []

    # --- paciente e data (bloco "Plano Alimentar / Nome / Data: ...") ---
    paciente, data = "", ""
    for i, l in enumerate(linhas):
        m = RE_DATA.match(l)
        if m:
            data = m.group(1)
            if i > 0 and linhas[i - 1].lower() != "plano alimentar":
                paciente = linhas[i - 1]
            break
    if not paciente:
        m = re.search(r"ativo de (.+)$", caminho.stem, re.I)
        paciente = m.group(1).strip() if m else caminho.stem
        avisos.append("Nome do paciente não encontrado no texto; usado o nome do arquivo.")
    if not data:
        avisos.append("Data do plano não encontrada no PDF.")

    # --- refeições ---
    refeicoes: list[Refeicao] = []
    atual: Refeicao | None = None
    em_obs = False
    for l in linhas:
        m = RE_REFEICAO.match(l)
        if m:
            atual = Refeicao(horario=m.group(1), nome=m.group(2).strip(), opcoes=[Opcao()])
            refeicoes.append(atual)
            em_obs = False
            continue
        if atual is None:
            continue  # cabeçalho do PDF (nutricionista, endereço, etc.)

        opcao = atual.opcoes[-1]
        if RE_SUBST.match(l):
            atual.opcoes.append(Opcao())
            em_obs = False
        elif l.startswith("•"):
            opcao.itens.append("• " + l.lstrip("•").strip())
            em_obs = False
        elif m := RE_OBS.match(l):
            opcao.obs = (opcao.obs + " " + m.group(1)).strip()
            em_obs = True
        elif em_obs:
            opcao.obs += " " + l  # observação quebrada em várias linhas
        elif opcao.itens:
            opcao.itens[-1] += " " + l  # item quebrado em várias linhas
        else:
            avisos.append(f"Linha ignorada em '{atual.nome}': {l}")

    for r in refeicoes:
        r.opcoes = [o for o in r.opcoes if o.itens or o.obs]
        if len(r.opcoes) > NUM_OPCOES:
            avisos.append(
                f"'{r.nome}' tem {len(r.opcoes)} opções; o modelo comporta {NUM_OPCOES}. "
                f"As opções excedentes NÃO foram incluídas."
            )
            r.opcoes = r.opcoes[:NUM_OPCOES]
        for n, o in enumerate(r.opcoes[1:], start=2):
            if o.obs:
                avisos.append(f"'{r.nome}', opção {n}: observação incluída dentro da célula.")

    if not refeicoes:
        raise ValueError("Nenhuma refeição encontrada. O PDF é um 'Plano alimentar ativo' do Dietbox?")
    return Plano(paciente=paciente, data=data, refeicoes=refeicoes, avisos=avisos)


# --------------------------------------------------------------------------- #
# Montagem do Word
# --------------------------------------------------------------------------- #
def _rpr(negrito: bool = False) -> str:
    b = "<w:b/><w:bCs/>" if negrito else ""
    return (f'<w:rPr><w:rFonts w:cstheme="minorHAnsi"/>{b}'
            f'<w:sz w:val="{TAM_FONTE}"/><w:szCs w:val="{TAM_FONTE}"/></w:rPr>')


def _run(texto: str, negrito: bool = False) -> str:
    return f'<w:r>{_rpr(negrito)}<w:t xml:space="preserve">{escape(texto)}</w:t></w:r>'


def _runs_com_ou(texto: str, negrito: bool = False) -> str:
    """Divide o texto destacando a palavra 'OU' em negrito."""
    partes = re.split(r"\b(OU)\b", texto)
    return "".join(_run(p, negrito or p == "OU") for p in partes if p)


def _par(conteudo: str, alinhamento: str | None = None, lista: bool = False,
         negrito: bool = False, manter_com_proximo: bool = False) -> str:
    ppr = ""
    if lista:
        ppr += '<w:pStyle w:val="PargrafodaLista"/>'
    if manter_com_proximo:
        ppr += "<w:keepNext/>"
    if not lista:
        ppr += '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>'
    if lista:
        ppr += '<w:ind w:left="0"/>'
    if alinhamento:
        ppr += f'<w:jc w:val="{alinhamento}"/>'
    return f"<w:p><w:pPr>{ppr}{_rpr(negrito)}</w:pPr>{conteudo}</w:p>"


def _celula(paragrafos: str, largura: int, span: int = 1, vmerge: str | None = None,
            fundo: str | None = None, centro_vertical: bool = False,
            bordas: str = "") -> str:
    tcpr = f'<w:tcW w:w="{largura}" w:type="dxa"/>'
    if span > 1:
        tcpr += f'<w:gridSpan w:val="{span}"/>'
    if vmerge == "restart":
        tcpr += '<w:vMerge w:val="restart"/>'
    elif vmerge == "continue":
        tcpr += "<w:vMerge/>"
    tcpr += bordas
    if fundo:
        tcpr += f'<w:shd w:val="clear" w:color="auto" w:fill="{fundo}"/>'
    if centro_vertical:
        tcpr += '<w:vAlign w:val="center"/>'
    return f"<w:tc><w:tcPr>{tcpr}</w:tcPr>{paragrafos or _par('')}</w:tc>"


def _linha(celulas: str, altura: int | None = None, cabecalho: bool = False,
           nao_dividir: bool = False) -> str:
    trpr = ""
    if nao_dividir:
        trpr += "<w:cantSplit/>"
    if altura:
        trpr += f'<w:trHeight w:val="{altura}"/>'
    if cabecalho:
        trpr += "<w:tblHeader/>"
    return f"<w:tr><w:trPr>{trpr}</w:trPr>{celulas}</w:tr>"


def _paragrafos_opcao(opcao: Opcao, manter: bool) -> str:
    pars = [_par(_runs_com_ou(item), lista=True, manter_com_proximo=manter) for item in opcao.itens]
    return "".join(pars)


def _obs_paragrafo(obs: str, manter: bool = False) -> str:
    return _par(_run("Obs: ", negrito=True) + _runs_com_ou(obs), lista=True,
                manter_com_proximo=manter)


def montar_tabela(plano: Plano) -> str:
    L = LARG_COLUNAS
    sem_borda = ('<w:tcBorders><w:top w:val="nil"/><w:left w:val="nil"/>'
                 '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
                 '<w:right w:val="nil"/></w:tcBorders>')

    # Linha 1: PACIENTE / Data  (repetida no topo de cada página)
    data_txt = _run("Data: ", True) + _run(plano.data, True) if plano.data else ""
    l1 = (_celula(_par(_run("PACIENTE:", True), "center", negrito=True), L[0],
                  bordas=sem_borda, centro_vertical=True)
          + _celula(_par(_run(plano.paciente, True), negrito=True), L[1] + L[2], span=2,
                    bordas=sem_borda, centro_vertical=True)
          + _celula(_par(data_txt, negrito=True), L[3], bordas=sem_borda, centro_vertical=True)
          + _celula(_par("", negrito=True), L[4], bordas=sem_borda))
    # Linha 2: REFEIÇÕES / OPÇÃO 1..4  (repetida no topo de cada página)
    titulos = ["REFEIÇÕES"] + [f"OPÇÃO {i}" for i in range(1, NUM_OPCOES + 1)]
    l2 = "".join(_celula(_par(_run(t, True), "center", negrito=True), L[i],
                         fundo=COR_CABECALHO, centro_vertical=True)
                 for i, t in enumerate(titulos))
    linhas = [_linha(l1, 550, cabecalho=True), _linha(l2, 557, cabecalho=True)]

    for ref in plano.refeicoes:
        tem_obs = bool(ref.obs_principal)
        titulo = _par(_run(ref.titulo, True), "center", negrito=True, manter_com_proximo=tem_obs)
        celulas = _celula(titulo, L[0], vmerge="restart" if tem_obs else None,
                          centro_vertical=True)
        if len(ref.opcoes) == 1:
            # opção única: ocupa as 4 colunas
            celulas += _celula(_paragrafos_opcao(ref.opcoes[0], tem_obs), sum(L[1:]),
                               span=NUM_OPCOES)
        else:
            for i in range(NUM_OPCOES):
                if i < len(ref.opcoes):
                    op = ref.opcoes[i]
                    conteudo = _paragrafos_opcao(op, tem_obs)
                    if i > 0 and op.obs:  # obs de substituição fica dentro da célula
                        conteudo += _obs_paragrafo(op.obs, tem_obs)
                else:
                    conteudo = ""
                celulas += _celula(conteudo, L[i + 1])
        linhas.append(_linha(celulas, 771, nao_dividir=True))

        if tem_obs:
            obs_row = (_celula(_par("", "center", negrito=True), L[0], vmerge="continue",
                               centro_vertical=True)
                       + _celula(_obs_paragrafo(ref.obs_principal), sum(L[1:]),
                                 span=NUM_OPCOES))
            linhas.append(_linha(obs_row, None, nao_dividir=True))

    borda = '<w:{0} w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
    bordas = "".join(borda.format(b) for b in
                     ("top", "left", "bottom", "right", "insideH", "insideV"))
    grid = "".join(f'<w:gridCol w:w="{w}"/>' for w in L)
    return (f"<w:tbl {nsdecls('w')}><w:tblPr><w:tblW w:w=\"{sum(L)}\" w:type=\"dxa\"/>"
            f"<w:tblBorders>{bordas}</w:tblBorders><w:tblLayout w:type=\"fixed\"/>"
            "<w:tblCellMar><w:left w:w=\"70\" w:type=\"dxa\"/><w:right w:w=\"70\" w:type=\"dxa\"/>"
            "</w:tblCellMar><w:tblLook w:val=\"0000\"/></w:tblPr>"
            f"<w:tblGrid>{grid}</w:tblGrid>{''.join(linhas)}</w:tbl>")


def gerar_docx(plano: Plano, modelo: Path, destino: Path) -> None:
    doc = Document(str(modelo))
    body = doc.element.body
    sect = body.find(qn("w:sectPr"))

    # Se o modelo já tiver tabelas (ex.: um plano antigo), remove da 1ª tabela em diante
    primeira = body.find(qn("w:tbl"))
    if primeira is not None:
        apagar = False
        for el in list(body):
            if el is primeira:
                apagar = True
            if apagar and el is not sect:
                body.remove(el)

    tabela = parse_xml(montar_tabela(plano))
    final = parse_xml(f'<w:p {nsdecls("w")}><w:pPr><w:spacing w:after="0"/>'
                      '<w:rPr><w:sz w:val="2"/></w:rPr></w:pPr></w:p>')
    sect.addprevious(tabela)
    sect.addprevious(final)  # o Word exige um parágrafo após a última tabela

    doc.core_properties.title = f"Plano Alimentar - {plano.paciente}"
    doc.save(str(destino))


# --------------------------------------------------------------------------- #
# Conversão para PDF
# --------------------------------------------------------------------------- #
def _achar_libreoffice() -> str | None:
    for nome in ("soffice", "libreoffice"):
        if achado := shutil.which(nome):
            return achado
    for p in (r"C:\Program Files\LibreOffice\program\soffice.exe",
              r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
              "/Applications/LibreOffice.app/Contents/MacOS/soffice"):
        if Path(p).exists():
            return p
    return None


def converter_pdf(docx_path: Path, pdf_path: Path) -> str:
    """Converte usando o Microsoft Word (docx2pdf); se não houver, usa o LibreOffice."""
    erros = []
    if sys.platform in ("win32", "darwin"):
        try:
            from docx2pdf import convert
            convert(str(docx_path), str(pdf_path))
            if pdf_path.exists():
                return "Microsoft Word"
        except Exception as e:  # Word ausente, docx2pdf não instalado, etc.
            erros.append(f"Word: {e}")

    soffice = _achar_libreoffice()
    if soffice:
        with tempfile.TemporaryDirectory() as tmp:
            subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir", tmp,
                            str(docx_path)], check=True, capture_output=True, timeout=180)
            gerado = Path(tmp) / (docx_path.stem + ".pdf")
            shutil.move(str(gerado), pdf_path)
        return "LibreOffice"
    erros.append("LibreOffice não encontrado")
    raise RuntimeError("Não foi possível gerar o PDF (" + "; ".join(erros) + ")")


# --------------------------------------------------------------------------- #
# Programa principal
# --------------------------------------------------------------------------- #
def nome_arquivo_seguro(texto: str) -> str:
    return re.sub(r'[<>:"/\\|?*]', "-", texto).strip()


def processar(pdf: Path, modelo: Path, pasta_saida: Path | None, gerar_pdf: bool) -> Path | None:
    """Gera os arquivos e devolve a pasta onde foram salvos (ou None em caso de erro)."""
    print(f"\n>> {pdf.name}")
    try:
        plano = ler_plano_dietbox(pdf)
    except Exception as e:
        print(f"   [ERRO] Falha ao ler o PDF: {e}")
        return None

    print(f"   Paciente: {plano.paciente} | Data: {plano.data or '-'}")
    for r in plano.refeicoes:
        obs = " + obs" if r.obs_principal else ""
        print(f"   - {r.titulo}: {len(r.opcoes)} opção(ões){obs}")
    for a in plano.avisos:
        print(f"   [AVISO] {a}")

    saida = (pasta_saida or SAIDA_PADRAO).resolve()
    saida.mkdir(parents=True, exist_ok=True)
    base = nome_arquivo_seguro(
        f"Plano Alimentar - {plano.paciente}" + (f" - {plano.data.replace('/', '-')}" if plano.data else ""))
    docx_path = saida / f"{base}.docx"
    try:
        gerar_docx(plano, modelo, docx_path)
    except PermissionError:
        print(f"   [ERRO] Não foi possível salvar {docx_path.name}. O arquivo está aberto no Word?")
        return None
    print(f"   [OK] Word: {docx_path}")

    if gerar_pdf:
        pdf_path = saida / f"{base}.pdf"
        try:
            via = converter_pdf(docx_path, pdf_path)
            print(f"   [OK] PDF ({via}): {pdf_path}")
        except Exception as e:
            print(f"   [ERRO] {e}")
            return saida  # o Word foi gerado mesmo sem o PDF
    return saida


def abrir_pasta(pasta: Path) -> None:
    """Abre a pasta no Explorador de Arquivos (Windows), Finder (Mac) ou gerenciador (Linux)."""
    try:
        if sys.platform == "win32":
            os.startfile(pasta)  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(pasta)])
        elif shutil.which("xdg-open"):
            subprocess.Popen(["xdg-open", str(pasta)], stdout=subprocess.DEVNULL,
                             stderr=subprocess.DEVNULL)
    except Exception:
        pass  # abrir a pasta é só uma conveniência


def escolher_arquivos() -> list[Path]:
    try:
        import tkinter as tk
        from tkinter import filedialog
    except ImportError:
        return []
    root = tk.Tk()
    root.withdraw()
    arquivos = filedialog.askopenfilenames(
        title="Selecione o(s) Plano(s) alimentar(es) ativo(s) do Dietbox",
        filetypes=[("PDF", "*.pdf")])
    root.destroy()
    return [Path(a) for a in arquivos]


def main() -> None:
    ap = argparse.ArgumentParser(description="Converte o PDF do Dietbox no plano alimentar personalizado.")
    ap.add_argument("entradas", nargs="*", help="PDF(s) do Dietbox ou pasta(s) contendo os PDFs")
    ap.add_argument("--saida", type=Path, help="Pasta de saída (padrão: plano_alimentar/output)")
    ap.add_argument("--modelo", type=Path, default=MODELO_PADRAO, help="Modelo Word (.docx)")
    ap.add_argument("--sem-pdf", action="store_true", help="Gera apenas o Word")
    ap.add_argument("--nao-abrir", action="store_true", help="Não abre a pasta de saída no final")
    args = ap.parse_args()

    entradas = [Path(e) for e in args.entradas] or escolher_arquivos()
    pdfs: list[Path] = []
    for e in entradas:
        if e.is_dir():
            pdfs += sorted(p for p in e.glob("*.pdf") if "ativo" in p.name.lower())
        elif e.suffix.lower() == ".pdf" and e.exists():
            pdfs.append(e)
        else:
            print(f"[AVISO] Ignorado (não é PDF ou não existe): {e}")

    if not pdfs:
        print("Nenhum PDF para processar.")
    elif not args.modelo.exists():
        print(f"Modelo não encontrado: {args.modelo}")
    else:
        pastas = [processar(p, args.modelo, args.saida, not args.sem_pdf) for p in pdfs]
        geradas = list(dict.fromkeys(p for p in pastas if p))
        print(f"\nConcluído: {sum(1 for p in pastas if p)} de {len(pdfs)} plano(s) gerado(s).")
        for pasta in geradas:
            print(f"Arquivos salvos em: {pasta}")
            if not args.nao_abrir:
                abrir_pasta(pasta)

    if not args.entradas:  # executado com duplo clique: mantém a janela aberta
        input("\nPressione Enter para sair...")


if __name__ == "__main__":
    main()
