#!/usr/bin/env python3
"""
Gerador de Propostas — Solar In Focco
=====================================
Lê a conta de energia (Enel, PDF) e o orçamento da distribuidora de kits (PDF,
texto ou escaneado) e preenche o modelo Word padrão, alterando SOMENTE os
campos combinados. A tabela de payback é calculada pela planilha Excel
(aba "PAY BACK", com fórmulas) e a tabela de orçamento pela aba "ORÇAMENTO".

Uso:
  python gerar_proposta.py --conta conta.pdf --orcamento orcamento.pdf \
      --modelo modelo.docx --planilha payback.xlsx --saida pasta_saida

Requisitos: python-docx não é necessário; usa lxml, openpyxl, pdfplumber,
Pillow, pypdfium2, LibreOffice e, para PDFs escaneados, Tesseract (por).
"""
import argparse, copy, datetime as dt, json, os, re, shutil, subprocess, sys, tempfile, zipfile
from pathlib import Path

APP_DIR = Path(__file__).resolve().parent
NO_WINDOW = 0x08000000 if os.name == "nt" else 0  # evita janelas pretas no Windows

def _run(cmd, **kw):
    return subprocess.run(cmd, creationflags=NO_WINDOW, **kw)

def _achar_exe(nome, candidatos):
    achado = shutil.which(nome)
    if achado:
        return achado
    for c in candidatos:
        if c and Path(c).exists():
            return str(c)
    return None

def tesseract_exe():
    pf = os.environ.get("ProgramFiles", r"C:\Program Files")
    la = os.environ.get("LOCALAPPDATA", "")
    return _achar_exe("tesseract", [Path(pf) / "Tesseract-OCR" / "tesseract.exe",
                                     Path(la) / "Programs" / "Tesseract-OCR" / "tesseract.exe"])

def soffice_exe():
    pf = os.environ.get("ProgramFiles", r"C:\Program Files")
    pf86 = os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")
    return _achar_exe("soffice", [Path(pf) / "LibreOffice" / "program" / "soffice.exe",
                                   Path(pf86) / "LibreOffice" / "program" / "soffice.exe",
                                   "/Applications/LibreOffice.app/Contents/MacOS/soffice"]) or shutil.which("libreoffice")

def pdf_para_pngs(pdf, pasta, prefixo, dpi=300):
    """Converte cada página do PDF em PNG (pypdfium2, sem programas externos)."""
    import pypdfium2 as pdfium
    doc = pdfium.PdfDocument(str(pdf))
    saidas = []
    for i in range(len(doc)):
        img = doc[i].render(scale=dpi / 72).to_pil()
        f = Path(pasta) / f"{prefixo}-{i + 1:02d}.png"
        img.save(f)
        saidas.append(f)
    doc.close()
    return saidas

MESES = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho",
         "agosto", "setembro", "outubro", "novembro", "dezembro"]

# --------------------------------------------------------------------------
# Formatação BR
# --------------------------------------------------------------------------
def br(v, casas=2):
    s = f"{v:,.{casas}f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")

def num_br(s):
    """'8.262,26' -> 8262.26 ; '4.83' -> 4.83"""
    s = s.strip()
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    return float(s)

def fmt_como(valor, antigo, casas=2):
    """Formata 'valor' no mesmo estilo do texto antigo (vírgula ou ponto decimal)."""
    if "," in antigo:
        return br(valor, casas)
    if "." in antigo and len(antigo.split(".")[-1]) != 3:
        return f"{valor:.{casas}f}"
    return br(valor, casas)

PARTICULAS = {"de", "da", "do", "das", "dos", "e"}
def titulo(s):
    out = []
    for i, w in enumerate(s.lower().split()):
        out.append(w if (i > 0 and w in PARTICULAS) else w[:1].upper() + w[1:])
    return " ".join(out)

CIDADES = {
    "SAO BERNARDO DO CAMPO": ("São Bernardo do Campo", "SBC"),
    "SANTO ANDRE": ("Santo André", "Santo André"),
    "SAO CAETANO DO SUL": ("São Caetano do Sul", "SCS"),
    "DIADEMA": ("Diadema", "Diadema"),
    "MAUA": ("Mauá", "Mauá"),
    "RIBEIRAO PIRES": ("Ribeirão Pires", "Ribeirão Pires"),
    "RIO GRANDE DA SERRA": ("Rio Grande da Serra", "Rio Grande da Serra"),
    "SAO PAULO": ("São Paulo", "São Paulo"),
}
TIPOS_LOGRADOURO = {"R": "Rua", "AV": "Avenida", "AL": "Alameda", "TV": "Travessa",
                    "EST": "Estrada", "ROD": "Rodovia", "PC": "Praça", "PCA": "Praça",
                    "ESTR": "Estrada", "VL": "Vila"}

def formatar_endereco(raw):
    m = re.match(r"^\s*([A-Z]{1,4})\.?\s+(.+?)\s+(\d+\w*)\s*(?:-\s*(.+))?$", raw.strip())
    if not m:
        return titulo(raw)
    tipo, nome, numero, bairro = m.groups()
    tipo = TIPOS_LOGRADOURO.get(tipo, titulo(tipo))
    s = f"{tipo} {titulo(nome)}, {numero.lower()}"
    if bairro:
        s += f" - {titulo(bairro)}"
    return s

# --------------------------------------------------------------------------
# Leitura de PDFs
# --------------------------------------------------------------------------
def texto_pdf(caminho, ocr_se_vazio=True):
    import pdfplumber
    with pdfplumber.open(caminho) as pdf:
        txt = "\n".join((p.extract_text(layout=True) or "") for p in pdf.pages)
    if len(txt.strip()) > 80 or not ocr_se_vazio:
        return txt, "texto"
    # PDF escaneado -> OCR
    exe = tesseract_exe()
    if not exe:
        raise RuntimeError("O orçamento é uma imagem escaneada e o leitor de texto (Tesseract) não foi encontrado. "
                           "Rode o INSTALAR novamente.")
    tmp = Path(tempfile.mkdtemp())
    partes = []
    tessdata = APP_DIR / "tessdata"
    extra = ["--tessdata-dir", str(tessdata)] if (tessdata / "por.traineddata").exists() else []
    for img in pdf_para_pngs(caminho, tmp, "p"):
        r = _run([exe, str(img), "stdout", "-l", "por", "--psm", "4", *extra],
                 capture_output=True, text=True, encoding="utf-8", errors="ignore")
        partes.append(r.stdout)
    shutil.rmtree(tmp, ignore_errors=True)
    return "\n".join(partes), "ocr"

def ler_conta(caminho):
    txt, modo = texto_pdf(caminho)
    d = {"_modo": modo}
    # Nome/endereço: bloco do pagador (débito automático "PAGADOR / CPF:" ou boleto "PAGADOR:")
    m = re.search(r"PAGADOR(?:\s*/\s*CPF)?:\s*\n\s*(.+?)\s+-\s+CPF", txt)
    if m:
        d["nome"] = titulo(m.group(1))
        linha = txt[m.end():].split("\n", 1)[1].strip().split("\n")[0]
        m2 = re.match(r"(.+?)\s+-\s+([A-ZÀ-Ú ]+)/([A-Z]{2})\s+CEP:\s*(\d{5}-\d{3})", linha.strip())
        if m2:
            d["endereco_raw"], cid, d["uf"], d["cep"] = m2.groups()
            d["cidade_raw"] = cid.strip()
    if "nome" not in d or "endereco_raw" not in d:
        # Alternativa: cabeçalho da 1ª página (nome, endereço e "CEP: ... - CIDADE/UF")
        m = re.search(r"\n\s*([A-ZÀ-Ú][A-ZÀ-Ú .'-]{5,})\s*\n\s*\n?\s*(\S.+?)\s*\n(?:.*\n){0,3}?\s*CEP:\s*(\d{5}-\d{3})\s*-\s*([A-ZÀ-Ú ]+?)\s*[/-]\s*([A-Z]{2})\b", txt)
        if m:
            d.setdefault("nome", titulo(m.group(1).strip()))
            if "endereco_raw" not in d:
                d["endereco_raw"] = re.split(r"\s{3,}", m.group(2).strip())[0]
                d["cep"], d["cidade_raw"], d["uf"] = m.group(3), m.group(4).strip(), m.group(5)
    m = re.search(r"\b(B\d?)\s*-\s*(B\d)\s*-\s*\w+\s*-\s*(\w+)", txt)
    if m:
        d["subgrupo"], d["classe_tarifaria"] = m.group(2), m.group(3)
    m = re.search(r"(Monof[áa]sico|Bif[áa]sico|Trif[áa]sico)", txt, re.I)
    if m:
        lig = m.group(1).capitalize()
        d["ligacao"] = re.sub(r"fasico$", "fásico", lig)
    hist = re.findall(r"\b(JAN|FEV|MAR|ABR|MAI|JUN|JUL|AGO|SET|OUT|NOV|DEZ)/(\d{2})\s+([\d.]+,\d{3})\s+(\d{2})\s+[A-Z]{2,4}", txt)
    d["historico"] = [(f"{a}/{b}", num_br(c)) for a, b, c, _ in hist]
    def valor(rotulo):
        m = re.search(rotulo + r"\s+(-?[\d.]+,\d{2})", txt)
        return num_br(m.group(1)) if m else None
    d["subtotal_faturamento"] = valor(r"Subtotal Faturamento")
    m = re.search(r"^\s*TOTAL\s+([\d.]+,\d{2})", txt, re.M)
    d["total"] = num_br(m.group(1)) if m else None
    m = re.search(r"CIP[^\n]*?\s([\d.]+,\d{2})\s", txt)
    d["cip"] = num_br(m.group(1)) if m else None
    m = re.search(r"\(TUSD\)\s+KWH\s+[\d.,]+\s+([\d,]+)", txt)
    t1 = num_br(m.group(1)) if m else None
    m = re.search(r"ENERGIA \(TE\)\s+KWH\s+[\d.,]+\s+([\d,]+)", txt)
    t2 = num_br(m.group(1)) if m else None
    d["tarifa_conta_com_tributos"] = (t1 + t2) if (t1 and t2) else None
    m = re.search(r"REFER[ÊE]NCIA:.*?\n\s*\S+\s+\S+\s+(\d{2}/\d{4})", txt, re.S)
    d["referencia"] = m.group(1) if m else None
    return d

MARCAS_MODULO = ["RISEN", "OSDA", "JINKO", "ASTRONERGY", "ASTROENERGY", "CANADIAN", "TRINA", "LONGI",
                 "JA SOLAR", "BYD", "SUNOVA", "HONOR", "DAH", "ZNSHINE", "AE SOLAR", "TSUN", "LEAPTON", "RONMA", "BELENERGY"]
MARCAS_INVERSOR = ["DEYE", "HOYMILES", "APSYSTEMS", "GROWATT", "SOLIS", "SAJ", "FRONIUS", "GOODWE",
                   "SUNGROW", "HUAWEI", "PHB", "TSUNESS", "ENPHASE", "SOFAR", "CHINT", "WEG"]

def _marca(linha, lista, prox=""):
    for mk in lista:
        if mk in linha.upper():
            return mk
    for mk in lista:
        if mk in prox.upper():
            return mk
    return ""

def _qtd_apos(linhas, i):
    for j in range(i, min(i + 6, len(linhas))):
        if j > i and re.search(r"^[A-ZÁ-Ú0-9 .,/()-]{25,}$", linhas[j].strip()) and "CÓD" not in linhas[j].upper():
            break  # chegou no próximo produto
        m = re.search(r"\b(\d{1,4})\s*(PC|PÇ|UN|UND|PÇS)\b", linhas[j].upper())
        if m:
            return int(m.group(1))
    return None

def ler_orcamento(caminho):
    txt, modo = texto_pdf(caminho)
    d = {"_modo": modo}
    m = re.search(r"Cota[çc][ãa]o\s+([A-Z]+-\d+)", txt)
    d["cotacao"] = m.group(1) if m else ""
    m = re.search(r"Pot[êe]ncia do sistema:?\s*([\d.,]+)\s*kWp", txt, re.I)
    d["kwp"] = float(m.group(1).replace(",", ".")) if m else None
    m = re.search(r"Cliente:\s*(.+?)(?:\s{2,}|\n|$)", txt)
    d["cliente"] = m.group(1).strip() if m else ""
    tipos = r"(Colonial|Fibrocimento|Met[áa]lic[ao]|Laje|Solo|Cer[âa]mic[ao]|T[ée]gula|Shingle|Romana|Portuguesa|Francesa)"
    trecho = txt[txt.lower().find("estrutura"):] if "estrutura" in txt.lower() else txt
    # 1º: tipo no início da linha (coluna "Tipo estrutura"); 2º: qualquer menção que não seja "não indicado em ..."
    m = re.search(r"^\s*" + tipos + r"\b", trecho, re.I | re.M)
    if not m:
        m = next((x for x in re.finditer(r"\b" + tipos + r"\b", trecho, re.I)
                  if "não indicado" not in trecho[max(0, x.start() - 40):x.start()].lower()), None)
    d["estrutura"] = m.group(1).capitalize() if m else ""

    linhas = txt.split("\n")
    for i, l in enumerate(linhas):
        U = l.upper()
        if "MODULO" in U.replace("Ó", "O") and re.search(r"\d{3}\s*W\b", U) and "modulo_w" not in d:
            d["modulo_w"] = int(re.search(r"(\d{3})\s*W\b", U).group(1))
            d["modulo_marca"] = titulo(_marca(U, MARCAS_MODULO, " ".join(linhas[i:i + 4])))
            d["modulo_bifacial"] = "BIFACIAL" in U
            d["modulo_qtd"] = _qtd_apos(linhas, i)
            d["modulo_desc_original"] = l.strip()
        if "INVERSOR" in U and re.search(r"(\d+[.,]?\d*)\s*KW", U) and "inversor_kw" not in d:
            d["inversor_kw"] = float(re.search(r"(\d+[.,]?\d*)\s*KW", U).group(1).replace(",", "."))
            d["inversor_marca"] = titulo(_marca(U, MARCAS_INVERSOR, " ".join(linhas[i:i + 4])))
            d["inversor_micro"] = "MICRO" in U
            d["inversor_qtd"] = _qtd_apos(linhas, i)
            d["inversor_desc_original"] = l.strip()

    def valor(rot):
        m = re.search(rot + r"[^\n]*?R\$\s*([\d.]+,\d{2})", txt, re.I)
        return num_br(m.group(1)) if m else 0.0
    comp = {k: valor(r) for k, r in {
        "produtos": r"Total de Produtos", "frete": r"Valor do Frete", "seguro": r"Valor do Seguro",
        "icms": r"Total de ICMS", "difal": r"Diferencial de Al[íi]quota", "ipi": r"Total de IPI", "st": r"Total de ST"}.items()}
    d["componentes"] = comp
    soma = round(sum(comp.values()), 2)
    todos = [num_br(x) for x in re.findall(r"R\$\s*([\d.]+,\d{2})", txt)]
    maior = max(todos) if todos else 0.0
    d["valor_total"] = soma if soma > 0 else maior
    d["valor_total_conferido"] = abs(soma - maior) < 0.02
    return d

# --------------------------------------------------------------------------
# Cálculos
# --------------------------------------------------------------------------
def calcular(conta, orc, cfg, data):
    ov = cfg.get("overrides", {})
    r = {}
    r["data"] = data
    r["cliente"] = ov.get("cliente", conta.get("nome", ""))
    r["endereco"] = ov.get("endereco", formatar_endereco(conta.get("endereco_raw", "")))
    cid = conta.get("cidade_raw", "")
    r["municipio"] = ov.get("municipio", f"{CIDADES.get(cid, (titulo(cid), titulo(cid)))[1]} / {conta.get('uf', 'SP')}")
    classe = titulo(conta.get("classe_tarifaria", "Residencial"))
    r["classe"] = ov.get("classe", f"{classe} {conta.get('ligacao', '')}".strip())
    hist = [v for _, v in conta.get("historico", [])][:12]
    r["consumo_medio"] = ov.get("consumo_medio", round(sum(hist) / len(hist)) if hist else 0)
    if cfg.get("valor_conta_fora_taxa") == "total":
        conta_val = conta.get("total")
    else:
        conta_val = conta.get("subtotal_faturamento") or conta.get("total")
    r["valor_conta"] = ov.get("valor_conta", conta_val or 0)

    r["kwp"] = ov.get("kwp", orc.get("kwp") or (orc["modulo_qtd"] * orc["modulo_w"] / 1000))
    r["qtd_paineis"] = ov.get("qtd_paineis", orc.get("modulo_qtd") or round(r["kwp"] * 1000 / orc["modulo_w"]))
    r["modulo_w"] = ov.get("modulo_w", orc.get("modulo_w"))
    r["modulo_marca"] = ov.get("modulo_marca", orc.get("modulo_marca", ""))
    r["modulo_bifacial"] = orc.get("modulo_bifacial", False)
    r["inversor_kw"] = ov.get("inversor_kw", orc.get("inversor_kw"))
    r["inversor_marca"] = ov.get("inversor_marca", orc.get("inversor_marca", ""))
    r["inversor_micro"] = orc.get("inversor_micro", True)
    r["inversor_qtd"] = ov.get("inversor_qtd", orc.get("inversor_qtd") or 1)
    r["estrutura"] = ov.get("estrutura", orc.get("estrutura", "Colonial"))

    r["hsp"] = cfg["hsp"]
    r["tarifa"] = cfg["tarifa_kwh"]
    r["reajuste"] = cfg["reajuste_anual"]
    r["geracao"] = ov.get("geracao", round(r["kwp"] * r["hsp"] * cfg["dias_mes"] * cfg["performance_ratio"]))
    base = cfg.get("base_kwh_payback", "geracao")
    r["kwh_payback"] = {"geracao": r["geracao"], "consumo": r["consumo_medio"],
                        "menor": min(r["geracao"], r["consumo_medio"] or r["geracao"])}[base]
    r["area_m2"] = ov.get("area_m2", round(r["qtd_paineis"] * cfg["area_por_painel_m2"]))
    r["peso_kg"] = ov.get("peso_kg", round(r["qtd_paineis"] * cfg["peso_por_painel_kg"]))

    r["custo_distribuidora"] = orc["valor_total"]
    preco = r["custo_distribuidora"] * cfg["margem_sobre_distribuidora"]
    arr = cfg.get("arredondar_preco_para", 0)
    if arr:
        preco = round(preco / arr) * arr
    r["investimento"] = round(ov.get("investimento", preco), 2)

    # Projeção (espelha as fórmulas da aba PAY BACK)
    tar, acum, anos = r["tarifa"], 0.0, []
    for ano in range(1, 26):
        if ano > 1:
            tar = tar + tar * r["reajuste"]
        mensal = tar * r["kwh_payback"]
        anual = mensal * 12
        acum += anual
        anos.append({"ano": ano, "tarifa": tar, "mensal": mensal, "anual": anual, "acumulado": acum})
    r["projecao"] = anos
    r["economia_anual"] = anos[0]["anual"]
    r["ano_payback"] = next((a["ano"] for a in anos if a["acumulado"] >= r["investimento"]), 25)

    # Textos dos itens
    r["txt_modulo_tabela"] = "Módulo Fotovoltaico " + ("Bifacial " if r["modulo_bifacial"] else "") + f"{r['modulo_marca']} {r['modulo_w']} W"
    if r["inversor_marca"].lower() == "deye" and r["inversor_micro"]:
        r["txt_inversor_tabela"] = f"Micro Inversor Deye {br(r['inversor_kw'] * 1000, 0)} SUN"
    else:
        r["txt_inversor_tabela"] = ("Micro Inversor " if r["inversor_micro"] else "Inversor Fotovoltaico ") + \
            f"{r['inversor_marca']} {br(r['inversor_kw'], 2)} kW"
    kw = f"{r['inversor_kw']:.2f}".rstrip("0").rstrip(".") if r["inversor_kw"] % 1 else f"{r['inversor_kw']:.0f}"
    r["txt_inversor_detalhe"] = ("Micro Inversor " if r["inversor_micro"] else "Inversor ") + f"{r['inversor_marca']} {kw}kW"
    r["txt_modulo_detalhe"] = f"{r['modulo_w']} w {r['modulo_marca']}"

    # Alertas
    al = []
    if orc.get("modulo_qtd") and orc.get("kwp") and abs(orc["modulo_qtd"] * orc["modulo_w"] / 1000 - orc["kwp"]) > 0.05:
        al.append(f"Potência do orçamento ({orc['kwp']} kWp) não bate com {orc['modulo_qtd']} × {orc['modulo_w']} W.")
    if not orc.get("valor_total_conferido", True):
        al.append("Valor total do orçamento não confere com a soma dos componentes — confira.")
    nc, no = (r["cliente"] or "").split(" ")[0].lower(), (orc.get("cliente") or "").split(" ")[0].lower()
    if no and nc and no != nc:
        al.append(f"Cliente do orçamento ('{orc.get('cliente')}') é diferente do titular da conta ('{r['cliente']}').")
    if r["consumo_medio"] and r["geracao"] > r["consumo_medio"] * 1.1 and base == "geracao":
        al.append(f"Geração estimada ({r['geracao']} kWh/mês) é maior que o consumo médio da conta "
                  f"({r['consumo_medio']} kWh/mês). O payback usa a geração (como no modelo) e pode superestimar a economia.")
    if conta.get("tarifa_conta_com_tributos"):
        al.append(f"Tarifa na conta (TUSD+TE c/ tributos): R$ {br(conta['tarifa_conta_com_tributos'], 5)}/kWh; "
                  f"a proposta usa R$ {br(r['tarifa'])} (config).")
    if not orc.get("modulo_qtd"):
        al.append("Quantidade de módulos não lida no orçamento — calculada pela potência.")
    if not orc.get("inversor_qtd"):
        al.append("Quantidade de inversores não lida no orçamento — usando 1.")
    if not r["cliente"]:
        al.insert(0, "NOME DO CLIENTE não foi encontrado na conta — o campo ficou em branco. Preencha no Word.")
    if not r["endereco"]:
        al.insert(0, "ENDEREÇO não foi encontrado na conta — confira no Word.")
    if not conta.get("historico"):
        al.insert(0, "Histórico de consumo não foi encontrado na conta — o consumo médio ficou 0.")
    r["alertas"] = al
    return r

# --------------------------------------------------------------------------
# Planilha (edição direta do XML, preservando o arquivo original)
# --------------------------------------------------------------------------
def _xml_esc(s):
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

def set_cell(xml, ref, valor):
    pat = re.compile(rf'<c r="{ref}"([^>]*?)(/>|>.*?</c>)', re.S)
    m = pat.search(xml)
    attrs = re.sub(r'\s+t="[^"]*"', "", m.group(1)) if m else ""
    if isinstance(valor, str):
        novo = f'<c r="{ref}"{attrs} t="inlineStr"><is><t xml:space="preserve">{_xml_esc(valor)}</t></is></c>'
    else:
        novo = f'<c r="{ref}"{attrs}><v>{valor}</v></c>'
    if m:
        return xml[:m.start()] + novo + xml[m.end():]
    raise ValueError(f"Célula {ref} não encontrada")

def limpar_cache_formulas(xml):
    return re.sub(r"(<f[^>]*?(?:/>|>.*?</f>))<v>[^<]*</v>", r"\1", xml, flags=re.S)

def mover_destaque(xml, ano_destino):
    """A linha destacada (altura maior) acompanha o ano do payback."""
    rows = {int(m.group(1)): m for m in re.finditer(r'<row r="(\d+)"[^>]*>.*?</row>', xml, re.S)}
    atual = next((n for n in range(6, 31) if n in rows and 'customHeight="1"' in rows[n].group(0)
                  and re.search(r'ht="2\d', rows[n].group(0))), None)
    destino = 5 + ano_destino
    if atual is None or atual == destino or destino not in rows:
        return xml
    def estilos(txt):
        return re.findall(r'<c r="([A-F])\d+" s="(\d+)"', txt)
    def abre(txt):
        return re.match(r"<row [^>]*>", txt).group(0)
    a, b = rows[atual].group(0), rows[destino].group(0)
    ea, eb = dict(estilos(a)), dict(estilos(b))
    def aplica(txt, est, origem_open, n):
        for col, s in est.items():
            txt = re.sub(rf'(<c r="{col}{n}" s=")\d+(")', rf"\g<1>{s}\2", txt)
        attrs_ht = re.findall(r'\s(?:ht|customHeight)="[^"]*"', origem_open)
        op = re.sub(r'\s(?:ht|customHeight)="[^"]*"', "", abre(txt))
        op = op.replace(f'<row r="{n}"', f'<row r="{n}"' + "".join(attrs_ht), 1)
        return op + txt[len(abre(txt)):]
    na = aplica(a, eb, abre(b), atual)
    nb = aplica(b, ea, abre(a), destino)
    out = xml
    out = out.replace(a, "\0A\0").replace(b, "\0B\0")
    return out.replace("\0A\0", na).replace("\0B\0", nb)

def preencher_planilha(src, dst, r):
    tmp = Path(tempfile.mkdtemp())
    with zipfile.ZipFile(src) as z:
        z.extractall(tmp)
    wbx = (tmp / "xl/workbook.xml").read_text(encoding="utf8")
    rels = (tmp / "xl/_rels/workbook.xml.rels").read_text(encoding="utf8")
    def arq(nome):
        rid = re.search(rf'<sheet name="{nome}"[^>]*r:id="(rId\d+)"', wbx).group(1)
        alvo = re.search(rf'Id="{rid}"[^>]*Target="([^"]+)"', rels) or re.search(rf'Target="([^"]+)"[^>]*Id="{rid}"', rels)
        return tmp / "xl" / alvo.group(1)
    pb = arq("PAY BACK")
    x = pb.read_text(encoding="utf8")
    x = set_cell(x, "B6", r["tarifa"])
    x = set_cell(x, "C6", r["kwh_payback"])
    x = set_cell(x, "D5", r["reajuste"])
    x = set_cell(x, "D31", r["investimento"])
    x = set_cell(x, "A31", f"Valor do Sistema {r['kwp']:.2f} kWp")
    x = mover_destaque(x, r["ano_payback"])
    pb.write_text(limpar_cache_formulas(x), encoding="utf8")

    oc = arq("ORÇAMENTO")
    x = oc.read_text(encoding="utf8")
    x = set_cell(x, "C4", r["qtd_paineis"]); x = set_cell(x, "D4", r["txt_modulo_tabela"])
    x = set_cell(x, "C5", r["inversor_qtd"]); x = set_cell(x, "D5", r["txt_inversor_tabela"])
    x = set_cell(x, "C8", r["investimento"])
    oc.write_text(x, encoding="utf8")
    # estrutura: só troca se o tipo não estiver contemplado no texto atual
    if r["estrutura"]:
        import openpyxl  # leitura apenas
        wb = openpyxl.load_workbook(src)
        atual = str(wb["ORÇAMENTO"]["D6"].value or "")
        if r["estrutura"].lower() not in atual.lower():
            x = oc.read_text(encoding="utf8")
            x = set_cell(x, "D6", f"Kit Estrutura para superficie {r['estrutura']}")
            oc.write_text(x, encoding="utf8")
    wbx = re.sub(r"<calcPr([^>]*)/>", lambda m: "<calcPr" + re.sub(r'\s*fullCalcOnLoad="[^"]*"', "", m.group(1)) + ' fullCalcOnLoad="1"/>', wbx)
    (tmp / "xl/workbook.xml").write_text(wbx, encoding="utf8")
    cc = tmp / "xl/calcChain.xml"
    if cc.exists():  # força o Excel a refazer a cadeia de cálculo
        cc.unlink()
        ct = (tmp / "[Content_Types].xml").read_text(encoding="utf8")
        ct = re.sub(r'<Override PartName="/xl/calcChain.xml"[^>]*/>', "", ct)
        (tmp / "[Content_Types].xml").write_text(ct, encoding="utf8")
        rels2 = re.sub(r'<Relationship [^>]*calcChain[^>]*/>', "", rels)
        (tmp / "xl/_rels/workbook.xml.rels").write_text(rels2, encoding="utf8")
    _zipar(tmp, dst)
    shutil.rmtree(tmp, ignore_errors=True)

def _zipar(pasta, destino):
    destino = Path(destino)
    if destino.exists():
        destino.unlink()
    with zipfile.ZipFile(destino, "w", zipfile.ZIP_DEFLATED) as z:
        ct = pasta / "[Content_Types].xml"
        z.write(ct, "[Content_Types].xml")
        for f in sorted(pasta.rglob("*")):
            if f.is_file() and f != ct:
                z.write(f, f.relative_to(pasta).as_posix())

# --------------------------------------------------------------------------
# LibreOffice (perfil pt-BR, recálculo sempre)
# --------------------------------------------------------------------------
PERFIL = Path(tempfile.gettempdir()) / "lo_perfil_solarinfocco"
def _perfil():
    u = PERFIL / "user"
    u.mkdir(parents=True, exist_ok=True)
    (u / "registrymodifications.xcu").write_text("""<?xml version="1.0" encoding="UTF-8"?>
<oor:items xmlns:oor="http://openoffice.org/2001/registry" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<item oor:path="/org.openoffice.Setup/L10N"><prop oor:name="ooSetupSystemLocale" oor:op="fuse"><value>pt-BR</value></prop></item>
<item oor:path="/org.openoffice.Setup/L10N"><prop oor:name="ooLocale" oor:op="fuse"><value>pt-BR</value></prop></item>
<item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>0</value></prop></item>
</oor:items>""", encoding="utf8")
    return PERFIL.as_uri()

def soffice_pdf(arquivo, pasta):
    exe = soffice_exe()
    if not exe:
        raise RuntimeError("LibreOffice não encontrado. Rode o INSTALAR novamente.")
    _run([exe, f"-env:UserInstallation={_perfil()}", "--headless", "--convert-to", "pdf",
                    "--outdir", str(pasta), str(arquivo)], check=True, capture_output=True, timeout=300)
    return Path(pasta) / (Path(arquivo).stem + ".pdf")

def word_pdf(docx, pasta):
    """No Windows com Microsoft Word: gera o PDF pelo próprio Word (diagramação idêntica)."""
    if os.name != "nt":
        return None
    docx = Path(docx).resolve(); pdf = Path(pasta).resolve() / (docx.stem + ".pdf")
    q = lambda p: str(p).replace("'", "''")
    ps = (
        "$ErrorActionPreference='Stop';"
        "$w=New-Object -ComObject Word.Application;"
        "$w.DisplayAlerts=0;"
        f"$d=$w.Documents.Open('{q(docx)}',$false,$true,$false);"
        f"$d.ExportAsFixedFormat('{q(pdf)}',17);"
        "$d.Close($false);"
        "if($w.Documents.Count -eq 0){$w.Quit()}"
    )
    try:
        _run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps],
             capture_output=True, timeout=180)
    except Exception:
        return None
    return pdf if pdf.exists() else None

# --------------------------------------------------------------------------
# Renderização das tabelas (a partir da planilha já preenchida)
# --------------------------------------------------------------------------
def renderizar_tabelas(xlsx, pasta, aspectos):
    import openpyxl, warnings
    from PIL import Image
    warnings.simplefilter("ignore")
    wb = openpyxl.load_workbook(xlsx)
    for n in list(wb.sheetnames):
        if n not in ("PAY BACK", "ORÇAMENTO"):
            del wb[n]
    for nome, area in (("PAY BACK", "A1:F41"), ("ORÇAMENTO", "C3:D9")):
        ws = wb[nome]
        ws.print_area = area
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 1
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.print_options.gridLines = False
    tmpx = Path(pasta) / "_render.xlsx"
    wb.save(tmpx)
    pdf = soffice_pdf(tmpx, pasta)
    pags = pdf_para_pngs(pdf, pasta, "_tab")
    saida = {}
    for chave, png in zip(("payback", "orcamento"), pags):
        im = Image.open(png).convert("RGB")
        cinza = im.convert("L").point(lambda p: 0 if p > 245 else 255)
        bbox = cinza.getbbox()
        im = im.crop((bbox[0] - 4, bbox[1] - 4, bbox[2] + 4, bbox[3] + 4))
        alvo = aspectos[chave]  # largura/altura do quadro no Word
        w, h = im.size
        if w / h > alvo:
            nh = round(w / alvo); fundo = Image.new("RGB", (w, nh), "white"); fundo.paste(im, (0, (nh - h) // 2))
        else:
            nw = round(h * alvo); fundo = Image.new("RGB", (nw, h), "white"); fundo.paste(im, ((nw - w) // 2, 0))
        destino = Path(pasta) / f"tabela_{chave}.png"
        fundo.save(destino, dpi=(300, 300))
        saida[chave] = destino
    for f in [tmpx, pdf, *pags]:
        Path(f).unlink(missing_ok=True)
    return saida

# --------------------------------------------------------------------------
# Word
# --------------------------------------------------------------------------
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XMLNS = "http://www.w3.org/XML/1998/namespace"

def _ts_do_paragrafo(p):
    out = []
    for t in p.iter(f"{{{W}}}t"):
        anc = next(t.iterancestors(f"{{{W}}}p"), None)
        if anc is p:
            out.append(t)
    return out

def _substituir_span(ts, a, b, novo):
    pos = 0
    primeiro = True
    for t in ts:
        txt = t.text or ""
        ini, fim = pos, pos + len(txt)
        pos = fim
        if fim <= a or ini >= b:
            continue
        la, lb = max(a, ini) - ini, min(b, fim) - ini
        if primeiro:
            t.text = txt[:la] + novo + txt[lb:]
            primeiro = False
        else:
            t.text = txt[:la] + txt[lb:]
        t.set(f"{{{XMLNS}}}space", "preserve")

class EditorWord:
    def __init__(self, xml_bytes):
        from lxml import etree
        self.etree = etree
        self.root = etree.fromstring(xml_bytes)
        self.log = []

    def paragrafos(self):
        return list(self.root.iter(f"{{{W}}}p"))

    def sub(self, padrao, func, rotulo):
        """padrao: regex com grupos; func(match) -> dict {grupo: novo_texto}."""
        rx = re.compile(padrao)
        n = 0
        for p in self.paragrafos():
            ts = _ts_do_paragrafo(p)
            texto = "".join(t.text or "" for t in ts)
            ms = list(rx.finditer(texto))
            for m in reversed(ms):
                novos = func(m)
                for g in sorted(novos, reverse=True):
                    a, b = m.span(g)
                    _substituir_span(ts, a, b, novos[g])
                n += 1
        self.log.append((rotulo, n))
        return n

    def bytes(self):
        return self.etree.tostring(self.root, xml_declaration=True, encoding="UTF-8", standalone=True)

def preencher_word(modelo, destino, r, imagens):
    tmp = Path(tempfile.mkdtemp())
    with zipfile.ZipFile(modelo) as z:
        z.extractall(tmp)
    doc = tmp / "word/document.xml"
    ed = EditorWord(doc.read_bytes())
    d = r["data"]
    mes = MESES[d.month - 1]

    # Página 1 e 2 — cliente e datas
    ed.sub(r"^Cliente:\s*(.+?)\s*$", lambda m: {1: r["cliente"]}, "cliente")
    ed.sub(r"^Endereço:\s*(.+?)\s*$", lambda m: {1: r["endereco"]}, "endereço")
    ed.sub(r"(\d{1,2}) de ([a-zç]+) (\d{4})", lambda m: {1: str(d.day), 2: mes, 3: str(d.year)}, "data capa")
    ed.sub(r"(\d{1,2})/([A-ZÇ][a-zç]+)/(\d{4})", lambda m: {1: f"{d.day:02d}", 2: mes.capitalize(), 3: str(d.year)}, "data dd/Mês/aaaa")
    ed.sub(r"(\d{1,2}) */ *([A-ZÇ][a-zç]+) */ *(\d{4})", lambda m: {1: f"{d.day:02d}", 2: mes.capitalize(), 3: str(d.year)}, "data dd / Mês / aaaa")

    # Página 4 — detalhamento
    ed.sub(r"^Potência do Sistema:\s*([\d.,]+)", lambda m: {1: fmt_como(r["kwp"], m.group(1))}, "potência kWp")
    ed.sub(r"^Produção Média Esperada:\s*([\d.,]+)", lambda m: {1: br(r["geracao"], 0)}, "produção")
    ed.sub(r"^Área Estimada de ocupação:\s*([\d.,]+)", lambda m: {1: br(r["area_m2"], 0)}, "área")
    ed.sub(r"^Quantidade de Painéis:\s*(\d+)", lambda m: {1: str(r["qtd_paineis"])}, "qtd painéis")
    ed.sub(r"^Potência dos módulos:\s*(.+?)\s*$", lambda m: {1: r["txt_modulo_detalhe"]}, "potência módulos")
    ed.sub(r"^Inversor:\s*(.+?)\s*$", lambda m: {1: r["txt_inversor_detalhe"]}, "inversor")
    ed.sub(r"^Economia Anual:\s*R\$\s*([\d.,]+)", lambda m: {1: br(r["economia_anual"])}, "economia anual")
    ed.sub(r"^Tio? de Estrutura:\s*(.+?)\s*$", lambda m: {1: r["estrutura"].lower()}, "estrutura")
    ed.sub(r"^Peso aproximado:\s*([\d.,]+)", lambda m: {1: br(r["peso_kg"], 0)}, "peso")
    ed.sub(r"^Consumo médio mensal:\s*([\d.,]+)", lambda m: {1: br(r["consumo_medio"], 0)}, "consumo médio")
    ed.sub(r"^Valor da tarifa R\$\s*([\d.,]+)", lambda m: {1: br(r["tarifa"])}, "tarifa")
    ed.sub(r"^Classe:\s*(.+?)\s*$", lambda m: {1: r["classe"]}, "classe")
    ed.sub(r"^Municipio de Instalação:\s*(.+?)\s*$", lambda m: {1: r["municipio"]}, "município")
    ed.sub(r"^HSP Conforme Crescesc Sun Data:\s*([\d.,]+)", lambda m: {1: br(r["hsp"])}, "HSP")

    # Página do orçamento
    ed.sub(r"^Orçamento do Sistema\s+([\d.,]+)\s*kWp", lambda m: {1: fmt_como(r["kwp"], m.group(1))}, "título orçamento")
    ed.sub(r"geração média de ([\d.,]+) kWh/m alimentando uma conta de R\$\s*([\d.,]+)",
           lambda m: {1: br(r["geracao"], 0), 2: br(r["valor_conta"])}, "texto orçamento")
    # Condições de pagamento (parágrafo só com o valor)
    ed.sub(r"^\s*R\$\s*([\d.]+,\d{2})\s*$", lambda m: {1: br(r["investimento"])}, "condições de pagamento")

    doc.write_bytes(ed.bytes())

    # Imagens das tabelas (substitui os EMF colados do Excel)
    relp = tmp / "word/_rels/document.xml.rels"
    rels = relp.read_text(encoding="utf8")
    for chave, rid in (("orcamento", IMG_ORC_RID), ("payback", IMG_PB_RID)):
        m = re.search(rf'Id="{rid}"[^>]*Target="media/([^"]+)"', rels)
        antigo = m.group(1)
        novo = f"tabela_{chave}.png"
        shutil.copy(imagens[chave], tmp / "word/media" / novo)
        rels = rels.replace(f'Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/{antigo}"',
                            f'Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/{novo}"')
        if rels.count(f"media/{antigo}") == 0:
            (tmp / "word/media" / antigo).unlink(missing_ok=True)
    relp.write_text(rels, encoding="utf8")
    _zipar(tmp, destino)
    shutil.rmtree(tmp, ignore_errors=True)
    return ed.log

IMG_ORC_RID, IMG_PB_RID = "rId30", "rId31"

def aspectos_no_word(modelo):
    with zipfile.ZipFile(modelo) as z:
        d = z.read("word/document.xml").decode("utf8")
    out = {}
    for chave, rid in (("orcamento", IMG_ORC_RID), ("payback", IMG_PB_RID)):
        i = d.find(f'r:embed="{rid}"')
        s = d.rfind("<w:drawing>", 0, i)
        cx, cy = map(int, re.search(r'<wp:extent cx="(\d+)" cy="(\d+)"', d[s:i]).groups())
        out[chave] = cx / cy
    return out

# --------------------------------------------------------------------------
def resumo_md(conta, orc, r, log):
    L = [f"# Resumo da proposta — {r['cliente']}", "",
         f"Gerada em {r['data'].strftime('%d/%m/%Y')}", "", "## Dados lidos",
         f"- Conta ({conta['_modo']}): {r['cliente']} · {r['endereco']} · {r['municipio']} · {r['classe']}",
         f"- Consumo médio (últimos 12 meses): {br(r['consumo_medio'], 0)} kWh/mês · conta fora taxa: R$ {br(r['valor_conta'])}",
         f"- Orçamento ({orc['_modo']}): {orc.get('cotacao')} · {br(r['kwp'])} kWp · {r['qtd_paineis']} × {r['modulo_w']} W {r['modulo_marca']} · "
         f"{r['inversor_qtd']} × {r['txt_inversor_detalhe']} · estrutura {r['estrutura']}",
         f"- Valor da distribuidora: R$ {br(r['custo_distribuidora'])}", "",
         "## Calculado",
         f"- Investimento (distribuidora × margem): **R$ {br(r['investimento'])}**",
         f"- Geração estimada: {br(r['geracao'], 0)} kWh/mês (kWp × HSP {br(r['hsp'])} × 30 × PR)",
         f"- kWh usado no payback: {br(r['kwh_payback'], 0)} · tarifa R$ {br(r['tarifa'])} · reajuste {r['reajuste']*100:.0f}% a.a.",
         f"- Economia anual (ano 1): R$ {br(r['economia_anual'])} · retorno no ano {r['ano_payback']}",
         f"- Área ≈ {r['area_m2']} m² · peso ≈ {r['peso_kg']} kg", "",
         "## Campos alterados no Word"]
    L += [f"- {k}: {n} ocorrência(s)" + ("  ⚠️ NÃO ENCONTRADO" if n == 0 else "") for k, n in log]
    L += ["", "## Pontos de atenção"] + ([f"- {a}" for a in r["alertas"]] or ["- Nenhum."])
    return "\n".join(L) + "\n"

def _padrao(*nomes):
    for n in nomes:
        for base in (APP_DIR, APP_DIR.parent):
            for pasta in ("modelo", "Modelo"):
                f = base / pasta / n
                if f.exists():
                    return f
    return APP_DIR / "modelo" / nomes[0]

def gerar(conta_pdf, orcamento_pdf, saida, data=None, modelo=None, planilha=None, config=None, progresso=print):
    """Gera tudo e devolve dict com caminhos, resumo e alertas."""
    modelo = Path(modelo or _padrao("modelo_proposta.docx"))
    planilha = Path(planilha or _padrao("payback.xlsx"))
    cfg = json.loads(Path(config or APP_DIR / "config.json").read_text(encoding="utf8"))
    data = data or dt.date.today()
    saida = Path(saida); saida.mkdir(parents=True, exist_ok=True)

    progresso("Lendo a conta de energia...")
    conta = ler_conta(conta_pdf)
    progresso("Lendo o orçamento da distribuidora...")
    orc = ler_orcamento(orcamento_pdf)
    r = calcular(conta, orc, cfg, data)

    nome = re.sub(r"[^\w]+", "_", r["cliente"]).strip("_") or "cliente"
    progresso("Calculando o payback na planilha...")
    xlsx = saida / f"Payback_{nome}.xlsx"
    preencher_planilha(planilha, xlsx, r)
    progresso("Montando as tabelas...")
    imgs = renderizar_tabelas(xlsx, saida, aspectos_no_word(modelo))
    progresso("Preenchendo a proposta...")
    docx = saida / f"Proposta_{nome}.docx"
    log = preencher_word(modelo, docx, r, imgs)
    for f in imgs.values():
        Path(f).unlink(missing_ok=True)
    pdf = None
    progresso("Gerando o PDF...")
    pdf = word_pdf(docx, saida)          # 1º: Microsoft Word (fica igual ao modelo)
    if not pdf:
        try:
            pdf = soffice_pdf(docx, saida)   # 2º: LibreOffice (pode variar a diagramação)
        except Exception:
            pdf = None
    resumo = resumo_md(conta, orc, r, log)
    texto = re.sub(r"^#+\s*", "", resumo, flags=re.M).replace("**", "")
    (saida / f"Resumo_{nome}.txt").write_text(texto, encoding="utf-8-sig")
    return {"docx": docx, "pdf": pdf, "xlsx": xlsx, "resumo": resumo, "alertas": r["alertas"],
            "calculo": r, "nao_encontrados": [k for k, n in log if n == 0]}

def main():
    ap = argparse.ArgumentParser(description="Gera a proposta Solar In Focco")
    ap.add_argument("--conta", required=True)
    ap.add_argument("--orcamento", required=True)
    ap.add_argument("--modelo")
    ap.add_argument("--planilha")
    ap.add_argument("--config")
    ap.add_argument("--saida", default=str(APP_DIR / "saida"))
    ap.add_argument("--data", help="dd/mm/aaaa (padrão: hoje)")
    a = ap.parse_args()
    data = dt.datetime.strptime(a.data, "%d/%m/%Y").date() if a.data else None
    res = gerar(a.conta, a.orcamento, a.saida, data, a.modelo, a.planilha, a.config)
    print(res["resumo"])

if __name__ == "__main__":
    main()
