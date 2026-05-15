from pathlib import Path
from datetime import datetime
from collections import defaultdict
import csv
import json
import re
import unicodedata
import zipfile
import xml.etree.ElementTree as ET

ARQUIVO_EMPRESA = "Programa Conecta - Acompanhamento interno.xlsx"
ARQUIVO_CONSULTORIA = "AN.020 - Acompanhamento RICEs.xlsx"
ABA_DADOS = "Dados Consolidados"
SAIDA_CSV = "dados_consilidados_rices.csv"
SAIDA_JSON = "dados_consilidados_rices.json"
SAIDA_RESUMO = "resumo_comparacao_rices.csv"

COLUNAS_PADRAO = [
    "Identificação da tarefa", "Nome da tarefa", "Categoria", "Meta", "Status", "Prioridade", "Atribuído a",
    "Criado por", "Criado em", "Data de conclusão", "Data de início", "É Recorrente", "Atrasados",
    "Concluído em", "Concluída por", "Itens concluídos da lista de verificação", "Itens da lista de verificação",
    "Rótulos", "Notas"
]

PREFIXOS_VALIDOS = {
    "ACE", "ADJ", "AP", "APO", "CEP", "CON", "CPO", "CPR", "CRF", "CUS", "CVP", "FAB", "FFO", "FRG",
    "GAF", "GCI", "GCO", "GES", "GM", "GMS", "GQ", "GRF", "INC", "INDU", "INV", "ITG", "KT", "LOP",
    "MAN", "NSI", "OIC", "OTM", "PFC", "PPE", "PPO", "PPX", "PVF", "RPT", "SUP", "TAX", "TES", "TRP", "W"
}

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
RNS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


def remover_acentos(valor):
    texto = str(valor or "")
    texto = unicodedata.normalize("NFD", texto)
    return "".join(c for c in texto if unicodedata.category(c) != "Mn")


def normalizar_texto(valor):
    texto = remover_acentos(valor).upper().strip()
    texto = texto.replace("–", "-").replace("—", "-")
    texto = re.sub(r"\s+", " ", texto)
    return texto


def prefixo_codigo(codigo):
    texto = normalizar_texto(codigo)
    m = re.match(r"([A-Z]+)", texto)
    return m.group(1) if m else ""


def chave_rice(codigo):
    texto = normalizar_texto(codigo)
    texto = re.sub(r"\s*#\s*", "#", texto)
    texto = re.sub(r"[^A-Z0-9#._]", "", texto)
    return texto


def codigo_exibicao(codigo):
    texto = normalizar_texto(codigo)
    texto = re.sub(r"\s*#\s*", "#", texto)
    texto = re.sub(r"\s*\.\s*", ".", texto)
    texto = re.sub(r"\s*-\s*", "-", texto)
    texto = re.sub(r"\s+", " ", texto)
    return texto.strip(" -|;/,")


def preparar_texto_codigos(nome_tarefa):
    texto = normalizar_texto(nome_tarefa)
    texto = re.sub(r"(?<=\d)I(?=[A-Z]{2,}[.#]?#?\d)", "|", texto)
    texto = texto.replace(" I ", "|")
    return texto


def extrair_rices(nome_tarefa):
    texto = preparar_texto_codigos(nome_tarefa)
    candidatos = []

    for m in re.finditer(r"([A-Z]{2,}(?:[.\-][A-Z0-9]+)*(?:\s+\d+)?)#(\d+)((?:\s*/\s*#?\d+)+)", texto):
        prefixo = m.group(1)
        candidatos.append(f"{prefixo}#{m.group(2)}")
        for numero in re.findall(r"#?(\d+)", m.group(3)):
            candidatos.append(f"{prefixo}#{numero}")

    padrao = re.compile(
        r"(?<![A-Z0-9])"
        r"(?:[A-Z]{1,}(?:[.\-][A-Z0-9]+)*|[A-Z]{1,3})"
        r"(?:\s*-\s*|\s+)?#?\d+[A-Z0-9]*"
        r"(?:[._\-]\d+[A-Z0-9]*)*"
        r"(?:#\d+)?"
        r"(?:\s+[A-Z]{1,3})?"
    )

    for m in padrao.finditer(texto):
        candidatos.append(m.group(0))

    resultado = []
    vistos = set()
    for candidato in candidatos:
        display = codigo_exibicao(candidato)
        pref = prefixo_codigo(display)
        key = chave_rice(display)
        if not key or key in vistos:
            continue
        if pref not in PREFIXOS_VALIDOS:
            continue
        if not re.search(r"\d", key):
            continue
        vistos.add(key)
        resultado.append((key, display))
    return resultado


def coluna_para_indice(ref):
    m = re.match(r"([A-Z]+)", ref or "")
    if not m:
        return 0
    total = 0
    for ch in m.group(1):
        total = total * 26 + ord(ch) - 64
    return total - 1


def carregar_shared_strings(zip_xlsx):
    if "xl/sharedStrings.xml" not in zip_xlsx.namelist():
        return []
    raiz = ET.fromstring(zip_xlsx.read("xl/sharedStrings.xml"))
    valores = []
    for si in raiz.findall(NS + "si"):
        partes = []
        for t in si.iter(NS + "t"):
            partes.append(t.text or "")
        valores.append("".join(partes))
    return valores


def caminho_aba(zip_xlsx, nome_aba):
    ns = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main", "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
    workbook = ET.fromstring(zip_xlsx.read("xl/workbook.xml"))
    rels = ET.fromstring(zip_xlsx.read("xl/_rels/workbook.xml.rels"))
    mapa_rels = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels}
    primeira = None
    for sheet in workbook.find("a:sheets", ns):
        if primeira is None:
            primeira = sheet
        if sheet.attrib.get("name") == nome_aba:
            target = mapa_rels[sheet.attrib.get(RNS + "id")]
            return "xl/" + target if not target.startswith("xl/") else target
    target = mapa_rels[primeira.attrib.get(RNS + "id")]
    return "xl/" + target if not target.startswith("xl/") else target


def ler_planilha(caminho, origem):
    with zipfile.ZipFile(caminho) as zip_xlsx:
        shared = carregar_shared_strings(zip_xlsx)
        sheet_path = caminho_aba(zip_xlsx, ABA_DADOS)
        raiz = ET.fromstring(zip_xlsx.read(sheet_path))
        linhas = []
        for row in raiz.iter(NS + "row"):
            temporario = []
            max_col = -1
            for cell in row.findall(NS + "c"):
                idx = coluna_para_indice(cell.attrib.get("r", "A1"))
                max_col = max(max_col, idx)
                tipo = cell.attrib.get("t")
                valor = ""
                v = cell.find(NS + "v")
                if tipo == "s":
                    valor = shared[int(v.text)] if v is not None and v.text is not None else ""
                elif tipo == "inlineStr":
                    valor = "".join(t.text or "" for t in cell.iter(NS + "t"))
                else:
                    valor = v.text if v is not None and v.text is not None else ""
                temporario.append((idx, valor))
            valores = [""] * (max_col + 1)
            for idx, valor in temporario:
                valores[idx] = valor
            linhas.append(valores)

    if not linhas:
        return []

    cabecalho = [str(c or "").strip() for c in linhas[0]]
    dados = []
    for linha in linhas[1:]:
        item = {cabecalho[i]: linha[i] if i < len(linha) else "" for i in range(len(cabecalho)) if cabecalho[i]}
        if not any(str(v).strip() for v in item.values()):
            continue
        item["__origem"] = origem
        item["__arquivo"] = Path(caminho).name
        dados.append(item)
    return dados


def ordem_fase(categoria):
    m = re.match(r"\s*(\d+)", str(categoria or ""))
    return int(m.group(1)) if m else 999


def status_norm(status):
    return normalizar_texto(status)


def status_concluido(status):
    return "CONCLUID" in status_norm(status)


def status_andamento(status):
    return "ANDAMENTO" in status_norm(status)


def status_nao_iniciado(status):
    return "NAO INICIADO" in status_norm(status)


def flag_atrasado(valor):
    return normalizar_texto(valor) in {"TRUE", "VERDADEIRO", "SIM", "1", "YES"}


def linha_mais_representativa(linhas):
    if not linhas:
        return {}
    abertas = [l for l in linhas if not status_concluido(l.get("Status"))]
    base = abertas if abertas else linhas
    return sorted(
        base,
        key=lambda l: (
            ordem_fase(l.get("Categoria")) if abertas else -ordem_fase(l.get("Categoria")),
            0 if flag_atrasado(l.get("Atrasados")) else 1,
            str(l.get("Identificação da tarefa", ""))
        )
    )[0]


def consolidar_origem(linhas_por_rice, origem):
    saida = {}
    for key, registros in linhas_por_rice.items():
        representante = linha_mais_representativa(registros)
        fases = sorted({str(r.get("Categoria", "")).strip() for r in registros if str(r.get("Categoria", "")).strip()}, key=ordem_fase)
        responsaveis = []
        for r in registros:
            for pessoa in str(r.get("Atribuído a", "")).split(";"):
                pessoa = pessoa.strip()
                if pessoa and pessoa not in responsaveis:
                    responsaveis.append(pessoa)
        display = registros[0].get("__rice_display", key)
        linha = {col: representante.get(col, "") for col in COLUNAS_PADRAO}
        linha.update({
            "RICE": display,
            "RICE_KEY": key,
            "Origem": origem,
            "__origem": origem,
            "__arquivo": representante.get("__arquivo", ""),
            "Qtd registros na origem": len(registros),
            "Fases encontradas na origem": " | ".join(fases),
            "Responsáveis consolidados": "; ".join(responsaveis),
            "Status consolidado origem": representante.get("Status", ""),
            "Fase consolidada origem": representante.get("Categoria", "")
        })
        if not linha.get("Atribuído a") and responsaveis:
            linha["Atribuído a"] = "; ".join(responsaveis)
        if display and display not in str(linha.get("Nome da tarefa", ""))[:80].upper():
            linha["Nome da tarefa"] = f"{display} - {linha.get('Nome da tarefa', '')}".strip(" -")
        saida[key] = linha
    return saida


def montar_base():
    base_dir = Path(__file__).resolve().parent
    empresa_path = base_dir / ARQUIVO_EMPRESA
    consultoria_path = base_dir / ARQUIVO_CONSULTORIA

    if not empresa_path.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {empresa_path.name}")
    if not consultoria_path.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {consultoria_path.name}")

    empresa_linhas = ler_planilha(empresa_path, "Empresa")
    consultoria_linhas = ler_planilha(consultoria_path, "Consultoria")

    agrupado_empresa = defaultdict(list)
    agrupado_consultoria = defaultdict(list)

    for origem, linhas, destino in [
        ("Empresa", empresa_linhas, agrupado_empresa),
        ("Consultoria", consultoria_linhas, agrupado_consultoria)
    ]:
        for linha in linhas:
            codigos = extrair_rices(linha.get("Nome da tarefa", ""))
            for key, display in codigos:
                nova = dict(linha)
                nova["__rice_key"] = key
                nova["__rice_display"] = display
                destino[key].append(nova)

    empresa = consolidar_origem(agrupado_empresa, "Empresa")
    consultoria = consolidar_origem(agrupado_consultoria, "Consultoria")

    keys_empresa = set(empresa)
    keys_consultoria = set(consultoria)
    iguais = keys_empresa & keys_consultoria
    so_empresa = keys_empresa - keys_consultoria
    so_consultoria = keys_consultoria - keys_empresa
    todos = sorted(keys_empresa | keys_consultoria)

    linhas_csv = []
    for key in todos:
        tipo = "RICE igual" if key in iguais else "Só na Empresa" if key in so_empresa else "Só na Consultoria"
        presente_empresa = "Sim" if key in keys_empresa else "Não"
        presente_consultoria = "Sim" if key in keys_consultoria else "Não"
        display = (empresa.get(key) or consultoria.get(key) or {}).get("RICE", key)
        for origem, mapa in [("Empresa", empresa), ("Consultoria", consultoria)]:
            if key not in mapa:
                continue
            linha = dict(mapa[key])
            linha.update({
                "Tipo comparação": tipo,
                "Presente Empresa": presente_empresa,
                "Presente Consultoria": presente_consultoria,
                "Total registros Empresa": len(agrupado_empresa.get(key, [])),
                "Total registros Consultoria": len(agrupado_consultoria.get(key, [])),
                "RICE base": display
            })
            linhas_csv.append(linha)

    resumo = {
        "gerado_em": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "arquivo_empresa": ARQUIVO_EMPRESA,
        "arquivo_consultoria": ARQUIVO_CONSULTORIA,
        "linhas_empresa_origem": len(empresa_linhas),
        "linhas_consultoria_origem": len(consultoria_linhas),
        "rices_empresa_sem_repetir": len(keys_empresa),
        "rices_consultoria_sem_repetir": len(keys_consultoria),
        "rices_iguais": len(iguais),
        "rices_so_empresa": len(so_empresa),
        "rices_so_consultoria": len(so_consultoria),
        "total_rices_sem_repetir_duas_bases": len(todos),
        "linhas_csv_geradas": len(linhas_csv)
    }
    return linhas_csv, resumo


def salvar_csv(caminho, linhas):
    extras = [
        "RICE", "RICE_KEY", "RICE base", "Tipo comparação", "Presente Empresa", "Presente Consultoria", "Origem",
        "Qtd registros na origem", "Fases encontradas na origem", "Responsáveis consolidados", "Status consolidado origem",
        "Fase consolidada origem", "Total registros Empresa", "Total registros Consultoria", "__origem", "__arquivo"
    ]
    colunas = extras + [c for c in COLUNAS_PADRAO if c not in extras]
    existentes = []
    for c in colunas:
        if c not in existentes:
            existentes.append(c)
    for linha in linhas:
        for c in linha:
            if c not in existentes:
                existentes.append(c)
    with open(caminho, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=existentes, delimiter=";")
        writer.writeheader()
        writer.writerows(linhas)


def salvar_resumo(caminho, resumo):
    with open(caminho, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f, delimiter=";")
        writer.writerow(["Indicador", "Valor"])
        for k, v in resumo.items():
            writer.writerow([k, v])


def main():
    base_dir = Path(__file__).resolve().parent
    linhas, resumo = montar_base()
    salvar_csv(base_dir / SAIDA_CSV, linhas)
    salvar_resumo(base_dir / SAIDA_RESUMO, resumo)
    with open(base_dir / SAIDA_JSON, "w", encoding="utf-8") as f:
        json.dump({"metadata": resumo, "rows": linhas}, f, ensure_ascii=False, indent=2)

    print("Conversão concluída.")
    print(f"CSV site: {SAIDA_CSV}")
    print(f"JSON apoio: {SAIDA_JSON}")
    print(f"Resumo: {SAIDA_RESUMO}")
    print("\nResumo executivo:")
    for k, v in resumo.items():
        print(f"- {k}: {v}")


if __name__ == "__main__":
    main()
