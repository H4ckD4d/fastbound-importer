#!/usr/bin/env python3
"""
fastbound_importer.py
---------------------
Lê dados de um A&D (aba/planilha fonte) e preenche uma planilha no formato FastBound
(aba/planilha destino), preservando os cabeçalhos e a ordem das colunas da planilha
FastBound. Gera também um relatório de mapeamento.

Compatível com macOS, Linux e Windows.
Requer: pandas, openpyxl, xlsxwriter, pyyaml (opcional para YAML)

Uso básico:
  python fastbound_importer.py \
    --atf /caminho/ATF-Firearms-AD-Record.xlsx --atf-sheet "ATF A&D Record" \
    --fastbound "/caminho/FastBoundImport Live - By Chris.xlsx" --fastbound-sheet "FastBoundImport Live - By Chris" \
    --out "/caminho/FastBoundImport_Populado.xlsx"

Com mapeamento manual (CSV/YAML/JSON):
  python fastbound_importer.py ... --map overrides.csv

Formato de overrides (CSV):
  FastBound Column,ATF Source
  Manufacturer,Maker
  Model,Model
  Serial Number,Serial

Autor: você 😄
"""

import argparse
import logging
from difflib import get_close_matches

import pandas as pd
import numpy as np
from pathlib import Path

try:
    import yaml  # opcional
    HAS_YAML = True
except Exception:
    HAS_YAML = False

def norm(s: str) -> str:
    return ''.join(ch for ch in str(s).lower() if ch.isalnum())

ALIASES = {
    # identidade
    "serialnumber": ["serial", "serialnumber", "s/n", "sn"],
    "manufacturer": ["manufacturer", "maker", "make"],
    "importer": ["importer"],
    "model": ["model", "mdl"],
    "type": ["type", "firearmtype", "guntype"],
    "caliber": ["caliber", "calibre", "gauge"],
    "barrellength": ["barrellength", "barrel", "bbl", "lengthofbarrel", "barrellength(in)"],
    "overalllength": ["overalllength", "oal"],
    "finish": ["finish", "color", "colour"],
    "countryofmanufacture": ["countryofmanufacture", "country", "manufacturecountry"],
    "upc": ["upc", "barcode"],
    "sku": ["sku", "item#", "itemnumber", "pn", "partnumber"],
    # aquisição
    "acquisitiondate": ["acquisitiondate", "dateacquired", "dateofacquisition", "receiveddate"],
    "acquiredfromname": ["acquiredfromname", "supplier", "vendor", "acquiredfrom"],
    "acquiredfromaddress": ["acquiredfromaddress", "supplieraddress", "vendoraddress"],
    "acquiredfromffl": ["acquiredfromffl", "supplierffl", "vendorffl", "ffl"],
    "acquiredfromlicensetype": ["acquiredfromlicensetype", "supplierlicensetype"],
    "acquiredfromcity": ["acquiredfromcity", "suppliercity"],
    "acquiredfromstate": ["acquiredfromstate", "supplierstate"],
    "acquiredfromzip": ["acquiredfromzip", "supplierzip", "zipcode", "zip"],
    "acquisitiondocument": ["acquisitiondocument","invoice","ponumber","po","bo"],
    # disposição
    "dispositiondate": ["dispositiondate","dateofdisposition","datesold","transferdate","disposeddate"],
    "disposedtoname": ["disposedtoname","customername","buyername","transfereename"],
    "disposedtoaddress": ["disposedtoaddress","customeraddress","buyeraddress"],
    "disposedtocity": ["disposedtocity","customercity"],
    "disposedtostate": ["disposedtostate","customerstate"],
    "disposedtozip": ["disposedtozip","customerzip","zip"],
    "disposedtoffl": ["disposedtoffl","destinationffl","receiverffl","ffl"],
    "form4473": ["4473","form4473","4473number","4473#"],
    "nicsnumber": ["nicsnumber","nics","ntn","backgroundchecknumber"],
    "nicsstatus": ["nicsstatus","backgroundstatus","status"],
    "nicsexpiration": ["nicsexpiration","nicsvaliduntil","ntnexpire"],
    "transactionnumber": ["transactionnumber","trans#","ttn","ntn"],
    "permitnumber": ["permit","permitnumber","cwfl","ccw"],
    "permitexpiration": ["permitexpiration","permitexpires"],
    "birthdate": ["dob","dateofbirth","birthdate"],
    "idnumber": ["idnumber","driverlicense","dl","identificationnumber"],
    "idstate": ["idstate","dlstate"],
    # valores
    "cost": ["cost","unitcost"],
    "price": ["price","saleprice","sellingprice","amount"],
}

def read_overrides(path: Path):
    """
    Lê overrides de mapeamento:
    - CSV com colunas: FastBound Column,ATF Source
    - JSON: { "FastBound Column": "ATF Source", ... }
    - YAML: idem JSON
    Retorna dict[str, str]
    """
    if not path:
        return {}
    p = Path(path)
    if not p.exists():
        raise FileNotFoundError(f"Arquivo de mapeamento não encontrado: {p}")
    if p.suffix.lower() == ".csv":
        df = pd.read_csv(p)
        if not set(df.columns).issuperset({"FastBound Column","ATF Source"}):
            raise ValueError("CSV de override deve ter colunas 'FastBound Column' e 'ATF Source'.")
        return {str(r["FastBound Column"]).strip(): str(r["ATF Source"]).strip() for _, r in df.iterrows()}
    elif p.suffix.lower() in (".json",):
        return json.loads(p.read_text(encoding="utf-8"))
    elif p.suffix.lower() in (".yml",".yaml"):
        if not HAS_YAML:
            raise RuntimeError("pyyaml não instalado. Rode: pip install pyyaml")
        return yaml.safe_load(p.read_text(encoding="utf-8"))
    else:
        raise ValueError("Formato de override não suportado. Use CSV, JSON ou YAML.")

def build_mapping(atf_cols, fb_cols, overrides=None, fuzzy_cutoff=0.84, logger=None):
    atf_norm_map = {norm(c): c for c in atf_cols}
    mapping = {}
    details = []

    # Aplicar overrides primeiro
    overrides = overrides or {}
    for fb_col, atf_src in overrides.items():
        # permitir apontar por nome "natural" da coluna ATF
        if atf_src in atf_cols:
            mapping[fb_col] = atf_src
            details.append((fb_col, atf_src, "OVERRIDE"))
        else:
            # tentar via normalização
            key = norm(atf_src)
            if key in atf_norm_map:
                mapping[fb_col] = atf_norm_map[key]
                details.append((fb_col, atf_norm_map[key], "OVERRIDE(norm)"))
            else:
                mapping[fb_col] = None
                details.append((fb_col, "", "OVERRIDE-NOTFOUND"))

    # Preencher demais via match direto, alias e fuzzy
    for fb_col in fb_cols:
        if fb_col in mapping:
            continue
        fb_key = norm(fb_col)
        # match direto
        if fb_key in atf_norm_map:
            mapping[fb_col] = atf_norm_map[fb_key]
            details.append((fb_col, atf_norm_map[fb_key], "DIRECT"))
            continue
        # alias
        hit = None
        for target, alias_list in ALIASES.items():
            if fb_key == target or fb_key in map(norm, alias_list+[target]):
                for a in alias_list+[target]:
                    k = norm(a)
                    if k in atf_norm_map:
                        hit = atf_norm_map[k]
                        break
            if hit:
                break
        if hit:
            mapping[fb_col] = hit
            details.append((fb_col, hit, "ALIAS"))
            continue
        # fuzzy
        cand = get_close_matches(fb_key, list(atf_norm_map.keys()), n=1, cutoff=fuzzy_cutoff)
        if cand:
            mapping[fb_col] = atf_norm_map[cand[0]]
            details.append((fb_col, atf_norm_map[cand[0]], "FUZZY"))
        else:
            mapping[fb_col] = None
            details.append((fb_col, "", "MISSING"))
    # log resumo
    if logger:
        total = len(fb_cols)
        ok = sum(1 for c in fb_cols if mapping.get(c))
        logger.info(f"Mapeados: {ok}/{total} colunas FastBound")
    return mapping, details

def main():
    ap = argparse.ArgumentParser(description="Preencher planilha FastBound a partir de A&D (ATF).")
    ap.add_argument("--atf", required=True, help="Caminho do Excel ATF (entrada).")
    ap.add_argument("--atf-sheet", default=None, help="Nome da aba no ATF (padrão: primeira).")
    ap.add_argument("--fastbound", required=True, help="Caminho do Excel FastBound (template).")
    ap.add_argument("--fastbound-sheet", default=None, help="Nome da aba no FastBound (padrão: primeira).")
    ap.add_argument("--out", required=True, help="Caminho do Excel de saída preenchido.")
    ap.add_argument("--map", dest="overrides", default=None, help="CSV/JSON/YAML com overrides de mapeamento.")
    ap.add_argument("--strict", action="store_true", help="Falhar (exit 2) se houver colunas FastBound sem origem.")
    ap.add_argument("--fuzzy-cutoff", type=float, default=0.84, help="Corte de similaridade para fuzzy (0-1).")
    ap.add_argument("--verbose", action="store_true", help="Log detalhado.")
    args = ap.parse_args()

    logging.basicConfig(level=logging.DEBUG if args.verbose else logging.INFO, format="%(levelname)s: %(message)s")
    log = logging.getLogger("fastbound")

    atf_path = Path(args.atf)
    fb_path = Path(args.fastbound)
    out_path = Path(args.out)

    if not atf_path.exists():
        log.error(f"ATF não encontrado: {atf_path}")
        sys.exit(1)
    if not fb_path.exists():
        log.error(f"Template FastBound não encontrado: {fb_path}")
        sys.exit(1)

    # Ler planilhas
    if args.atf_sheet:
        atf_df = pd.read_excel(atf_path, sheet_name=args.atf_sheet)
    else:
        atf_df = pd.read_excel(atf_path)  # primeira aba
        log.info(f"ATF sheet não especificada; usando primeira aba: {atf_df.shape}")

    if args.fastbound_sheet:
        fastbound_df = pd.read_excel(fb_path, sheet_name=args.fastbound_sheet, nrows=0)
        fb_sheetname = args.fastbound_sheet
    else:
        # usar primeira aba como layout
        with pd.ExcelFile(fb_path) as xls:
            fb_sheetname = xls.sheet_names[0]
        fastbound_df = pd.read_excel(fb_path, sheet_name=fb_sheetname, nrows=0)
        log.info(f"FastBound sheet não especificada; usando '{fb_sheetname}'")

    fb_columns = list(fastbound_df.columns)
    atf_columns = list(atf_df.columns)

    # Overrides
    overrides = read_overrides(Path(args.overrides)) if args.overrides else {}

    mapping, details = build_mapping(atf_columns, fb_columns, overrides=overrides, fuzzy_cutoff=args.fuzzy_cutoff, logger=log)

    # Construir saída
    out_df = pd.DataFrame(columns=fb_columns)
    for fb_col, src in mapping.items():
        if src and src in atf_df.columns:
            out_df[fb_col] = atf_df[src]
        else:
            out_df[fb_col] = np.nan

    # Salvar com sheets: FastBoundImport, Mapping Report, Missing & Guidance
    with pd.ExcelWriter(out_path, engine="xlsxwriter") as writer:
        out_df.to_excel(writer, sheet_name="FastBoundImport", index=False)
        rep_df = pd.DataFrame(details, columns=["FastBound Column","ATF Source","Match Type"])
        rep_df.to_excel(writer, sheet_name="Mapping Report", index=False)

        # Missing guidance
        missing = [c for c in fb_columns if not mapping.get(c)]
        guidance_rows = []
        for col in missing:
            key = col.lower()
            hints = []
            if any(k in key for k in ["serial", "sn"]):
                hints.append("Número de série – marcação na arma / 4473 / registro anterior.")
            if any(k in key for k in ["manufacturer","mfr","maker","make"]):
                hints.append("Fabricante – marcação do frame/receiver; nota do fornecedor.")
            if any(k in key for k in ["importer"]):
                hints.append("Importador – marcação no cano/frame; invoice de entrada.")
            if any(k in key for k in ["model"]):
                hints.append("Modelo – marcação na arma; invoice/packing slip.")
            if any(k in key for k in ["caliber","gauge"]):
                hints.append("Calibre/ga – gravado no cano/slide; ficha do fabricante.")
            if any(k in key for k in ["type"]):
                hints.append("Tipo – pistol/revolver/rifle/shotgun/receiver/other.")
            if any(k in key for k in ["barrel","length","oal"]):
                hints.append("Comprimento de cano/OAL – ficha técnica; inspeção.")
            if any(k in key for k in ["finish","color"]):
                hints.append("Acabamento/cor – inspeção / descrição do fabricante.")
            if any(k in key for k in ["upc","sku"]):
                hints.append("UPC/SKU – caixa do produto; invoice.")
            if any(k in key for k in ["acq","acquisition","received","source","supplier","vendor"]):
                hints.append("Dados de aquisição – fornecedor, data, invoice/PO.")
            if any(k in key for k in ["dispo","dispose","transferee","customer","buyer","4473"]):
                hints.append("Dados de disposição – cliente/FFL, data, 4473, NICS.")
            if any(k in key for k in ["nics","ttn","poc","background"]):
                hints.append("Background – NICS/POC (número/status/expiração).")
            if any(k in key for k in ["ffl","license"]):
                hints.append("FFL – nº/expiração do destinatário; cópia arquivada.")
            if any(k in key for k in ["cost","price","amount","msrp"]):
                hints.append("Valores – custo/preço; ERP/nota fiscal.")
            if not hints:
                hints.append("Verificar notas de entrada, 4473, cópia de FFL e marcações físicas.")
            guidance_rows.append({"Missing FastBound Column": col, "Como obter": " | ".join(hints)})
        gd = pd.DataFrame(guidance_rows) if guidance_rows else pd.DataFrame(columns=["Missing FastBound Column","Como obter"])
        gd.to_excel(writer, sheet_name="Missing & Guidance", index=False)

    # Saída de status para automações/CI
    missing_count = sum(1 for c in fb_columns if not mapping.get(c))
    if missing_count and args.strict:
        print(f"ERRO: {missing_count} colunas do FastBound não foram mapeadas. Veja 'Mapping Report'.", file=sys.stderr)
        sys.exit(2)

    log.info(f"Concluído. Saída: {out_path}")
    log.info(f"Colunas FastBound: {len(fb_columns)} | Mapeadas: {len(fb_columns)-missing_count} | Faltando: {missing_count}")

if __name__ == "__main__":
    main()
