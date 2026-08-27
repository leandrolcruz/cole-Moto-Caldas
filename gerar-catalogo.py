#!/usr/bin/env python3
"""Gera catalogo.csv + catalogo.meta.json a partir da Listagem de Mercadoria do
MicroWork Cloud (CSV ';' com colunas Código Mercadoria; Descrição Mercadoria;
Localizações).

Consolida linhas repetidas por loja (fica a descrição mais longa e a união dos
locais), remove colchetes/[SEM LOCAL] e grava no formato que o app consome:
uma peça por linha, "CODIGO;DESCRICAO;LOC1,LOC2".

Uso: python3 gerar-catalogo.py "Listagem de Mercadoria.csv" [versao]
     (versao default = data de hoje AAAA-MM-DD)
"""
import csv, json, re, sys, datetime, os

if len(sys.argv) < 2:
    sys.exit("uso: gerar-catalogo.py <listagem.csv> [versao]")
src = sys.argv[1]
versao = (sys.argv[2] if len(sys.argv) > 2
          else datetime.datetime.now().strftime("%Y-%m-%d %H:%M"))
dest_dir = os.path.dirname(os.path.abspath(__file__))

pecas = {}
with open(src, encoding="utf-8-sig") as fh:
    r = csv.reader(fh, delimiter=";")
    next(r)  # cabeçalho
    for row in r:
        if len(row) < 3:
            continue
        cod = row[0].strip()
        desc = row[1].strip().replace(";", ",").replace("\t", " ")
        loc = row[2].strip()
        if not cod or len(cod) < 3:
            continue
        toks = [t for t in re.split(r"[\[\],]+", loc) if t and t != "SEM LOCAL"]
        d, ls = pecas.get(cod, ("", []))
        if len(desc) > len(d):
            d = desc
        for t in toks:
            if t not in ls:
                ls.append(t)
        pecas[cod] = (d, ls)

if len(pecas) < 1000:
    sys.exit(f"apenas {len(pecas)} peças — arquivo de entrada suspeito, abortando")

out_csv = os.path.join(dest_dir, "catalogo.csv")
with open(out_csv, "w", encoding="utf-8", newline="\n") as fh:
    for cod in sorted(pecas):
        d, ls = pecas[cod]
        fh.write(f"{cod};{d};{','.join(ls)}\n")

meta = {"versao": versao, "pecas": len(pecas),
        "com_local": sum(1 for _, ls in pecas.values() if ls),
        "hifen": sum(1 for c in pecas if "-" in c)}
with open(os.path.join(dest_dir, "catalogo.meta.json"), "w", encoding="utf-8") as fh:
    json.dump(meta, fh, ensure_ascii=False)

print(f"catalogo.csv: {len(pecas)} peças ({meta['com_local']} com local, "
      f"{meta['hifen']} com hífen) — versão {versao}")
