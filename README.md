# FastBound Importer (ATF A&D → FastBound)

Script CLI para preencher um template do **FastBound** usando dados da planilha **ATF A&D**.

## Instalação (macOS)
```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
pip install -r requirements.txt
```

## Uso básico
```bash
python fastbound_importer.py \
  --atf "/caminho/ATF-Firearms-AD-Record.xlsx" --atf-sheet "ATF A&D Record" \
  --fastbound "/caminho/FastBoundImport Live - By Chris.xlsx" --fastbound-sheet "FastBoundImport Live - By Chris" \
  --out "/caminho/FastBoundImport_Populado.xlsx"
```

## Overrides de mapeamento (opcional)
Crie um CSV (ou JSON/YAML) com cabeçalhos:
```
FastBound Column,ATF Source
Manufacturer,Maker
Model,Model
Serial Number,Serial
```
E rode:
```bash
python fastbound_importer.py ... --map overrides.csv
```

## Opções úteis
- `--strict` : falha se houver colunas do FastBound sem origem (útil para QA).
- `--fuzzy-cutoff 0.80` : relaxa (ou endurece) o pareamento fuzzy.
- `--verbose` : logging detalhado.

## Saída
Um Excel com 3 abas:
- **FastBoundImport**: dados já no layout do FastBound.
- **Mapping Report**: tabela com origem de cada coluna e tipo de pareamento.
- **Missing & Guidance**: colunas não preenchidas + como obter a informação.

> Dica: se o FastBound aceitar CSV, basta exportar a aba **FastBoundImport** para CSV.
