import pandas as pd
import math
from pathlib import Path

BASE = Path(__file__).parent

def fix(v):
    if isinstance(v, float) and math.isnan(v): return ''
    if isinstance(v, str) and v.strip().lower() in ('nan', 'none'): return ''
    return v

def clean(df):
    return df.map(fix)

# ── DH Tempos ──
dh_file = next((f for f in [
    BASE / 'fLista_DH_tempos.xlsx',
    BASE / 'fLista_DH_tempos.xls',
    *BASE.glob('*DH*tempos*.xls*'),
] if f.exists()), None)

if dh_file:
    df = pd.read_excel(dh_file)
    df = df.rename(columns={'Unnamed: 3': 'Fornecedor'})
    df = clean(df)
    df.to_csv(BASE / 'dh_tempos.csv', index=False, encoding='utf-8-sig')
    print(f'✓ dh_tempos.csv — {len(df):,} registros')
else:
    raise FileNotFoundError('fLista_DH_tempos.xls(x) não encontrado')

# ── OBs Pagas ──
ob_file = next((f for f in [
    BASE / 'OBs_Pagas_Novo.xlsx',
    BASE / 'OBs_Pagas_Novo.xls',
    *BASE.glob('*OBs*Pagas*.xls*'),
] if f.exists()), None)

if ob_file:
    df = pd.read_excel(ob_file, skiprows=2, header=0)
    df.columns = ['Emissão','Empenho','NP','Documento',
                  'Natureza_Despesa_Cod','Natureza_Despesa','Observacao','Valor']
    df = df[df['Emissão'] != 'Total'].dropna(subset=['Emissão'])
    df = clean(df)
    df.to_csv(BASE / 'obs_pagas.csv', index=False, encoding='utf-8-sig')
    print(f'✓ obs_pagas.csv — {len(df):,} registros')
else:
    raise FileNotFoundError('OBs_Pagas_Novo.xls(x) não encontrado')

print('✅ Conversão concluída!')
