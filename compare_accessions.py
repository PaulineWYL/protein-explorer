import pandas as pd
from pathlib import Path

EXCEL_FILE = Path('26042903-Protein Peptide Report.xlsx')
SHEETS = ['MAD', 'MADEV', 'MADEVNP']
OUTPUT_CSV = Path('compare_MAD_MADEV_MADEVNP_common.csv')


def load_sheet(sheet_name: str) -> pd.DataFrame:
    df = pd.read_excel(EXCEL_FILE, sheet_name=sheet_name, engine='openpyxl')
    df['Accession'] = df['Accession'].astype(str).str.strip()
    return df


def main() -> None:
    if not EXCEL_FILE.exists():
        raise FileNotFoundError(f'找不到 Excel 檔案: {EXCEL_FILE}')

    dfs = {sheet: load_sheet(sheet) for sheet in SHEETS}

    accession_sets = {
        sheet: set(df['Accession'].dropna().astype(str).str.strip())
        for sheet, df in dfs.items()
    }

    print('--- 檔案與工作表信息 ---')
    print(f'Excel 檔案: {EXCEL_FILE}')
    for sheet, df in dfs.items():
        print(f'{sheet}: {len(df)} rows, {len(accession_sets[sheet])} unique Accession')

    common_mad_madev = accession_sets['MAD'] & accession_sets['MADEV']
    common_mad_madevnp = accession_sets['MAD'] & accession_sets['MADEVNP']
    common_madev_madevnp = accession_sets['MADEV'] & accession_sets['MADEVNP']
    common_all = accession_sets['MAD'] & accession_sets['MADEV'] & accession_sets['MADEVNP']

    print('\n--- 比較結果 ---')
    print(f'MAD ∩ MADEV: {len(common_mad_madev)}')
    print(f'MAD ∩ MADEVNP: {len(common_mad_madevnp)}')
    print(f'MADEV ∩ MADEVNP: {len(common_madev_madevnp)}')
    print(f'MAD ∩ MADEV ∩ MADEVNP: {len(common_all)}')

    if len(common_all) == 0:
        print('\n沒有三者共有的 Accession。')
        return

    print('\n共有三者的 Accession 範例 (前 50 個):')
    print(', '.join(sorted(common_all)[:50]))

    combined = pd.DataFrame({'Accession': sorted(common_all)})
    for sheet, df in dfs.items():
        sub = df[['Accession', f'Abundance: {sheet}', f'Abundances Count: {sheet}', 'Relative abundance (%)']].copy()
        sub.columns = [
            'Accession',
            f'Abundance_{sheet}',
            f'Count_{sheet}',
            f'Relative_{sheet}',
        ]
        combined = combined.merge(sub, on='Accession', how='left')

    combined.to_csv(OUTPUT_CSV, index=False)
    print(f'\n已生成比較表: {OUTPUT_CSV}')
    print('\n前三筆比較結果:')
    print(combined.head(10).to_string(index=False))


if __name__ == '__main__':
    main()
