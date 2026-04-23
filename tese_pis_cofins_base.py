"""
JCR Sistema — Módulo de Recuperação Tributária
Tese: Exclusão do PIS/COFINS da Própria Base (Gross Up Reverso) — JUD

Lógica extraída da planilha: Exclusão_do_PIS_COFINS_da_Própria_Base__JUD__pag1.xlsx
Total analisado na planilha de referência: 65.248 itens | R$ 4.241.380,49 identificados

Fundamento legal:
- RE 574.706/PR (STF) — Tese do Século (extensão lógica)
- PIS e COFINS não devem compor sua própria base de cálculo
- Tese em consolidação jurisprudencial — requer ação judicial (JUD)

Arquivos necessários:
- SPED Contribuições (EFD PIS/COFINS) — registros C100/C170
- Períodos a analisar: últimos 5 anos (prazo decadencial)

Responsabilidade do usuário:
- Validar itens com regime especial, receitas com deduções
- Verificar CSTs antes de protocolar
- Sistema não identifica exceções automaticamente
- Valores sem correção SELIC
"""

import pandas as pd
import numpy as np
from pathlib import Path


# ─── CSTs que permitem a tese ────────────────────────────────────────────────
# Conforme validação da planilha de referência: CSTs 01, 02, 49, 50
CST_VALIDAS = {1, 2, 49, 50}

# ─── Colunas esperadas do SPED Contribuições (C100/C170) ────────────────────
COLUNAS_SPED = [
    'Período', 'Registro', 'CNPJ', 'Nome Empresa',
    'Código Participante', 'Nome Participante', 'CNPJ Participante',
    'CPF Participante', 'UF Origem/Destino', 'Modelo Documento',
    'Situação Documento', 'Núm. Documento', 'Série', 'Chave Documento',
    'Data Documento', 'Data Escrituração Docum.', 'Valor Documento',
    'Valor Desconto', 'Valor Mercadoria/Operação', 'Valor Frete',
    'Valor Seguro', 'Valor Outras DA', 'Núm. Item', 'Cód. Item',
    'Descrição Item', 'NCM Item', 'Cód. Serviço', 'Descrição Serviço',
    'Tipo Item', 'CFOP Item', 'Descrição CFOP', 'Tipo Faturamento',
    'Qtidade Item', 'Unidade Item', 'Valor Item', 'Valor Desconto Item',
    'Valor ICMS Item', 'Valor ICMS-ST Item', 'Valor IPI Item',
    'CST PIS/COFINS Item', 'Valor BC PIS/COFINS Item',
    'Alíquota PIS Item', 'Valor PIS Item',
    'Alíquota COFINS Item', 'Valor COFINS Item',
]


# ═══════════════════════════════════════════════════════════════════════════════
# FUNÇÕES PRINCIPAIS
# ═══════════════════════════════════════════════════════════════════════════════

def carregar_planilha(caminho: str) -> pd.DataFrame:
    """
    Carrega a planilha de saída do sistema (já processada) ou
    um arquivo SPED bruto convertido para DataFrame.
    
    Aceita .xlsx ou .csv (pipe-delimitado para SPED bruto).
    """
    path = Path(caminho)
    if path.suffix in ['.xlsx', '.xlsm']:
        df = pd.read_excel(caminho)
    elif path.suffix == '.csv':
        df = pd.read_csv(caminho, sep='|', encoding='latin-1', low_memory=False)
    else:
        raise ValueError(f"Formato não suportado: {path.suffix}")
    
    print(f"✓ Arquivo carregado: {len(df):,} linhas | {len(df.columns)} colunas")
    return df


def validar_colunas(df: pd.DataFrame) -> bool:
    """Verifica se as colunas necessárias estão presentes."""
    colunas_necessarias = [
        'CST PIS/COFINS Item', 'Valor BC PIS/COFINS Item',
        'Alíquota PIS Item', 'Valor PIS Item',
        'Alíquota COFINS Item', 'Valor COFINS Item',
    ]
    faltando = [c for c in colunas_necessarias if c not in df.columns]
    if faltando:
        print(f"✗ Colunas faltando: {faltando}")
        return False
    return True


def aplicar_criterios(df: pd.DataFrame) -> pd.DataFrame:
    """
    Aplica os 3 critérios de elegibilidade da tese.
    
    Critério 1 — CST válida: 01, 02, 49 ou 50
    Critério 2 — PIS/COFINS não excluído da base (flag da planilha)
    Critério 3 — Confirmação adicional (Não Exclui = Sim)
    
    Soma Critérios = 3 → item elegível para recuperação
    """
    df = df.copy()
    
    # Critério 1: CST válida
    df['CST_num'] = pd.to_numeric(df['CST PIS/COFINS Item'], errors='coerce')
    df['CST Válida?'] = df['CST_num'].isin(CST_VALIDAS).map({True: 'Sim', False: 'Não'})
    
    # Critério 2: Exclui PIS/COFINS?
    # Se Diferença BC = PIS + COFINS pagos, o imposto está na base
    df['Diferença BC Item'] = df['Valor PIS Item'] + df['Valor COFINS Item']
    
    # Tolerância de R$ 0,05 para arredondamentos
    bc_com_tributo = abs(
        df['Valor BC PIS/COFINS Item'] - 
        (df['Valor Item'] - df['Valor Desconto Item'].fillna(0))
    ) < 0.05
    
    df['Exclui PIS/COFINS?'] = bc_com_tributo.map(
        {True: 'Exclui', False: 'Não Exclui'}
    )
    
    # Critério 3: Não Exclui = Sim (derivado do critério 2)
    df['Não Exclui PIS/COFINS?'] = (df['Exclui PIS/COFINS?'] == 'Não Exclui').map(
        {True: 'Sim', False: 'Não'}
    )
    
    # Soma dos critérios
    df['Soma Critérios'] = (
        (df['CST Válida?'] == 'Sim').astype(int) +
        (df['Exclui PIS/COFINS?'] == 'Não Exclui').astype(int) +
        (df['Não Exclui PIS/COFINS?'] == 'Sim').astype(int)
    )
    
    return df


def calcular_recuperacao(df: pd.DataFrame) -> pd.DataFrame:
    """
    Calcula o crédito a recuperar para itens elegíveis (Soma Critérios = 3).
    
    Lógica do Gross Up Reverso:
    - BC atual já contém PIS + COFINS embutidos
    - Nova Base = BC atual - (PIS pago + COFINS pago)
    - PIS correto = Nova Base × alíquota PIS / 100
    - COFINS correto = Nova Base × alíquota COFINS / 100
    - A recuperar = valor pago - valor correto
    """
    df = df.copy()
    
    # Inicializa colunas de resultado com zero
    df['Nova Base PIS/COFINS'] = 0.0
    df['PIS/COFINS Recalculado'] = 0.0
    df['Valor PIS a Recuperar'] = 0.0
    df['Valor COFINS a Recuperar'] = 0.0
    df['Total a Recuperar'] = 0.0
    
    # Filtra apenas elegíveis
    elegivel = df['Soma Critérios'] == 3
    
    # Nova base = BC atual - (PIS + COFINS pagos)
    df.loc[elegivel, 'Nova Base PIS/COFINS'] = (
        df.loc[elegivel, 'Valor BC PIS/COFINS Item'] -
        df.loc[elegivel, 'Diferença BC Item']
    ).clip(lower=0)
    
    # PIS/COFINS recalculado sobre a nova base
    df.loc[elegivel, 'PIS/COFINS Recalculado'] = (
        df.loc[elegivel, 'Nova Base PIS/COFINS'] *
        (df.loc[elegivel, 'Alíquota PIS Item'] + df.loc[elegivel, 'Alíquota COFINS Item']) / 100
    )
    
    # PIS a recuperar
    pis_correto = (
        df.loc[elegivel, 'Nova Base PIS/COFINS'] *
        df.loc[elegivel, 'Alíquota PIS Item'] / 100
    )
    df.loc[elegivel, 'Valor PIS a Recuperar'] = (
        df.loc[elegivel, 'Valor PIS Item'] - pis_correto
    ).clip(lower=0).round(2)
    
    # COFINS a recuperar
    cofins_correto = (
        df.loc[elegivel, 'Nova Base PIS/COFINS'] *
        df.loc[elegivel, 'Alíquota COFINS Item'] / 100
    )
    df.loc[elegivel, 'Valor COFINS a Recuperar'] = (
        df.loc[elegivel, 'Valor COFINS Item'] - cofins_correto
    ).clip(lower=0).round(2)
    
    # Total
    df.loc[elegivel, 'Total a Recuperar'] = (
        df.loc[elegivel, 'Valor PIS a Recuperar'] +
        df.loc[elegivel, 'Valor COFINS a Recuperar']
    ).round(2)
    
    return df


def gerar_resumo(df: pd.DataFrame) -> dict:
    """
    Gera resumo executivo da análise — por período e totais.
    """
    elegivel = df[df['Soma Critérios'] == 3]
    
    total_pis = elegivel['Valor PIS a Recuperar'].sum()
    total_cofins = elegivel['Valor COFINS a Recuperar'].sum()
    total_geral = elegivel['Total a Recuperar'].sum()
    
    # Resumo por período
    por_periodo = (
        elegivel.groupby('Período')
        .agg(
            itens=('Total a Recuperar', 'count'),
            pis=('Valor PIS a Recuperar', 'sum'),
            cofins=('Valor COFINS a Recuperar', 'sum'),
            total=('Total a Recuperar', 'sum')
        )
        .reset_index()
        .sort_values('Período')
    )
    
    resumo = {
        'tese': 'Exclusão do PIS/COFINS da Própria Base (Gross Up Reverso)',
        'tipo': 'JUD',
        'fundamento': 'RE 574.706/PR — extensão lógica',
        'total_itens_analisados': len(df),
        'total_itens_elegiveis': len(elegivel),
        'percentual_elegivel': round(len(elegivel) / len(df) * 100, 1) if len(df) > 0 else 0,
        'total_pis': round(total_pis, 2),
        'total_cofins': round(total_cofins, 2),
        'total_geral': round(total_geral, 2),
        'por_periodo': por_periodo.to_dict('records'),
        'observacao': 'Valores SEM correção SELIC. Aplicar SELIC acumulada para valor atualizado.',
        'alerta': 'Validar itens com regime especial antes de protocolar ação judicial.'
    }
    
    return resumo


def imprimir_resumo(resumo: dict) -> None:
    """Imprime o resumo de forma legível."""
    print("\n" + "═" * 60)
    print(f"  RESUMO — {resumo['tese']}")
    print("═" * 60)
    print(f"  Tipo: {resumo['tipo']} | {resumo['fundamento']}")
    print(f"  Itens analisados: {resumo['total_itens_analisados']:,}")
    print(f"  Itens elegíveis:  {resumo['total_itens_elegiveis']:,} ({resumo['percentual_elegivel']}%)")
    print(f"\n  {'PIS a recuperar:':<25} R$ {resumo['total_pis']:>15,.2f}")
    print(f"  {'COFINS a recuperar:':<25} R$ {resumo['total_cofins']:>15,.2f}")
    print(f"  {'TOTAL (sem SELIC):':<25} R$ {resumo['total_geral']:>15,.2f}")
    print(f"\n  ⚠  {resumo['observacao']}")
    print(f"  ⚠  {resumo['alerta']}")
    print("═" * 60)
    
    print("\n  Detalhamento por período:")
    print(f"  {'Período':<12} {'Itens':>6} {'PIS':>14} {'COFINS':>14} {'Total':>14}")
    print("  " + "-" * 64)
    for p in resumo['por_periodo']:
        print(
            f"  {str(p['Período']):<12} {p['itens']:>6,} "
            f"R$ {p['pis']:>11,.2f} R$ {p['cofins']:>11,.2f} R$ {p['total']:>11,.2f}"
        )


def exportar_resultado(df: pd.DataFrame, caminho_saida: str) -> None:
    """
    Exporta o resultado completo para Excel.
    Inclui aba de dados e aba de resumo por período.
    """
    elegivel = df[df['Soma Critérios'] == 3].copy()
    
    por_periodo = (
        elegivel.groupby('Período')
        .agg(
            Itens=('Total a Recuperar', 'count'),
            PIS_Recuperar=('Valor PIS a Recuperar', 'sum'),
            COFINS_Recuperar=('Valor COFINS a Recuperar', 'sum'),
            Total_Recuperar=('Total a Recuperar', 'sum')
        )
        .reset_index()
    )
    
    with pd.ExcelWriter(caminho_saida, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Dados Completos', index=False)
        elegivel.to_excel(writer, sheet_name='Apenas Elegíveis', index=False)
        por_periodo.to_excel(writer, sheet_name='Resumo por Período', index=False)
    
    print(f"\n✓ Resultado exportado: {caminho_saida}")


# ═══════════════════════════════════════════════════════════════════════════════
# EXECUÇÃO PRINCIPAL
# ═══════════════════════════════════════════════════════════════════════════════

def analisar(caminho_arquivo: str, caminho_saida: str = None) -> dict:
    """
    Função principal — executa a análise completa.
    
    Args:
        caminho_arquivo: caminho para o arquivo .xlsx ou .csv do SPED
        caminho_saida:   caminho para salvar o resultado (opcional)
    
    Returns:
        dict com o resumo da análise
    
    Exemplo de uso:
        resumo = analisar("sped_contribuicoes_2024.xlsx", "resultado_jud_2024.xlsx")
    """
    print(f"\n{'─'*60}")
    print(f"  JCR Sistema — Tese: Exclusão PIS/COFINS da Própria Base")
    print(f"{'─'*60}")
    
    # 1. Carregar
    df = carregar_planilha(caminho_arquivo)
    
    # 2. Validar
    if not validar_colunas(df):
        raise ValueError("Arquivo não possui as colunas necessárias para esta tese.")
    
    # 3. Aplicar critérios
    print("→ Aplicando critérios de elegibilidade...")
    df = aplicar_criterios(df)
    
    # 4. Calcular recuperação
    print("→ Calculando créditos a recuperar...")
    df = calcular_recuperacao(df)
    
    # 5. Gerar resumo
    resumo = gerar_resumo(df)
    imprimir_resumo(resumo)
    
    # 6. Exportar (opcional)
    if caminho_saida:
        exportar_resultado(df, caminho_saida)
    
    return resumo, df


# ─── Teste direto ──────────────────────────────────────────────────────────────
if __name__ == "__main__":
    import sys
    
    if len(sys.argv) < 2:
        print("Uso: python tese_pis_cofins_base.py <arquivo.xlsx> [saida.xlsx]")
        print("\nExemplo:")
        print("  python tese_pis_cofins_base.py sped_2024.xlsx resultado_2024.xlsx")
        sys.exit(0)
    
    arquivo = sys.argv[1]
    saida = sys.argv[2] if len(sys.argv) > 2 else None
    
    resumo, df_resultado = analisar(arquivo, saida)
