"""
Vórtex — atualizar_dados.py
Roda diariamente via Task Scheduler.
Pulls ERP data and saves to dados/dados_sap.xlsx
O Streamlit lê desse arquivo.
"""
import pandas as pd
import pyodbc
import datetime
import os
import re

# capacidade_producao.py is in the same directory
from capacidade_producao import (
    conectar, buscar_pas_ativos, buscar_estoque_insumos, buscar_bom,
    buscar_consumo_3m, buscar_carteira_por_pa, buscar_carteira_detalhada,
    buscar_compras_pa, buscar_pedidos_compra, buscar_todos_pedidos_futuros,
    buscar_ordens_producao, buscar_historico_detalhado, buscar_todos_itens,
    buscar_mps_ativos, buscar_consumo_3m_mp, buscar_historico_detalhado_mp,
    buscar_sc_aberta, buscar_sc_linhas, buscar_esbocos_po,
    buscar_esbocos_linhas, buscar_sc_esboco_link, buscar_aprovacoes_wf,
    calcular_curva_abc, calcular_capacidade, calcular_calendario,
    calcular_mrp, calcular_mrp_revenda, calcular_forecast_compra,
    expandir_bom, extrair_familia
)

DADOS_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "dados", "dados_sap.xlsx")
os.makedirs(os.path.dirname(DADOS_PATH), exist_ok=True)

def buscar_faturado_mes_atual(conn):
    """Faturamento no mês corrente. Mesmos filtros de buscar_consumo_3m."""
    print("📆 Buscando faturado do mês atual...")
    query = """
        DECLARE @ini DATE = DATEFROMPARTS(YEAR(GETDATE()), MONTH(GETDATE()), 1);
        DECLARE @fim DATE = DATEADD(MONTH, 1, @ini);
        SELECT L.ItemCode, SUM(L.Quantity) as FaturadoMes
        FROM OINV H WITH(NOLOCK)
        INNER JOIN INV1 L WITH(NOLOCK) ON H.DocEntry = L.DocEntry
        INNER JOIN ORDR O WITH(NOLOCK) ON L.BaseEntry = O.DocEntry AND L.BaseType = 17
        WHERE H.DocDate >= @ini AND H.DocDate < @fim
          AND H.CANCELED = 'N'
          AND ISNULL(H.Comments, '') NOT LIKE '%VD FUTURA%'
        GROUP BY L.ItemCode
    """
    return pd.read_sql(query, conn)

def buscar_historico_mensal(conn):
    """Histórico mensal por PA. Mesmos filtros de buscar_consumo_3m.
    Alinha com OTD (NFs com pedido, exclui VD Futura 1ª NF)."""
    print("📈 Buscando histórico mensal...")
    query = """
        DECLARE @inicio DATE = DATEADD(MONTH, -3, DATEFROMPARTS(YEAR(GETDATE()), MONTH(GETDATE()), 1));
        DECLARE @fim    DATE = DATEFROMPARTS(YEAR(GETDATE()), MONTH(GETDATE()), 1);
        SELECT
            L.ItemCode,
            YEAR(H.DocDate)  as Ano,
            MONTH(H.DocDate) as Mes,
            SUM(L.Quantity)  as Consumo
        FROM OINV H WITH(NOLOCK)
        INNER JOIN INV1 L WITH(NOLOCK) ON H.DocEntry = L.DocEntry
        INNER JOIN ORDR O WITH(NOLOCK) ON L.BaseEntry = O.DocEntry AND L.BaseType = 17
        WHERE H.DocDate >= @inicio AND H.DocDate < @fim
          AND H.CANCELED = 'N'
          AND ISNULL(H.Comments, '') NOT LIKE '%VD FUTURA%'
        GROUP BY L.ItemCode, YEAR(H.DocDate), MONTH(H.DocDate)
        ORDER BY L.ItemCode, Ano, Mes
    """
    return pd.read_sql(query, conn)

def buscar_historico_mensal_mp(conn):
    """Histórico mensal de consumo de MP em produção. Mesmos filtros de buscar_consumo_3m_mp."""
    print("📈 Buscando histórico mensal MP (Issue for Production)...")
    query = """
        DECLARE @inicio DATE = DATEADD(MONTH, -3, DATEFROMPARTS(YEAR(GETDATE()), MONTH(GETDATE()), 1));
        DECLARE @fim    DATE = DATEFROMPARTS(YEAR(GETDATE()), MONTH(GETDATE()), 1);
        SELECT L.ItemCode,
               YEAR(H.DocDate)  as Ano,
               MONTH(H.DocDate) as Mes,
               SUM(L.Quantity)  as Consumo
        FROM OIGE H WITH(NOLOCK)
        INNER JOIN IGE1 L WITH(NOLOCK) ON H.DocEntry = L.DocEntry
        WHERE H.DocDate >= @inicio AND H.DocDate < @fim
          AND H.CANCELED = 'N'
          AND L.BaseType = 202
        GROUP BY L.ItemCode, YEAR(H.DocDate), MONTH(H.DocDate)
        ORDER BY L.ItemCode, Ano, Mes
    """
    return pd.read_sql(query, conn)

def buscar_consumido_mes_atual_mp(conn):
    """MP consumida em produção no mês corrente parcial."""
    print("📆 Buscando consumido MP do mês atual...")
    query = """
        DECLARE @ini DATE = DATEFROMPARTS(YEAR(GETDATE()), MONTH(GETDATE()), 1);
        DECLARE @fim DATE = DATEADD(MONTH, 1, @ini);
        SELECT L.ItemCode, SUM(L.Quantity) as ConsumidoMes
        FROM OIGE H WITH(NOLOCK)
        INNER JOIN IGE1 L WITH(NOLOCK) ON H.DocEntry = L.DocEntry
        WHERE H.DocDate >= @ini AND H.DocDate < @fim
          AND H.CANCELED = 'N'
          AND L.BaseType = 202
        GROUP BY L.ItemCode
    """
    return pd.read_sql(query, conn)

def montar_historico_mp(df_mps_abc, df_hist_mp_mensal):
    """Monta aba historico_mp: MP + consumo em produção por mês + média.
    Espelho de montar_historico mas pra MP."""
    hoje = datetime.date.today()
    meses = []
    for i in range(3, 0, -1):
        total = hoje.year * 12 + (hoje.month - 1) - i
        a, m = divmod(total, 12)
        m += 1
        meses.append((a, m, f"{m:02d}/{a}"))

    TIPO_MP = {119: "Mat. Prima", 120: "Componente", 121: "Embalagem", 122: "Consumível"}
    df = df_mps_abc[["Familia","ItmsGrpCod","ItemCode","ItemName","ABC","MediaMensal","Consumo3m"]].copy()
    df.columns = ["Família","ItmsGrpCod","Código MP","Descrição MP","Classe ABC","Média/Mês","Consumo3m"]
    df["Categoria"] = df["ItmsGrpCod"].map(TIPO_MP).fillna("Outro")
    df = df[["Família","Categoria","Código MP","Descrição MP","Classe ABC","Média/Mês","Consumo3m"]]
    df["Média/Mês"] = df["Média/Mês"].round().astype(int)

    for (ano, mes, label) in meses:
        hist_mes = df_hist_mp_mensal[(df_hist_mp_mensal["Ano"]==ano) & (df_hist_mp_mensal["Mes"]==mes)]
        mapa = hist_mes.set_index("ItemCode")["Consumo"].to_dict()
        df[label] = df["Código MP"].map(mapa).fillna(0).astype(int)
    return df

def montar_historico(df_pas_abc, df_hist_mensal):
    """Monta aba histórico: PA + saídas por mês + média"""
    hoje = datetime.date.today()
    # Últimos 3 meses FECHADOS (exclui o mês atual)
    meses = []
    for i in range(3, 0, -1):
        total = hoje.year * 12 + (hoje.month - 1) - i
        a, m = divmod(total, 12)
        m += 1
        meses.append((a, m, f"{m:02d}/{a}"))

    df = df_pas_abc[["Familia","ItmsGrpCod","ItemCode","ItemName","ABC","MediaMensal","Consumo3m"]].copy()
    df.columns = ["Família","ItmsGrpCod","Código PA","Descrição PA","Classe ABC","Média/Mês","Consumo3m"]
    df["Categoria"] = df["ItmsGrpCod"].apply(lambda g: "Revenda" if g == 123 else "PA")
    df = df[["Família","Categoria","Código PA","Descrição PA","Classe ABC","Média/Mês","Consumo3m"]]
    df["Média/Mês"] = df["Média/Mês"].round().astype(int)

    for (ano, mes, label) in meses:
        hist_mes = df_hist_mensal[(df_hist_mensal["Ano"]==ano) & (df_hist_mensal["Mes"]==mes)]
        mapa = hist_mes.set_index("ItemCode")["Consumo"].to_dict()
        df[label] = df["Código PA"].map(mapa).fillna(0).astype(int)

    return df

def montar_bom_explodida(df_pas_abc, df_bom, estoque_dict, nome_dict, grupo_dict):
    """Monta aba BOM explodida em formato flat (com coluna Nível)"""
    bom_dict = {}
    for _, row in df_bom.iterrows():
        f = row["Father"]
        if f not in bom_dict:
            bom_dict[f] = []
        bom_dict[f].append((row["Insumo"], float(row["Quantity"])))

    GRUPOS_RECURSIVOS = (124, 126)
    GRUPOS_IGNORAR    = (127, 132, 100)

    linhas = []

    def explodir(item_code, qtd_pai, nivel, pa_code, pa_nome, familia, abc):
        grupo   = grupo_dict.get(item_code, 0)
        estoque = estoque_dict.get(item_code, 0)
        nome    = str(nome_dict.get(item_code, "")).strip()

        tipo_map = {125:"PA",123:"Revenda",126:"Conjunto",124:"Prod. Interm.",
                    120:"Componente",121:"Embalagem",122:"Consumível",119:"Mat. Prima"}
        tipo = tipo_map.get(grupo, f"Grp {grupo}")

        nivel_str = "PA" if nivel==0 else ("PI/CJ" if grupo in (124,126) else "INS")

        if nivel > 0:
            from capacidade_producao import calcular_curva_abc
            media_pa = df_pas_abc[df_pas_abc["ItemCode"]==pa_code]["MediaDiaria"].values
            media_pa = float(media_pa[0]) if len(media_pa) > 0 else 0
            if media_pa > 0 and qtd_pai > 0:
                cob = round(estoque / (media_pa * qtd_pai), 1)
            else:
                cob = "-"

            if estoque == 0 and grupo not in (125,123,126,124):
                status = "🔴 ZERADO"
            elif isinstance(cob, float) and cob < 7:
                status = "🟡 BAIXO"
            elif isinstance(cob, float) and cob >= 7:
                status = "🟢 OK"
            else:
                status = "—"
        else:
            cob    = "-"
            status = ""

        linhas.append({
            "Família":      familia,
            "Código PA":    pa_code,
            "Nível":        nivel_str,
            "Profundidade": nivel,                         # 0=PA, 1=primeiro filho, 2=neto, ...
            "Código":       item_code,
            "Descrição":    ("  " * nivel) + nome[:55],
            "ABC / Tipo":   abc if nivel==0 else tipo,
            "Qtd p/ 1 PA":  round(qtd_pai, 4) if nivel > 0 else 1,
            "Estoque Atual": int(estoque),
            "Cob. (dias)":  cob,
            "Status":       status,
        })

        if item_code in bom_dict:
            for (filho, qtd_filho) in bom_dict[item_code]:
                explodir(filho, qtd_pai * qtd_filho if nivel > 0 else qtd_filho,
                         nivel+1, pa_code, pa_nome, familia, abc)

    df_sorted = df_pas_abc.sort_values(["Familia","ABC","ItemCode"])
    for _, pa in df_sorted.iterrows():
        if pa["ItemCode"] not in bom_dict:
            continue
        explodir(pa["ItemCode"], 1, 0, pa["ItemCode"], pa["ItemName"], pa["Familia"], pa["ABC"])

    return pd.DataFrame(linhas)

def consolidar_compras(df_sc, df_link, df_esbocos, df_wf):
    """Consolida status de aprovação no nível de SC e de Esboço.

    Adiciona em `df_sc`:
      - EsbocosCount, EsbocosNums (lista string), AprovacaoStatus (string)
    Adiciona em `df_esbocos`:
      - SCsOrigemNums, AprovacaoStatus, JaAprovaram, Aguardando, Rejeitaram

    Retorna (df_sc_enriched, df_esbocos_enriched).
    """
    df_sc = df_sc.copy()
    df_esbocos = df_esbocos.copy()

    # ── Status de workflow por esboço ────────────────────────────────
    # Agrupa aprovadores por esboço
    wf_por_esboco = {}   # DocEntry_esboco -> {"status": "...", "ja": [], "aguard": [], "rej": []}
    if len(df_wf) > 0:
        for esboco_id, grupo in df_wf.groupby("Esboco_DocEntry"):
            wf_status = grupo["WorkflowStatus"].iloc[0]
            ja      = grupo.loc[grupo["StatusAprovador"] == "Y", "Aprovador"].dropna().tolist()
            aguard  = grupo.loc[grupo["StatusAprovador"] == "W", "Aprovador"].dropna().tolist()
            rej     = grupo.loc[grupo["StatusAprovador"] == "N", "Aprovador"].dropna().tolist()
            wf_por_esboco[esboco_id] = {
                "wf_status": wf_status,
                "ja":        ja,
                "aguard":    aguard,
                "rej":       rej,
            }

    def _resumo_aprovacao(esboco_id):
        info = wf_por_esboco.get(esboco_id)
        if not info:
            return "Sem aprovação requerida", "", "", ""
        if info["wf_status"] == "Y":
            status = "✅ Aprovado"
        elif info["wf_status"] == "N":
            status = "🔴 Rejeitado"
        else:
            status = "📝 Em aprovação"
        return (
            status,
            ", ".join(info["ja"]),
            ", ".join(info["aguard"]),
            ", ".join(info["rej"]),
        )

    # Enriquece df_esbocos
    esboco_status_map = {}   # docentry -> (status, ja, aguard, rej)
    status_l, ja_l, aguard_l, rej_l = [], [], [], []
    for _, row in df_esbocos.iterrows():
        s, j, a, r = _resumo_aprovacao(row["DocEntry"])
        esboco_status_map[row["DocEntry"]] = (s, j, a, r)
        status_l.append(s); ja_l.append(j); aguard_l.append(a); rej_l.append(r)
    df_esbocos["AprovacaoStatus"] = status_l
    df_esbocos["JaAprovaram"]     = ja_l
    df_esbocos["Aguardando"]      = aguard_l
    df_esbocos["Rejeitaram"]      = rej_l

    # SCs origem de cada esboço
    sc_de_esboco = {}   # esboco_docentry -> [sc_docnum, ...]
    esboco_de_sc = {}   # sc_docentry     -> [(esboco_docentry, esboco_docnum), ...]
    if len(df_link) > 0:
        for _, row in df_link.iterrows():
            sc_de_esboco.setdefault(row["Esboco_DocEntry"], []).append(str(row["SC_DocNum"]))
            esboco_de_sc.setdefault(row["SC_DocEntry"], []).append(
                (row["Esboco_DocEntry"], str(row["Esboco_DocNum"]))
            )
    df_esbocos["SCsOrigemNums"] = df_esbocos["DocEntry"].map(
        lambda e: ", ".join(sorted(set(sc_de_esboco.get(e, []))))
    )

    # Fornecedor por esboço (CardName) — pra consolidar nas SCs
    card_de_esboco = df_esbocos.set_index("DocEntry")["CardName"].to_dict()

    # ── Enriquece df_sc ──────────────────────────────────────────────
    esb_count, esb_nums, sc_status, sc_fornecedores = [], [], [], []
    for _, row in df_sc.iterrows():
        esbocos = esboco_de_sc.get(row["DocEntry"], [])
        esb_count.append(len(esbocos))
        nums = sorted({n for (_e, n) in esbocos})
        esb_nums.append(", ".join(nums))
        cards = []
        for (e_id, _e_num) in esbocos:
            c = card_de_esboco.get(e_id)
            if isinstance(c, str) and c.strip() and c not in cards:
                cards.append(c.strip())
        sc_fornecedores.append(", ".join(cards))

        if not esbocos:
            sc_status.append("—")
            continue

        # Consolida status do(s) esboço(s) associado(s):
        # Se algum aprovado → ✅ Aprovado (esboço X)
        # senão se algum em aprovação → 📝 Em aprovação
        # senão se algum rejeitado → 🔴 Rejeitado
        # senão → 📋 Esboço criado (sem workflow)
        flags = [esboco_status_map.get(e, ("Sem aprovação requerida","","",""))[0] for (e, _n) in esbocos]
        if any("Aprovado" in f for f in flags):
            sc_status.append("✅ Aprovado")
        elif any("Em aprovação" in f for f in flags):
            sc_status.append("📝 Em aprovação")
        elif any("Rejeitado" in f for f in flags):
            sc_status.append("🔴 Rejeitado")
        else:
            sc_status.append("📋 Esboço criado")

    df_sc["EsbocosCount"]    = esb_count
    df_sc["EsbocosNums"]     = esb_nums
    df_sc["Fornecedores"]    = sc_fornecedores
    df_sc["AprovacaoStatus"] = sc_status
    df_sc["TemEsboco"]       = df_sc["EsbocosCount"] > 0

    return df_sc, df_esbocos


def main():
    print(f"\n{'='*55}")
    print(f"Vórtex — Atualizar Dados — {datetime.datetime.now():%d/%m/%Y %H:%M}")
    print(f"{'='*55}\n")

    conn = conectar()

    df_pas             = buscar_pas_ativos(conn)
    df_insumos         = buscar_estoque_insumos(conn)
    df_bom             = buscar_bom(conn)
    df_consumo         = buscar_consumo_3m(conn)
    df_carteira        = buscar_carteira_por_pa(conn)
    df_carteira_det    = buscar_carteira_detalhada(conn)
    df_compras_pa      = buscar_compras_pa(conn)
    df_pedidos         = buscar_pedidos_compra(conn)
    df_ped_futuros     = buscar_todos_pedidos_futuros(conn)
    df_ordens          = buscar_ordens_producao(conn)
    df_hist_mensal     = buscar_historico_mensal(conn)
    df_hist_det        = buscar_historico_detalhado(conn)
    df_fat_mes         = buscar_faturado_mes_atual(conn)
    df_todos_itens     = buscar_todos_itens(conn)
    df_mps             = buscar_mps_ativos(conn)
    df_consumo_mp      = buscar_consumo_3m_mp(conn)
    df_hist_mp_mensal  = buscar_historico_mensal_mp(conn)
    df_cons_mes_mp     = buscar_consumido_mes_atual_mp(conn)
    df_hist_mp_det     = buscar_historico_detalhado_mp(conn)

    # Compras (SC + esboços + workflow)
    df_sc              = buscar_sc_aberta(conn)
    df_sc_linhas       = buscar_sc_linhas(conn)
    df_esbocos         = buscar_esbocos_po(conn)
    df_esbocos_linhas  = buscar_esbocos_linhas(conn)
    df_sc_esboco_link  = buscar_sc_esboco_link(conn)
    df_aprovacoes_wf   = buscar_aprovacoes_wf(conn)

    conn.close()

    df_pas_abc = calcular_curva_abc(df_pas, df_consumo)
    df_mps_abc = calcular_curva_abc(df_mps, df_consumo_mp)

    (df_capacidade, pa_insumos, estoque_dict, pedidos_dict,
     carteira_dict, nome_dict) = calcular_capacidade(
        df_pas_abc, df_bom, df_insumos, df_pedidos, df_carteira)

    df_cal, _ = calcular_calendario(df_pas_abc, df_bom, df_insumos, df_ped_futuros,
                                    df_carteira=df_carteira,
                                    df_compras_pa=df_compras_pa,
                                    df_ordens=df_ordens,
                                    df_carteira_det=df_carteira_det)

    # bom_dict e grupo_dict são passados pra cascata (PI/CJ propagam demanda pros filhos)
    _bom_dict_mrp = {}
    for _b_row in df_bom.itertuples(index=False):
        _bom_dict_mrp.setdefault(_b_row.Father, []).append((_b_row.Insumo, float(_b_row.Quantity)))
    _grupo_dict_mrp = df_insumos.set_index("ItemCode")["ItmsGrpCod"].to_dict()
    df_mrp = calcular_mrp(df_pas_abc, pa_insumos, estoque_dict, pedidos_dict,
                          carteira_dict, nome_dict, df_ordens,
                          bom_dict=_bom_dict_mrp, grupo_dict=_grupo_dict_mrp)

    pedidos_dict_todos = df_pedidos.set_index("ItemCode").to_dict("index") if len(df_pedidos) > 0 else {}
    df_mrp_revenda = calcular_mrp_revenda(df_pas_abc, pedidos_dict_todos, carteira_dict)

    grupo_dict = df_insumos.set_index("ItemCode")["ItmsGrpCod"].to_dict()

    # Pra renderizar a BOM corretamente (incluindo itens fantasma com prefixo F),
    # usamos nome/grupo de TODOS os itens — não só os com estoque (InvntItem='Y').
    nome_dict_full  = df_todos_itens.set_index("ItemCode")["ItemName"].to_dict()
    grupo_dict_full = df_todos_itens.set_index("ItemCode")["ItmsGrpCod"].to_dict()
    # Mantém fallback nos dicts originais (estoque-based) caso item esteja só lá
    nome_dict_full  = {**nome_dict,  **nome_dict_full}
    grupo_dict_full = {**grupo_dict, **grupo_dict_full}

    # Forecast Compra (similar à planilha "Forecast Mensal China")
    fonte_dict = df_todos_itens.set_index("ItemCode")["CountryOrg"].to_dict() if "CountryOrg" in df_todos_itens.columns else {}
    df_forecast_compra = calcular_forecast_compra(
        df_pas_abc, df_bom, df_insumos, df_pedidos,
        nome_dict_full, grupo_dict_full,
        fonte_dict=fonte_dict,
    )

    df_historico   = montar_historico(df_pas_abc, df_hist_mensal)
    fat_dict = df_fat_mes.set_index("ItemCode")["FaturadoMes"].to_dict() if not df_fat_mes.empty else {}
    df_historico["Faturado Mês Atual"] = df_historico["Código PA"].map(fat_dict).fillna(0).astype(int)

    df_historico_mp = montar_historico_mp(df_mps_abc, df_hist_mp_mensal)
    cons_dict_mp = df_cons_mes_mp.set_index("ItemCode")["ConsumidoMes"].to_dict() if not df_cons_mes_mp.empty else {}
    df_historico_mp["Consumido Mês Atual"] = df_historico_mp["Código MP"].map(cons_dict_mp).fillna(0).astype(int)

    df_bom_explod  = montar_bom_explodida(df_pas_abc, df_bom, estoque_dict,
                                           nome_dict_full, grupo_dict_full)

    # Compras — consolida status de aprovação no nível de SC e de esboço
    df_sc, df_esbocos = consolidar_compras(df_sc, df_sc_esboco_link, df_esbocos, df_aprovacoes_wf)

    # Salva em arquivo temporário e faz rename atômico para evitar leitura de arquivo parcial
    print(f"\n📊 Salvando dados em {DADOS_PATH}...")
    tmp_path = os.path.join(os.path.dirname(DADOS_PATH), "_dados_sap_tmp.xlsx")
    with pd.ExcelWriter(tmp_path, engine="openpyxl") as writer:
        df_capacidade.to_excel(writer,    sheet_name="capacidade",         index=False)
        df_cal.to_excel(writer,           sheet_name="calendario",          index=False)
        df_mrp.to_excel(writer,           sheet_name="mrp",                 index=False)
        df_mrp_revenda.to_excel(writer,   sheet_name="mrp_revenda",         index=False)
        df_historico.to_excel(writer,     sheet_name="historico",           index=False)
        df_historico_mp.to_excel(writer,  sheet_name="historico_mp",        index=False)
        df_hist_mp_det.to_excel(writer,   sheet_name="historico_detalhado_mp", index=False)
        df_bom_explod.to_excel(writer,    sheet_name="bom_explodida",       index=False)
        df_bom.to_excel(writer,           sheet_name="bom",                 index=False)
        df_carteira_det.to_excel(writer,  sheet_name="carteira_detalhada",  index=False)
        df_compras_pa.to_excel(writer,    sheet_name="compras_pa",          index=False)
        df_ordens.to_excel(writer,        sheet_name="ordens_producao",     index=False)
        df_hist_det.to_excel(writer,      sheet_name="historico_detalhado", index=False)
        df_forecast_compra.to_excel(writer, sheet_name="forecast_compra",   index=False)
        df_sc.to_excel(writer,            sheet_name="sc_aberta",           index=False)
        df_sc_linhas.to_excel(writer,     sheet_name="sc_linhas",           index=False)
        df_esbocos.to_excel(writer,       sheet_name="esbocos_po",          index=False)
        df_esbocos_linhas.to_excel(writer, sheet_name="esbocos_linhas",     index=False)
        df_sc_esboco_link.to_excel(writer, sheet_name="sc_esboco_link",     index=False)
        df_aprovacoes_wf.to_excel(writer, sheet_name="aprovacoes_wf",       index=False)
    os.replace(tmp_path, DADOS_PATH)

    print(f"\n✅ Dados salvos com sucesso!")
    print(f"   Capacidade:         {len(df_capacidade)} PAs")
    print(f"   Calendário:         {len(df_cal)} PAs")
    print(f"   MRP:                {len(df_mrp)} insumos")
    print(f"   Histórico:          {len(df_historico)} PAs")
    print(f"   Histórico MP:       {len(df_historico_mp)} MPs")
    print(f"   BOM:                {len(df_bom_explod)} linhas")
    print(f"   Carteira detalhada: {len(df_carteira_det)} linhas de pedidos")
    print(f"   Compras PA:         {len(df_compras_pa)} linhas de PO")
    print(f"   Ordens produção:    {len(df_ordens)} OPs abertas")
    print(f"   Histórico det.:     {len(df_hist_det)} linhas de NF")
    print(f"   Forecast compra:    {len(df_forecast_compra)} insumos")
    print(f"   SC abertas:         {len(df_sc)} solicitações")
    print(f"   Esboços de PO:      {len(df_esbocos)} esboços")
    print(f"   Aprovações:         {len(df_aprovacoes_wf)} linhas de workflow")

if __name__ == "__main__":
    main()
