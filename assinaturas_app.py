import streamlit as st
import pandas as pd
import os
from decimal import Decimal
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import textwrap
import urllib.parse

# Configuração da página - DEVE SER A PRIMEIRA CHAMADA DO STREAMLIT
st.set_page_config(page_title="Gestão de Assinaturas", page_icon="📊", layout="wide")

# Caminho do arquivo de dados (banco de dados)
DATA_FILE = "assinaturas_v4.xlsx"

# Ações válidas (dropdown)
ACOES = ["Manter", "Cancelar", "Não Renovar", "Migrar", "Avaliar", "Em Revisão"]

# Dados extraídos do PDF (analise_assinaturas_paulo) estruturados com Prós, Contras, Tags e Decisão
DADOS_INICIAIS = [
    {
        "Serviço": "LinkedIn Premium",
        "Empresa/Conta": "PV Móveis",
        "Custo (R$)": 179.99,
        "Frequência": "Mensal",
        "Ação": "Manter",
        "Tags": "Alto ROI",
        "Prós": "• Principal canal de networking B2B\n• Credibilidade profissional\n• Acesso a InMail",
        "Contras": "• Assinatura mais cara da lista\n• ROI depende de uso ativo",
        "Decisão": "Manter — usar ativamente. Se não prospectar em 30 dias, reavalie."
    },
    {
        "Serviço": "Google One (Principal)",
        "Empresa/Conta": "PV Móveis",
        "Custo (R$)": 96.99,
        "Frequência": "Mensal",
        "Ação": "Manter",
        "Tags": "Essencial",
        "Prós": "• Armazenamento essencial para o negócio\n• Backup de emails, Drive e Fotos",
        "Contras": "• Verificar se plano atual é o menor necessário",
        "Decisão": "Manter, mas conferir se há plano menor que atenda."
    },
    {
        "Serviço": "Facebook (Verificação)",
        "Empresa/Conta": "PV Móveis",
        "Custo (R$)": 53.90,
        "Frequência": "Mensal",
        "Ação": "Manter",
        "Tags": "Imagem de marca",
        "Prós": "• Credibilidade e verificação do perfil\n• Custo baixo para o benefício de imagem",
        "Contras": "• Depende de postagens regulares para render",
        "Decisão": "Manter. Vence 07/05 — renovar com atenção."
    },
    {
        "Serviço": "YouTube Premium",
        "Empresa/Conta": "PV Móveis",
        "Custo (R$)": 26.90,
        "Frequência": "Mensal",
        "Ação": "Manter",
        "Tags": "Fluxo AI",
        "Prós": "• Integrado ao fluxo com Gemmy AI\n• Upload de vídeos privados para análise\n• Melhor custo-benefício",
        "Contras": "• Uso pouco visível externamente",
        "Decisão": "Manter — custo-benefício excelente dado o uso estratégico."
    },
    {
        "Serviço": "Fuelin",
        "Empresa/Conta": "PV Móveis",
        "Custo (R$)": 50.00,
        "Frequência": "Mensal",
        "Ação": "Manter",
        "Tags": "Uso ativo",
        "Prós": "• Serviço ativo e em uso\n• Custo dentro da média",
        "Contras": "• Verificar frequência real de uso\n• Confirmar se entrega resultado mensurável",
        "Decisão": "Manter por ora — avaliar ROI concreto antes da próxima renovação."
    },
    {
        "Serviço": "Avast Antivírus",
        "Empresa/Conta": "PV Móveis",
        "Custo (R$)": 213.00,
        "Frequência": "Anual",
        "Ação": "Não Renovar",
        "Tags": "Já pago",
        "Prós": "• Proteção ativa até março/2027",
        "Contras": "• Windows Defender nativo é gratuito e eficaz\n• Custo desnecessário na renovação",
        "Decisão": "Não cancelar (já pago). Em março/2027, substituir pelo Windows Defender."
    },
    {
        "Serviço": "Gemmy AI",
        "Empresa/Conta": "PV Móveis",
        "Custo (R$)": 129.99,
        "Frequência": "Anual",
        "Ação": "Avaliar",
        "Tags": "Já pago, Produtividade AI",
        "Prós": "• Ferramenta de IA a custo muito baixo\n• Já pago até março/2027",
        "Contras": "• Na renovação, comparar com outras ferramentas AI",
        "Decisão": "Manter até 2027. Avaliar ROI real antes de renovar."
    },
    {
        "Serviço": "Google One (Secundário)",
        "Empresa/Conta": "Grupo Paluto",
        "Custo (R$)": 96.99,
        "Frequência": "Mensal",
        "Ação": "Cancelar",
        "Tags": "Centralizar",
        "Prós": "• Repositório separado por empresa",
        "Contras": "• Duplica custo do Google One principal\n• Objetivo de centralizar torna redundante",
        "Decisão": "Ao centralizar, migrar dados e cancelar. Economia de R$ 96,99/mês."
    },
    {
        "Serviço": "CapCut Pro",
        "Empresa/Conta": "Grupo Paluto",
        "Custo (R$)": 32.90,
        "Frequência": "Mensal",
        "Ação": "Avaliar",
        "Tags": "Centralizar",
        "Prós": "• Edição de vídeo com IA integrada",
        "Contras": "• Versão gratuita pode ser suficiente",
        "Decisão": "Testar versão gratuita. Se renovar, migrar para conta principal."
    },
    {
        "Serviço": "Instagram Business",
        "Empresa/Conta": "V4 Company",
        "Custo (R$)": 120.90,
        "Frequência": "Mensal",
        "Ação": "Migrar",
        "Tags": "Imagem de marca, Centralizar",
        "Prós": "• Verificação profissional no Instagram\n• Credibilidade para marca e clientes",
        "Contras": "• Custo mensal elevado\n• Conta separada dificulta gestão",
        "Decisão": "Manter. Vence 04/05 — ao renovar, avaliar migrar para conta principal."
    },
    {
        "Serviço": "Kommo (CRM)",
        "Empresa/Conta": "V4 Company",
        "Custo (R$)": 62.00,
        "Frequência": "Mensal",
        "Ação": "Manter",
        "Tags": "Gestão clientes",
        "Prós": "• Essencial para gestão comercial\n• Centraliza acompanhamento de clientes",
        "Contras": "• Só vale se usado consistentemente pela equipe",
        "Decisão": "Manter — CRM ativo e fundamental para escalar vendas."
    },
    {
        "Serviço": "Canva",
        "Empresa/Conta": "V4 Company",
        "Custo (R$)": 0.00,
        "Frequência": "Mensal",
        "Ação": "Manter",
        "Tags": "Gratuito",
        "Prós": "• Criação de artes e materiais sem custo\n• Fácil de usar e colaborativo",
        "Contras": "• Recursos avançados limitados na versão free",
        "Decisão": "Manter — gratuito e útil."
    },
]

# Ordem das colunas na planilha/tabela
COLUNAS = ["Serviço", "Empresa/Conta", "Custo (R$)", "Frequência", "Tags", "Prós", "Contras", "Ação", "Decisão"]

def init_data():
    if not os.path.exists(DATA_FILE):
        df = pd.DataFrame(DADOS_INICIAIS, columns=COLUNAS)
        df.to_excel(DATA_FILE, index=False)

def load_data():
    if os.path.exists(DATA_FILE):
        df = pd.read_excel(DATA_FILE, dtype={"Custo (R$)": float})
        # Garante que colunas novas existam (retrocompatibilidade se alguém usar arquivo velho)
        for col in COLUNAS:
            if col not in df.columns:
                df[col] = ""
        return df[COLUNAS]
    return pd.DataFrame(DADOS_INICIAIS, columns=COLUNAS)

def save_data(df):
    df.to_excel(DATA_FILE, index=False)

def fmt_brl(valor: float) -> str:
    """Formata valor monetário sem arredondamento."""
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def calc_mensal(row):
    if row["Frequência"] == "Anual":
        return row["Custo (R$)"] / 12
    return row["Custo (R$)"]

def calc_anual(row):
    if row["Frequência"] == "Mensal":
        return row["Custo (R$)"] * 12
    return row["Custo (R$)"]

# Inicializa o arquivo se não existir
init_data()

# =======================
# INJECT CUSTOM CSS
# =======================
def local_css(file_name):
    if os.path.exists(file_name):
        with open(file_name) as f:
            st.markdown(f'<style>{f.read()}</style>', unsafe_allow_html=True)

local_css("style.css")

# Carregar os dados
df = load_data()

# Calcula custos normalizados para sumarização
df["Custo_Mensal_Calc"] = df.apply(calc_mensal, axis=1)
df["Custo_Anual_Calc"] = df.apply(calc_anual, axis=1)

total_mensal = df["Custo_Mensal_Calc"].sum()
projecao_anual = df["Custo_Anual_Calc"].sum()
acoes_corte = ["Cancelar", "Não Renovar", "Avaliar", "Em Revisão"]
economia_df = df[df["Ação"].isin(acoes_corte)]
economia_mensal = economia_df["Custo_Mensal_Calc"].sum()

ativas = len(df[df["Ação"] == "Manter"])
total_subs = len(df)

# =======================
# HERO SECTION
# =======================
st.markdown("<h1 style='margin-bottom: 2.5rem;'>Gestão de Assinaturas</h1>", unsafe_allow_html=True)

# =======================
# KPI CARDS
# =======================
kpi_html = f"""
<div style="display: grid; grid-template-columns: repeat(3, 1fr); gap: 1.5rem; margin-bottom: 3rem;">
    <div class="kpi-card">
        <div class="kpi-title">Gasto Mensal Total</div>
        <div class="kpi-value">{fmt_brl(total_mensal)}</div>
        <div class="kpi-subtext">Comprometido este mês</div>
    </div>
    <div class="kpi-card">
        <div class="kpi-title">Projeção Anual</div>
        <div class="kpi-value">{fmt_brl(projecao_anual)}</div>
        <div class="kpi-subtext">Impacto financeiro a longo prazo</div>
    </div>
    <div class="kpi-card">
        <div class="kpi-title">Economia Potencial</div>
        <div class="kpi-value">{fmt_brl(economia_mensal)}</div>
        <div class="kpi-subtext warning">Em revisão / cancelamento</div>
    </div>
</div>
"""
st.markdown(kpi_html, unsafe_allow_html=True)

# =======================
# FILTER & ACTION BAR
# =======================
col_search, col_filter, col_toggle = st.columns([2, 2, 1])
with col_search:
    search_query = st.text_input("Buscar assinatura...", placeholder="Ex: Spotify, Google...")
with col_filter:
    status_filter = st.selectbox("Filtrar por Status", ["Todos"] + ACOES)
with col_toggle:
    st.markdown("<div style='height: 28px;'></div>", unsafe_allow_html=True) # spacer
    edit_mode = st.toggle("Modo Edição (Planilha)")
    show_list = st.toggle("Mostrar Lista", value=True)

# Aplicar filtros
filtered_df = df.copy()
if search_query:
    filtered_df = filtered_df[filtered_df["Serviço"].str.contains(search_query, case=False, na=False) | filtered_df["Tags"].str.contains(search_query, case=False, na=False)]
if status_filter != "Todos":
    filtered_df = filtered_df[filtered_df["Ação"] == status_filter]

# =======================
# MAIN TABLE / LIST VIEW
# =======================
st.markdown("<h2 style='margin-top:0; margin-bottom: 1rem;'>Workspace de Assinaturas</h2>", unsafe_allow_html=True)

def get_badge_class(acao):
    if acao == "Manter": return "badge-green"
    if acao == "Cancelar" or acao == "Não Renovar": return "badge-red"
    if acao == "Avaliar" or acao == "Em Revisão": return "badge-yellow"
    if acao == "Migrar": return "badge-purple"
    return "badge-gray"

if show_list:
    if edit_mode:
        st.info("Modo de edição ativado. Alterações feitas na tabela abaixo podem ser salvas.", icon="✏️")
        edited_df = st.data_editor(
            df[COLUNAS], # Mostra todos para edição, não apenas os filtrados, para evitar bugs de index
            num_rows="dynamic",
            width="stretch",
            height=500,
            column_config={
                "Serviço": st.column_config.TextColumn("Serviço", required=True),
                "Empresa/Conta": st.column_config.SelectboxColumn("Conta", options=["PV Móveis", "Grupo Paluto", "V4 Company", "Geral/Outro"]),
                "Custo (R$)": st.column_config.NumberColumn("Custo (R$)", format="R$ %.2f", min_value=0.0, step=0.01),
                "Frequência": st.column_config.SelectboxColumn("Frequência", options=["Mensal", "Anual"]),
                "Ação": st.column_config.SelectboxColumn("Ação", options=ACOES),
            }
        )
        
        col_save, col_export = st.columns([1, 2])
        with col_save:
            if st.button("💾 Salvar Alterações", type="primary"):
                save_data(edited_df)
                st.success("Alterações salvas!")
                st.rerun()
        with col_export:
            st.download_button(label="📥 Baixar CSV", data=edited_df.to_csv(index=False, sep=';', decimal=',').encode("utf-8-sig"), file_name="assinaturas.csv", mime="text/csv")
    else:
        # Premium List View
        if filtered_df.empty:
            st.markdown("<div class='list-row' style='justify-content:center; color:var(--text-secondary);'>Nenhuma assinatura encontrada.</div>", unsafe_allow_html=True)
        else:
            # Header Row (Visual only)
            header_html = """
            <div style="display: flex; padding: 0 1.5rem 0.5rem 1.5rem; color: var(--text-secondary); font-size: 0.8rem; font-weight: 600; text-transform: uppercase; letter-spacing: 0.05em; border-bottom: 1px solid var(--border-color); margin-bottom: 1rem;">
                <div style="flex: 2;">Serviço & Categoria</div>
                <div style="flex: 1;">Status</div>
                <div style="flex: 1;">Frequência</div>
                <div style="flex: 1; text-align: right;">Custo</div>
            </div>
            """
            st.markdown("".join(line.strip() for line in header_html.split('\n')), unsafe_allow_html=True)

            for _, row in filtered_df.iterrows():
                badge_cls = get_badge_class(row['Ação'])
                tags_display = row['Tags'] if pd.notna(row['Tags']) and row['Tags'] else row['Empresa/Conta']
                
                freq_display = f"/{row['Frequência'].lower()[:3]}"
                custo_display = fmt_brl(row['Custo (R$)'])

                row_html = f"""
                <div class="list-row">
                    <div style="flex: 2; display: flex; align-items: center; gap: 1rem;">
                        <div style="width: 40px; height: 40px; border-radius: 8px; background: #2D3748; display: flex; align-items: center; justify-content: center; font-weight: 700; color: #fff; font-size: 1.2rem;">
                            {str(row['Serviço'])[0]}
                        </div>
                        <div>
                            <div class="service-name">{row['Serviço']}</div>
                            <div class="service-meta">{tags_display}</div>
                        </div>
                    </div>
                    <div style="flex: 1;">
                        <span class="badge {badge_cls}">{row['Ação']}</span>
                    </div>
                    <div style="flex: 1; color: var(--text-secondary); font-size: 0.9rem;">
                        {row['Frequência']}
                    </div>
                    <div style="flex: 1; text-align: right;">
                        <div class="cost-highlight">{custo_display}</div>
                        <div class="cost-freq">{freq_display}</div>
                    </div>
                </div>
                """
                st.markdown("".join(line.strip() for line in row_html.split('\n')), unsafe_allow_html=True)
else:
    st.info("A lista de assinaturas está oculta. Ative 'Mostrar Lista' acima para visualizar.", icon="👁️")

# =======================
# ANALYTICS SECTION
# =======================
st.markdown("<h2 style='margin-top: 3rem !important;'>Analytics & Insights</h2>", unsafe_allow_html=True)

view_annual = st.toggle("📊 Visualizar Valores Anuais")
value_col = "Custo_Anual_Calc" if view_annual else "Custo_Mensal_Calc"
y_axis_title = "Custo Anual (R$)" if view_annual else "Custo Mensal (R$)"

col_chart1, col_chart2 = st.columns(2)

with col_chart1:
    with st.container(border=True):
        st.markdown("<h4 style='text-align: center; margin-top:0; color:var(--text-secondary); font-size:0.9rem; text-transform:uppercase;'>Account Allocation</h4>", unsafe_allow_html=True)
        
        df_conta = df.groupby("Empresa/Conta")[value_col].sum().reset_index()
        colors = ['#4F46E5', '#10B981', '#F59E0B', '#6366F1']
        
        fig1 = go.Figure()
        
        total_spend = df_conta[value_col].sum()
        total_str = f"R$ {total_spend:,.0f}".replace(",", "X").replace(".", ",").replace("X", ".")
        
        fig1.add_trace(go.Pie(
            labels=df_conta['Empresa/Conta'],
            values=df_conta[value_col],
            hole=0.75,
            marker_colors=colors,
            textinfo='none',
            hovertemplate="<b>%{label}</b><br>R$ %{value:,.2f} (%{percent})<extra></extra>",
            sort=False,
            showlegend=False
        ))
        
        fig1.update_layout(
            paper_bgcolor='rgba(0,0,0,0)', 
            plot_bgcolor='rgba(0,0,0,0)',
            font_color='#E2E8F0',
            margin=dict(t=20, b=20, l=20, r=20),
            showlegend=False,
            annotations=[dict(
                text=f"<span style='font-size:24px; font-weight:700;'>{total_str}</span><br><span style='font-size:12px; color:#94A3B8;'>{'Anual' if view_annual else 'Mensal'}</span>",
                x=0.5, y=0.5, font_size=20, showarrow=False
            )]
        )
        st.plotly_chart(fig1, width="stretch")

with col_chart2:
    with st.container(border=True):
        st.markdown("<h4 style='text-align: center; margin-top:0; color:var(--text-secondary); font-size:0.9rem; text-transform:uppercase;'>Top Active Subscriptions</h4>", unsafe_allow_html=True)
        
        df_servico = df.groupby(["Serviço", "Empresa/Conta"])[value_col].sum().reset_index()
        df_servico = df_servico.sort_values(by=value_col, ascending=True)
        
        # Salva a ordem original das empresas para que as cores não mudem
        original_company_order = df_servico['Empresa/Conta'].drop_duplicates().tolist()
        
        # Remove itens com custo 0 (ex: Canva)
        df_servico = df_servico[df_servico[value_col] > 0]
        
        fig2 = px.bar(df_servico, y='Serviço', x=value_col, 
                      color='Empresa/Conta',
                      orientation='h',
                      category_orders={"Empresa/Conta": original_company_order},
                      color_discrete_sequence=['#4F46E5', '#10B981', '#F59E0B', '#6366F1'])
        
        fig2.update_traces(hovertemplate="<b>%{y}</b><br>%{data.name}<br>R$ %{x:,.2f}<extra></extra>")
        
        fig2.update_layout(
            paper_bgcolor='rgba(0,0,0,0)', 
            plot_bgcolor='rgba(0,0,0,0)',
            font_color='#E2E8F0',
            margin=dict(t=20, b=20, l=10, r=20),
            showlegend=False,
            xaxis=dict(showgrid=False, visible=False),
            yaxis=dict(showgrid=False, title="")
        )
        st.plotly_chart(fig2, width="stretch")

# =======================
# INSIGHTS (SMART UX)
# =======================
if economia_mensal > 0:
    insight_html = f"""
    <div class="smart-insight-card">
        <div class="smart-insight-icon">✨</div>
        <div class="smart-insight-content">
            <h4>Smart Insight</h4>
            <p>Você pode economizar até <b>{fmt_brl(economia_mensal)}/mês</b> ao revisar as {len(economia_df)} assinaturas que precisam de atenção.</p>
        </div>
    </div>
    """
    st.markdown(insight_html, unsafe_allow_html=True)
