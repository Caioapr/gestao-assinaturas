import streamlit as st
import pandas as pd
import os
from decimal import Decimal

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

st.title("📊 Painel de Gestão Centralizada de Assinaturas")
st.markdown("Aplicação avançada baseada na análise de prós/contras para gestão do plano corporativo.")

# Carregar os dados
df = load_data()

# Calcula custos normalizados para sumarização
df["Custo_Mensal_Calc"] = df.apply(calc_mensal, axis=1)
df["Custo_Anual_Calc"] = df.apply(calc_anual, axis=1)

# =======================
# MÉTRICAS E RESUMO
# =======================
total_mensal = df["Custo_Mensal_Calc"].sum()
projecao_anual = df["Custo_Anual_Calc"].sum()

# Economia potencial: assinaturas com ação de Cancelar, Não Renovar ou Avaliar
acoes_corte = ["Cancelar", "Não Renovar", "Avaliar"]
economia_df = df[df["Ação"].isin(acoes_corte)]
economia_mensal = economia_df["Custo_Mensal_Calc"].sum()

st.subheader("Resumo Financeiro Atual", divider="blue")
col1, col2, col3, col4 = st.columns(4)
col1.metric("Custo Total Mensal (Equivalente)", fmt_brl(total_mensal))
col2.metric("Projeção Anual Total", fmt_brl(projecao_anual))
col3.metric("Economia Potencial / Mês", fmt_brl(economia_mensal),
            delta=f"- {fmt_brl(economia_mensal)}", delta_color="inverse")
col4.metric("Qtd. de Assinaturas", f"{len(df)} ativas")

# =======================
# TABELA EDITÁVEL (ESTILO EXCEL)
# =======================
st.subheader("📝 Tabela Completa: Prós, Contras e Decisões", divider="blue")
st.markdown("Analise os pontos fortes e fracos de cada assinatura. Clique duas vezes em qualquer campo para editar.")

edited_df = st.data_editor(
    df[COLUNAS],
    num_rows="dynamic",
    use_container_width=True,
    height=500,
    column_config={
        "Serviço": st.column_config.TextColumn("Serviço", required=True, width="medium"),
        "Empresa/Conta": st.column_config.SelectboxColumn("Conta", options=["PV Móveis", "Grupo Paluto", "V4 Company", "Geral/Outro"], required=True, width="medium"),
        "Custo (R$)": st.column_config.NumberColumn("Custo (R$)", format="R$ %.2f", min_value=0.0, step=0.01, width="small"),
        "Frequência": st.column_config.SelectboxColumn("Frequência", options=["Mensal", "Anual"], required=True, width="small"),
        "Tags": st.column_config.TextColumn("Tags/Categoria", width="small"),
        "Prós": st.column_config.TextColumn("Prós (+)", width="large"),
        "Contras": st.column_config.TextColumn("Contras (-)", width="large"),
        "Ação": st.column_config.SelectboxColumn("Ação", options=ACOES, required=True, width="small"),
        "Decisão": st.column_config.TextColumn("Decisão Final", width="large"),
    },
    column_order=COLUNAS
)

col_save, col_export = st.columns([1, 2])
with col_save:
    if st.button("💾 Salvar Alterações no Excel", type="primary"):
        save_data(edited_df)
        st.success("Alterações salvas com sucesso em 'assinaturas_v4.xlsx'!")
        st.rerun()
with col_export:
    st.download_button(
        label="📥 Baixar Relatório (CSV)",
        data=edited_df.to_csv(index=False, sep=';', decimal=',').encode("utf-8-sig"),
        file_name="relatorio_gestao_assinaturas.csv",
        mime="text/csv",
        help="Baixa a tabela atualizada num formato leve e fácil de abrir no Excel para compartilhar."
    )



# =======================
# VISUALIZAÇÃO DOS PRÓS E CONTRAS (CARDS)
# =======================
st.subheader("🔍 Resumo Executivo das Contas em Revisão/Cancelamento")
st.markdown("Um foco detalhado nas assinaturas que exigem atenção imediata.")

revisar_df = edited_df[edited_df["Ação"].isin(["Cancelar", "Avaliar", "Não Renovar", "Migrar"])]
if not revisar_df.empty:
    for index, row in revisar_df.iterrows():
        with st.expander(f"⚠️ {row['Serviço']} ({row['Empresa/Conta']}) - Ação: {row['Ação']}", expanded=True):
            col_pro, col_con = st.columns(2)
            with col_pro:
                st.success(f"**Prós:**\n{row['Prós']}")
            with col_con:
                st.error(f"**Contras:**\n{row['Contras']}")
            st.info(f"**Decisão:** {row['Decisão']}")
else:
    st.success("Não há assinaturas pendentes para revisar ou cancelar!")

st.caption("Os dados são salvos localmente em `assinaturas_v4.xlsx` e podem ser abertos no Excel a qualquer momento.")
