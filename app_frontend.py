import streamlit as st
import pandas as pd
import pdfplumber
import re
import requests

# --- CONFIGURAÇÃO DA PÁGINA (FORÇANDO TEMA CLARO) ---
st.set_page_config(page_title="Raichu Pro ⚡", layout="wide", page_icon="⚡", initial_sidebar_state="auto")

# --- ESTILIZAÇÃO CUSTOMIZADA (MODO CLARO + UPLOAD ESTILIZADO) ---
st.markdown(
    """
    <style>
        /* Fundo geral da aplicação limpo e claro */
        .stApp {
            background-color: #F8F9FA !important;
            color: #2C3E50 !important;
        }
        
        /* Textos e títulos principais em cor escura legível */
        h1, h2, h3, h4, h5, h6, p, span, label {
            color: #2C3E50 !important;
        }

        /* Botões principais com o Laranja Elétrico do Raichu */
        .stButton > button {
            background-color: #FF8C00 !important;
            color: white !important;
            font-weight: bold !important;
            border-radius: 8px !important;
            border: 2px solid #E67E22 !important;
        }
        .stButton > button:hover {
            background-color: #E67E22 !important;
            border-color: #D35400 !important;
            color: white !important;
        }

        /* Estilização Profissional da Caixa de Upload de PDF */
        [data-testid="stFileUploader"] {
            background-color: #FFFFFF !important;
            border: 2px dashed #D35400 !important;
            border-radius: 10px !important;
            padding: 15px !important;
        }
        [data-testid="stFileUploader"] section {
            background-color: #FFFFFF !important;
        }
        [data-testid="stFileUploader"] section div span, 
        [data-testid="stFileUploader"] section div small {
            color: #2C3E50 !important;
        }

        /* Cards e balões informativos em fundo branco com borda elegante */
        .raichu-card {
            background-color: #FFFFFF;
            border-left: 5px solid #FF8C00;
            padding: 20px;
            border-radius: 10px;
            margin-bottom: 20px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.08);
            border-top: 1px solid #EAECEE;
            border-right: 1px solid #EAECEE;
            border-bottom: 1px solid #EAECEE;
        }
        .raichu-title {
            color: #D35400 !important;
            font-weight: bold;
        }
    </style>
    """,
    unsafe_allow_html=True
)

# --- FUNÇÕES DE APOIO ---
def limpar_texto_bloco(txt):
    if not txt: return ""
    linhas = txt.split('\n')
    linhas_limpas = []
    for l in linhas:
        l_strip = l.strip()
        if re.search(r'(?i)Página \d+ de \d+', l_strip): continue
        if re.search(r'(?i)UNIVERSIDADE FEDERAL DE SANTA MARIA', l_strip): continue
        if re.search(r'(?i)PROJETO NA ÍNTEGRA', l_strip): continue
        if re.search(r'(?i)Consulte em http', l_strip): continue
        if re.search(r'\d{2}/\d{2}/\d{4}\s+\d{2}:\d{2}', l_strip): continue
        if re.search(r'[A-F0-9]{4}(?:\.[A-F0-9]{4}){7}', l_strip): continue
        linhas_limpas.append(l)

    txt_final = " ".join([l.strip() for l in linhas_limpas if l.strip()])
    if txt_final.strip() in ["", "-", ".", "Não se aplica"]:
        return ""
    return txt_final.strip()

def identificar_instrumento_juridico(texto):
    texto = texto or ""
    m_instr = re.search(r'Instrumento jurídico celebrado\s*:\s*(.*?)(?:\n|$)', texto, re.IGNORECASE)
    if m_instr:
        valor = m_instr.group(1).strip().lower()
        if "acordo de cooperação técnica" in valor or "cooperação técnica" in valor or "cooperacao tecnica" in valor:
            return "Acordo de Cooperação Técnica (ACT)"
        if "contrato global" in valor:
            return "Contrato Global (CG)"
        if "acordo de parceria" in valor or "parceria" in valor:
            return "Acordo de Parceria (AP)"

    texto_low = texto.lower()
    if "acordo de cooperação técnica" in texto_low or "cooperação técnica" in texto_low or "cooperacao tecnica" in texto_low:
        return "Acordo de Cooperação Técnica (ACT)"
    if "contrato global" in texto_low:
        return "Contrato Global (CG)"
    if "acordo de parceria" in texto_low or "parceria" in texto_low:
        return "Acordo de Parceria (AP)"

    return "Acordo de Cooperação Técnica (ACT)"

# ==============================================================================
# NAVEGAÇÃO PRINCIPAL (ABAS)
# ==============================================================================
aba_home, aba_gerador = st.tabs(["⚡ Início & Sobre", "🚀 Gerador de Documentos"])

# ==============================================================================
# TELA 1: HOME / APRESENTAÇÃO DO RAICHU PRO
# ==============================================================================
with aba_home:
    st.markdown("<h1 style='color: #FF8C00;'>⚡ Bem-vindo ao Raichu Pro</h1>", unsafe_allow_html=True)
    st.markdown("### Automatização Inteligente de Documentação de Projetos")
    
    st.markdown("---")
    
    col_info1, col_info2 = st.columns([2, 1])
    
    with col_info1:
        st.markdown(
            """
            <div class="raichu-card">
                <h3 class="raichu-title">🎯 Objetivo do Sistema</h3>
                <p>O <b>Raichu Pro</b> foi desenvolvido para eliminar o trabalho manual e repetitivo na criação de documentações acadêmicas e administrativas de projetos. Através da leitura inteligente de relatórios em PDF, o sistema extrai dados de títulos, equipes, resumos e prazos, gerando instantaneamente pacotes completos em formatos <b>Word (.docx)</b> e <b>Excel (.xlsx)</b> perfeitamente formatados.</p>
            </div>
            """,
            unsafe_allow_html=True
        )
        
        st.markdown("### 📊 Opções de Instrumentos Jurídicos Suportados")
        
        b_col1, b_col2, b_col3 = st.columns(3)
        with b_col1:
            st.markdown(
                """
                <div class="raichu-card" style="border-left-color: #F1C40F;">
                    <h4 style='color: #D35400;'>ACT</h4>
                    <p style='font-size: 13px;'><b>Acordo de Cooperação Técnica</b><br>Foco em cooperações acadêmicas sem repasse financeiro direto ou fundações obrigatórias.</p>
                </div>
                """,
                unsafe_allow_html=True
            )
        with b_col2:
            st.markdown(
                """
                <div class="raichu-card" style="border-left-color: #FF8C00;">
                    <h4 style='color: #D35400;'>CG</h4>
                    <p style='font-size: 13px;'><b>Contrato Global</b><br>Gerenciamento integrado com fundações de apoio parceiras (FATEC, FUNDEP, etc.).</p>
                </div>
                """,
                unsafe_allow_html=True
            )
        with b_col3:
            st.markdown(
                """
                <div class="raichu-card" style="border-left-color: #E67E22;">
                    <h4 style='color: #D35400;'>AP</h4>
                    <p style='font-size: 13px;'><b>Acordo de Parceria</b><br>Projetos voltados à inovação, P&D e parcerias estratégicas institucionais.</p>
                </div>
                """,
                unsafe_allow_html=True
            )

    with col_info2:
        st.markdown(
            """
            <div class="raichu-card">
                <h4 class="raichu-title">ℹ️ Informações da Versão</h4>
                <p><b>Versão:</b> 2.1.0 (FastAPI Edition)</p>
                <p><b>Desenvolvido por:</b> Julio Maia dos Santos - Estudante de graduação em Engenharia Elétrica  👨‍💻⚡</p>
                <p><b>Arquitetura:</b> Decoupled (Front-End Streamlit + Back-End FastAPI)</p>
                <hr style='border-color: #EAECEE;'>
                <p style='font-size: 12px; color: #666;'>⚡ Sistema otimizado para alta performance e precisão em relatórios institucionais.</p>
            </div>
            """,
            unsafe_allow_html=True
        )

# ==============================================================================
# TELA 2: GERADOR DE DOCUMENTOS (O SISTEMA PRINCIPAL)
# ==============================================================================
with aba_gerador:
    st.markdown("### ⚡ Painel de Operações - Raichu Pro")
    st.markdown("Insira o seu relatório em PDF abaixo para iniciar o preenchimento automático dos dados.")
    
    arquivo_pdf = st.file_uploader("Insira o seu relatório do projeto (.pdf)", type=["pdf"])

    fundacoes_dados = {
        "FATEC": {"fundacao": "FATEC - Fundação de Apoio à Tecnologia e Ciência", "sigla_fundacao": "FATEC"},
        "FUNDEP": {"fundacao": "FUNDEP - Fundação de Desenvolvimento da Pesquisa", "sigla_fundacao": "FUNDEP"},
        "FAURGS": {"fundacao": "FAURGS - Fundação de Apoio à Universidade Federal do Rio Grande do Sul", "sigla_fundacao": "FAURGS"},
        "FDMS": {"fundacao": "FDMS - Fundação Delfim Mendes Silveira", "sigla_fundacao": "FDMS"}
    }

    dados_extraidos = {
        "titulo": "", "numero": "", "empresa": "", "data_inicio_proj": "", "data_termino_proj": "",
        "resumo": "", "objetivos": "", "justificativa_proj": "", "resultados": "", "importancia_projeto": "",
        "plano_gestao": "", "objetivo_estrategico": "", "inovacao_bool": "", "inovacao_potencial": "",
        "instrumento_juridico_pdf": "",
        "classificacoes_raw": [], "equipe_raw": [], "unidades_raw": [], "regioes_raw": [],
        "fundacao_sugerida": "FATEC", "tipo_processo_sugerido": "Acordo de Cooperação Técnica (ACT)"
    }

    # ==============================================================================
    # MOTOR DE LEITURA DO PDF
    # ==============================================================================
    if arquivo_pdf:
        try:
            texto_completo = ""
            with pdfplumber.open(arquivo_pdf) as pdf:
                for page in pdf.pages:
                    extraido = page.extract_text()
                    if extraido: texto_completo += extraido + "\n"

            texto_limpo = re.sub(r'---\s*PAGE\s*\d+\s*---', '\n', texto_completo, flags=re.IGNORECASE)
            texto_limpo = re.sub(r'\d{2}/\d{2}/\d{4}\s\d{2}:\d{2}', '', texto_limpo)
            texto_limpo = re.sub(r'[A-F0-9]{4}(?:\.[A-F0-9]{4}){7}', '', texto_limpo, flags=re.IGNORECASE)
            texto_limpo = re.sub(r'Consulte em http[^\n]+', '', texto_limpo, flags=re.IGNORECASE)
            texto_limpo = re.sub(r'Registrado em:\s*\d{2}/\d{2}/\d{4}', '', texto_limpo, flags=re.IGNORECASE)
            texto_limpo = re.sub(r'UNIVERSIDADE FEDERAL DE SANTA MARIA - UFSM', '', texto_limpo, flags=re.IGNORECASE)
            texto_limpo = re.sub(r'PROJETO NA ÍNTEGRA', '', texto_limpo, flags=re.IGNORECASE)

            lixos_para_apagar = [
                r'(?i)PARTICIPANTE\s+V[ÍI]NCULO\s+CURSO/LOTA[ÇC][ÃA]O\s+FUN[ÇC][ÃA]O\s*\$\$\$',
                r'(?i)PARTICIPANTE\s+V[ÍI]NCULO\s+CURSO/LOTA[ÇC][ÃA]O\s+FUN[ÇC][ÃA]O',
                r'(?i)CH\s+DENTRO\s+CH\s+FORA\s+IN[ÍI]CIO\s+T[ÉE]RMINO\s+OBSERVA[ÇC][ÃA]O',
                r'(?i)UNIDADE\s+FUN[ÇC][ÃA]O\s+VALOR\s+IN[ÍI]CIO\s+T[ÉE]RMINO',
                r'(?i)TIPO\s+DE\s+CLASSIFICA[ÇC][ÃA]O\s+CLASSIFICA[ÇC][ÃA]O'
            ]
            for lixo in lixos_para_apagar:
                texto_limpo = re.sub(lixo, ' ', texto_limpo)

            for sigla in ["FATEC", "FUNDEP", "FAURGS", "FDMS"]:
                if re.search(r'\b' + sigla + r'\b', texto_limpo, re.IGNORECASE):
                    dados_extraidos["fundacao_sugerida"] = sigla
                    break

            dados_extraidos["tipo_processo_sugerido"] = identificar_instrumento_juridico(texto_limpo)

            def extrair(regex, group=1):
                m = re.search(regex, texto_limpo, re.IGNORECASE)
                return m.group(group).strip() if m else ""

            dados_extraidos["titulo"] = extrair(r'Título:\s*(.*?)\n')
            dados_extraidos["numero"] = extrair(r'Número:\s*(\d+)')
            dados_extraidos["data_inicio_proj"] = extrair(r'Início:\s*(\d{2}/\d{2}/\d{4})')
            dados_extraidos["data_termino_proj"] = extrair(r'Término:\s*(\d{2}/\d{2}/\d{4})')
            dados_extraidos["classificacao"] = extrair(r'Classificação:\s*(.*?)\n')
            dados_extraidos["empresa"] = extrair(r'(?:Financiador[a]?|Empresa|Cooperante|Financiador|Instituição):\s*(.*?)\n')
            dados_extraidos["instrumento_juridico_pdf"] = extrair(r'Instrumento jurídico celebrado:\s*([^\n]+)')

            m_coord = re.search(r'Responsável pelo projeto:\s*(.*?)\s*\(\s*(\d+)\s*\)', texto_limpo, re.IGNORECASE)
            if m_coord: dados_extraidos["coord_geral_pdf"] = {"nome": m_coord.group(1).strip(), "siape": m_coord.group(2).strip()}

            m_fisc = re.search(r'Fiscal:\s*(\d+)\s*-\s*(.*?)\s*\(', texto_limpo, re.IGNORECASE)
            if m_fisc: dados_extraidos["fiscal_pdf"] = {"siape": m_fisc.group(1).strip(), "nome": m_fisc.group(2).strip()}

            def extrair_bloco(inicio_regex, fins_regex):
                m_inicio = re.search(inicio_regex, texto_limpo, re.IGNORECASE)
                if not m_inicio: return ""
                idx = m_inicio.end()
                end_idx = len(texto_limpo)
                for f in fins_regex:
                    mf = re.search(f, texto_limpo[idx:], re.IGNORECASE)
                    if mf:
                        pos = idx + mf.start()
                        if pos < end_idx: end_idx = pos
                return texto_limpo[idx:end_idx].strip()

            dados_extraidos["resumo"] = limpar_texto_bloco(extrair_bloco(r'Resumo:', [r'Objetivos:']))
            dados_extraidos["objetivos"] = limpar_texto_bloco(extrair_bloco(r'Objetivos:', [r'Justificativa:']))
            dados_extraidos["justificativa_proj"] = limpar_texto_bloco(extrair_bloco(r'Justificativa:', [r'Resultados esperados:']))
            dados_extraidos["resultados"] = limpar_texto_bloco(extrair_bloco(r'Resultados esperados:', [r'PARTICIPANTES', r'PLANO DE GESTÃO', r'UNIDADES VINCULADAS']))

            classif_blk = extrair_bloco(r'CLASSIFICAÇÕES', [r'PLANO DE GESTÃO'])
            for line in classif_blk.split('\n'):
                line = line.strip()
                if not line or "TIPO DE CLASSIFICAÇÃO" in line.upper() or line.upper() == "CLASSIFICAÇÃO" or "CLASSIFICAÇÕES" in line.upper(): continue

                m = re.search(r'\s+(\d{1,5}\.\d.*|\d{1,5}\s+-.*)', line)
                if m:
                    tipo = line[:m.start()].strip()
                    valor = line[m.start():].strip()
                    dados_extraidos["classificacoes_raw"].append({"Tipo de Classificação": tipo, "Classificação": valor})
                else:
                    if len(line) > 5:
                        dados_extraidos["classificacoes_raw"].append({"Tipo de Classificação": line, "Classificação": ""})

            bloco_participantes = extrair_bloco(r'PARTICIPANTES', [r'UNIDADES VINCULADAS\s*\n', r'CLASSIFICAÇÕES', r'REGIÕES DE ATUAÇÃO'])
            if not bloco_participantes:
                bloco_participantes = texto_limpo

            matches_participantes = list(re.finditer(r'(\d{5,15})\s*-\s*([A-ZÀ-Ÿ\s\']+?)\s*(?=[A-ZÀ-Ÿ][a-zà-ÿ]|UNIDADES VINCULADAS|CLASSIFICAÇÕES|$)', bloco_participantes))
            
            for i, match in enumerate(matches_participantes):
                siape = match.group(1).strip()
                nome_bruto = match.group(2).strip()
                nome_limpo = re.sub(r'\s+', ' ', nome_bruto).strip()
                corte_idx = len(nome_limpo)
                palavras_corte = ["VÍNCULO", "VINCULO", "CURSO", "LOTAÇÃO", "LOTACAO", "FUNÇÃO", "FUNCAO", "UNIDADE", "CLASSIFICA", "PARTICIPANTE", "CH DENTRO", "CH FORA", "INÍCIO", "INICIO", "TÉRMINO", "TERMINO", "DEPARTAMENTO", "OBSERVA", "VALOR", "TIPO", "$$$"]
                for p in palavras_corte:
                    idx = nome_limpo.upper().find(p)
                    if idx != -1 and idx < corte_idx: corte_idx = idx
                nome = nome_limpo[:corte_idx].strip(" -/")
                
                if not nome or len(nome) < 2: continue

                pos_inicio = match.end()
                pos_fim = matches_participantes[i+1].start() if i + 1 < len(matches_participantes) else pos_inicio + 400
                janela_texto = bloco_participantes[pos_inicio:pos_fim]
                janela_limpa = " ".join(janela_texto.split())
                vinculo, lotacao, funcao, bolsa, ch_d, ch_f, data_ini, data_fim = "Outro", "", "Participante", "Não", "0", "0", "", ""

                m_info = re.search(r'(Coordenador Administrativo|Coordenador|Estagiário|Colaborador|Fiscal|Participante|Membro|Pesquisador|Responsável Técnico|Responsável|Técnico|Bolsista)\s+(Sim|Não|Nao)[\s\S]*?(\d+)\s+(\d+)\s+(\d{2}/\d{2}/\d{4})\s+(\d{2}/\d{2}/\d{4})', janela_limpa, re.IGNORECASE)

                prefixos_ufsm_regex = [
                    r"Estudante de Pós-\s*graduação", r"Estudante de Pós-Graduação", r"Estudante de Graduação", r"Estudante de graduação",
                    r"Estudante de Ensino Médio", r"Técnico[- ]Administrativo em Educação", r"Técnico[- ]Administrativo", r"Tecnico Administrativo",
                    r"Docente", r"Pesquisador", r"Participante Externo", r"Visitante", r"Estudante", r"Servidor", r"Outro"
                ]

                if m_info:
                    text_before = janela_limpa[:m_info.start()].strip()
                    text_after = janela_limpa[m_info.end():].strip()
                    vinculo_encontrado = False
                    for prefixo in prefixos_ufsm_regex:
                        match_vinculo = re.match(r'(?i)^' + prefixo, text_before)
                        if match_vinculo:
                            vinculo = match_vinculo.group(0).strip()
                            vinculo = re.sub(r'\s+', ' ', vinculo).replace('- ', '-') 
                            lotacao = text_before[match_vinculo.end():].strip()
                            vinculo_encontrado = True
                            break
                    if not vinculo_encontrado:
                        partes = text_before.split(" ", 1)
                        vinculo = partes[0] if len(partes) > 0 else "Outro"
                        lotacao = partes[1] if len(partes) > 1 else ""

                    if "técnico" in vinculo.lower() or "técnico" in janela_limpa.lower() or "administrativo" in janela_limpa.lower():
                        if "administrativo" in janela_limpa.lower() or "educação" in janela_limpa.lower() or "educacao" in janela_limpa.lower():
                            vinculo = "Técnico-Administrativo em Educação"

                    lotacao_limpa = re.sub(r'\s+', ' ', lotacao).strip()
                    corte_lot = len(lotacao_limpa)
                    palavras_corte_lotacao = ["UNIDADE", "CLASSIFICA", "PARTICIPANTE", "FUNÇÃO", "FUNCAO", "VALOR", "INÍCIO", "INICIO", "TÉRMINO", "TERMINO", "OBSERVA", "TIPO", "$$$"]
                    for p in palavras_corte_lotacao:
                        idx = lotacao_limpa.upper().find(p)
                        if idx != -1 and idx < corte_lot: corte_lot = idx
                    lotacao = lotacao_limpa[:corte_lot].strip(" -/")

                    funcao = m_info.group(1).title()
                    bolsa = m_info.group(2).title()
                    ch_d = m_info.group(3)
                    ch_f = m_info.group(4)
                    data_ini = m_info.group(5)
                    data_fim = m_info.group(6)
                else:
                    for prefixo in prefixos_ufsm_regex:
                        match_vinculo = re.match(r'(?i)^' + prefixo, janela_limpa)
                        if match_vinculo:
                            vinculo = match_vinculo.group(0).strip()
                            vinculo = re.sub(r'\s+', ' ', vinculo).replace('- ', '-')
                            break

                dados_extraidos["equipe_raw"].append({
                    "Nome": nome, "SIAPE": siape, "Vínculo": vinculo.title(), "Lotação": lotacao,
                    "Função": funcao, "Bolsa": bolsa, "CH_D": ch_d, "CH_F": ch_f, "Início": data_ini, "Término": data_fim,
                    "Chefia Imediata": "", "SIAPE Chefia": ""
                })

        except Exception as e:
            st.error(f"❌ Erro no processamento do PDF: {str(e)}")

    # ==============================================================================
    # VALIDAÇÃO E FORMULÁRIOS
    # ==============================================================================
    if arquivo_pdf:
        st.markdown("---")
        st.markdown("### 2️⃣ Passo 2: Validação do Processo")

        tipo_sugerido = dados_extraidos.get("tipo_processo_sugerido", "Acordo de Cooperação Técnica (ACT)")
        opcoes_processo = ["Acordo de Parceria (AP)", "Contrato Global (CG)", "Acordo de Cooperação Técnica (ACT)"]
        try:
            indice_padrao = opcoes_processo.index(tipo_sugerido)
        except ValueError:
            indice_padrao = 2

        st.info(f"📌 Instrumento jurídico identificado automaticamente: **{tipo_sugerido}**.")

        tipo_processo = st.radio(
            "Selecione o Tipo de Processo:",
            opcoes_processo, index=indice_padrao, horizontal=True, key="tipo_processo_radio"
        )

        if tipo_processo == "Acordo de Cooperação Técnica (ACT)":
            st.info("💡 Processos do tipo **ACT** não necessitam de Fundação de Apoio.")
            status_fund, fund_sigla = "Não possui", "ACT"
        else:
            fund_sugerida = dados_extraidos.get("fundacao_sugerida", "FATEC")
            st.info(f"🤖 O robô identificou que este projeto está vinculado à fundação: **{fund_sugerida}**.")
            fundacao_correta = st.radio("A fundação identificada está correta?", ["Sim", "Não"], index=0, horizontal=True)

            if fundacao_correta == "Sim":
                fund_sigla = fund_sugerida
                status_fund = "Já definida"
            else:
                fund_sigla = st.selectbox("Selecione a fundação correta:", list(fundacoes_dados.keys()))
                status_fund = "Já definida"

        st.markdown("---")
        st.markdown("### 3️⃣ Passo 3: Conferência e Edição de Dados")

        with st.expander("📝 Detalhes do Projeto e Textos Longos", expanded=True):
            c1, c2 = st.columns(2)
            with c1:
                tit_proj = st.text_input("Nome do Projeto (Título)", value=dados_extraidos.get("titulo", ""))
                n_proj = st.text_input("Número do Registro GAP", value=dados_extraidos.get("numero", ""))
                resumo = st.text_area("Resumo do Projeto", value=dados_extraidos.get("resumo", ""), height=120)
                objetivos = st.text_area("Objetivos do Projeto", value=dados_extraidos.get("objetivos", ""), height=120)
                justificativa = st.text_area("Justificativa do Projeto", value=dados_extraidos.get("justificativa_proj", ""), height=120)
                importancia = st.text_area("Importância do Projeto", value=dados_extraidos.get("importancia_projeto", ""), height=80)
                justificativa_fund = st.text_area("Justificativa para escolha da Fundação", height=80)
            with c2:
                diretor_unidade = st.text_input("Diretor da Unidade")
                siape_diretor = st.text_input("SIAPE do Diretor")
                classificacao_final = st.text_input("Classificação", value=dados_extraidos.get("classificacao", ""))
                data_termino_edit = st.text_input("Data de Término", value=dados_extraidos.get("data_termino_proj", ""))
                instrumento_juridico_edit = st.text_input("Instrumento Jurídico (Excel)", value=dados_extraidos.get("instrumento_juridico_pdf", ""))
                resultados = st.text_area("Resultados Esperados", value=dados_extraidos.get("resultados", ""), height=120)
                metas = st.text_area("Metas do Projeto (Opcional)", height=120)

        st.markdown("---")
        st.subheader("👨‍🏫 Coordenador e Fiscal do Projeto")
        col_c1, col_c2, col_f1, col_f2 = st.columns(4)
        c_g_n = col_c1.text_input("Coordenador", value=dados_extraidos.get("coord_geral_pdf", {}).get("nome", "") if dados_extraidos.get("coord_geral_pdf") else "")
        c_g_s = col_c2.text_input("SIAPE Coord.", value=dados_extraidos.get("coord_geral_pdf", {}).get("siape", "") if dados_extraidos.get("coord_geral_pdf") else "")
        f_nome = col_f1.text_input("Fiscal", value=dados_extraidos.get("fiscal_pdf", {}).get("nome", "") if dados_extraidos.get("fiscal_pdf") else "")
        f_siape = col_f2.text_input("SIAPE Fiscal", value=dados_extraidos.get("fiscal_pdf", {}).get("siape", "") if dados_extraidos.get("fiscal_pdf") else "")

        st.markdown("---")
        col_adm1, col_adm2 = st.columns(2)
        nome_coord_adm = col_adm1.text_input("Coordenador Administrativo")
        siape_coord_adm = col_adm2.text_input("SIAPE Coord. Adm.")

        st.markdown("---")
        st.subheader("🏢 Empresas / Parceiras")
        num_empresas = st.number_input("Quantas empresas/instituições parceiras participam deste projeto?", min_value=1, max_value=10, value=1)

        nomes_empresas_validas = []
        for i in range(num_empresas):
            key_emp = f"nome_empresa_simples_{i}"
            if key_emp not in st.session_state:
                st.session_state[key_emp] = dados_extraidos.get("empresa", "") if i == 0 else ""
            nome_emp = st.text_input(f"Nome da Empresa {i+1}", key=key_emp)
            if nome_emp and nome_emp.strip() != "":
                nomes_empresas_validas.append(nome_emp.strip())

        st.markdown("---")
        st.write("### 📊 Tabelas Estruturadas Consolidadas")
        t1 = st.tabs(["👥 Equipe"])[0]

        with t1:
            st.warning("⚠️ Preencha as colunas 'Chefia Imediata' e 'SIAPE Chefia' para cada participante.")
            equipe_final = dados_extraidos["equipe_raw"].copy()
            if f_nome and not any(e["SIAPE"] == f_siape for e in equipe_final):
                equipe_final.append({"Nome": f_nome, "SIAPE": f_siape, "Vínculo": "Docente", "Lotação": "DEPARTAMENTO", "Função": "Fiscal", "CH_D": "0", "CH_F": "0", "Bolsa": "Não", "Início": "", "Término": "", "Chefia Imediata": "", "SIAPE Chefia": ""})

            df_equipe = pd.DataFrame(equipe_final).fillna("")
            df_equipe_edit = st.data_editor(df_equipe, num_rows="dynamic", key="ed_equipe", use_container_width=True)
            equipe_final = df_equipe_edit.fillna("").to_dict(orient="records")

        st.markdown("---")
        st.markdown("### 4️⃣ Passo 4: Geração de Documentos")

        if st.button("⚡ Processar Documentos na API e Gerar Pacote", type="primary", use_container_width=True):
            payload = {
                "dados_projeto": {
                    "titulo": tit_proj, "numero": n_proj, "resumo": resumo, "objetivos": objetivos,
                    "justificativa": justificativa, "importancia": importancia, "justificativa_fund": justificativa_fund,
                    "resultados": resultados, "metas": metas, "classificacao": classificacao_final,
                    "instrumento_juridico": instrumento_juridico_edit, "data_termino": data_termino_edit,
                    "tipo_processo": tipo_processo, "fundacao_sigla": fund_sigla, "status_fund": status_fund,
                    "classificacoes_raw": dados_extraidos["classificacoes_raw"]
                },
                "pessoas": {
                    "coordenador": c_g_n, "siape_coord": c_g_s, "fiscal": f_nome, "siape_fiscal": f_siape,
                    "coord_adm": nome_coord_adm, "siape_adm": siape_coord_adm, "diretor": diretor_unidade, "siape_diretor": siape_diretor
                },
                "empresas": nomes_empresas_validas,
                "equipe": equipe_final
            }

            with st.spinner("⚡ O Raichu Pro está processando os documentos na API..."):
                try:
                    url_api = "https://codautom.onrender.com/gerar-zip-completo/"
                    resposta = requests.post(url_api, json=payload)
                    
                    if resposta.status_code == 200:
                        st.success("🔥 Pacote ZIP gerado com sucesso absoluto e pronto para download!")
                        # Balões removidos para manter um visual corporativo e profissional
                        
                        st.download_button(
                            label="⬇️ BAIXAR PACOTE DE DOCUMENTOS (.ZIP)",
                            data=resposta.content,
                            file_name=f"Documentos_Gerados_{fund_sigla}_{n_proj or 'Projeto'}.zip",
                            mime="application/zip",
                            type="primary",
                            use_container_width=True
                        )
                    else:
                        st.error(f"❌ Erro na API (Código {resposta.status_code}): {resposta.text}")
                except requests.exceptions.ConnectionError:
                    st.error("❌ Erro de conexão! Certifique-se de que o FastAPI está rodando na rede.")

st.markdown("<br><hr>", unsafe_allow_html=True)
st.markdown("<div style='text-align: center; color: #666666; padding: 10px; font-size: 14px;'>⚡ <b>Raichu Pro V2.1.0 (FastAPI Edition)</b> | Desenvolvido por Julio Maia 👨‍💻</div>", unsafe_allow_html=True)
