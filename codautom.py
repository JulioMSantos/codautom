import streamlit as st
import pdfplumber
import re
import os
import zipfile
import openpyxl
from datetime import datetime
from docxtpl import DocxTemplate

# --- 1. O RADAR DE ARQUIVOS (Busca automática nas pastas) ---
def encontrar_arquivo(nome_arquivo):
    """Vasculha todas as pastas do projeto até achar o arquivo exato"""
    for raiz, diretorios, arquivos in os.walk("."):
        if nome_arquivo in arquivos:
            return os.path.join(raiz, nome_arquivo)
    return None

def limpar_assinatura(valor, linha_padrao):
    """Se o usuário deixar em branco, injeta a linha para assinatura manual"""
    return str(valor) if valor.strip() != "" else linha_padrao

def obter_cnpj(fundacao):
    """Memória de CNPJs das Fundações da UFSM"""
    f = fundacao.upper()
    if "FATEC" in f: return "89.252.431/0001-59"
    if "FAURGS" in f: return "74.704.008/0001-75"
    if "FUNDEP" in f: return "18.720.938/0001-41"
    if "FDMS" in f: return "03.946.068/0001-46"
    return ""

# --- 2. MANIPULAÇÃO DO EXCEL ---
def preencher_excel_ficha_gestao(caminho_base, contexto, nome_saida):
    """Abre o Excel em branco, preenche as células exatas e salva"""
    try:
        wb = openpyxl.load_workbook(caminho_base)
        ws = wb.active 
        
        # Injeção nas células (respeitando células mescladas pela superior-esquerda)
        ws['C6'] = contexto.get("titulo_projeto", "")
        ws['J7'] = contexto.get("vigencia", "")
        ws['H12'] = contexto.get("fundacao", "")
        ws['K12'] = contexto.get("cnpj_fundacao", "")
        
        wb.save(nome_saida)
        return True
    except Exception as e:
        st.error(f"Erro ao tentar preencher o Excel: {e}")
        return False

# --- 3. FUNÇÃO DE EXTRAÇÃO DO PDF COM FILTRO DE ESTUDANTES ---
def ler_relatorio_gap(arquivo_upload):
    texto_completo = ""
    try:
        with pdfplumber.open(arquivo_upload) as pdf:
            for page in pdf.pages:
                extraido = page.extract_text()
                if extraido: texto_completo += extraido + "\n"
    except Exception as e:
        st.error(f"Erro ao ler PDF: {e}")
        return {}

    dados = {}
    
    # Extrações básicas de Metadados
    m_titulo = re.search(r'Título:\s*(.*?)\n', texto_completo, re.IGNORECASE)
    dados['titulo_projeto'] = m_titulo.group(1).strip() if m_titulo else ""

    m_num = re.search(r'Número:\s*(\d+)', texto_completo, re.IGNORECASE)
    dados['n_projeto'] = m_num.group(1).strip() if m_num else ""

    m_coord = re.search(r'Responsável pelo projeto:\s*(.*?)\s*\(\s*(\d+)\s*\)', texto_completo, re.IGNORECASE)
    if m_coord:
        dados['nome_coord'] = m_coord.group(1).strip()
        dados['siape_coord'] = m_coord.group(2).strip()
    else:
        dados['nome_coord'], dados['siape_coord'] = "", ""

    m_fisc = re.search(r'Fiscal:\s*(\d+)\s*-\s*(.*?)\s*\(', texto_completo, re.IGNORECASE)
    if m_fisc:
        dados['siape_fiscal_at'] = m_fisc.group(1).strip()
        dados['nome_fiscal_at'] = m_fisc.group(2).strip()
    else:
        dados['siape_fiscal_at'], dados['nome_fiscal_at'] = "", ""

    m_ini = re.search(r'Início:\s*(\d{2}/\d{2}/\d{4})', texto_completo, re.IGNORECASE)
    dados['inicio'] = m_ini.group(1).strip() if m_ini else ""
    
    m_fim = re.search(r'Término:\s*(\d{2}/\d{2}/\d{4})', texto_completo, re.IGNORECASE)
    dados['termino'] = m_fim.group(1).strip() if m_fim else ""

    for sigla in ["FATEC", "FUNDEP", "FAURGS", "FDMS"]:
        if re.search(r'\b' + sigla + r'\b', texto_completo, re.IGNORECASE):
            dados['fundacao'] = sigla
            break
    dados['fundacao'] = dados.get('fundacao', "")

    # 🛑 FILTRO DE ESTUDANTES: Isolamento do bloco de equipe e triagem de Carga Horária
    dados['equipe_servidores'] = []
    bloco_part = ""
    m_inicio_bloco = re.search(r'PARTICIPANTES', texto_completo, re.IGNORECASE)
    if m_inicio_bloco:
        idx = m_inicio_bloco.end()
        end_idx = len(texto_completo)
        for f in [r'UNIDADES VINCULADAS', r'CLASSIFICAÇÕES', r'PLANO DE GESTÃO', r'DECLARAÇÃO']:
            mf = re.search(f, texto_completo[idx:], re.IGNORECASE)
            if mf:
                pos = idx + mf.start()
                if pos < end_idx: end_idx = pos
        bloco_part = texto_completo[idx:end_idx]

    if bloco_part:
        linhas = bloco_part.split('\n')
        atual_part = None
        
        for linha in linhas:
            # Captura a linha inicial do participante (SIAPE - Nome)
            m_p = re.search(r'(\d{5,15})\s*-\s*([A-ZÀ-Ÿ\s]{5,100})', linha)
            if m_p:
                if atual_part:
                    # Analisa o vínculo do participante anterior antes de mudar o ponteiro
                    contexto_txt = " ".join(atual_part['contexto']).lower()
                    if not any(x in contexto_txt for x in ["estudante", "discente", "graduação", "graduacao", "médio", "medio"]):
                        dados['equipe_servidores'].append(atual_part)
                
                atual_part = {
                    "SIAPE": m_p.group(1).strip(),
                    "Nome": m_p.group(2).strip(),
                    "contexto": [linha]
                }
            elif atual_part:
                atual_part['contexto'].append(linha)
        
        # Validação do último integrante do bloco
        if atual_part:
            contexto_txt = " ".join(atual_part['contexto']).lower()
            if not any(x in contexto_txt for x in ["estudante", "discente", "graduação", "graduacao", "médio", "medio"]):
                dados['equipe_servidores'].append(atual_part)

    return dados

# --- 4. INTERFACE VISUAL DO SITE ---
st.set_page_config(page_title="Adequações do NAP", layout="wide")
st.title("Gerador de Adequações do NAP")
st.write("Faça o upload do relatório do projeto (GAP/Integra) para extrair os dados. Em seguida, verifique e complete as caixas de texto.")

arquivo_pdf = st.file_uploader("Anexar Relatório de Projetos (PDF)", type=['pdf'])
dados_extraidos = {}

if arquivo_pdf is not None:
    st.success("Relatório lido com sucesso! Os dados foram preenchidos nas caixas abaixo para sua conferência.")
    dados_extraidos = ler_relatorio_gap(arquivo_pdf)

st.markdown("---")

# --- 5. SELEÇÃO DO TIPO DE DOCUMENTO ---
tipo_documento = st.selectbox(
    "Selecione o tipo de processo:",
    ["1. Adendo (adequação de valores entre rubricas)", 
     "2. Aditivo (alteração de prazo e-ou valor)",
     "3. Ficha de ajuste (substituição de equipe)",
     "4. Ficha de Gestão (ajuste geral de item)"]
)

st.markdown("---")

# Painel Informativo do Filtro de Carga Horária
if dados_extraidos.get("equipe_servidores"):
    with st.expander("👥 Servidores Habilitados para Carga Horária (Docentes/TAEs)", expanded=False):
        st.warning("⚠️ Nota: Estudantes e bolsistas (Graduação/Pós) foram omitidos automaticamente desta lista conforme as regras de restrição de Carga Horária.")
        for s in dados_extraidos["equipe_servidores"]:
            st.write(f"• **{s['Nome']}** (SIAPE: {s['SIAPE']})")

# --- 6. CAMPOS DINÂMICOS ---
st.subheader("1. Dados Gerais do Projeto")
col1, col2 = st.columns(2)
with col1:
    titulo = st.text_area("Título do Projeto", value=dados_extraidos.get("titulo_projeto", ""))
    
    vigencia_padrao = ""
    if dados_extraidos.get("inicio") and dados_extraidos.get("termino"):
        vigencia_padrao = f"{dados_extraidos.get('inicio')} a {dados_extraidos.get('termino')}"
    vigencia = st.text_input("Data de Vigência", value=vigencia_padrao)
    
with col2:
    n_projeto = st.text_input("Nº do Projeto", value=dados_extraidos.get("n_projeto", ""))
    tipo_contrato = st.text_input("Tipo de Contrato", value="", placeholder="Ex: Contrato, Convênio, TED, Acordo...")

st.subheader("2. Pessoal Envolvido e Assinaturas")
st.info("💡 Dica: Se quiser que o documento saia com uma linha em branco para ser assinada à mão, basta deixar as caixas vazias.")

c_coord1, c_coord2 = st.columns(2)
coord = c_coord1.text_input("Coordenador(a) Atual (Nome)", value=dados_extraidos.get("nome_coord", ""))
siape = c_coord2.text_input("Coordenador(a) Atual (SIAPE)", value=dados_extraidos.get("siape_coord", ""))

c_dir1, c_dir2 = st.columns(2)
nome_diretor = c_dir1.text_input("Diretor(a) do Centro (Nome)", placeholder="Deixe em branco para imprimir linha")
siape_diretor = c_dir2.text_input("Diretor(a) do Centro (SIAPE)", placeholder="Vazio = linha")

if "1. Adendo" in tipo_documento:
    c_fisc1, c_fisc2 = st.columns(2)
    nome_fiscal_at = c_fisc1.text_input("Fiscal Atual (Nome)", value=dados_extraidos.get("nome_fiscal_at", ""))
    siape_fiscal_at = c_fisc2.text_input("Fiscal Atual (SIAPE)", value=dados_extraidos.get("siape_fiscal_at", ""))
else:
    nome_fiscal_at = dados_extraidos.get("nome_fiscal_at", "")
    siape_fiscal_at = dados_extraidos.get("siape_fiscal_at", "")

if "2. Aditivo" in tipo_documento or "3. Ficha de ajuste" in tipo_documento or "4. Ficha de Gestão" in tipo_documento:
    st.markdown("##### Dados da Fundação")
    c_fund1, c_fund2, c_fund3 = st.columns(3)
    fundacao = c_fund1.text_input("Nome da Fundação", value=dados_extraidos.get("fundacao", ""))
    
    cnpj_padrao = obter_cnpj(dados_extraidos.get("fundacao", ""))
    cnpj_fundacao = c_fund2.text_input("CNPJ da Fundação", value=cnpj_padrao)
    
    diretor_fund = c_fund3.text_input("Diretor(a) da Fundação", placeholder="Deixe em branco para imprimir linha")
    
    st.markdown("##### Dados do Setor Financeiro")
    c_sup1, c_sup2 = st.columns(2)
    nome_supfin = c_sup1.text_input("Supervisor(a) Financeiro (Nome)", placeholder="Deixe em branco para imprimir linha")
    siape_supfin = c_sup2.text_input("Supervisor(a) Financeiro (SIAPE)", placeholder="Vazio = linha")
else:
    fundacao, cnpj_fundacao, diretor_fund, nome_supfin, siape_supfin = "", "", "", "", ""

st.markdown("---")
st.subheader("3. Detalhes da Adequação")

# ==================== OPÇÃO 1: ADENDO RUBRICA ====================
if "1. Adendo" in tipo_documento:
    col3, col4 = st.columns(2)
    with col3:
        rubrica_ret = st.text_input("Rubrica Retirada (De onde sai o dinheiro)")
        valor_ret = st.text_input("Valor a ser remanejado (Ex: 1.500,00)")
    with col4:
        rubrica_col = st.text_input("Rubrica Colocada (Para onde vai o dinheiro)")
        
    if st.button("Gerar Documentos para Adendo"):
        contexto = {
            "titulo_projeto": str(titulo), "n_projeto": str(n_projeto), "tipo_contrato": str(tipo_contrato),
            "nome_coord_at": str(coord), "siape_coord_at": str(siape), "nome_coord_ant": str(coord), 
            "nome_fiscal_at": str(nome_fiscal_at), "siape_fiscal_at": str(siape_fiscal_at),
            "data_atual": str(datetime.now().strftime("%d/%m/%Y")),
            "rubrica_retirada": str(rubrica_ret), "valor_retirada": str(valor_ret), "rubrica_colocada": str(rubrica_col),
            "nome_diretor": limpar_assinatura(nome_diretor, "_________________"),
            "siape_diretor": limpar_assinatura(siape_diretor, "_______"), "cargo": "Coordenador",
            "nome_fiscal_ant": str(nome_fiscal_at), "siape_fiscal_ant": str(siape_fiscal_at)
        }
        caminho_word = encontrar_arquivo("adendo_alteracao_rubrica.docx")
        if not caminho_word: st.error("❌ Arquivo 'adendo_alteracao_rubrica.docx' não encontrado.")
        else:
            try:
                doc = DocxTemplate(caminho_word)
                doc.render(contexto)
                n_proj_limpo = str(n_projeto).replace('/', '-')
                nome_word = f"Adendo_Rubrica_{n_proj_limpo}.docx"
                doc.save(nome_word)
                
                nome_zip = f"Pacote_Adequacao_{n_proj_limpo}.zip"
                with zipfile.ZipFile(nome_zip, 'w') as zipf:
                    zipf.write(nome_word)
                    caminho_excel = encontrar_arquivo("ficha_ajuste_rubrica.xlsx")
                    if caminho_excel: zipf.write(caminho_excel, arcname="ficha_ajuste_rubrica.xlsx") 
                with open(nome_zip, "rb") as f:
                    st.download_button("📥 Baixar Pacote Completo (ZIP)", data=f, file_name=nome_zip, mime="application/zip")
            except Exception as e: st.error(f"Erro: {e}")

# ==================== OPÇÃO 2: ADITIVO ====================
elif "2. Aditivo" in tipo_documento:
    tipo_aditivo = st.radio("Qual será a alteração?", ["Alteração de Prazo", "Alteração de Valor"])
    
    if tipo_aditivo == "Alteração de Prazo":
        data_prorrog = st.text_input("Nova data de prorrogação (Ex: 31/12/2026)")
        just_prorrog = st.text_area("Justificativa da prorrogação")
        
        if st.button("Gerar Aditivo de Prazo"):
            contexto = {
                "titulo_projeto": str(titulo), "n_projeto": str(n_projeto), "tipo_contrato": str(tipo_contrato),
                "nome_coord": str(coord), "siape_coord": str(siape), "data_atual": str(datetime.now().strftime("%d/%m/%Y")),
                "data_prorrog": str(data_prorrog), "just_prorrog": str(just_prorrog),
                "nome_supfin": limpar_assinatura(nome_supfin, "_________________"), "siape_supfin": limpar_assinatura(siape_supfin, "_______"),
                "nome_diretor": limpar_assinatura(nome_diretor, "_________________"), "siape_diretor": limpar_assinatura(siape_diretor, "_______"),
                "fundacao": str(fundacao), "diretor_fund": limpar_assinatura(diretor_fund, "_________________")
            }
            caminho_word = encontrar_arquivo("aditivo_alteracao_prazo.docx")
            if not caminho_word: st.error("❌ Arquivo Word não encontrado.")
            else:
                try:
                    doc = DocxTemplate(caminho_word)
                    doc.render(contexto)
                    n_proj_limpo = str(n_projeto).replace('/', '-')
                    nome_word = f"Aditivo_Prazo_{n_proj_limpo}.docx"
                    doc.save(nome_word)
                    nome_zip = f"Pacote_Aditivo_Prazo_{n_proj_limpo}.zip"
                    with zipfile.ZipFile(nome_zip, 'w') as zipf:
                        zipf.write(nome_word)
                        caminho_anexo = encontrar_arquivo("aditivo_prazo_cronograma.docx")
                        if caminho_anexo: zipf.write(caminho_anexo, arcname="aditivo_prazo_cronograma.docx")
                    with open(nome_zip, "rb") as f:
                        st.download_button("📥 Baixar Aditivo de Prazo (ZIP)", data=f, file_name=nome_zip, mime="application/zip")
                except Exception as e: st.error(f"Erro: {e}")

    elif tipo_aditivo == "Alteração de Valor":
        col3, col4 = st.columns(2)
        with col3:
            aumento_reducao = st.selectbox("Tipo de alteração", ["aumento", "redução"])
            valor_colocado = st.text_input("Valor da alteração (R$)")
        with col4:
            valor_total = st.text_input("Valor TOTAL final do projeto (R$)")
        just_aumento = st.text_area("Justificativa da alteração de valor")
        
        if st.button("Gerar Aditivo de Valor"):
            contexto = {
                "titulo_projeto": str(titulo), "n_projeto": str(n_projeto), "tipo_contrato": str(tipo_contrato),
                "nome_coord": str(coord), "siape_coord": str(siape), "data_atual": str(datetime.now().strftime("%d/%m/%Y")),
                "aumento_reducao": str(aumento_reducao), "valor_colocado": str(valor_colocado),
                "valor_total": str(valor_total), "just_aumento_reducao": str(just_aumento),
                "nome_supfin": limpar_assinatura(nome_supfin, "_________________"), "siape_supfin": limpar_assinatura(siape_supfin, "_______"),
                "nome_diretor": limpar_assinatura(nome_diretor, "_________________"), "siape_diretor": limpar_assinatura(siape_diretor, "_______"),
                "fundacao": str(fundacao), "diretor_fund": limpar_assinatura(diretor_fund, "_________________")
            }
            caminho_word = encontrar_arquivo("aditivo_alteracao_valor.docx")
            if not caminho_word: st.error("❌ Arquivo Word não encontrado.")
            else:
                try:
                    doc = DocxTemplate(caminho_word)
                    doc.render(contexto)
                    n_proj_limpo = str(n_projeto).replace('/', '-')
                    nome_word = f"Aditivo_Valor_{n_proj_limpo}.docx"
                    doc.save(nome_word)
                    nome_zip = f"Pacote_Aditivo_Valor_{n_proj_limpo}.zip"
                    with zipfile.ZipFile(nome_zip, 'w') as zipf:
                        zipf.write(nome_word)
                        caminho_anexo = encontrar_arquivo("aditivo_valor_plano_aplicacao.docx")
                        if caminho_anexo: zipf.write(caminho_anexo, arcname="aditivo_valor_plano_aplicacao.docx")
                    with open(nome_zip, "rb") as f:
                        st.download_button("📥 Baixar Aditivo de Valor (ZIP)", data=f, file_name=nome_zip, mime="application/zip")
                except Exception as e: st.error(f"Erro: {e}")

# ==================== OPÇÃO 3: SUBSTITUIÇÃO DE EQUIPE ====================
elif "3. Ficha de ajuste" in tipo_documento:
    st.write("Serão gerados: **Solicitação de Substituição (Word) + Ficha de Gestão (Excel) preenchidos**.")
    cargo = st.selectbox("Quem será substituído?", ["Fiscal", "Coordenador"])
    data_subst = st.text_input("A partir de qual data? (Ex: 01/08/2026)")
    
    col5, col6 = st.columns(2)
    with col5:
        st.write("**Membro ANTERIOR (Quem sai):**")
        nome_ant = st.text_input(f"Nome do {cargo} Anterior", value=coord if cargo == "Coordenador" else nome_fiscal_at)
        siape_ant = st.text_input(f"SIAPE do {cargo} Anterior", value=siape if cargo == "Coordenador" else siape_fiscal_at)
    with col6:
        st.write("**Membro ATUAL (Quem entra):**")
        nome_at = st.text_input(f"Nome do {cargo} Atual")
        siape_at = st.text_input(f"SIAPE do {cargo} Atual")
        
    if st.button("Gerar Solicitação e Ficha de Equipe"):
        # Verificação preventiva extra antes de gerar arquivos
        if any(x in str(nome_saida_teste:=nome_at).lower() for x in ["estudante", "discente"]):
            st.error("❌ Operação abortada: Não é permitido gerar documentos de adequação de carga horária para discentes.")
        else:
            contexto = {
                "titulo_projeto": str(titulo), "n_projeto": str(n_projeto), "tipo_contrato": str(tipo_contrato),
                "vigencia": str(vigencia), "fundacao": str(fundacao), "cnpj_fundacao": str(cnpj_fundacao),
                "data_atual": str(datetime.now().strftime("%d/%m/%Y")), "cargo": str(cargo), "data_subst": str(data_subst),
                "nome_diretor": limpar_assinatura(nome_diretor, "_________________"), "siape_diretor": limpar_assinatura(siape_diretor, "_______"),
                "nome_coord_ant": str(nome_ant) if cargo == "Coordenador" else "", "siape_coord_ant": str(siape_ant) if cargo == "Coordenador" else "",
                "nome_coord_at": str(nome_at) if cargo == "Coordenador" else "", "siape_coord_at": str(siape_at) if cargo == "Coordenador" else "",
                "nome_fiscal_ant": str(nome_ant) if cargo == "Fiscal" else "", "siape_fiscal_ant": str(siape_ant) if cargo == "Fiscal" else "",
                "nome_fiscal_at": str(nome_at) if cargo == "Fiscal" else "", "siape_fiscal_at": str(siape_at) if cargo == "Fiscal" else "",
            }
            caminho_word = encontrar_arquivo("adendo_fiscal_coordenador.docx")
            if not caminho_word: st.error("❌ Arquivo 'adendo_fiscal_coordenador.docx' não encontrado.")
            else:
                try:
                    doc = DocxTemplate(caminho_word)
                    doc.render(contexto)
                    n_proj_limpo = str(n_projeto).replace('/', '-')
                    nome_word = f"Substituicao_Equipe_{n_proj_limpo}.docx"
                    doc.save(nome_word)
                    
                    nome_zip = f"Pacote_Substituicao_Equipe_{n_proj_limpo}.zip"
                    with zipfile.ZipFile(nome_zip, 'w') as zipf:
                        zipf.write(nome_word)
                        
                        caminho_excel = encontrar_arquivo("ficha_gestao_ajuste_item.xlsx")
                        if caminho_excel:
                            nome_excel_gerado = f"Ficha_Equipe_Preenchida_{n_proj_limpo}.xlsx"
                            if preencher_excel_ficha_gestao(caminho_excel, contexto, nome_excel_gerado):
                                zipf.write(nome_excel_gerado, arcname="ficha_gestao_ajuste_item.xlsx")
                                
                    with open(nome_zip, "rb") as f:
                        st.download_button("📥 Baixar Pacote Substituição (ZIP)", data=f, file_name=nome_zip, mime="application/zip")
                except Exception as e: st.error(f"Erro: {e}")

# ==================== OPÇÃO 4: FICHA DE GESTÃO (Ajuste de Item) ====================
elif "4. Ficha de Gestão" in tipo_documento:
    st.write("Serão gerados: **Solicitação de Ajuste de Item (Word) + Ficha de Gestão (Excel) preenchidos**.")
    virtude = st.text_area("Descreva o motivo da alteração (Essa frase entrará após 'em virtude de...'):")
    
    if st.button("Gerar Ficha de Gestão (Word + Excel)"):
        contexto = {
            "titulo_projeto": str(titulo), "n_projeto": str(n_projeto), "tipo_contrato": str(tipo_contrato),
            "vigencia": str(vigencia), "fundacao": str(fundacao), "cnpj_fundacao": str(cnpj_fundacao),
            "nome_coord": str(coord), "siape_coord": str(siape), "data_atual": str(datetime.now().strftime("%d/%m/%Y")),
            "virtude": str(virtude),
            "nome_supfin": limpar_assinatura(nome_supfin, "_________________"), "siape_supfin": limpar_assinatura(siape_supfin, "_______"),
            "nome_diretor": limpar_assinatura(nome_diretor, "_________________"), "siape_diretor": limpar_assinatura(siape_diretor, "_______"),
            "diretor_fund": limpar_assinatura(diretor_fund, "_________________")
        }
        caminho_word = encontrar_arquivo("ajuste_item_solicitacao.docx")
        if not caminho_word: st.error("❌ Arquivo 'ajuste_item_solicitacao.docx' não encontrado.")
        else:
            try:
                doc = DocxTemplate(caminho_word)
                doc.render(contexto)
                n_proj_limpo = str(n_projeto).replace('/', '-')
                nome_word = f"Ajuste_Item_{n_proj_limpo}.docx"
                doc.save(nome_word)
                
                nome_zip = f"Pacote_Ficha_Gestao_{n_proj_limpo}.zip"
                with zipfile.ZipFile(nome_zip, 'w') as zipf:
                    zipf.write(nome_word)
                    
                    caminho_excel = encontrar_arquivo("ficha_gestao_ajuste_item.xlsx")
                    if caminho_excel:
                        nome_excel_gerado = f"Ficha_Gestao_Preenchida_{n_proj_limpo}.xlsx"
                        if preencher_excel_ficha_gestao(caminho_excel, contexto, nome_excel_gerado):
                            zipf.write(nome_excel_gerado, arcname="ficha_gestao_ajuste_item.xlsx")
                            
                with open(nome_zip, "rb") as f:
                    st.download_button("📥 Baixar Pacote Ficha de Gestão (ZIP)", data=f, file_name=nome_zip, mime="application/zip")
            except Exception as e: st.error(f"Erro: {e}")
