from fastapi import FastAPI, File, UploadFile, HTTPException
from pydantic import BaseModel, field_validator
from typing import List
import pdfplumber
import re
import io
import uvicorn
from fastapi.responses import StreamingResponse
import openpyxl
from openpyxl.utils import coordinate_to_tuple
from openpyxl.styles import Alignment

app = FastAPI(title="Raichu Pro - Back-end Seguro (Pilar 2 - Equipe Completa)")

# ==========================================
# MODELOS DE SEGURANÇA (PYDANTIC)
# ==========================================

# Atualizamos o modelo para receber TODOS os dados que sua regex extrai
class ParticipanteSeguro(BaseModel):
    Nome: str
    SIAPE: str
    Vínculo: str = "Outro"
    Lotação: str = ""
    Função: str = "Participante"
    Bolsa: str = "Não"
    CH_D: str = "0"
    CH_F: str = "0"
    Início: str = ""
    Término: str = ""
    Chefia_Imediata: str = ""
    SIAPE_Chefia: str = ""

    # Limpa possíveis ataques (XSS) em qualquer campo de texto
    @field_validator('Nome', 'Vínculo', 'Lotação', 'Função')
    def limpar_texto(cls, valor):
        texto_limpo = re.sub(r'<[^>]*>', '', str(valor))
        return texto_limpo.strip()

class DadosProjetoSeguros(BaseModel):
    titulo: str
    coordenador: str
    classificacao: str
    equipe: List[ParticipanteSeguro] = [] 

    @field_validator('titulo', 'coordenador', 'classificacao')
    def limpar_texto(cls, valor):
        texto_limpo = re.sub(r'<[^>]*>', '', valor)
        return texto_limpo.strip()

# ==========================================
# FUNÇÕES DE EXTRAÇÃO DO SEU CÓDIGO ORIGINAL
# ==========================================

def extrair_bloco(inicio, fins, texto):
    match_inicio = re.search(inicio, texto, re.IGNORECASE)
    if not match_inicio:
        return ""
    pos_ini = match_inicio.start()
    pos_fim = len(texto)
    for f in fins:
        match_fim = re.search(f, texto[pos_ini:], re.IGNORECASE)
        if match_fim:
            pos_temp = pos_ini + match_fim.start()
            if pos_temp < pos_fim:
                pos_fim = pos_temp
    return texto[pos_ini:pos_fim]

def extrair_dados_pdf(arquivo_bytes: bytes) -> dict:
    dados = {
        "titulo": "Não encontrado",
        "coordenador": "Não encontrado",
        "classificacao": "Não encontrado",
        "equipe": []
    }
    
    with pdfplumber.open(io.BytesIO(arquivo_bytes)) as pdf:
        texto_completo = ""
        for pagina in pdf.pages:
            texto = pagina.extract_text()
            if texto:
                texto_completo += texto + "\n"
                
    # --- SUAS REGEX BÁSICAS ---
    # --- SUAS REGEX BÁSICAS (Extraídas do codautom.py) ---
    def extrair(regex, group=1):
        m = re.search(regex, texto_completo, re.IGNORECASE)
        return m.group(group).strip() if m else ""

    # Usando exatamente a sua lógica de extração
    dados["titulo"] = extrair(r'Título:\s*(.*?)\n')
    dados["classificacao"] = extrair(r'Classificação:\s*(.*?)\n')
    
    # A sua regex brilhante para pegar nome e SIAPE do responsável
    m_coord = re.search(r'Responsável pelo projeto:\s*(.*?)\s*\(\s*(\d+)\s*\)', texto_completo, re.IGNORECASE)
    if m_coord:
        dados["coordenador"] = m_coord.group(1).strip()
        # O SIAPE do coordenador ficaria no group(2) se precisássemos
    match_titulo = re.search(r'Título do projeto:\s*(.*?)\n', texto_completo, re.IGNORECASE)
    if match_titulo: dados["titulo"] = match_titulo.group(1).strip()
        
    match_coord = re.search(r'Nome do Coordenador:\s*(.*?)\n', texto_completo, re.IGNORECASE)
    if match_coord: dados["coordenador"] = match_coord.group(1).strip()
        
    match_classificacao = re.search(r'Classificação:\s*(.*?)\n', texto_completo, re.IGNORECASE)
    if match_classificacao: dados["classificacao"] = match_classificacao.group(1).strip()
        
    # --- A SUA LÓGICA BRILHANTE DA EQUIPE ---
    bloco_participantes = extrair_bloco(r'PARTICIPANTES', [r'UNIDADES VINCULADAS\s*\n', r'CLASSIFICAÇÕES', r'REGIÕES DE ATUAÇÃO'], texto_completo)
    
    if not bloco_participantes:
        bloco_participantes = texto_completo # Fallback se não achar o bloco

    matches_participantes = list(re.finditer(r'(\d{5,15})\s*-\s*([A-ZÀ-Ÿ\s\']+?)\s*(?=[A-ZÀ-Ÿ][a-zà-ÿ]|UNIDADES VINCULADAS|CLASSIFICAÇÕES|$)', bloco_participantes))
    
    for i, match in enumerate(matches_participantes):
        siape = match.group(1).strip()
        nome_bruto = match.group(2).strip()

        nome_limpo = re.sub(r'\s+', ' ', nome_bruto).strip()
        corte_idx = len(nome_limpo)

        palavras_corte = [
            "VÍNCULO", "VINCULO", "CURSO", "LOTAÇÃO", "LOTACAO",
            "FUNÇÃO", "FUNCAO", "UNIDADE", "CLASSIFICA", "PARTICIPANTE",
            "CH DENTRO", "CH FORA", "INÍCIO", "INICIO", "TÉRMINO", "TERMINO",
            "DEPARTAMENTO", "OBSERVA", "VALOR", "TIPO", "$$$"
        ]

        for p in palavras_corte:
            idx = nome_limpo.upper().find(p)
            if idx != -1 and idx < corte_idx:
                corte_idx = idx

        nome = nome_limpo[:corte_idx].strip(" -/")
        if not nome or len(nome) < 2:
            continue

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

            partes_nome_perdidas = []
            for w in text_after.split():
                if w.lower() in ["administrativo", "em", "educação", "educacao", "técnico", "tecnico", "responsável", "responsavel"]:
                    break
                if w.isupper() and len(w) > 1 and w not in ["SIM", "NÃO", "NAO"]:
                    partes_nome_perdidas.append(w)

            if partes_nome_perdidas:
                nome_final_teste = nome + " " + " ".join(partes_nome_perdidas)
                corte_idx_2 = len(nome_final_teste)
                for p in palavras_corte:
                    idx = nome_final_teste.upper().find(p)
                    if idx != -1 and idx < corte_idx_2:
                        corte_idx_2 = idx
                nome = nome_final_teste[:corte_idx_2].strip(" -/")

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

            palavras_extras_lotacao = []
            for w in text_after.split():
                if w.isupper() and w not in partes_nome_perdidas and w not in ["SIM", "NÃO", "NAO", "FISCAL", "COORDENADOR", "PESQUISADOR", "TÉCNICO", "TECNICO", "BOLSISTA"]:
                    palavras_extras_lotacao.append(w)
            if palavras_extras_lotacao:
                lotacao = lotacao + " " + " ".join(palavras_extras_lotacao)

            lotacao_limpa = re.sub(r'\s+', ' ', lotacao).strip()
            corte_lot = len(lotacao_limpa)
            palavras_corte_lotacao = ["UNIDADE", "CLASSIFICA", "PARTICIPANTE", "FUNÇÃO", "FUNCAO", "VALOR", "INÍCIO", "INICIO", "TÉRMINO", "TERMINO", "OBSERVA", "TIPO", "$$$"]
            
            for p in palavras_corte_lotacao:
                idx = lotacao_limpa.upper().find(p)
                if idx != -1 and idx < corte_lot:
                    corte_lot = idx
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
                    remainder = janela_limpa[match_vinculo.end():]
                    lotacao = re.split(r'(?i)(Coordenador|Participante|Pesquisador|Estagiário|Colaborador|Membro|Fiscal|Responsável|Técnico|Bolsista)', remainder)[0].strip()
                    break

        # A GRANDE MUDANÇA: Em vez de adicionar em "equipe_raw", adicionamos na lista de equipe do Pydantic
        dados["equipe"].append({
            "Nome": nome, "SIAPE": siape, "Vínculo": vinculo.title(), "Lotação": lotacao,
            "Função": funcao, "Bolsa": bolsa, "CH_D": ch_d, "CH_F": ch_f, "Início": data_ini, "Término": data_fim,
            "Chefia_Imediata": "", "SIAPE_Chefia": ""
        })
        
    return dados

# ==========================================
# ROTAS DA API
# ==========================================

PDF_MAGIC_NUMBER = b"%PDF-"

@app.post("/extrair-dados/", response_model=DadosProjetoSeguros)
async def processar_pdf(arquivo: UploadFile = File(...)):
    
    if not arquivo.filename.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="Erro: O arquivo deve ter a extensão .pdf")

    arquivo_bytes = await arquivo.read()
    
    if arquivo_bytes[:5] != PDF_MAGIC_NUMBER:
        raise HTTPException(status_code=400, detail="Acesso Negado: O arquivo não é um PDF válido.")
    
    try:
        dados_puros = extrair_dados_pdf(arquivo_bytes)
        dados_seguros = DadosProjetoSeguros(**dados_puros)
        return dados_seguros
        
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Erro ao processar o PDF: {str(e)}")

if __name__ == "__main__":
    uvicorn.run(app, host="127.0.0.1", port=8000)
    # ==========================================
# PILAR 3: GERAÇÃO DE DOCUMENTOS (EXCEL)
# ==========================================

@app.post("/gerar-excel/")
async def gerar_plano_trabalho(dados: DadosProjetoSeguros):
    try:
        # ATENÇÃO: Coloque aqui o caminho exato de um modelo Excel seu para o teste
        caminho_template = "Modelos/ACT/Seu_Modelo_ACT.xlsx" 
        
        wb = openpyxl.load_workbook(caminho_template)
        ws = wb["Plano de Trabalho"] if "Plano de Trabalho" in wb.sheetnames else wb.worksheets[0]

        # --- O SEU MOTOR DE AUTOAJUSTE DE CÉLULAS ---
        def escrever_excel(celula, valor):
            val_str = str(valor).strip() if valor is not None else ""
            if val_str in ["", "-", "None", "Não se aplica"]: val_str = None
            try:
                r_row, r_col = coordinate_to_tuple(celula)
                
                # Cálculo de altura de linha
                if val_str and len(val_str) > 0:
                    qtd_quebras = val_str.count('\n')
                    linhas_estimadas = (len(val_str) / 110.0) + qtd_quebras
                    if linhas_estimadas < 1: linhas_estimadas = 1
                    altura_calculada = (linhas_estimadas * 15) + 10
                    
                    altura_atual = ws.row_dimensions[r_row].height
                    if altura_atual is None or altura_calculada > altura_atual:
                        ws.row_dimensions[r_row].height = altura_calculada

                # Tratamento de mesclagem
                for merged_range in list(ws.merged_cells.ranges):
                    min_col, min_row, max_col, max_row = merged_range.bounds
                    if min_col <= r_col <= max_col and min_row <= r_row <= max_row:
                        intervalo = str(merged_range)
                        ws.unmerge_cells(intervalo)
                        cel_alvo = ws.cell(row=min_row, column=min_col)
                        cel_alvo.value = val_str
                        cel_alvo.alignment = Alignment(wrap_text=True, vertical='top')
                        ws.merge_cells(intervalo)
                        return
                        
                cel_alvo = ws.cell(row=r_row, column=r_col)
                cel_alvo.value = val_str
                cel_alvo.alignment = Alignment(wrap_text=True, vertical='top')
            except Exception as err:
                print(f"Aviso na célula {celula}: {str(err)}")

        # --- PREENCHENDO COM OS DADOS DA API ---
        # Aqui nós pegamos o JSON validado (dados) e enviamos para a planilha
        escrever_excel("C17", dados.titulo)
        escrever_excel("C20", dados.coordenador)
        escrever_excel("C27", dados.classificacao)
        
        # (Opcional) Podemos fazer um loop para listar os nomes da equipe se houver um campo para isso
        # escrever_excel("A50", dados.equipe[0].Nome se houver membros)

        # --- SALVANDO EM MEMÓRIA E ENVIANDO O DOWNLOAD ---
        excel_buffer = io.BytesIO()
        wb.save(excel_buffer)
        excel_buffer.seek(0) # Volta o ponteiro da memória para o começo
        
        return StreamingResponse(
            excel_buffer, 
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": "attachment; filename=Plano_de_Trabalho_Raichu.xlsx"}
        )
        
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Erro crítico ao gerar o Excel: {str(e)}")
