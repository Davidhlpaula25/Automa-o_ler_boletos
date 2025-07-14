import fitz  # PyMuPDF
import pandas as pd
import os
import re

# --- FUNÇÃO PARA LIMPAR E CONVERTER VALORES MONETÁRIOS ---
def limpar_valor(texto_valor):
    if texto_valor:
        valor_limpo = texto_valor.strip().replace('R$', '').replace('.', '').replace(',', '.')
        try:
            return float(valor_limpo)
        except ValueError:
            return 0.0
    return 0.0

# --- FUNÇÃO INTELIGENTE PARA ENCONTRAR VALORES POR COORDENADAS ---
def encontrar_valor_ao_lado(pagina, rotulo):
    try:
        tolerancia_vertical = 10 
        areas_rotulo = pagina.search_for(rotulo, quads=True)
        if not areas_rotulo: return None
        primeira_area = fitz.Rect(areas_rotulo[0])
        zona_busca = fitz.Rect(
            primeira_area.x1,
            primeira_area.y0 - tolerancia_vertical,
            pagina.rect.width,
            primeira_area.y1 + tolerancia_vertical
        )
        palavras_encontradas = pagina.get_text("words", clip=zona_busca)
        if palavras_encontradas:
            valor_completo = " ".join([palavra[4] for palavra in palavras_encontradas])
            return valor_completo
        return None
    except Exception:
        return None

# --- FUNÇÃO PRINCIPAL DE EXTRAÇÃO ---
def extrair_dados_finais(caminho_pdf):
    try:
        doc = fitz.open(caminho_pdf)
        pagina = doc[0]

        # --- TENTATIVA 1: MÉTODO PRECISO POR COORDENADAS ---
        agrupador_str = encontrar_valor_ao_lado(pagina, "Número de seu telefone:")
        bruto_str = encontrar_valor_ao_lado(pagina, "VALOR BRUTO DA FATURA")
        retencao_str = encontrar_valor_ao_lado(pagina, "VALOR DA RETENCAO IMPOSTOS")
        if not retencao_str:
            retencao_str = encontrar_valor_ao_lado(pagina, "RETENCOES")
        pagar_str = encontrar_valor_ao_lado(pagina, "Valor a pagar")

        # --- TENTATIVA 2: PLANO B COM REGEX ---
        if not bruto_str or not pagar_str or not retencao_str or not agrupador_str:
            print(f"INFO: Usando Plano B (Busca por Texto) para {os.path.basename(caminho_pdf)}")
            texto_completo = pagina.get_text("text")
            
            def buscar_com_regex(padrao, texto):
                match = re.search(padrao, texto)
                return match.group(1) if match else None

            # CORREÇÃO FINAL: O padrão agora aceita letras (A-Z) no número do telefone
            if not agrupador_str: agrupador_str = buscar_com_regex(re.compile(r"Número de seu telefone:\s*([A-Z\d\s]+)"), texto_completo)
            if not bruto_str: bruto_str = buscar_com_regex(re.compile(r"VALOR BRUTO DA FATURA\s*([\d.,]+)"), texto_completo)
            if not pagar_str: pagar_str = buscar_com_regex(re.compile(r"Valor a pagar\s*([\d.,]+)"), texto_completo)
            if not retencao_str:
                retencao_str = buscar_com_regex(re.compile(r"VALOR DA RETENCAO IMPOSTOS\s*(-?[\d.,]+)"), texto_completo)
                if not retencao_str:
                    retencao_str = buscar_com_regex(re.compile(r"RETENCOES\s*(-?[\d.,]+)"), texto_completo)

        # --- MONTAGEM E AUTOCORREÇÃO DOS DADOS ---
        terminal = agrupador_str.replace(" ", "") if agrupador_str else "Não Encontrado"
        bruto_limpo = limpar_valor(bruto_str)
        retencao_limpa = abs(limpar_valor(retencao_str))
        pagar_limpo = limpar_valor(pagar_str)
        
        if bruto_limpo == 0 and pagar_limpo != 0 and retencao_limpa != 0:
            bruto_final = pagar_limpo + retencao_limpa
        else:
            bruto_final = bruto_limpo
            
        if retencao_limpa == 0 and bruto_final != 0 and pagar_limpo != 0:
            retencao_final = bruto_final - pagar_limpo
        else:
            retencao_final = retencao_limpa

        dados = {
            "TERMINAL": terminal,
            "VL_BRUTO": bruto_final,
            "VL_RETEN": abs(retencao_final),
            "VL_LIQUIDO": pagar_limpo,
        }
        
        doc.close()
        return dados

    except Exception as e:
        print(f"ERRO CRÍTICO ao processar o arquivo {caminho_pdf}: {e}")
        return None

# --- SCRIPT PRINCIPAL ---
if __name__ == "__main__":
    pasta_pdfs = 'boletos_pdf'
    lista_de_dados = []
    print("--- Iniciando automação (Versão Finalíssima) ---")

    if not os.path.isdir(pasta_pdfs):
        print(f"ERRO: A pasta '{pasta_pdfs}' não foi encontrada.")
    else:
        for nome_arquivo in os.listdir(pasta_pdfs):
            if nome_arquivo.lower().endswith('.pdf'):
                caminho_completo = os.path.join(pasta_pdfs, nome_arquivo)
                print(f"Processando arquivo: {nome_arquivo}...")
                dados_extraidos = extrair_dados_finais(caminho_completo)
                if dados_extraidos:
                    lista_de_dados.append(dados_extraidos)
        
        if not lista_de_dados:
            print("\nNenhum dado foi extraído.")
        else:
            df = pd.DataFrame(lista_de_dados)
            total_bruto = df['VL_BRUTO'].sum()
            total_reten = df['VL_RETEN'].sum()
            total_liquido = df['VL_LIQUIDO'].sum()
            linha_totais = pd.DataFrame([{'TERMINAL': 'TOTAL', 'VL_BRUTO': total_bruto, 'VL_RETEN': total_reten, 'VL_LIQUIDO': total_liquido}])
            df_final = pd.concat([df, linha_totais], ignore_index=True)

            nome_planilha_saida = 'Relatorio_Final.xlsx'
            with pd.ExcelWriter(nome_planilha_saida, engine='openpyxl') as writer:
                df_final.to_excel(writer, index=False, sheet_name='Faturas')
                workbook  = writer.book
                worksheet = writer.sheets['Faturas']
                formato_brl = 'R$ #,##0.00'
                for col_num, col_name in enumerate(df_final.columns):
                    letra_coluna = chr(ord('A') + col_num)
                    worksheet.column_dimensions[letra_coluna].width = 18
                    if 'VL_' in col_name:
                        for cell in worksheet[letra_coluna]:
                            if isinstance(cell.value, (int, float)):
                                cell.number_format = formato_brl
            
            print(f"\n✅ Automação Concluída com Sucesso! ---")
            print(f"Planilha '{nome_planilha_saida}' criada.")