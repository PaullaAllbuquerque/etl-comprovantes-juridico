import os
import glob
import pandas as pd
from dotenv import load_dotenv
from PIL import Image
from google import genai
from google.genai import types

# Carrega as variáveis do arquivo .env (GEMINI_API_KEY)
load_dotenv()

api_key = os.getenv("GEMINI_API_KEY")
if not api_key:
    raise ValueError("A variável GEMINI_API_KEY não foi encontrada no arquivo .env!")

client = genai.Client(api_key=api_key)

# 1. Mapear todos os arquivos de imagem, incluindo os que estão
# dentro de todas as subpastas

from pathlib import Path

pasta_fotos = Path("dados/fotos")

extensoes = {".jpg", ".jpeg", ".png"}

arquivos = [
    arquivo
    for arquivo in pasta_fotos.rglob("*")
    if arquivo.is_file() and arquivo.suffix.lower() in extensoes
]

print(f"\nTotal de imagens encontradas: {len(arquivos)}")

print("\n--- Arquivos encontrados ---")
for arquivo in arquivos:
    print(f"ENCONTRADO: {arquivo}")
print("--- Fim da lista ---\n")

print("--- Teste de informações do primeiro arquivo ---")
primeiro = arquivos[0]

print(f"Nome do arquivo: {primeiro.name}")
print(f"Pasta de origem: {primeiro.parent.name}")
print(f"Caminho completo: {primeiro}")
print("--- Fim do teste ---\n")


prompt_extracao = """
Analise cuidadosamente esta imagem de um documento financeiro.

Classifique a imagem em APENAS uma destas categorias:
- "PIX": comprovante de transferência/pagamento via PIX.
- "RECIBO": recibo, nota/declaração de pagamento ou documento que comprove uma despesa, mas que não seja um comprovante de PIX.
- "DESCONHECIDO": quando não for possível identificar com segurança como PIX ou recibo.

Extraia os dados visíveis na imagem.

Regras importantes:
1. Não invente informações.
2. Se uma informação não estiver legível ou não existir, use "NÃO ENCONTRADO" para textos ou null para valor.
3. O valor deve ser um número decimal usando ponto, por exemplo: 134.96.
4. A data deve estar no formato DD/MM/AAAA.
5. Em "pagador_ou_emissor", informe:
   - para PIX: preferencialmente o nome do pagador/remetente ou a identificação disponível na transação;
   - para recibo: o nome da pessoa, empresa ou estabelecimento que emitiu o recibo.
6. Em "descricao", faça um resumo curto do que o documento representa.
7. Não confunda data de emissão, data de vencimento ou outras datas com a data efetiva do pagamento, quando for possível diferenciá-las.
8. Se houver dúvida entre PIX e RECIBO, prefira "DESCONHECIDO" em vez de inventar uma classificação.

Retorne APENAS um JSON válido, exatamente neste formato:

{
  "tipo": "PIX",
  "data": "DD/MM/AAAA",
  "valor": 150.00,
  "pagador_ou_emissor": "Nome identificado",
  "descricao": "Breve descrição do documento"
}
"""

dados_extraidos = []

print("\n--- Iniciando leitura dos comprovantes via IA ---")
for idx, arq in enumerate(arquivos, 1):
    try:
        img = Image.open(arq)
        response = client.models.generate_content(
            model='gemini-2.5-flash',
            contents=[img, prompt_extracao],
            config=types.GenerateContentConfig(
                response_mime_type="application/json"
            )
        )
        res_text = response.text.strip()
        import json
        info = json.loads(res_text)
        info["arquivo_origem"] = arq.name
        info["pasta_origem"] = arq.parent.name
        info["caminho_completo"] = str(arq)
        dados_extraidos.append(info)
        print(f"[{idx}/{len(arquivos)}] Processado: {arq.name} -> Tipo: {info.get('tipo')}, Data: {info.get('data')}, Valor: R$ {info.get('valor')}")
    except Exception as e:
        print(f"[{idx}/{len(arquivos)}] Erro ao processar {arq.name}: {e}")

# 2. Separar Pix e Recibos
lista_pix = [d for d in dados_extraidos if d.get("tipo") == "PIX"]
lista_recibos = [d for d in dados_extraidos if d.get("tipo") == "RECIBO"]

print(f"\nResumo da Extração: {len(lista_pix)} Pix encontrados e {len(lista_recibos)} Recibos encontrados.")

# 3. Executar o cruzamento (Pix é a chave primária)
correlacionados = []
recibos_usados = set()

for pix in lista_pix:
    p_valor = pix.get("valor")
    p_data = pix.get("data")
    
    recibo_match = None
    
    # Busca por recibo com mesmo valor e mesma data
    for idx_r, rec in enumerate(lista_recibos):
        if idx_r in recibos_usados:
            continue
        
        # Correlação por valor e data iguais
        if rec.get("valor") == p_valor and rec.get("data") == p_data and p_valor is not None:
            recibo_match = rec
            recibos_usados.add(idx_r)
            break
            
    # Caso não ache com data exata, tenta por valor igual
    if not recibo_match and p_valor is not None:
        for idx_r, rec in enumerate(lista_recibos):
            if idx_r in recibos_usados:
                continue
            if rec.get("valor") == p_valor:
                recibo_match = rec
                recibos_usados.add(idx_r)
                break

    correlacionados.append({
        "Pix - Data": p_data,
        "Pix - Valor (R$)": p_valor,
        "Pix - Pagador/Detalhes": pix.get("pagador_ou_emissor"),
        "Pix - Arquivo": pix.get("arquivo_origem"),
        "Recibo Correlacionado": recibo_match.get("arquivo_origem") if recibo_match else "Nenhum Recibo Correspondente",
        "Recibo - Data": recibo_match.get("data") if recibo_match else "-",
        "Recibo - Valor (R$)": recibo_match.get("valor") if recibo_match else "-",
        "Recibo - Emissor/Detalhes": recibo_match.get("pagador_ou_emissor") if recibo_match else "-",
        "Status Correlação": "Correlacionado" if recibo_match else "Pix Sem Recibo"
    })

# 4. Gerar Tabela e exportar para CSV/Excel
df_resultado = pd.DataFrame(correlacionados)
caminho_csv = "output/planilha_pix_recibos_correlacionados.csv"
caminho_excel = "output/planilha_pix_recibos_correlacionados.xlsx"

df_resultado.to_csv(caminho_csv, index=False, encoding="utf-8-sig")
try:
    df_resultado.to_excel(caminho_excel, index=False)
    print(f"\n--- SUCESSO! ---")
    print(f"Planilhas geradas com sucesso nas pastas:")
    print(f" - CSV: {caminho_csv}")
    print(f" - Excel: {caminho_excel}")
except Exception as e:
    print(f"CSV gerado em {caminho_csv}. (Aviso Excel: {e})")