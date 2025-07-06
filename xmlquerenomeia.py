import pandas as pd
from lxml import etree
import os
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox

# ===================================================================
# 1. LÓGICA DE PROCESSAMENTO (GERA UM XML E SALVA COMO .XML E .RPI)
# ===================================================================

def gerar_arquivos_do_prestador(caminho_da_planilha):
    """
    Lê uma planilha, gera uma estrutura XML e a salva em dois arquivos:
    um .xml e uma cópia com a extensão .rpi.
    Retorna uma tupla (sucesso, mensagem).
    """
    try:
        # Leitura da planilha a partir do caminho fornecido
        # O .fillna("") é crucial aqui, pois garante que células vazias se tornem strings vazias.
        df = pd.read_excel(caminho_da_planilha, dtype=str).fillna("")
        if df.empty:
            return (False, "Erro: A planilha está vazia ou não pôde ser lida.")

        # Pega dados fixos da primeira linha
        linha0 = df.iloc[0]

        # --- Início da Geração do XML ---
        root = etree.Element("operadora", nsmap={"xsi": "http://www.w3.org/2001/XMLSchema-instance"})
        root.attrib["{http://www.w3.org/2001/XMLSchema-instance}noNamespaceSchemaLocation"] = "SolicitacaoInclusaoPrestador.xsd"
        etree.SubElement(root, "registroANS").text = linha0["registroANS"]
        etree.SubElement(root, "cnpjOperadora").text = linha0["cnpjOperadora"]
        solicitacao = etree.SubElement(root, "solicitacao")
        etree.SubElement(solicitacao, "nossoNumero").text = linha0["nossoNumero"]
        etree.SubElement(solicitacao, "isencaoOnus").text = linha0["isencaoOnus"]
        
        def formatar_data_xml(data_raw):
            """Formata uma data para o padrão dd/mm/aaaa, se for válida."""
            if not data_raw: return ""
            try: 
                return pd.to_datetime(data_raw).strftime("%d/%m/%Y")
            except (ValueError, TypeError): 
                return ""

        # Loop para incluir todos os prestadores no XML
        for _, linha in df.iterrows():
            inclusao = etree.SubElement(solicitacao, "inclusaoPrestador")
            
            # Estas tags já estavam corretas, pois não tinham condição 'if'
            etree.SubElement(inclusao, "classificacao").text = linha["classificacao"]
            etree.SubElement(inclusao, "cnpjCpf").text = linha["cnpjCpf"]
            etree.SubElement(inclusao, "cnes").text = linha["cnes"]
            etree.SubElement(inclusao, "uf").text = linha["uf"]
            etree.SubElement(inclusao, "codigoMunicipioIBGE").text = linha["codigoMunicipioIBGE"]
            etree.SubElement(inclusao, "razaoSocial").text = linha["razaoSocial"]
            etree.SubElement(inclusao, "relacaoOperadora").text = linha["relacaoOperadora"]

            # Este bloco condicional está correto, pois depende de uma lógica (se é 'C' ou não)
            if linha["relacaoOperadora"] == "C":
                etree.SubElement(inclusao, "tipoContratualizacao").text = linha["tipoContratualizacao"]
                etree.SubElement(inclusao, "registroANSOperadoraIntermediaria").text = linha["registroANSOperadoraIntermediaria"]

            etree.SubElement(inclusao, "dataContratualizacao").text = formatar_data_xml(linha["dataContratualizacao"])
            etree.SubElement(inclusao, "dataInicioPrestacaoServico").text = formatar_data_xml(linha["dataInicioPrestacaoServico"])
            etree.SubElement(inclusao, "disponibilidadeServico").text = linha["disponibilidadeServico"]
            etree.SubElement(inclusao, "urgenciaEmergencia").text = linha["urgenciaEmergencia"]

            # =========================================================================
            # MODIFICAÇÃO PRINCIPAL: Bloco 'vinculacao'
            # As condições 'if' foram removidas para garantir que as tags sejam sempre
            # criadas, mesmo que o conteúdo da célula na planilha seja vazio.
            # =========================================================================
            vinculacao = etree.SubElement(inclusao, "vinculacao")
            etree.SubElement(vinculacao, "numeroRegistroPlanoVinculacao").text = linha["numeroRegistroPlanoVinculacao"]
            etree.SubElement(vinculacao, "numeroRegistroPlanoVinculacao").text = linha["numeroRegistroPlanoVinculacao1"]
            etree.SubElement(vinculacao, "codigoPlanoOperadoraVinculacao").text = linha["codigoPlanoOperadoraVinculacao"]

        # --- Salvando os arquivos ---
        os.makedirs("saida", exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        base_arquivo_saida = f"saida/prestadores_lote_{timestamp}"
        
        # Define os nomes dos dois arquivos
        arquivo_xml = f"{base_arquivo_saida}.xml"
        arquivo_rpi = f"{base_arquivo_saida}.rpi"

        # Cria a árvore XML para ser salva
        tree = etree.ElementTree(root)
        
        # Salva o mesmo conteúdo XML em dois arquivos com extensões diferentes
        write_params = {"encoding": "utf-8", "pretty_print": True, "xml_declaration": True}
        tree.write(arquivo_xml, **write_params)
        tree.write(arquivo_rpi, **write_params)

        # Mensagem de sucesso com os dois arquivos
        mensagem_sucesso = (f"Arquivos gerados com sucesso:\n\n"
                            f"-> {os.path.abspath(arquivo_xml)}\n"
                            f"-> {os.path.abspath(arquivo_rpi)}")
        
        return (True, mensagem_sucesso)

    except KeyError as e:
        # Erro específico para coluna não encontrada
        return (False, f"Erro: A coluna '{e}' não foi encontrada na planilha.\n\nVerifique se o nome da coluna está exatamente correto no arquivo Excel.")
    except Exception as e:
        # Captura outros erros que possam ocorrer durante o processo
        return (False, f"Ocorreu um erro inesperado: {e}\n\nVerifique se a planilha está formatada corretamente.")


# ===================================================================
# 2. FUNÇÕES DA INTERFACE GRÁFICA (ATUALIZADAS)
# ===================================================================
def selecionar_arquivo():
    """Abre uma janela para o usuário selecionar um arquivo Excel."""
    caminho_arquivo = filedialog.askopenfilename(
        title="Selecione a planilha de prestadores",
        filetypes=[("Planilhas Excel", "*.xlsx *.xls"), ("Todos os arquivos", "*.*")]
    )
    if caminho_arquivo:
        caminho_planilha.set(caminho_arquivo)
        label_arquivo_selecionado.config(text=os.path.basename(caminho_arquivo))
        status_var.set("Planilha selecionada. Clique em 'Gerar Arquivos'.")

def iniciar_processamento():
    """Valida a seleção e inicia a geração dos arquivos."""
    arquivo_selecionado = caminho_planilha.get()
    
    if not arquivo_selecionado:
        messagebox.showwarning("Aviso", "Por favor, selecione uma planilha primeiro.")
        return

    status_var.set("Processando... Aguarde.")
    btn_gerar.config(state="disabled")
    root.update_idletasks()

    # Chama a função principal atualizada
    sucesso, mensagem = gerar_arquivos_do_prestador(arquivo_selecionado)
    
    btn_gerar.config(state="normal")
    status_var.set("Pronto.")

    if sucesso:
        messagebox.showinfo("Sucesso!", mensagem)
    else:
        messagebox.showerror("Erro!", mensagem)

# ===================================================================
# 3. CRIAÇÃO E CONFIGURAÇÃO DA JANELA PRINCIPAL
# ===================================================================
root = tk.Tk()
root.title("Gerador de Arquivos para Prestadores (XML e RPI)") # Título atualizado
root.geometry("450x250")

caminho_planilha = tk.StringVar()
status_var = tk.StringVar(value="Selecione uma planilha para começar.")

main_frame = tk.Frame(root, padx=20, pady=20)
main_frame.pack(expand=True, fill="both")

btn_selecionar = tk.Button(main_frame, text="1. Selecionar Planilha...", command=selecionar_arquivo, font=("Helvetica", 10))
btn_selecionar.pack(fill="x", pady=5)

label_arquivo_selecionado = tk.Label(main_frame, text="Nenhum arquivo selecionado", fg="blue", wraplength=400)
label_arquivo_selecionado.pack(pady=5)

# Botão atualizado
btn_gerar = tk.Button(main_frame, text="2. Gerar Arquivos (XML e RPI)", command=iniciar_processamento, font=("Helvetica", 10, "bold"))
btn_gerar.pack(fill="x", pady=(15, 5))

label_status = tk.Label(main_frame, textvariable=status_var, fg="gray")
label_status.pack(side="bottom", fill="x", pady=(10, 0))

# --- Rótulo de Créditos ---
label_creditos = tk.Label(main_frame, 
                          text="Feito por Tucano Soluções TI", 
                          font=("Helvetica", 8, "italic"), 
                          fg="darkgray")
label_creditos.pack(side="bottom", pady=(0, 5))
root.mainloop()