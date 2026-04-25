import threading
import time
import os
import shutil
import pythoncom
from core.corel_integration import CorelEngine

class RenderThread(threading.Thread):
    def __init__(self, ui_callback, data_list, template_path="", output_dir="", finance_data=None, usar_ia=False):
        super().__init__()
        self.ui_callback = ui_callback
        self.data_list = data_list
        self.template_path = template_path
        self.output_dir = output_dir
        self.finance_data = finance_data or {}
        self.usar_ia = usar_ia
        self.is_running = True
        
    def run(self):
        """
        Main execution loop for rendering.
        Novo Pipeline (Fase 6): Em vez de gerar JPG e desfazer,
        Nós multiplicamos as páginas, preenchemos, e salvamos um grande arquivo MASTER.
        """
        pythoncom.CoInitialize() 
        corel = CorelEngine()
        
        if not corel.start():
             self.ui_callback("[FALHA] CorelDRAW não está aberto ou acessível no Windows.", 0.0)
             pythoncom.CoUninitialize()
             return

        total = len(self.data_list)
        
        # Cria cópia física de segurança para não corromper o arquivo Base
        if not self.template_path or not os.path.exists(self.template_path):
            self.ui_callback("[ERRO] Template base não encontrado.", 0.0)
            pythoncom.CoUninitialize()
            return
            
        # O nome do arquivo temporário que vai virar o mestre.
        temp_file = os.path.join(self.output_dir, "_TEMP_MASTER.cdr")
        shutil.copy2(self.template_path, temp_file)
        
        # Limpa o log de debug de clonagem de cada execução
        log_path = os.path.join(self.output_dir, '..', 'clone_debug.log') if self.output_dir else 'clone_debug.log'
        try:
            open(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'clone_debug.log'), 'w').close()
        except: pass

        doc_master = corel.open_template(temp_file)
        
        # Guarda os nomes das páginas originais para apagá-las no final (Ex: 'P', 'M', 'G')
        paginas_matrizes = []
        for i in range(1, doc_master.Pages.Count + 1):
            paginas_matrizes.append(doc_master.Pages(i).Name.upper())
            
        # --- FASE 9: TIRA A "FOTO" DE ESTILOS DA PÁGINA MESTRA ---
        self.ui_callback(f"Extraindo DNA de Estilos da página referência...", 0.0)
        master_styles = {"@NUMERO_COSTAS": None, "@NOME_JOGADOR": None}
        
        # Prioridade total para a página "P" como mestre (conforme feedback do usuário)
        # Fallback para "M" ou para a primeira página disponível
        if "P" in paginas_matrizes:
            pag_rainha = "P"
        elif "M" in paginas_matrizes:
            pag_rainha = "M"
        else:
            pag_rainha = paginas_matrizes[0] if paginas_matrizes else None
            
        self.ui_callback(f"Usando página '{pag_rainha}' como referência de arte.", 0.0)
        
        try:
             corel.switch_page(doc_master, pag_rainha)
             
             # Scanner de Tags para ajudar o usuário
             tags_vistas = corel.scan_page_tags(doc_master.ActivePage)
             self.ui_callback(f"Scanner na página {pag_rainha}: Encontrei {len(tags_vistas)} etiquetas: {', '.join(tags_vistas)}", 0.0)
             
             # Busca Robusta: Tenta na mestre, se não der, varre TODAS as páginas
             for tag in ["@NUMERO_COSTAS", "@NOME_JOGADOR"]:
                 found = corel._find_shape_recursive(doc_master.ActivePage.Shapes, tag)
                 if not found:
                     self.ui_callback(f"🔍 {tag} não está na página {pag_rainha}. Varrendo o documento...", 0.0)
                     for p_idx in range(1, doc_master.Pages.Count + 1):
                         found = corel._find_shape_recursive(doc_master.Pages(p_idx).Shapes, tag)
                         if found:
                             self.ui_callback(f"✅ Achei {tag} na página '{doc_master.Pages(p_idx).Name}'", 0.0)
                             break
                 master_styles[tag] = found

             if not master_styles["@NUMERO_COSTAS"]:
                 self.ui_callback("⚠️ Aviso: @NUMERO_COSTAS não encontrado em NENHUMA página.", 0.0)
        except Exception as e_style:
             self.ui_callback(f"⚠️ Erro ao extrair estilos: {e_style}", 0.0)

            
        for index, item in enumerate(self.data_list):
            if not self.is_running:
                break
                
            nome_pequeno = item.get('NOME', f'Peca_{index}')
            tamanho = str(item.get('TAMANHO', 'P')).upper()
            
            nome_nova_pagina = f"{index+1:03d}_{nome_pequeno}_{tamanho}"
            self.ui_callback(f"Operando clonagem: {nome_nova_pagina}", (index) / total)
            
            try:
                # 0. Limpeza Neural de Brasão
                caminho_logo = item.get('LOGOMARCA', '')
                if self.usar_ia and caminho_logo and str(caminho_logo).endswith(('.jpg', '.png', '.jpeg')):
                    try:
                        from rembg import remove
                        from PIL import Image
                        self.ui_callback("IA Removendo fundo branco do Escudo...", (index) / total)
                        
                        img_orig = Image.open(caminho_logo)
                        img_sem_fundo = remove(img_orig)
                        # Salva copia na temp
                        caminho_limpo = os.path.join(self.output_dir, f"limpo_{index}.png")
                        img_sem_fundo.save(caminho_limpo)
                        caminho_logo = caminho_limpo # Muda o ponteiro pra imagem transparente
                    except Exception as e_ia:
                        print(f"Erro IA: {e_ia}")
                        
                # 1. Pula para a página matriz do tamanho correto
                if not corel.switch_page(doc_master, tamanho):
                    self.ui_callback(f"⚠️ Tamanho {tamanho} não tem molde no arquivo! Pulando.", 0.0)
                    continue
                    
                pagina_matriz_ativa = doc_master.ActivePage
                
                # Desbloqueia todas as camadas para garantir que tudo será copiado
                for layer in pagina_matriz_ativa.Layers:
                    try:
                        layer.Editable = True
                        layer.Visible = True
                    except: pass

                # 2. CLONAGEM — 3 métodos em cascata (do mais confiável ao mais agressivo)
                #    MÉTODO 1: Duplicate() nativo (Corel 2018+)
                nova_pag = corel.clone_page_native(doc_master, pagina_matriz_ativa)
                
                if nova_pag is None:
                    # MÉTODO 2: Copy/Paste COM com múltiplos pastes
                    self.ui_callback(f"   [Aviso] Método 1 falhou. Tentando Copy/Paste COM...", 0.0)
                    nova_pag = corel.clone_page_fallback(doc_master, pagina_matriz_ativa)

                if nova_pag is None:
                    # MÉTODO 3: Ctrl+A / Ctrl+C / Ctrl+V via Win32 (último recurso)
                    self.ui_callback(f"   [Aviso] Método 2 falhou. Tentando SendKeys Win32...", 0.0)
                    nova_pag = corel.clone_page_sendkeys(doc_master, pagina_matriz_ativa)

                if nova_pag is None:
                    self.ui_callback(f"⚠️ TODOS OS MÉTODOS DE CLONAGEM FALHARAM para {nome_nova_pagina}. Pulando.", 0.0)
                    continue

                nova_pag.Name = nome_nova_pagina
                nova_pag.Activate()
                
                # 3. Injeta todos os dados nessa aba final
                corel.inject_text(doc_master, "@NUMERO_COSTAS", item.get('NUMERO', '10'), master_shape=master_styles.get("@NUMERO_COSTAS"))
                corel.inject_text(doc_master, "@NOME_JOGADOR", item.get('NOME', 'ATLETA'), master_shape=master_styles.get("@NOME_JOGADOR"))
                corel.process_logo_logic(doc_master, caminho_logo)
                
                # FASE 10: AUTO-GRADING DE ARTE VETORIAL MATEMÁTICO (Clona artes da P)
                self.ui_callback(f"Graduando arte para {tamanho}...", 0.0)
                corel.auto_grade_artwork(doc_master, pag_rainha, nova_pag, item)

                
                # 4. Fabrica 4.0 (Crop Marks de Mira)
                corel.draw_factory_guides(doc_master, item)
                
            except Exception as e:
                self.ui_callback(f"⚠️ Erro ao clonar {nome_pequeno}: {e}", 0.0)
                 
        self.ui_callback("Limpando páginas matriz...", 0.95)

        # Deleta as páginas originais de TRÁS PRA FRENTE por índice
        # (evita deslocamento de índice ao deletar)
        try:
            paginas_para_deletar = []
            for i in range(1, doc_master.Pages.Count + 1):
                nome_pag = doc_master.Pages(i).Name.upper()
                if nome_pag in paginas_matrizes:
                    paginas_para_deletar.append(i)

            # Deleta do maior índice pro menor
            for idx in sorted(paginas_para_deletar, reverse=True):
                try:
                    doc_master.Pages(idx).Activate()
                    doc_master.ActivePage.Delete()
                except Exception as e_del:
                    self.ui_callback(f"   Aviso ao deletar pág {idx}: {e_del}", 0.96)
        except Exception as e_clean:
            self.ui_callback(f"   Aviso no cleanup: {e_clean}", 0.96)

        if doc_master:
            pedido_id = self.data_list[0].get("PEDIDO", str(int(time.time()))) if self.data_list else str(int(time.time()))
            nome_lote_final = f"O.P_{pedido_id}_Master.cdr"
            caminho_final = os.path.join(self.output_dir, nome_lote_final)

            self.ui_callback(f"Salvando arquivo em: {caminho_final}", 0.97)

            salvo = False
            # MÉTODO 1: SaveAs simples (sem options) — funciona no Corel 2022+
            try:
                doc_master.SaveAs(caminho_final)
                salvo = True
                self.ui_callback(f"   Salvo com SaveAs() simples.", 0.97)
            except Exception as e1:
                self.ui_callback(f"   SaveAs simples falhou ({e1}), tentando com options...", 0.97)

            # MÉTODO 2: SaveAs com CreateStructSaveAsOptions (versões antigas)
            if not salvo:
                try:
                    opt = corel.app.CreateStructSaveAsOptions()
                    doc_master.SaveAs(caminho_final, opt)
                    salvo = True
                    self.ui_callback(f"   Salvo com SaveAs+Options.", 0.97)
                except Exception as e2:
                    self.ui_callback(f"   SaveAs+Options falhou ({e2}), tentando Save()...", 0.97)

            # MÉTODO 3: Save() — salva no caminho atual (temp_file)
            if not salvo:
                try:
                    doc_master.Save()
                    # Move o temp pra pasta de saída
                    import shutil as _sh
                    _sh.copy2(os.path.join(self.output_dir, "_TEMP_MASTER.cdr"), caminho_final)
                    salvo = True
                    self.ui_callback(f"   Salvo via Save() + cópia.", 0.97)
                except Exception as e3:
                    self.ui_callback(f"⚠️ TODOS OS MÉTODOS DE SAVE FALHARAM: {e3}", 1.0)

            try:
                doc_master.Close()
            except: pass

            if salvo:
                # --- FASE 7: PDF de Ordem de Produção ---
                self.ui_callback(f"Imprimindo Ordem de Serviço da Fábrica P.D.F...", 0.98)
                try:
                    from core.report_generator import ReportGenerator
                    rg = ReportGenerator(self.output_dir)
                    rg.generate_pdf(f"PEDIDO_{pedido_id}", self.data_list, self.finance_data)
                except Exception as ex_pdf:
                    self.ui_callback(f"⚠️ Aviso: Não rolou gerar PDF: {ex_pdf}", 0.99)

                self.ui_callback(f"100% Finalizado! Salvo como {nome_lote_final}", 1.0)

        # Não fechamos o CorelDRAW — o operador pode ter outros arquivos abertos.
        # Apenas garantimos que não há referência pendurada.
        pythoncom.CoUninitialize()
        
        if self.is_running:
            self.ui_callback("✅ ARQUIVO MESTRE CDR GERADO E PRONTO PARA ENCAIXE MANUAL!", 1.0)
            
    def stop(self):
        self.is_running = False

class ExportThread(threading.Thread):
    def __init__(self, ui_callback, master_cdr_path, output_dir):
        super().__init__()
        self.ui_callback = ui_callback
        self.master_cdr_path = master_cdr_path
        self.output_dir = output_dir
        
    def run(self):
        """ Loop que apenas lê um CDR já encaixado manualmente e cospe JPEGs para Ploter """
        pythoncom.CoInitialize() 
        corel = CorelEngine()
        if not corel.start():
             self.ui_callback("[FALHA] CorelDRAW inacessível.", 0.0)
             pythoncom.CoUninitialize()
             return
             
        doc_master = corel.open_template(self.master_cdr_path)
        total_paginas = doc_master.Pages.Count
        
        self.ui_callback(f"Iniciando plotagem em lote de {total_paginas} pranchetas...", 0.1)
        
        for i in range(1, total_paginas + 1):
             aba = doc_master.Pages(i)
             aba.Activate()
             nome_aba = aba.Name
             
             self.ui_callback(f"Exportando JPEG: {nome_aba}...", i / total_paginas)
             
             # Salva o pedaço
             out_path = os.path.join(self.output_dir, f"{nome_aba}.jpg")
             corel.export_jpg(doc_master, out_path)
             
        doc_master.Close()
        corel.app.Quit()
        pythoncom.CoUninitialize()
        self.ui_callback("✅ TODAS AS IMAGENS ESTÃO NA PASTA DA PLOTTER!", 1.0)
