import win32com.client
import pythoncom
import os
import time
import pandas as pd

class CorelEngine:
    def __init__(self, version="CorelDRAW.Application"): # Pega a versão instalada no PC
        self.app = None
        self.version = version
        
    def start(self):
        pythoncom.CoInitialize() # Necessário para rodar em Threads sem Crash
        
        versoes = [
            "CorelDRAW.Application.24", # Corel 2022
            "CorelDRAW.Application.23", # Corel 2021
            "CorelDRAW.Application.22", # Corel 2020
            "CorelDRAW.Application.21", # Corel 2019
            "CorelDRAW.Application.20", # Corel 2018
            "CorelDRAW.Application",
            "VGCore.Application" 
        ]
        
        for v in versoes:
            try:
                self.app = win32com.client.Dispatch(v)
                self.app.Visible = True
                print(f"Conectado ao CorelDRAW Versão: {v}")
                return True
            except:
                pass
                
        print("Nenhuma versão do CorelDRAW registrada no sistema respondeu ao COM.")
        return False

    def open_template(self, template_path):
        if not self.app:
            return None
        doc = self.app.OpenDocument(template_path)
        return doc
        
    def switch_page(self, doc, size_name):
        """
        No novo modelo, as páginas se chamam P, M, G, GG.
        O Python acessa a página correspondente do tamanho pedido na planilha.
        """
        try:
            for i in range(1, doc.Pages.Count + 1):
                if doc.Pages(i).Name.upper() == size_name.upper():
                    doc.Pages(i).Activate()
                    return True
            return False
        except Exception as e:
            print(f"Erro ao trocar página: {e}")
            return False

    def _log(self, msg):
        """Grava log em arquivo E no console — visível mesmo sem terminal aberto."""
        import os
        print(msg)
        try:
            log_path = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                    '..', 'clone_debug.log')
            with open(log_path, 'a', encoding='utf-8') as f:
                f.write(msg + '\n')
        except: pass

    def clone_page_native(self, doc, source_page):
        """MÉTODO 1: page.Duplicate() nativo do CorelDRAW (Corel 2018+)."""
        try:
            count_before = doc.Pages.Count
            source_page.Duplicate()
            time.sleep(0.5)
            if doc.Pages.Count > count_before:
                nova = doc.Pages(doc.Pages.Count)
                self._log(f"[M1-NATIVE] OK — nova pag shapes={nova.Shapes.Count}")
                return nova
            self._log("[M1-NATIVE] Duplicate() não criou nova página.")
            return None
        except Exception as e:
            self._log(f"[M1-NATIVE] FALHOU: {e}")
            return None

    def clone_page_fallback(self, doc, source_page):
        """MÉTODO 2: Seleciona tudo, copia, cria página, cola."""
        try:
            # Passo 1 — desbloquear camadas e ativar página fonte
            source_page.Activate()
            time.sleep(0.3)
            for layer in source_page.Layers:
                try:
                    layer.Editable = True
                    layer.Visible = True
                except: pass

            # Passo 2 — selecionar tudo por camada (SelectAll não existe no Corel 2022)
            doc.ClearSelection()
            shapes_encontrados = 0
            for layer in source_page.Layers:
                if not layer.IsSpecialLayer:
                    try:
                        layer.Shapes.All().AddToSelection()
                        shapes_encontrados += layer.Shapes.Count
                    except: pass
            self._log(f"[M2] Seleção manual por camada: {shapes_encontrados} shapes encontrados")

            # Passo 3 — checar seleção via Selection (mais confiável no Corel 2022)
            try:
                sel_count = doc.Selection.Count
            except:
                sel_count = shapes_encontrados  # usa o que contamos manualmente
            self._log(f"[M2] Selection.Count confirmado: {sel_count}")

            # Passo 4 — copiar ANTES de trocar de página
            try:
                self.app.ActiveSelection.Copy()
                time.sleep(0.8)  # tempo generoso para o clipboard estabilizar
                self._log("[M2] Copy() executado.")
            except Exception as e_copy:
                self._log(f"[M2] Copy() FALHOU: {e_copy}")
                return None

            # Passo 5 — criar nova página
            doc.AddPages(1)
            time.sleep(0.4)
            nova_pag = doc.Pages(doc.Pages.Count)
            nova_pag.SizeWidth  = source_page.SizeWidth
            nova_pag.SizeHeight = source_page.SizeHeight
            nova_pag.Activate()
            time.sleep(0.5)
            self._log(f"[M2] Nova página criada. Total: {doc.Pages.Count}")

            # Passo 6 — tentar colar com múltiplos métodos, aceitar o primeiro que não exception
            colou = False
            for nome, fn in [
                ("doc.ActiveLayer.Paste()", lambda: doc.ActiveLayer.Paste()),
                ("doc.Paste()",             lambda: doc.Paste()),
                ("self.app.Paste()",        lambda: self.app.Paste()),
            ]:
                try:
                    fn()
                    time.sleep(0.5)
                    # Conta shapes direto na página (mais simples e confiável)
                    n = nova_pag.Shapes.Count
                    self._log(f"[M2] {nome} → Shapes.Count={n}")
                    if n > 0:
                        colou = True
                        break
                except Exception as ep:
                    self._log(f"[M2] {nome} → ERRO: {ep}")

            if not colou:
                self._log("[M2] Nenhum paste funcionou ou retornou 0 shapes. Deletando página vazia.")
                try: nova_pag.Delete()
                except: pass
                return None

            doc.ClearSelection()
            return nova_pag

        except Exception as e:
            self._log(f"[M2] EXCEÇÃO GERAL: {e}")
            return None


    def clone_page_sendkeys(self, doc, source_page):
        """
        MÉTODO 3 (Último Recurso): Simula Ctrl+A, Ctrl+C, cria página, Ctrl+V via Win32.
        Contorna problemas de COM com clipboard usando keystrokes reais.
        """
        import win32api
        import win32con
        import win32gui

        def send_keys(key, ctrl=False):
            if ctrl:
                win32api.keybd_event(win32con.VK_CONTROL, 0, 0, 0)
                time.sleep(0.05)
            win32api.keybd_event(key, 0, 0, 0)
            time.sleep(0.05)
            win32api.keybd_event(key, 0, win32con.KEYEVENTF_KEYUP, 0)
            if ctrl:
                time.sleep(0.05)
                win32api.keybd_event(win32con.VK_CONTROL, 0, win32con.KEYEVENTF_KEYUP, 0)
            time.sleep(0.1)

        try:
            # Garante que o CorelDRAW está em foco
            hwnd = win32gui.FindWindow(None, None)
            for h in self._get_corel_windows():
                win32gui.SetForegroundWindow(h)
                break

            source_page.Activate()
            time.sleep(0.3)

            send_keys(0x41, ctrl=True)  # Ctrl+A — Selecionar Tudo
            time.sleep(0.3)
            send_keys(0x43, ctrl=True)  # Ctrl+C — Copiar
            time.sleep(0.5)

            doc.AddPages(1)
            nova_pag = doc.Pages(doc.Pages.Count)
            nova_pag.SizeWidth = source_page.SizeWidth
            nova_pag.SizeHeight = source_page.SizeHeight
            nova_pag.Activate()
            time.sleep(0.4)

            send_keys(0x56, ctrl=True)  # Ctrl+V — Colar
            time.sleep(0.6)

            total_shapes = nova_pag.Shapes.Count
            print(f"[CLONE-SENDKEYS] Shapes colados: {total_shapes}")
            if total_shapes > 0:
                doc.ClearSelection()
                return nova_pag
            try: nova_pag.Delete()
            except: pass
            return None
        except Exception as e:
            print(f"[CLONE-SENDKEYS] ERRO: {e}")
            return None

    def _get_corel_windows(self):
        """Helper para localizar janelas do CorelDRAW para foco."""
        import win32gui
        hwnds = []
        def callback(h, _):
            title = win32gui.GetWindowText(h)
            if "CorelDRAW" in title:
                hwnds.append(h)
        win32gui.EnumWindows(callback, None)
        return hwnds


    def _find_shape_recursive(self, shapes, name_tag):
        """
        Varre todos os objetos e entra em grupos e PowerClips
        Tentando encontrar o objeto com o prefixo @NOME_JOGADOR, @LOGO_ESCUDO etc.
        """
        name_tag_upper = name_tag.upper().strip()
        for i in range(1, shapes.Count + 1):
            try:
                shape = shapes(i)
                curr_name = str(shape.Name).upper().strip()
                
                if curr_name == name_tag_upper:
                    return shape
                
                # Se for um Grupo, vasculha dentro
                if shape.Type == 7: # cdrGroupShape
                    found = self._find_shape_recursive(shape.Shapes, name_tag)
                    if found: return found
                
                # Se for um PowerClip, vasculha o conteúdo
                try:
                    if shape.PowerClip:
                        found = self._find_shape_recursive(shape.PowerClip.Shapes, name_tag)
                        if found: return found
                except: pass
            except: continue
        return None

    def scan_page_tags(self, page):
        """Lista todos os objetos que começam com @ para ajudar o usuário."""
        tags_encontradas = []
        def recursive_scan(shapes):
            for i in range(1, shapes.Count + 1):
                try:
                    s = shapes(i)
                    name = str(s.Name)
                    if name.startswith("@"):
                        tags_encontradas.append(name)
                    if s.Type == 7:
                        recursive_scan(s.Shapes)
                except: pass
        
        recursive_scan(page.Shapes)
        return list(set(tags_encontradas)) # Remove duplicatas

    def _find_all_shapes_recursive(self, shapes, name_prefix):
        """
        Retorna uma LISTA de todos os objetos que começam com o nome fornecido.
        Ideal para achar várias partes de roupa (@MOLDE_CORTE_FRENTE, @MOLDE_CORTE_COSTAS)
        """
        found = []
        for i in range(1, shapes.Count + 1):
            shape = shapes(i)
            name = str(shape.Name).upper()
            if name.startswith(name_prefix.upper()):
                found.append(shape)
            
            if shape.Type == 7: # group
                found.extend(self._find_all_shapes_recursive(shape.Shapes, name_prefix))
        return found

    def inject_text(self, doc, name_tag, new_text, master_shape=None):
        """ Localiza o placeholder (ex: @NUMERO) e troca pelo real com esteróides. """
        shape = self._find_shape_recursive(doc.ActivePage.Shapes, name_tag)
        if shape and hasattr(shape, 'Text'):
            
            # --- FASE 9: INJEÇÃO DE VACINA DE ESTILO MASTER (SYNC) ---
            if master_shape and hasattr(master_shape, 'Text'):
                 try:
                     # Copia a fonte e o alinhamento
                     shape.Text.Story.Font = master_shape.Text.Story.Font
                     shape.Text.Story.Alignment = master_shape.Text.Story.Alignment
                     # Copia as cores brutas e os contornos (Outlines)
                     shape.Fill.CopyAssign(master_shape.Fill)
                     if master_shape.Outline.Type != 0: # Não tem outline
                         shape.Outline.CopyAssign(master_shape.Outline)
                 except Exception as e_sync:
                     print(f"Aviso de Sync Estilo: {e_sync}")
                     
            shape.Text.Story = str(new_text)
            return True
        return False
        
    def process_logo_logic(self, doc, data_value):
        """ Resolve a regra complexa de negócios: Tem escudo? É bordado? É texto? """
        shape_logo = self._find_shape_recursive(doc.ActivePage.Shapes, "@LOGO_ESCUDO")
        
        if not shape_logo:
             return # Template não tem lugar pra escudo, ignora
             
        # Condição C: Célula vazia na planilha = Ser bordado
        if pd.isna(data_value) or str(data_value).strip() == "":
             shape_logo.Visible = False
             return
             
        # Condição B: A célula é um texto (Ex: VAI CURINTHIA) em vez de caminho C:/...
        data_str = str(data_value).strip()
        if not data_str.lower().endswith(('.png', '.jpg', '.jpeg', '.cdr')):
             # É um texto. Deleta shape grafico e escreve por cima na mesma posicao
             x, y = shape_logo.PositionX, shape_logo.PositionY
             shape_logo.Delete()
             # Criar texto (pseudo-código do objeto texto)
             new_text = doc.ActiveLayer.CreateArtisticText(x, y, data_str)
             return
             
        # Condição A: É o caminho da imagem de um escudo C:/escudos/time.png
        if os.path.exists(data_str):
             # Guarda posição e dimensões da caixa placeholder
             ph_x = shape_logo.PositionX
             ph_y = shape_logo.PositionY
             ph_w = shape_logo.SizeWidth
             ph_h = shape_logo.SizeHeight
             ph_cx = shape_logo.BoundingBox.X + ph_w / 2.0
             ph_cy = shape_logo.BoundingBox.Y + ph_h / 2.0
             shape_logo.Delete()

             # Importa a arte real
             import_filter = doc.ActiveLayer.Import(data_str)
             if import_filter:
                 try: import_filter.Finish()
                 except: pass

             # Pega o shape recém-importado e ajusta tamanho/posição
             try:
                 novo_logo = doc.ActiveShape
                 if novo_logo:
                     # Escala proporcional para caber no placeholder (mantendo aspect ratio)
                     ratio = min(ph_w / novo_logo.SizeWidth, ph_h / novo_logo.SizeHeight)
                     novo_logo.SetSize(novo_logo.SizeWidth * ratio, novo_logo.SizeHeight * ratio)
                     # Centraliza exatamente onde estava o placeholder
                     novo_logo.SetPositionEx(5, ph_cx, ph_cy)  # 5 = pcCenter
                     novo_logo.Name = "@LOGO_ESCUDO_REAL"
             except Exception as e_align:
                 print(f"Aviso: Não foi possível alinhar o logo importado: {e_align}")
             
    def draw_factory_guides(self, doc, data_dict):
        """
        [Melhoria Nível Industrial 4.0 - Dublagem Sublimática]
        Desenha as miras () para alinhamento manual de tecido na calandra
        em TODAS as partes da peça (Frente, Costas, Mangas) simultaneamente!
        E "Carimba" o Crachá da Peça gigante Fora do corte principal.
        """
        try:
             # Acha TODOS os contornos físicos do tecido no arquivo (Ex: Frente e Costas separados)
             moldes = self._find_all_shapes_recursive(doc.ActivePage.Shapes, "@MOLDE_CORTE")
             if not moldes:
                  return # Se o designer esqueceu, não trava. Apenas ignora.
             
             maior_y = -9999
             x_cracha = 0
             
             for molde in moldes:
                 x, y = molde.PositionX, molde.PositionY
                 w = molde.SizeWidth
                 h = molde.SizeHeight
                 
                 # Rastrea o ponto mais alto do papel pra escrever o crachá sem furar a arte
                 if (y + h) > maior_y:
                     maior_y = y + h
                     x_cracha = x
                 
                 # Desenha os Crop Marks (Cruzes) para CADA PEDAÇO DE TECIDO!
                 doc.ActiveLayer.CreateLineSegment(x - 1.0, y + h + 0.5, x - 1.0, y + h + 1.5)
                 doc.ActiveLayer.CreateLineSegment(x - 0.5, y + h + 1.0, x - 1.5, y + h + 1.0)
                 
                 doc.ActiveLayer.CreateLineSegment(x + w + 1.0, y + h + 0.5, x + w + 1.0, y + h + 1.5)
                 doc.ActiveLayer.CreateLineSegment(x + w + 0.5, y + h + 1.0, x + w + 1.5, y + h + 1.0)
                 
             # Vamos escrever o crachá gigante de informações apenas DUMA vez no Topo de tudo
             texto_info = f"[ PEDIDO {data_dict.get('PEDIDO', '...')} | NOME: {data_dict.get('NOME', '...')} | PEÇA: {data_dict.get('PRODUTO', '...')} | TAM: {data_dict.get('TAMANHO', '...')} ]"
             cracha = doc.ActiveLayer.CreateArtisticText(x_cracha, maior_y + 2.0, texto_info)
             cracha.Text.FontProperties.Size = 48 # Letra gigantesca de 48pt pra ler de longe
             
        except Exception as e:
             print(f"Erro ao desenhar guias de chão de fabrica: {e}")
             
    def export_jpg(self, doc, output_filepath):
        """ Exporta o render final pra pasta """
        try:
             # Opções de renderização industrial
             export_filter = doc.Export(output_filepath, 774) # 774 = JPEG
             export_filter.Finish()
        except Exception as e:
             print(f"Erro ao exportar JPG: {e}")

    def auto_grade_artwork(self, doc, ref_page_name, target_page, item_data=None):
        """ Copia artes da página P para os moldes da página alvo com telemetria """
        try:
             ref_page = None
             for p in doc.Pages:
                 if p.Name == ref_page_name:
                     ref_page = p
                     break
             
             if not ref_page:
                  self._log(f"[GRAD] Erro: Página de referência '{ref_page_name}' não encontrada.")
                  return
              
             partes = ["FRENTE", "COSTAS", "MANGA_DIREITA", "MANGA_ESQUERDA"]
             count_sucesso = 0
             
             for parte in partes:
                  tag_arte = f"@ARTE_{parte}"
                  tag_molde = f"@MOLDE_CORTE_{parte}"
                  
                  # BUSCA MANUAL (Garante que está na página correta)
                  arte_ref = self._find_shape_in_page_manual(ref_page, tag_arte)
                  molde_ref = self._find_shape_in_page_manual(ref_page, tag_molde)
                  molde_target = self._find_shape_in_page_manual(target_page, tag_molde)
                  
                  if not arte_ref or not molde_ref or not molde_target:
                      continue
                  
                  self._log(f"[DEBUG] {parte} em {target_page.Name} | X={molde_target.CenterX:.2f} Y={molde_target.CenterY:.2f}")

                  # 1. Copia a arte
                  ref_page.Activate()
                  doc.ClearSelection()
                  arte_ref.CreateSelection()
                  self.app.ActiveSelection.Copy()
                  
                  # 2. Cola no destino
                  target_page.Activate()
                  molde_target.Layer.Activate()
                  doc.ActiveLayer.Paste()
                  time.sleep(0.3)
                  nova_arte = doc.ActiveShape
                  
                  # 3. Rotação e Escala (Geometria Industrial)
                  ref_orient = molde_ref.SizeWidth > molde_ref.SizeHeight
                  tar_orient = molde_target.SizeWidth > molde_target.SizeHeight
                  
                  # Cálculo lógico de escala: garante que a arte acompanhe o molde independente do giro
                  if ref_orient != tar_orient:
                       fator_w = molde_target.SizeHeight / molde_ref.SizeWidth
                       fator_h = molde_target.SizeWidth / molde_ref.SizeHeight
                  else:
                       fator_w = molde_target.SizeWidth / molde_ref.SizeWidth
                       fator_h = molde_target.SizeHeight / molde_ref.SizeHeight
                  
                  # Redimensiona a arte baseada na proporção do molde
                  nova_arte.SetSize(arte_ref.SizeWidth * fator_w, arte_ref.SizeHeight * fator_h)
                  
                  # Se o molde alvo estiver "deitado" e o original "em pé" (ou vice-versa), gira a arte
                  if ref_orient != tar_orient:
                      nova_arte.Rotate(90)
                  
                  # 4. Centralização Absoluta
                  nova_arte.SetPositionEx(5, molde_target.CenterX, molde_target.CenterY)
                  
                  # 5. PowerClip com Robustez Industrial
                  time.sleep(0.5) # Estabiliza geometria no Corel
                  try:
                      doc.ClearSelection()
                      nova_arte.Visible = True
                      
                      # Se o molde for um grupo, tentamos o PowerClip no primeiro objeto (contorno)
                      alvo_pc = molde_target
                      if molde_target.Type == 7: # cdrGroupShape
                           alvo_pc = molde_target.Shapes(1)
                      
                      nova_arte.PlaceInto(alvo_pc)
                      # Tenta centralizar internamente no PowerClip
                      try: alvo_pc.PowerClip.Center()
                      except: pass
                      
                      count_sucesso += 1
                  except Exception as e_pc:
                      # Fallback: Apenas joga pra trás se o PowerClip falhar (evita travar o processo)
                      nova_arte.OrderToBack()
                      count_sucesso += 1
                      self._log(f"[!] {parte} sem PowerClip: {e_pc}")

             if count_sucesso == 0:
                 self._log(f"[GRAD] Aviso: Nenhuma arte processada em {target_page.Name}.")

        except Exception as e:
             self._log(f"[GRAD] Erro crítico no Auto-Grading: {e}")

    def _find_shape_in_page_manual(self, page, name):
        """ Busca um objeto pelo nome varrendo APENAS a página fornecida """
        return self._find_shape_recursive(page.Shapes, name)
