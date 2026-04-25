"""
TEST_CLONE_PAGE.PY — Diagnóstico de Clonagem de Páginas no CorelDRAW
=====================================================================
Execute com o arquivo .cdr BASE aberto no CorelDRAW.
Vai testar os dois métodos e dizer qual funciona no seu Corel.

Como rodar:
    .\\venv\\Scripts\\python.exe test_clone_page.py
"""
import win32com.client
import time

print("=" * 60)
print("  DIAGNÓSTICO DE CLONAGEM — MALHARIA ESPORTIVA")
print("=" * 60)

try:
    app = win32com.client.Dispatch('CorelDRAW.Application')
    doc = app.ActiveDocument
    if not doc:
        print("\n[ERRO] Nenhum documento aberto no CorelDRAW!")
        print("  -> Abra o arquivo .cdr BASE e rode novamente.")
        exit(1)
except Exception as e:
    print(f"\n[ERRO] Não conseguiu conectar ao CorelDRAW: {e}")
    exit(1)

print(f"\n[OK] Conectado ao CorelDRAW versão: {app.Version}")
print(f"[OK] Documento aberto: {doc.Name}")
print(f"[OK] Páginas encontradas: {doc.Pages.Count}")

for i in range(1, doc.Pages.Count + 1):
    print(f"     Aba {i}: '{doc.Pages(i).Name}'")

# ────────────────────────────────────────────
# TESTE 1: Duplicate() nativo
# ────────────────────────────────────────────
print("\n" + "-" * 40)
print("TESTE 1: page.Duplicate() (Método Nativo)")
print("-" * 40)

try:
    pagina_1 = doc.Pages(1)
    count_antes = doc.Pages.Count
    pagina_1.Duplicate()
    time.sleep(0.5)
    
    if doc.Pages.Count > count_antes:
        nova = doc.Pages.Last
        shapes_na_nova = nova.Shapes.Count
        print(f"[SUCESSO] Duplicate() funcionou!")
        print(f"          Nova página: '{nova.Name}' com {shapes_na_nova} objetos")
        
        if shapes_na_nova == 0:
            print("[AVISO] A nova página está VAZIA — Duplicate() copiou a estrutura mas não os objetos!")
            print("        -> Será necessário usar o FALLBACK.")
        else:
            print(f"[CONFIRMADO] Arte replicada corretamente. Use este método.")
        
        # Limpa o teste
        nova.Delete()
        print("          (Página de teste removida)")
    else:
        print("[FALHA] Duplicate() não criou nova página.")
except Exception as e:
    print(f"[FALHA] Duplicate() gerou exceção: {e}")

# ────────────────────────────────────────────
# TESTE 2: Copy/Paste com doc.Paste()
# ────────────────────────────────────────────
print("\n" + "-" * 40)
print("TESTE 2: Copy/Paste com doc.Paste()")
print("-" * 40)

try:
    pagina_1 = doc.Pages(1)
    pagina_1.Activate()
    doc.ClearSelection()
    
    total_sel = 0
    for layer in pagina_1.Layers:
        if not layer.IsSpecialLayer:
            try:
                layer.Editable = True
                layer.Visible = True
                layer.Shapes.All().AddToSelection()
                total_sel += layer.Shapes.Count
            except: pass
    
    print(f"          Shapes selecionados na página fonte: {total_sel}")
    
    if total_sel > 0:
        app.ActiveSelection.Copy()
        time.sleep(0.5)
        
        doc.AddPages(1)
        nova_pag = doc.Pages.Last
        nova_pag.Activate()
        nova_pag.SizeWidth = pagina_1.SizeWidth
        nova_pag.SizeHeight = pagina_1.SizeHeight
        time.sleep(0.3)
        
        doc.Paste()
        time.sleep(0.3)
        
        shapes_colados = nova_pag.Shapes.Count
        print(f"[RESULTADO] Shapes colados na nova página: {shapes_colados}")
        
        if shapes_colados > 0:
            print("[SUCESSO] Copy/Paste funcionou! Método FALLBACK é válido.")
        else:
            print("[FALHA] Paste() colou vazio — clipboard foi perdido ao trocar de página.")
        
        doc.ClearSelection()
        nova_pag.Delete()
        print("          (Página de teste removida)")
    else:
        print("[AVISO] Nenhum shape selecionado — verifique se a página tem arte.")
except Exception as e:
    print(f"[FALHA] Copy/Paste gerou exceção: {e}")

# ────────────────────────────────────────────
# RESULTADO FINAL
# ────────────────────────────────────────────
print("\n" + "=" * 60)
print("  DIAGNÓSTICO CONCLUÍDO")
print("  Verifique os resultados acima e informe qual SUCESSO apareceu.")
print("=" * 60)
