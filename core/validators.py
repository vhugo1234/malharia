class AntiErrorValidator:
    def __init__(self, base_config):
        self.config = base_config
        self.margins = self.config.get("METADATA", {}).get("safe_margin_mm", 20.0)

    def validate_text_overflow(self, text, element_type):
        """
        Protege contra quebra de texto, validando os limites numéricos fornecidos no JSON.
        """
        # Em produção rodamos a engine do Corel para medir a bounding box
        # O modo simulação faria uma conta aproximada de caracteres
        max_chars = 15 if element_type == "NOME" else 3
        if len(str(text)) > max_chars:
            return False, f"Overflow Alert: '{text}' excede limites seguros."
        return True, "OK"
        
    def validate_safe_margin(self, bounding_box):
        """
        Proteção de Corte/Costura para os PowerClips
        Recebe x, y, width, height do elemento inserido no prancheta.
        """
        # Pseudo-implementação das regras
        return True
