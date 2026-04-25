from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
import os
from datetime import datetime

class ReportGenerator:
    def __init__(self, output_dir):
        self.output_dir = output_dir
        os.makedirs(self.output_dir, exist_ok=True)
        
    def generate_pdf(self, lote_name, data, finance_data):
        """ Gera o relatório de OP com checklists e métricas. """
        filename = f"OP_FABRIL_{lote_name}_{int(datetime.now().timestamp())}.pdf"
        path = os.path.join(self.output_dir, filename)
        
        doc = SimpleDocTemplate(path, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=18)
        styles = getSampleStyleSheet()
        elements = []
        
        title_style = ParagraphStyle(
            name="TitleStyle",
            parent=styles["Heading1"],
            fontSize=18,
            spaceAfter=20,
            textColor=colors.HexColor("#1A2B4C")
        )
        
        # Header
        elements.append(Paragraph(f"<b>O.P. FÁBRICA 5.0 - GUIA DE CORTE E COSTURA</b>", title_style))
        elements.append(Paragraph(f"<b>Lote Sublimado:</b> {lote_name} - {datetime.now().strftime('%d/%m/%Y %H:%M')}", styles["Normal"]))
        elements.append(Spacer(1, 10))
        
        # Financeiro Calculator
        qnt_pecas = len(data)
        
        # Math Estimate
        p_roll_cost = finance_data.get("papel_valor_rolo_total", 0.0)
        p_len = finance_data.get("papel_comprimento_metros", 100.0)
        custo_papel_metro = (p_roll_cost / p_len) if p_len > 0 else 0
        
        c_v = finance_data.get("tinta_cyan_valor", 0)
        m_v = finance_data.get("tinta_magenta_valor", 0)
        y_v = finance_data.get("tinta_yellow_valor", 0)
        k_v = finance_data.get("tinta_black_valor", 0)
        t_ml = finance_data.get("tinta_frasco_ml", 1000)
        
        preco_cmyk_ml = ((c_v + m_v + y_v + k_v) / 4) / t_ml if t_ml > 0 else 0
        c_ink_m2 = finance_data.get("consumo_tinta_m2_ml", 6.0)
        custo_tinta_m2 = preco_cmyk_ml * c_ink_m2
        
        custo_total = (custo_papel_metro + custo_tinta_m2) * (qnt_pecas) # Estimativa grotesca de 1 peca = 1 metro/m2 linear
        
        fin_txt = f"""
        <b>Auditoria Comercial e Suprimentos:</b><br/>
        Total de Peças do Lote: {qnt_pecas}<br/>
        Custo Linear de Papel/Peça Estimado: R$ {custo_papel_metro:.2f}<br/>
        Custo CMYK/m² Estimado: R$ {custo_tinta_m2:.2f}<br/>
        <font color="red" size=12><b>Previsão do Gasto Operacional Base (Tinta e Papel): R$ {custo_total:.2f}</b></font>
        """
        elements.append(Paragraph(fin_txt, styles["Normal"]))
        elements.append(Spacer(1, 20))
        
        # Tabela (Planilha Checklist)
        elements.append(Paragraph("<b>Checklist Fabril (Inspeção Física das Cavas, Golas e Montagem):</b>", styles["Heading3"]))
        elements.append(Spacer(1, 10))
        
        table_data = [["[  ] PRONTO", "TAMANHO", "JOGADOR / PEDIDO", "NÚMERO", "OBSERVAÇÃO DA COSTUREIRA"]]
        for i, item in enumerate(data):
            sz = str(item.get("TAMANHO", "P")).upper()
            no = str(item.get("NOME", "ATLETA"))
            nm = str(item.get("NUMERO", "SV"))
            table_data.append(["[    ]", sz, no, nm, "_________________________"])
            
        # Pinta listras zebra para facilitar leitura na calandra
        style_list = [
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#1A2B4C")),
            ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('BOTTOMPADDING', (0,0), (-1,0), 8),
            ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
        ]
        
        for row in range(1, len(table_data)):
            if row % 2 == 0: style_list.append(('BACKGROUND', (0, row), (-1, row), colors.HexColor("#EAECEE")))
            else: style_list.append(('BACKGROUND', (0, row), (-1, row), colors.white))
            
        t = Table(table_data, colWidths=[80, 70, 180, 60, 150])
        t.setStyle(TableStyle(style_list))
        elements.append(t)
        
        doc.build(elements)
        return path
