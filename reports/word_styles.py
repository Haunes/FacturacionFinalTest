from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

class WordStyleManager:
    """Maneja los estilos para documentos Word."""
    
    # Colores del diseño
    COLOR_PRIMARY = RGBColor(0, 51, 102)   # #003366
    COLOR_ACCENT = RGBColor(226, 0, 116)   # #E20074
    COLOR_LIGHT_GRAY = RGBColor(240, 240, 240)
    COLOR_BLACK = RGBColor(0, 0, 0)
    COLOR_WHITE = RGBColor(255, 255, 255)
    
    FONT_FAMILY = 'Calibri Light'
    
    def setup_document_styles(self, doc: Document):
        """Configura los estilos básicos del documento."""
        normal_style = doc.styles['Normal']
        normal_style.font.name = self.FONT_FAMILY
        normal_style.font.size = Pt(11)
    
    def set_paragraph_border_bottom(self, paragraph, color="000000", size=6, space=1):
        """Agrega un borde inferior a un párrafo."""
        p = paragraph._p
        pPr = p.get_or_add_pPr()
        
        # Remover borde existente si existe
        pb = pPr.find(qn('w:pBdr'))
        if pb is not None:
            pPr.remove(pb)
        
        # Crear nuevo borde
        pb = OxmlElement('w:pBdr')
        bottom = OxmlElement('w:bottom')
        bottom.set(qn('w:val'), 'single')
        bottom.set(qn('w:sz'), str(size))
        bottom.set(qn('w:space'), str(space))
        bottom.set(qn('w:color'), color)
        
        pb.append(bottom)
        pPr.append(pb)
