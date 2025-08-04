from io import BytesIO
from datetime import datetime
from typing import Tuple
import pandas as pd
import unicodedata
from .word_report_generator import WordReportGenerator
from .excel_report_generator import ExcelReportGenerator

class ReportFactory:
    """Factory para crear diferentes tipos de reportes."""
    
    def __init__(self):
        self.word_generator = WordReportGenerator()
        self.excel_generator = ExcelReportGenerator()
    
    def create_report(self, data: pd.DataFrame, empresa: str, anio: int, mes: str, funcionarios: dict) -> Tuple[BytesIO, str]:
        """
        Crea un reporte según el tipo de empresa.
        
        Args:
            data: Datos filtrados
            empresa: Nombre de la empresa
            anio: Año del reporte
            mes: Mes del reporte
            funcionarios: Información de funcionarios
            
        Returns:
            Tuple con el buffer del archivo y el tipo MIME
        """
        if empresa == "Ravago Americas LLC":
            buffer = self.excel_generator.create_ravago_report(data, anio, mes, funcionarios)
            mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        else:
            buffer = self.word_generator.generate_report(data, empresa, anio, mes, funcionarios)
            mime_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        
        return buffer, mime_type
    
    def build_report_filename(self, empresa: str, date: datetime = None) -> str:
        """
        Construye el nombre del archivo de reporte.
        
        Args:
            empresa: Nombre de la empresa
            date: Fecha para el nombre del archivo
            
        Returns:
            Nombre del archivo sugerido
        """
        date = date or datetime.now()
        date_tag = date.strftime("%y%m%d")  # YYMMDD
        empresa_tag = self._slug_empresa(empresa)
        return f"{date_tag}-LV-{empresa_tag}-Facturación honorarios.docx"
    
    def _slug_empresa(self, nombre: str) -> str:
        """Convierte el nombre de empresa a un slug válido para archivos."""
        nfkd = unicodedata.normalize("NFKD", nombre)
        ascii_only = nfkd.encode("ascii", "ignore").decode("ascii")
        return "".join(ch for ch in ascii_only if ch.isalnum())
