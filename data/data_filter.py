import pandas as pd

class DataFilter:
    """Maneja el filtrado de datos."""
    
    def filter_data(self, df: pd.DataFrame, empresa: str, anio: str, mes: str) -> pd.DataFrame:
        """
        Filtra el DataFrame según los criterios especificados.
        
        Args:
            df: DataFrame a filtrar
            empresa: Empresa seleccionada
            anio: Año seleccionado
            mes: Mes seleccionado
            
        Returns:
            DataFrame filtrado
        """
        if df.empty:
            return pd.DataFrame()

        filtered = df.copy()

        if empresa and empresa != "Todas":
            filtered = filtered[filtered['EMPRESA'] == empresa]
        
        if anio and anio != "Todos":
            filtered = filtered[filtered['AÑO ASIGNACION'] == anio]

        if mes and mes != "Todos":
            filtered = filtered[filtered['MES ASIGNACION'] == mes]
            
        return filtered
