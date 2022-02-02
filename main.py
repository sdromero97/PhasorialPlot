# This is a sample Python script.
import numpy as np
import matplotlib.pyplot as plot
import pandas as pd
import os
from screeninfo import get_monitors

# Corresponde al Código Principal
if __name__ == '__main__':

    ######################### Configuraciones por el usuario ###################################
    ############################################################################################
    # Ruta completa del excel donde se tienen los resultados
    ruta_excel_resultados = r'D:\Proyectos Perú\12. 3474-COE001 COYA-YANA\Wampacs\Resultados\Ensayo DPL\Resultados.xlsx'
    # Se establece el tamaño de fuente del título
    font_title_fig=22
    # Se establecen los límites de los ejes de las gráficas
    lim_min_theta, lim_max_theta, intervalo_theta = [-90, 20, 10]         # Límites en grados
    lim_min_mag, lim_max_mag = [0.95,1.05]          # Límites en p.u.
    # Se establece la posición de las etiquetas del eje radial
    pos_label_radial=-90            # Posición en Grados
    # Guardar Gráficas Manualmente = 1; Guardar Gráficas de forma automática = 0
    save_manually=1
    # Ruta en la cual se almacenan las carpetas con las imágenes de forma automática
    ruta_archivos = r'D:\Proyectos Perú\12. 3474-COE001 COYA-YANA\Wampacs\Resultados\Ensayo DPL'
    ############################################################################################
    ########################### Terminan las configuraciones ###################################

    # Se obtiene el tamaño de la pantalla en pulgadas para crear las imagenes con dicho tamaño
    width = get_monitors()[0].width_mm * 0.0393701
    height = get_monitors()[0].height_mm * 0.0393701

    # Se obtiene los datos del libro de Microsoft Excel en un DataFrame
    workbook = pd.read_excel(ruta_excel_resultados, sheet_name='Barras')
    # Se obtiene el vector con todos los Barrajes encontrados en los datos
    Elementos = workbook['Elemento']
    # Se crea la Matriz Base que contiene todos los datos
    matriz_base=workbook[['Caso_de_Estudio', 'Año_de_Estudio', 'Hidrología', 'Demanda','Caso_de_Análisis',
                          'Condición_Operativa','Elemento','m:u','m:phiu']]

    # Se crea la carpeta que contendrá las gráficas en PDF
    ruta_pdf = (ruta_archivos + r'\Graficas en PDF')
    os.makedirs(ruta_pdf, exist_ok=True)
    # plot.savefig('scatter2.png', dpi=300, bbox_inches='tight', facecolor='w')
    # Se crea la carpeta que contendrá las gráficas en EMF
    ruta_imag = (ruta_archivos + r'\Graficas en EMF')
    os.makedirs(ruta_imag, exist_ok=True)

    # Se inicia el ciclo que permite crear las gráficas
    for elemento in Elementos.unique()[0:2]:
        # Se filtra la matriz con los resultados para cada elemento
        matriz_filtrada_1 = matriz_base.query("Elemento==@elemento")
        # Se obtiene la cantidad de despachos y demandas encontrada luego de filtrar para crear el gráfico
        num_dem=len(matriz_filtrada_1['Demanda'].unique())
        num_desp=len(matriz_filtrada_1['Hidrología'].unique())
        # Se crea la ventana del gráfico y se obtiene el arreglo de sub graficas "ax"
        Ventana, ax = plot.subplots(ncols=num_dem,nrows=num_desp,figsize=(width+2,height), constrained_layout=True,
                                    subplot_kw=dict(projection='polar'))
        # Se Agrega el Título a la Ventana, que corresponde al nombre del elemento
        Ventana.suptitle(str(elemento), fontsize=22)
        # Se obtiene el administrador de la figura y se agranda a pantalla completa
        mng = plot.get_current_fig_manager()
        mng.window.state('zoomed')
        #Se continua con la creación de gráficas
        for anio in matriz_filtrada_1['Año_de_Estudio'].unique():
            # Se filtra la matriz con los resultados para cada Año
            matriz_filtrada_2 = matriz_filtrada_1.query("Año_de_Estudio==@anio")
            # Se continua clasificando los datos a graficar y nos paramos en el subplot correspondiente con row
            for despacho, row in zip(matriz_filtrada_2['Hidrología'].unique(), range(num_desp)):
                # Se filtra la matriz con los resultados para cada despacho
                matriz_filtrada_3 = matriz_filtrada_2.query("Hidrología==@despacho")
                # Se continua clasificando los datos a graficar y nos paramos en el subplot correspondiente con col
                for demanda, col in zip(matriz_filtrada_3['Demanda'].unique(), range(num_dem)):
                    # Se filtra la matriz con los resultados para cada demanda
                    matriz_filtrada_4 = matriz_filtrada_3.query("Demanda==@demanda")
                    # Se continua clasificando los datos a graficar
                    for alternativa in matriz_filtrada_4['Caso_de_Análisis'].unique():
                        # Se filtra la matriz con los resultados para cada caso de análisis
                        matriz_filtrada_5 = matriz_filtrada_4.query("Caso_de_Análisis==@alternativa")
                        # Se establece una semilla para que ramdon arroje siempre los mismos valores pseudoaleatorios
                        np.random.seed(255)
                        for condicion in matriz_filtrada_5['Condición_Operativa'].unique():
                            # Se obtiene un vector con valores aleatorios para definir un color para el fasor
                            rgb = np.random.rand(3,)
                            # Se filtra la matriz con los resultados para cada condición operativa
                            matriz_filtrada_6 = matriz_filtrada_5.query("Condición_Operativa==@condicion")
                            # Se crea el fasor en la gráfica, definiendole el color y la etiqueta de la leyenda
                            ax[row,col].bar(np.radians(matriz_filtrada_6['m:phiu']), matriz_filtrada_6['m:u'],
                                            width=.015, bottom=0.0, color=rgb, label=condicion)
                        # Se ajustan los limites de los ejes polares
                        ax[row,col].set_xlim(np.radians(lim_min_theta),np.radians(lim_max_theta))
                        ax[row,col].set_ylim(lim_min_mag,lim_max_mag)
                        # Se gira las etiquetas del eje radial para que queden verticales
                        ax[row,col].set_rlabel_position(pos_label_radial)
                        # Se establece la separación radial de los grados en 10 deg
                        ax[row, col].set_xticks(np.arange(np.radians(lim_min_theta), np.radians(lim_max_theta+intervalo_theta), np.radians(intervalo_theta)))
                        # Se ajustan las etiquetas de los ejes
                        ax[row, col].set_xlabel('Theta [deg]', rotation=0)
                        ax[row, col].set_ylabel('Tensión [p.u.]')
                        # Se crea la leyenda de la gráfica y se posiciona en la esquina inferior derecha
                        ax[row,col].legend(loc=(1.1, 0))
                        # Se asigna el nombre de la sub grafica en la figura
                        ax[row,col].set_title('Año: '+str(anio)+' - Hidrología: '+str(despacho)+
                                              '\nDemanda: '+str(demanda)+' - Caso: '+str(alternativa))
        if not save_manually:
            plot.savefig(ruta_pdf+'\\'+str(elemento)+'.pdf', dpi=300, bbox_inches='tight', facecolor='w')
            plot.savefig(ruta_imag+'\\'+str(elemento)+'.png', dpi=300, bbox_inches='tight', facecolor='w')
    if save_manually:
        plot.show()


