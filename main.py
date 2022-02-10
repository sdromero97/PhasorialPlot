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
    # Se cuenta con multiples alternativas = 1; Sólo una Alternativa = 0
    multp_alternativas=0
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

    # Se crean todas las figuras
    num_per_max=0
    for elemento in Elementos.unique()[0:2]:
        num_desp_max = 0
        num_dem_max = 0
        matriz_filtrada_1 = matriz_base.query("Elemento==@elemento")
        num_per=len(matriz_filtrada_1['Año_de_Estudio'].unique())
        if num_per>num_per_max:
            num_per_max=num_per
        for anio in matriz_filtrada_1['Año_de_Estudio'].unique():
            matriz_filtrada_2 = matriz_filtrada_1.query("Año_de_Estudio==@anio")
            num_desp = len(matriz_filtrada_2['Hidrología'].unique())
            if num_desp > num_desp_max:
                num_desp_max = num_desp
            for despacho in matriz_filtrada_2['Hidrología'].unique():
                matriz_filtrada_3 = matriz_filtrada_2.query("Hidrología==@despacho")
                num_dem = len(matriz_filtrada_3['Demanda'].unique())
                if num_dem > num_dem_max:
                    num_dem_max = num_dem
        Ventana, _ = plot.subplots(ncols=num_dem_max, nrows=num_per_max*num_desp_max,
                                    figsize=(width + 2, height), constrained_layout=True,
                                    subplot_kw=dict(projection='polar'))
        Ventana.suptitle('Tensión Fasorial en la Subestación '+str(elemento), fontsize=22)
        mng = plot.get_current_fig_manager()
        mng.window.state('zoomed')
    #plot.show()

    # Se inicia el ciclo que permite crear los fasores de las gráficas
    for elemento, index_fig in zip(Elementos.unique()[0:2], plot.get_fignums()):
        fig=plot.figure(index_fig)
        ax=fig.get_axes()
        matriz_filtrada_1 = matriz_base.query("Elemento==@elemento")
        num_per = len(matriz_filtrada_1['Año_de_Estudio'].unique())
        grafica=0
        for anio in matriz_filtrada_1['Año_de_Estudio'].unique():
            matriz_filtrada_2 = matriz_filtrada_1.query("Año_de_Estudio==@anio")
            for despacho in matriz_filtrada_2['Hidrología'].unique():
                matriz_filtrada_3 = matriz_filtrada_2.query("Hidrología==@despacho")
                num_dem = len(matriz_filtrada_3['Demanda'].unique())
                for demanda in matriz_filtrada_3['Demanda'].unique():
                    matriz_filtrada_4 = matriz_filtrada_3.query("Demanda==@demanda")
                    for alternativa in matriz_filtrada_4['Caso_de_Análisis'].unique():
                        matriz_filtrada_5 = matriz_filtrada_4.query("Caso_de_Análisis==@alternativa")
                        np.random.seed(1315425)
                        for condicion in matriz_filtrada_5['Condición_Operativa'].unique():
                            rgb = np.random.rand(3,)
                            matriz_filtrada_6 = matriz_filtrada_5.query("Condición_Operativa==@condicion")
                            # Se crea el fasor en la gráfica a partir de los datos filtrados
                            if multp_alternativas:
                                ax[grafica].bar(np.radians(matriz_filtrada_6['m:phiu']), matriz_filtrada_6['m:u'],
                                                width=.015, bottom=0.0, color=rgb,
                                                label=str(alternativa)+' - '+str(condicion))
                            else:
                                ax[grafica].bar(np.radians(matriz_filtrada_6['m:phiu']), matriz_filtrada_6['m:u'],
                                                width=.015, bottom=0.0, color=rgb, label=condicion)
                        # Se Realizan ajustes al gráfico
                        ax[grafica].set_xlim(np.radians(lim_min_theta), np.radians(lim_max_theta))
                        ax[grafica].set_ylim(lim_min_mag, lim_max_mag)
                        ax[grafica].set_rlabel_position(pos_label_radial)
                        ax[grafica].set_xticks(np.arange(np.radians(lim_min_theta),
                                                         np.radians(lim_max_theta + intervalo_theta),
                                                         np.radians(intervalo_theta)))
                        ax[grafica].set_xlabel('Theta [deg]')
                        ax[grafica].set_ylabel('Tensión [p.u.]')
                        ax[grafica].legend(loc=(1.05, 0))
                        ax[grafica].set_title('\nAño ' + str(anio) + ' - Hidrología ' + str(despacho) +
                                              '\nDemanda ' + str(demanda))
                        grafica+=1
        if not save_manually:
            # Se crea la carpeta que contendrá las gráficas en PDF
            ruta_pdf = (ruta_archivos + r'\Graficas en PDF')
            os.makedirs(ruta_pdf, exist_ok=True)
            # Se crea la carpeta que contendrá las gráficas en EMF
            ruta_imag = (ruta_archivos + r'\Graficas en EMF')
            os.makedirs(ruta_imag, exist_ok=True)
            #Se guardan las imágenes
            plot.savefig(ruta_pdf+'\\'+str(elemento)+'.pdf', dpi=300, bbox_inches='tight', facecolor='w')
            plot.savefig(ruta_imag+'\\'+str(elemento)+'.png', dpi=300, bbox_inches='tight', facecolor='w')
    if save_manually:
        plot.show()


