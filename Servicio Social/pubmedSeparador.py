#Autor: Reyes Bolaños Martín

#Librerias utilizadas
import tkinter as tk
from tkinter import *
from tkinter import filedialog, messagebox, ttk
import pandas as pd 
import numpy as np
import os
import nltk
import fitz
import re
from nltk.text import ConcordanceIndex

guardarArchivo=[] #Coincidencias
registroID=''

#FUNCIÓN COINCIDENCIAS
def OpcionOraciones(Excel1,RutaPdf,TextoFinal,TextoKey,page_stuff2):
	try:
		TextoFinal.configure(state='normal')

		rutaExcel1=Excel1.get("1.0","end-1c")
		idPdf=RutaPdf.get("1.0","end-1c")
		rutaPDF='fetched_pdfs/'+idPdf+'.pdf'
		rutaKey=TextoKey.get("1.0","end-1c")	

		df = pd.read_excel(rutaExcel1)
		diccionario = df.set_index('Locus Tag').T.to_dict('dict')
		llaves = diccionario.keys()
		llaves = []
		nombres=[]
		sinonimos=[]

		for key in diccionario.keys():
		    llaves.append(key)
		    nombres.append(diccionario[key]['Gene Name'])
		    sinonimos.append(diccionario[key]['Product Synonyms'])

		#Abriendo PDF
		doc = fitz.open(rutaPDF)
		print ("number of pages: %i" % doc.pageCount) 

		#EXTRAYENDO ORACIONES CLAVE - 1 COINCIDENCIA
		#LEER TODO EL ARTICULO
		print("Encontrando 1 coincidencia")
		Lineas1=[]
		for pagina in range(doc.pageCount):
		   #EXTRAYENDO TEXTO DE ARTICULO
		    pagObj = doc.loadPage(pagina)
		    numPag = pagObj.getText("text")

		    #CONVIRTIENDO A OBTETO TEXTO
		    renglon = nltk.sent_tokenize(numPag)
			    
		    #En cada renglon de cada página
		    for oracion in renglon:
		        token= nltk.word_tokenize(oracion)
		        texto=nltk.Text(token)
			        
		        #Buscando llaves
		        for key in diccionario.keys():			            
		            if key in texto:
		                ci=ConcordanceIndex(texto)
		                results= concordance(ci,key)
		                for caso in results:
		                    Lineas1.append(oracion)
			            
		            if diccionario[key]['Gene Name'] in texto:
		                ci=ConcordanceIndex(texto)
		                results= concordance(ci,diccionario[key]['Gene Name'])
		                for caso in results:
		                    Lineas1.append(oracion)
			                    
		            if diccionario[key]['Product Synonyms'] in texto:
		                ci=ConcordanceIndex(texto)
		                results= concordance(ci,diccionario[key]['Product Synonyms'])
		                for caso in results:
		                    Lineas1.append(oracion)

		#EXTRAYENDO ORACIONES CLAVE - 2 COINCIDENCIA
		print("Encontrando 2 coincidencia")
		Lineas2=[]
		for renglon in Lineas1:
		    contador=0
		    renglon_token = nltk.word_tokenize(renglon)
		    for token in renglon_token:
		        if(token in llaves or 
		           token in nombres or 
		           token in sinonimos):
		            contador=contador+1
		            if(contador>1):
		                if(renglon not in Lineas2):
		                    Lineas2.append(renglon)

		#Lista de Palabras
		print("Detectando interaccion entre ellas")
		keyword = pd.read_excel(rutaKey)
		ListaKey=[]
		ListaAutokey=[]
		for palabra in keyword["KEYWORD"]:
			if (("auto-" in palabra) or ("self-" in palabra)):
				ListaAutokey.append(str(palabra))
			else:
				ListaKey.append(str(palabra))
			LineasInteraccion=[]
		LineasResultados=[]

		for renglon in Lineas2:
			for key in ListaAutokey:
				if key in renglon:
					if renglon not in LineasResultados:
						LineasResultados.append(renglon)

			for key in ListaKey: 
				if key in renglon:
					if renglon not in LineasInteraccion:
						LineasInteraccion.append(renglon)


	    #DATOS PSEUDOMONAS (FT,TG,SIGMA)	            
		factores = pd.read_excel(rutaExcel1)	
		factores = factores.replace(np.nan,'---')
		dFactores= factores.to_dict('list')

		if("Type1" and "Type2") in factores.columns:
			
			#COINCIDENCIAS CON FACTORES
			print("Buscando factores")
			for renglon in LineasInteraccion:
			    Sigma=False
			    TF=False
			    TG=False
			    contador=0
				    
			    renglon_token = nltk.word_tokenize(renglon)
			    for token in renglon_token:
			        if(token in dFactores['Locus Tag'] or
			           token in dFactores['Gene Name'] or
			           token in dFactores['Gene synonyns']):
			            TG=True
			            contador=contador+1
				            
			            if(token in dFactores['Locus Tag']):
			                index=dFactores['Locus Tag'].index(token)
			            if(token in dFactores['Gene Name']):
			                index=dFactores['Gene Name'].index(token)
			            if(token in dFactores['Gene synonyns']):
			                index=dFactores['Gene synonyns'].index(token)
				            
			            if(dFactores['Type1'][index] == 'TF'):
			                TF=True
			            if(dFactores['Type1'][index] == 'Sigma'):
			                Sigma=True
				            
			            if(contador>1):
			                if((Sigma==True or TF==True) and TG==True):
			                    if(renglon not in LineasResultados):
			                        LineasResultados.append(renglon)
			                        Sigma=False
			                        TF=False
			                        TG=False
			                        contador=0

			#Colocando resultados en variable final
			page_stuff2 = LineasResultados
			#Agregando texto a TextoFinal
			if (len(LineasResultados)==0):
				TextoFinal.insert(1.0, "No se encontraron Resultados")	 
			else:
				for frase in page_stuff2:
					TextoFinal.insert("insert", frase+"\n\n")	
		else:
			page_stuff2 = LineasResultados+LineasInteraccion
			#Agregando texto a TextoFinal
			if (len(page_stuff2)==0):
				TextoFinal.insert(1.0, "No se encontraron Resultados")	 
			else:
				for frase in page_stuff2:
					TextoFinal.insert("insert", frase+"\n\n")	
			TextoFinal.configure(state='disabled')
		print("Programa terminado")

		guardarArchivo=[]
		for i in page_stuff2:
			guardarArchivo.append(i)

	except:
		tk.messagebox.showerror("Informacion", "Error, parámetros incompletos o corruptos")

#Colocar coincidencias en la Listas de función
def concordance(ci, word, width=300, lines=25):
		half_width = (width - len(word) - 2) // 2
		context = width // 4
		results = []
		offsets = ci.offsets(word)
		if offsets:
			lines = min(lines, len(offsets))
			for i in offsets:
				if lines <= 0:
					break
				left = (' ' * half_width +
                    ' '.join(ci._tokens[i-context:i]))
				right = ' '.join(ci._tokens[i+1:i+context])
				left = left[-half_width:]
				right = right[:half_width]
				results.append('%s %s %s' % (left, ci._tokens[i], right))
				lines -= 1
		return results

#Botón - Limpiar textboxes
def limpiar_box(T1,Tkey,Tid,Tpdf,Tfinal):
	try:
		T1.configure(state='normal')
		Tkey.configure(state='normal')
		Tpdf.configure(state='normal')
		Tfinal.configure(state='normal')

		T1.delete("1.0", "end")
		Tkey.delete("1.0", "end")
		Tid.delete("1.0","end")
		Tpdf.delete("1.0","end")
		Tfinal.delete("1.0","end")

		T1.configure(state='disabled')
		Tkey.configure(state='disabled')
		Tpdf.configure(state='disabled')
		Tfinal.configure(state='disabled')
	except:
		tk.messagebox.showerror("Informacion", "Error al eliminar cuadros de Texto")

def obtenerID(Texto_id,Texto_pdf,TextoFinal):
	Texto_pdf.configure(state='normal')
	TextoFinal.configure(state='normal')

	Texto_pdf.delete("1.0","end")
	TextoFinal.delete("1.0","end")
	id_pdf= Texto_id.get("1.0","end-1c")
	registroID=id_pdf
	if os.path.isfile('fetched_pdfs/'+id_pdf+'.pdf'):
		print('Archivo existe')
		rutaPdf='fetched_pdfs/'+id_pdf+'.pdf'
		#Página a leer
		try:
			doc = fitz.open(str(rutaPdf))
			for pagina in range(doc.pageCount):
				page = doc.loadPage(pagina)
				page_stuff = page.getText("text")
				Texto_pdf.insert(1.0, page_stuff)
		except RuntimeError:
			print("Archivo corrupto, RunTime ERROR")
			tk.messagebox.showerror("Informacion", "Archivo corrupto, prueba con otro ID")
		except:
			print("ERROR")
	else:
		print("Descargando")
		try:
			os.system('cls')
			os.system('python fetch_pdfs.py -pmids '+id_pdf)
		except RuntimeError:
			print("Archivo corrupto, RunTime ERROR")
			tk.messagebox.showerror("Informacion", "Archivo corrupto, prueba con otro ID")
		except:
			print("Error")
		if os.path.isfile('fetched_pdfs/'+id_pdf+'.pdf'):
			print("exito")
			#Página a leer
			try:
				doc = fitz.open('fetched_pdfs/'+id_pdf+'.pdf')
				for pagina in range(doc.pageCount):
				    page = doc.loadPage(pagina)
				    page_stuff = page.getText("text")
				    Texto_pdf.insert(1.0, page_stuff)
			except RuntimeError:
				print("Archivo corrupto, RunTime ERROR")
				tk.messagebox.showerror("Informacion", "Archivo corrupto, prueba con otro ID")
			except:
				tk.messagebox.showerror("Informacion", "El archivo elegido no tiene formato xlsx o csv")
				return None

		else:
			Texto_pdf.delete("1.0","end")
			tk.messagebox.showerror("Informacion", "Archivo no encontrado en PUBMED, intente con otro ID")
	Texto_pdf.configure(state='disabled')
	TextoFinal.configure(state='disabled')


#Botón - Abriendo Archivo
def open_excel(Texto):
	Texto.configure(state='normal')
	Texto.delete("1.0", "end")
	#Tomando archivo
	rutaExcel = filedialog.askopenfilename(
		initialdir="C:/gui/",
		title="Selecciona un archivo",
		filetypes=(("xlsx files", "*.xlsx"),("All Files", "*.*"))
		)

	try:
		#Añadiendo a TextBox
		Texto.insert(1.0, rutaExcel)
		print(str(rutaExcel))
	#Errores
	except ValueError:
		tk.messagebox.showerror("Informacion", "El archivo elegido no tiene formato xlsx")
		return None
	except FileNotFoundError:
		tk.messagebox.showerror("Informacion", f"Archivo no encontrado como {rutaArchivo}")
		return None
	Texto.configure(state='disabled')

def exportarArchivo(RutaID,Coincidencias):
	cuadroId=RutaID.get("1.0","end-1c")
	cuadroCoincidencias=Coincidencias.get("1.0","end-1c")
	try:
		if os.path.isdir('Coincidencias'):
			print('Carpeta existe')
		else:
			os.mkdir('Coincidencias')

		with open('Coincidencias/'+cuadroId+'.txt','w',encoding="utf-8") as archivoID:
			archivoID.write(cuadroCoincidencias)
			#for i in guardarArchivo:
				#archivoID.write(i)
				#archivoID.write('\n\n')
		archivoID.close()
	except:
		print("ERROR AL EXPORTAR ARCHIVO")
		tk.messagebox.showerror("Informacion", "ERROR: Verificar que hay texto en el apartado Coincidencia")



#INTERFAZ GRAFICA
root = Tk()
root.title('Coincidencias SS')
root.geometry("1000x850")

#Direcciones Excel1
labelE1 = Label(root, text='Factores')
labelE1.place(relx=0.02,rely=0.01)
TextoE1 = Text(root, height=1.5, width=70)
TextoE1.place(relx=0.15)
TextoE1.configure(state='disabled')
botonE1 = tk.Button(root, text='Abrir', command=lambda: open_excel(TextoE1) )
botonE1.place(relx=0.1,rely=0.01)

#Direcciones Key
labelKey = Label(root, text='Keywords')
labelKey.place(relx=0.02,rely=0.06)
TextoKey = Text(root, height=1.5, width=70)
TextoKey.place(relx=0.15,rely=0.05)
TextoKey.configure(state='disabled')
botonKey = Button(root, text='Abrir', command= lambda: open_excel(TextoKey) )
botonKey.place(relx=0.1,rely=0.06)

scrolly = tk.Scrollbar(root)
scrolly.pack(side="right", fill="y")

#Leer PDF
labelText = Label(root, text='Contenido PDF')
labelText.place(relx=0.02,rely=0.18)
TextoPdf = Text(root, height=20, width=115,wrap=None, yscrollcommand=scrolly.set)
TextoPdf.place(relx=0.02,rely=0.21)
TextoPdf.configure(state='disabled')

scrolly.config(command=TextoPdf.yview)

#Palabras Finales
labelFinal = Label(root, text='Coincidencias')
labelFinal.place(relx=0.02,rely=0.62)
TextoFinal = Text(root, height=9, width=80,font=("Arial",15))
TextoFinal.place(relx=0.02,rely=0.65)
TextoFinal.configure(state='disabled')

#Direcciones ID
labelID = Label(root, text='ID - PDF')
labelID.place(relx=0.02,rely=0.11)
botonID = Button(root, text='Abrir', command= lambda: obtenerID(TextoID,TextoPdf,TextoFinal) )
botonID.place(relx=0.1,rely=0.11)

TextoID = Text(root, height=1.5, width=40)
TextoID.place(relx=0.15,rely=0.105)

#Acciones
botonLimpiar = Button(root, text='Limpiar', command= lambda: limpiar_box(TextoE1,TextoKey,TextoID,TextoPdf,TextoFinal) )
botonLimpiar.place(relx=0.85,rely=0.03)
botonEjecutar = Button(root, text='Ejecutar', command= lambda: OpcionOraciones(TextoE1,TextoID,TextoFinal,TextoKey,guardarArchivo) )
botonEjecutar.place(relx=0.85,rely=0.08)

#Exportar archivo
labelExportar = Label(root, text='Exportar')
labelExportar.place(relx=0.8,rely=0.96)
botonExportar = Button(root, text='Guardar', command= lambda: exportarArchivo(TextoID,TextoFinal) )
botonExportar.place(relx=0.85,rely=0.96)

root.mainloop()

#def File_dialog():
#	pass