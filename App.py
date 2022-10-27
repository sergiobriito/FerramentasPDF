import os
import time
import zipfile
import pikepdf
import tabula
from PyPDF2 import PdfFileReader, PdfFileMerger, PdfFileWriter
import glob
import csv
from pdf2docx import parse
from typing import Tuple
from xlsxwriter.workbook import Workbook
import streamlit as st


#----Funcionalidades---
def JuntarPDF(arquivosJuntar):
   
   pdf_editor = PdfFileMerger()
   
   for i in arquivosJuntar:
      with open(i.name,"wb") as x:
         x.write(i.getbuffer())

      pdf = pikepdf.open(i.name)
      name = i.name.replace(".pdf","") + "-Unlocked.pdf"
      pdf.save(name)
      pdf_editor.append(name)

   pdf_editor.write("Arquivo.pdf")
   pdf_editor.close()
   
   with open("Arquivo.pdf","rb") as arquivoFinal:
      st.download_button(label="📥 Download",data=arquivoFinal,file_name="Arquivo.pdf")
         

   for i in arquivosJuntar:
      os.remove(i.name)
      os.remove(i.name.replace(".pdf","") + "-Unlocked.pdf")

   os.remove("Arquivo.pdf")

   st.success('Concluído!', icon="✅")

   
   

def DividirPDF(arquivoDividir):
   pdf_conteudo = PdfFileReader(arquivoDividir, "rb")
   totalPaginas = pdf_conteudo.getNumPages()
   arquivoZIP = zipfile.ZipFile("Arquivos.zip", "w")
       
   for pagina in range(totalPaginas):
      pdf_editor = PdfFileWriter()
      pdf_editor.addPage(pdf_conteudo.getPage(pagina))
      nomePaginaPDF = "Página"+str(pagina+1)+".pdf"

      with open(nomePaginaPDF, "wb") as x:
            pdf_editor.write(x)
            
      arquivoZIP.write(nomePaginaPDF, nomePaginaPDF)

   arquivoZIP.close()
   with open("Arquivos.zip","rb") as arquivoFinal:
      st.download_button(label ="📥 Download",data = arquivoFinal,file_name="Arquivos.zip",mime="application/zip")

   for pagina in range(totalPaginas):
      nomePaginaPDF = "Página"+str(pagina+1)+".pdf"
      os.remove(nomePaginaPDF)

   os.remove("Arquivos.zip")

   st.success('Concluído!', icon="✅")

   

def ComprimirPDF(arquivoComprimir):
   
   st.info("Em desenvolvimento...")
   
   for i in arquivoComprimir:
      with open(i.name,"wb") as x:
         x.write(i.getbuffer())

   compress = "./pdfsizeopt.single"
   entrada = "./novo.pdf"
   saida = "./ArquivoCompress.pdf"


   os.system("chmod +x ./pdfsizeopt.single")
   os.system("chmod +x ./pdfsizeopt_libexec/avian")
   os.system("chmod +x ./pdfsizeopt_libexec/gs")
   os.system("chmod +x ./pdfsizeopt_libexec/jbig2")
   os.system("chmod +x ./pdfsizeopt_libexec/png22pnm")
   os.system("chmod +x ./pdfsizeopt_libexec/pngout")
   os.system("chmod +x ./pdfsizeopt_libexec/python")
   os.system("chmod +x ./pdfsizeopt_libexec/sam2p")
   os.system("dir")
   os.system("{} {} {}".format(compress,entrada,saida))
   
   #with open(saida,"rb") as arquivoFinal:
   #   st.download_button(label ="📥 Download",data = arquivoFinal,file_name=saida)
   #st.success('Concluído!', icon="✅")
       
   os.remove(entrada)
   os.remove(saida)
      
def ConverterPDF_EXCEL(arquivoConverter):
   
   for i in arquivoConverter:
         with open(i.name,"wb") as x:
            x.write(i.getbuffer())
   try:         
      dados = tabula.io.read_pdf(arquivoConverter[0].name, pages='all')[0]
      tabula.convert_into(arquivoConverter[0].name, "Planilha.csv", output_format="csv", pages='all',java_options="-Dfile.encoding=UTF8")
   
      for csvfile in glob.glob(os.path.join('.', '*.csv')):
       workbook = Workbook(csvfile[:-4] + '.xlsx')
       worksheet = workbook.add_worksheet()
       with open(csvfile, 'rt', encoding='utf8') as f:
           reader = csv.reader(f)
           for r, row in enumerate(reader):
               for c, col in enumerate(row):
                   worksheet.write(r, c, col)
       workbook.close()

      with open("Planilha.xlsx","rb") as arquivoFinal:
         st.download_button(label ="📥 Download",data = arquivoFinal,file_name="Planilha.xlsx")
         
      os.remove("Planilha.csv")
      os.remove("Planilha.xlsx")

      for i in arquivoConverter:
         os.remove(i.name)

      st.success('Concluído!', icon="✅")   
      
   except:
      st.info("Não foi possível converter!")


def ConverterPDF_WORD(arquivoConverter,pages: Tuple = None):

   for i in arquivoConverter:
         with open(i.name,"wb") as x:
            x.write(i.getbuffer())

   try: 
      entrada = arquivoConverter[0].name
      saida = str(arquivoConverter[0].name).replace(".pdf","") + ".docx"

      if pages:
            pages = [int(i) for i in list(pages) if i.isnumeric()]
               
      result = parse(pdf_file=entrada,docx_with_path=saida, pages=pages)

      summary = {
              "File": entrada, "Pages": str(pages), "Output File": saida
          }
               
      with open(saida,"rb") as arquivoFinal:
            st.download_button(label ="📥 Download",data = arquivoFinal,file_name=saida)

      os.remove(entrada)     
      os.remove(saida)

      st.success('Concluído!', icon="✅")     
   
   except:
      st.info("Não foi possível converter!")



#---Navegador---

st.set_page_config(page_icon="📄", page_title="Ferramentas para PDF")
st.title("📄 Ferramentas para PDF")

funcionalidaEscolhida = st.radio("Selecione uma opção:",("Juntar PDF", "Dividir PDF","Comprimir PDF","Converter PDF para Excel","Converter PDF para Word"))

if funcionalidaEscolhida == "Juntar PDF":
   arquivos = st.file_uploader("Escolha os arquivos:", accept_multiple_files=True)
   botaoExecutar = st.button("Executar")
   if botaoExecutar:
      with st.spinner('Processando...'):
         JuntarPDF(arquivos)


if funcionalidaEscolhida == "Dividir PDF":
   arquivo = st.file_uploader("Escolha o arquivo:", accept_multiple_files=False)
   if arquivo is not None:
      botaoExecutar = st.button("Executar")
      if botaoExecutar:
         with st.spinner('Processando...'):
            DividirPDF(arquivo)
            

if funcionalidaEscolhida == "Comprimir PDF":
   arquivo = st.file_uploader("Escolha o arquivo:", accept_multiple_files=True)
   if len(arquivo) == 1:
      botaoExecutar = st.button("Executar")
      if botaoExecutar:
         with st.spinner('Processando...'):
            ComprimirPDF(arquivo)


if funcionalidaEscolhida == "Converter PDF para Excel":
   arquivo = st.file_uploader("Escolha o arquivo:", accept_multiple_files=True)
   if len(arquivo) == 1:
      botaoExecutar = st.button("Executar")
      if botaoExecutar:
         with st.spinner('Processando...'):
            ConverterPDF_EXCEL(arquivo)
         

if funcionalidaEscolhida == "Converter PDF para Word":
   arquivo = st.file_uploader("Escolha o arquivo:", accept_multiple_files=True)
   if len(arquivo) == 1:
      botaoExecutar = st.button("Executar")
      if botaoExecutar:
         with st.spinner('Processando...'):
            ConverterPDF_WORD(arquivo)
         

style = """
<style>
#MainMenu {visibility: hidden;}
header {visibility: hidden;}
footer {visibility: hidden;}
footer:after {
visibility: visible;
content: 'Criado por Sérgio Brito';
display: block;
position: relative;
color: black;}
.css-12oz5g7 {padding: 2rem 1rem;}
.css-14xtw13 {visibility: hidden;}
</style>
"""

st.markdown(style, unsafe_allow_html=True)
 
