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
import streamlit.components.v1 as components


#----Funcionalidades---
def JuntarPDF(arquivosJuntar):

   try: 
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

   except:
      st.info("Não foi possível juntar!")
   

def DividirPDF(arquivoDividir):

   try: 
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

   except:
      st.info("Não foi possível dividir!")
   

def ComprimirPDF(arquivoComprimir):
   
   try:
      for i in arquivoComprimir:
         with open(i.name,"wb") as x:
            x.write(i.getbuffer())

      pdf = pikepdf.open(arquivoComprimir[0].name)
      name = (((arquivoComprimir[0].name.replace(".pdf","") + "-Unlocked.pdf").replace(" ","-")).replace("(","-")).replace(")","-")
      pdf.save(name)

      entrada = "./" + name
      saida = "./Arquivo_Compress.pdf"
      compress = "./pdfsizeopt.single"
      
      os.system("chmod +x ./pdfsizeopt.single")
      os.system("chmod +x ./pdfsizeopt_libexec/avian")
      os.system("chmod +x ./pdfsizeopt_libexec/gs")
      os.system("chmod +x ./pdfsizeopt_libexec/jbig2")
      os.system("chmod +x ./pdfsizeopt_libexec/png22pnm")
      os.system("chmod +x ./pdfsizeopt_libexec/pngout")
      os.system("chmod +x ./pdfsizeopt_libexec/python")
      os.system("chmod +x ./pdfsizeopt_libexec/sam2p")
      os.system("dir")
      os.system("{} --use-pngout=no {} {}".format(compress,entrada,saida))

      pdfFinal = pikepdf.open(saida)
      saidaFinal = "./Arquivo_Compress_Final.pdf"
      pdfFinal.save(saidaFinal)
      
      with open(saidaFinal,"rb") as arquivoFinal:
         st.download_button(label ="📥 Download",data = arquivoFinal,file_name=saida)

      for i in arquivoComprimir:
         os.remove(i.name)

      os.remove(entrada)
      os.remove(saida)
      os.remove(saidaFinal)

      st.success('Concluído!', icon="✅")

   except:
      st.info("Não foi possível comprimir!")
      
def ConverterPDF_EXCEL(arquivoConverter):
   
   try:     
      for i in arquivoConverter:
         with open(i.name,"wb") as x:
            x.write(i.getbuffer())

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
         
      for i in arquivoConverter:
         os.remove(i.name)

      os.remove("Planilha.csv")
      os.remove("Planilha.xlsx")

      st.success('Concluído!', icon="✅")   
      
   except:
      st.info("Não foi possível converter!")


def ConverterPDF_WORD(arquivoConverter,pages: Tuple = None): 

   try: 
      for i in arquivoConverter:
         with open(i.name,"wb") as x:
            x.write(i.getbuffer())

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
.css-12oz5g7 {padding: 2rem 1rem;}
.css-14xtw13 {visibility: hidden;}
span.css-9ycgxx.exg6vvm12 {
visibility: hidden;
white-space: nowrap;
}
span.css-9ycgxx.exg6vvm12::before {
    visibility: visible;
    content: "Selecionar os arquivos, ou arraste e solte os arquivos aqui";
    font-size: 1rem;
    font-family: "Source Sans Pro", sans-serif;
    font-weight: 400;
    line-height: 1.6;
    text-size-adjust: 100%;
    -webkit-tap-highlight-color: rgba(0, 0, 0, 0);
    -webkit-font-smoothing: auto;
    color: rgb(49, 51, 63);
    box-sizing: border-box;
    margin-bottom: 0.25rem;
}

small.css-1aehpvj.euu6i2w0 {visibility: hidden;}
small.css-1aehpvj.euu6i2w0::before {
    visibility: visible;
    content: "200MB por arquivo";
    font-family: "Source Sans Pro", sans-serif;
    font-weight: 400;
    text-size-adjust: 100%;
    -webkit-tap-highlight-color: rgba(0, 0, 0, 0);
    -webkit-font-smoothing: auto;
    box-sizing: border-box;
    color: rgba(49, 51, 63, 0.6);
    font-size: 14px;
    line-height: 1.25;
}

section.css-po3vlj.exg6vvm15 button{visibility:hidden;}

#Linkedin {margin-top: 75px;}
#desenvolvidoPor {color: black;}
#nome {color: black;}
</style>

<div id="Linkedin" class="badge-base LI-profile-badge" data-locale="pt_BR" data-size="medium" data-theme="light" data-type="VERTICAL" data-vanity="sérgio--brito" data-version="v1">
<a href="https://br.linkedin.com/in/s%C3%A9rgio--brito?trk=profile-badge"><img src="https://brand.linkedin.com/content/dam/me/business/en-us/amp/brand-site/v2/bg/LI-Bug.svg.original.svg" alt="Linkedin" style="width:42px;height:42px;"></a>
<a id="desenvolvidoPor">Desenvolvido por </a>
<a id="nome" class="badge-base__link LI-simple-link" href="https://br.linkedin.com/in/s%C3%A9rgio--brito?trk=profile-badge">Sérgio Brito</a>
</div>

"""

st.markdown(style, unsafe_allow_html=True)

my_js = """
document.getElementsByTagName('head')[0].appendChild( '<meta name="google-site-verification" content="eHqRAZ2787A7RfWNg4MTF_wCFXpt9i92UbQ-gh7BiRQ" />');
"""

my_html = f"<script>{my_js}</script>"

components.html(my_html)

