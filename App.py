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

   os.remove('01-101-Unlocked.pdf')
   os.remove('01-101.pdf')
   os.remove('01-102-Unlocked.pdf')
   os.remove('01-102.pdf')
   os.remove('01-103-Unlocked.pdf')
   os.remove('01-103.pdf')
   os.remove('09\ 01-Unlocked.pdf')
   os.remove('09\ 01.pdf')
   os.remove('104\ 03-Unlocked.pdf')
   os.remove('104\ 03.pdf')
   os.remove('11\ 03-Unlocked.pdf')
   os.remove('11\ 03.pdf')
   os.remove('12\ 07-Unlocked.pdf')
   os.remove('12\ 07.pdf')
   os.remove('13\ 03\ 10-2022-Unlocked.pdf')
   os.remove('13\ 03\ 10-2022.pdf')
   os.remove('13\ 05\ 09-2022-Unlocked.pdf')
   os.remove('13\ 05\ 09-2022.pdf')
   os.remove('13\ 05-Unlocked.pdf')
   os.remove('13\ 05.pdf')
   os.remove('201\ 02-Unlocked.pdf')
   os.remove('201\ 02.pdf')
   os.remove('201\ 04-Unlocked.pdf')
   os.remove('201\ 04.pdf')
   os.remove('206\ 03-Unlocked.pdf')
   os.remove('206\ 03.pdf')
   os.remove('22\ 13\ 09-2022-Unlocked.pdf')
   os.remove('22\ 13\ 09-2022.pdf')
   os.remove('22\ 13-Unlocked.pdf')
   os.remove('22\ 13.pdf')
   os.remove('23\ 01-Unlocked.pdf')
   os.remove('23\ 01.pdf')
   os.remove('24\ 14\ 09-2022-Unlocked.pdf')
   os.remove('24\ 14\ 09-2022.pdf')
   os.remove('24\ 14-Unlocked.pdf')
   os.remove('24\ 14.pdf')
   os.remove('25\ 0-Unlocked.pdf')
   os.remove('25\ 0.pdf')
   os.remove('25\ 02-Unlocked.pdf')
   os.remove('25\ 02.pdf')
   os.remove('26\ 05\ 10-2022-Unlocked.pdf')
   os.remove('26\ 05\ 10-2022.pdf')
   os.remove('30\ 07\ 09-2022-Unlocked.pdf')
   os.remove('30\ 07\ 09-2022.pdf')
   os.remove('30\ 07-Unlocked.pdf')
   os.remove('30\ 07.pdf')
   os.remove('302\ 15-Unlocked.pdf')
   os.remove('302\ 15.pdf')
   os.remove('30229676_10_04-0502_274-Unlocked.pdf')
   os.remove('30229676_10_04-0502_274.pdf')
   os.remove('30230149_10_05-0101_321-Unlocked.pdf')
   os.remove('30230149_10_05-0101_321.pdf')
   os.remove('303\ 17-Unlocked.pdf')
   os.remove('303\ 17.pdf')
   os.remove('32\ 15-Unlocked.pdf')
   os.remove('32\ 15.pdf')
   os.remove('36\ 07-Unlocked.pdf')
   os.remove('36\ 07.pdf')
   os.remove('62\ 03\ 10-2022-Unlocked.pdf')
   os.remove('62\ 03\ 10-2022.pdf')
   os.remove('ALBERICO\ DOS\ SANTOS_41782504_ADITIVO-Unlocked.pdf')
   os.remove('ALBERICO\ DOS\ SANTOS_41782504_ADITIVO.pdf')
   os.remove('ALBERICO\ DOS\ SANTOS_41782504_LAUDO-Unlocked.pdf')
   os.remove('ALBERICO\ DOS\ SANTOS_41782504_LAUDO.pdf')
   os.remove('ALESSANDRA\ MARTINS\ DE\ MOURA_41489150_ADITIVO-Unlocked.pdf')
   os.remove('ALESSANDRA\ MARTINS\ DE\ MOURA_41489150_ADITIVO.pdf')
   os.remove('ALESSANDRA\ MARTINS\ DE\ MOURA_41489150_LAUDO-Unlocked.pdf')
   os.remove('ALESSANDRA\ MARTINS\ DE\ MOURA_41489150_LAUDO.pdf')
   os.remove('ALESSANDRA-MARTINS-DE-MOURA_41489150_ADITIVO-Unlocked.pdf')
   os.remove('ALESSANDRA-MARTINS-DE-MOURA_41489150_ADITIVO.pdf')
   os.remove('ALESSANDRA-MARTINS-DE-MOURA_41489150_LAUDO-Unlocked.pdf')
   os.remove('ALESSANDRA-MARTINS-DE-MOURA_41489150_LAUDO.pdf')
   os.remove('Arquivo\ (2).pdf')
   os.remove('Arquivo\ (92).pdf')
   os.remove('Arquivo\ (93).pdf')
   os.remove('Boleto\ Forte\ Versalhes\ -\ 10\ 2022\ -\ entregues\ 08\ 2022\ (2)-Unlocked.pdf')
   os.remove('Boleto\ Forte\ Versalhes\ -\ 10\ 2022\ -\ entregues\ 08\ 2022\ (2).pdf')
   os.remove('Boleto\ Forte\ Versalhes\ -\ 10\ 2022\ -\ entregues\ 09\ 2022\ (3)-Unlocked.pdf')
   os.remove('Boleto\ Forte\ Versalhes\ -\ 10\ 2022\ -\ entregues\ 09\ 2022\ (3).pdf')
   os.remove('Boleto\ Forte\ Versalhes\ -\ 10\ 2022\ -\ entregues\ 10\ 2022\ (2)-Unlocked.pdf')
   os.remove('Boleto\ Forte\ Versalhes\ -\ 10\ 2022\ -\ entregues\ 10\ 2022\ (2).pdf')
   os.remove('Boleto\ Forte\ Versalhes\ -\ ent\ 08\ (1)-Unlocked.pdf')
   os.remove('Boleto\ Forte\ Versalhes\ -\ ent\ 08\ (1).pdf')
   os.remove('Boleto\ Forte\ Versalhes\ -\ ent\ 09\ (4)-Unlocked.pdf')
   os.remove('Boleto\ Forte\ Versalhes\ -\ ent\ 09\ (4).pdf')
   os.remove('Boleto\ Forte\ Versalhes\ -\ ent\ 10\ (5)-Unlocked.pdf')
   os.remove('Boleto\ Forte\ Versalhes\ -\ ent\ 10\ (5).pdf')
   os.remove('Boleto\ mês06\ -\ 108-02\ Braga-Unlocked.pdf')
   os.remove('Boleto\ mês06\ -\ 108-02\ Braga.pdf')
   os.remove('Boleto\ mês06\ -206-20\ Braga-Unlocked.pdf')
   os.remove('Boleto\ mês06\ -206-20\ Braga.pdf')
   os.remove('Boleto\ mês07\ -\ 108-02\ Braga-Unlocked.pdf')
   os.remove('Boleto\ mês07\ -\ 108-02\ Braga.pdf')
   os.remove('Boleto\ mês07\ -206-20\ Braga-Unlocked.pdf')
   os.remove('Boleto\ mês07\ -206-20\ Braga.pdf')
   os.remove('Boleto\ mês08\ -\ 108-02\ Braga-Unlocked.pdf')
   os.remove('Boleto\ mês08\ -\ 108-02\ Braga.pdf')
   os.remove('Boleto\ mês08\ -206-20\ Braga-Unlocked.pdf')
   os.remove('Boleto\ mês08\ -206-20\ Braga.pdf')
   os.remove('Boleto\ mês09\ -\ 108-02\ Braga-Unlocked.pdf')
   os.remove('Boleto\ mês09\ -\ 108-02\ Braga.pdf')
   os.remove('Boleto\ mês09\ -\ 206-20\ Braga-Unlocked.pdf')
   os.remove('Boleto\ mês09\ -\ 206-20\ Braga.pdf')
   os.remove('Boleto\ mês10\ -\ 108-02\ Braga-Unlocked.pdf')
   os.remove('Boleto\ mês10\ -\ 108-02\ Braga.pdf')
   os.remove('Boleto\ mês10\ -\ 206-20\ Braga-Unlocked.pdf')
   os.remove('Boleto\ mês10\ -\ 206-20\ Braga.pdf')
   os.remove('Boleto\ mês10\ -206-20\ Braga-Unlocked.pdf')
   os.remove('Boleto\ mês10\ -206-20\ Braga.pdf')
   os.remove('COND.\ RESID.\ PARQUE\ AQUARELLE\ RELATORIO\ DE\ DEBITOS\ MRV.pdf')
   os.remove('DEMONSTRATIVO\ DE\ DÉBITO\ -\ 25-10-2022.pdf')
   os.remove('EDIMAR\ ANTONIO\ DA\ SILVA_41201979_ADITIVO-Unlocked.pdf')
   os.remove('EDIMAR\ ANTONIO\ DA\ SILVA_41201979_ADITIVO.pdf')
   os.remove('EDIMAR\ ANTONIO\ DA\ SILVA_41201979_LAUDO-Unlocked.pdf')
   os.remove('EDIMAR\ ANTONIO\ DA\ SILVA_41201979_LAUDO.pdf')
   os.remove('F-104\ Relatorio\ de\ debitos\ .pdf')
   os.remove('PARQUE\ AMABILIE\ -21-103-Unlocked.pdf')
   os.remove('PARQUE\ AMABILIE\ -21-103.pdf')
   os.remove('PARQUE\ AMABILIE\ -21-1031-Unlocked.pdf')
   os.remove('PARQUE\ AMABILIE\ -21-1031.pdf')
   os.remove('Planilha.csv')
   os.remove('Prática\ 06-Unlocked.pdf')
   os.remove('Prática\ 06.pdf')
   os.remove('ROBERTO\ ROSSI\ JÚNIOR_41436080_LAUDO-Unlocked.pdf')
   os.remove('ROBERTO\ ROSSI\ JÚNIOR_41436080_LAUDO.pdf')
   os.remove('aula5-Unlocked.pdf')
   os.remove('aula5.pdf')
   os.remove('novo.pdf')
   os.remove('pdfsizeopt')
   os.remove('pdfsizeoptDir')
   os.remove('singular\ 01-506\ -\ 1-Unlocked.pdf')
   os.remove('singular\ 01-506\ -\ 1.pdf')
   os.remove('singular\ 01-506-Unlocked.pdf')
   os.remove('singular\ 01-506.pdf')
   os.remove('teste-Unlocked.pdf')
   os.remove('teste.pdf')
   os.remove('teste1-Unlocked.pdf')
   os.remove('teste1.pdf')
   os.remove('teste2-Unlocked.pdf')
   os.remove('teste2.pdf')


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
         
      os.remove("Planilha.xlsx")

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
 
