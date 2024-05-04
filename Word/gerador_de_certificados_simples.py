from docx import Document
from docx.shared  import Pt  # para aumentar e diminuir o texto

# abrindo o arquivo word
file_word = Document(r"caminho onde ta o docx")
# tamanho
style_file = file_word.styles["Normal"]

# subistintuindo o "Nome" e subistituir pelo nome desejado
for paragraph_file in file_word.paragraphs:
    if "@nome" in paragraph_file.text:
        paragraph_file.text = "Francisca Lorrayne"
        source = style_file.font
        source.name = "Calibri (Corpo)"
        style_file.size = Pt(24)

# Salva o certificado com o nome do aluno
file_word.save("Francisca Lorrayne.docx")
print("Arquivo gerado com sucesso!")