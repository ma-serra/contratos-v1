import streamlit as st
from datetime import datetime
from docx import Document
import io

def gerar_contrato(nome_cliente, cpf_cliente, endereco, cep, documento_upload):
    # Obtém a data atual
    data_atual = datetime.now()
    
    # Cria uma lista com os meses em português
    meses = [
        "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
        "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
    ]

    # Obtém o dia e o mês
    dia = data_atual.day
    mes = meses[data_atual.month - 1]  # O mês é indexado a partir de 0
    ano = data_atual.year

    # Formata a data no estilo desejado
    data_formatada = f"{dia} de {mes} de {ano}"

    # Remove espaços em branco indesejados
    nome_cliente = nome_cliente.strip()
    cpf_cliente = cpf_cliente.strip()
    endereco = endereco.strip()
    cep = cep.strip()

    # Lê o documento enviado pelo usuário
    doc = Document(documento_upload)

    # Dicionário de substituições
    substituicoes = {
        '{{nome_cliente}}': nome_cliente,
        '{{cpf_cliente}}': cpf_cliente,
        '{{endereco}}': endereco,
        '{{cep}}': cep,
        '{{data}}': data_formatada
    }

    # Percorre todos os parágrafos do documento
    for paragrafo in doc.paragraphs:
        for run in paragrafo.runs:  # Percorre cada "run" (parte do texto)
            for chave, valor in substituicoes.items():
                if chave in run.text:
                    # Substitui a chave pelo valor
                    run.text = run.text.replace(chave, valor)

                    # Aplica negrito somente para nome do cliente e CPF
                    if chave in ['{{nome_cliente}}', '{{cpf_cliente}}']:
                        run.bold = True
                    elif chave == '{{data}}':
                        run.bold = False  # Garante que a data não fique em negrito

    # Salva o documento modificado em um arquivo em memória
    output = io.BytesIO()
    doc.save(output)
    output.seek(0)  # Move o ponteiro para o início do arquivo em memória

    return output

# Interface em Streamlit
st.title("Gerador de Contratos")  
st.subheader("Envie seu arquivo Word para rápida personalização. Coloque as TAGS no molde para que o sistema possa substituir:")
st.subheader("'{{nome_cliente}}' , {{cpf_cliente}}, {{endereco}}, {{cep}}")
# Campos do formulário
nome_cliente = st.text_input("Nome do Cliente")
cpf_cliente = st.text_input("CPF do Cliente")
endereco = st.text_input("Endereço")
cep = st.text_input("CEP")

# Campo para upload do documento
documento_upload = st.file_uploader("Faça o upload do seu arquivo .docx", type="docx")

# Botão para gerar o contrato
if st.button("Gerar Contrato"):
    if nome_cliente and cpf_cliente and endereco and cep and documento_upload:
        # Gera o contrato e recebe o arquivo modificado
        arquivo_contrato = gerar_contrato(nome_cliente, cpf_cliente, endereco, cep, documento_upload)
        
        # Exibe um botão para baixar o arquivo gerado
        st.download_button(
            label="Baixar Contrato Modificado",
            data=arquivo_contrato,
            file_name=f'{nome_cliente}_Contrato_Modificado.docx',
            mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
    else:
        st.error("Por favor, preencha todos os campos e faça o upload do documento.")
