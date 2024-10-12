import streamlit as st
from datetime import datetime
from docx import Document
import io



def gerar_contrato(nome_cliente, cpf_cliente, endereco, cep):
    data_atual = datetime.now()

    meses = [
        "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
        "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
    ]

    dia = data_atual.day
    mes = meses[data_atual.month - 1] 
    ano = data_atual.year

    data_formatada = f"{dia} de {mes} de {ano}"

    nome_cliente = nome_cliente.strip()
    cpf_cliente = cpf_cliente.strip()
    endereco = endereco.strip()
    cep = cep.strip()

    doc = Document('Contrato_exemplo.docx')

    substituicoes = {
        '{{nome_cliente}}': nome_cliente,
        '{{cpf_cliente}}': cpf_cliente,
        '{{endereco}}': endereco,
        '{{cep}}': cep,
        '{{data}}': data_formatada
    }

    for paragrafo in doc.paragraphs:
        for run in paragrafo.runs: 
            for chave, valor in substituicoes.items():
                if chave in run.text:
                    # Substitui a chave pelo valor
                    run.text = run.text.replace(chave, valor)

                    if chave in ['{{nome_cliente}}', '{{cpf_cliente}}']:
                        run.bold = True
                    elif chave == '{{data}}':
                        run.bold = False 

    output = io.BytesIO()
    doc.save(output)
    output.seek(0) 

    return output


st.title("Gerador de Contratos")

nome_cliente = st.text_input("Nome do Cliente")
cpf_cliente = st.text_input("CPF do Cliente")
endereco = st.text_input("Endereço")
cep = st.text_input("CEP")

if st.button("Gerar Contrato"):
    if nome_cliente and cpf_cliente and endereco and cep:
        arquivo_contrato = gerar_contrato(
            nome_cliente, cpf_cliente, endereco, cep)

        st.download_button(
            label="Baixar Contrato",
            data=arquivo_contrato,
            file_name=f'{nome_cliente}_Contrato_de_Prestação_de_Serviços_Psicologicos.docx',
            mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
    else:
        st.error("Por favor, preencha todos os campos.")
