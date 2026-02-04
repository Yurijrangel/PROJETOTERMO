#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
App Streamlit - Gerador de Termos em PDF em Lote
Suporta: UNIANDRADE, UNIB, UNISMG
"""

import streamlit as st
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image
from reportlab.lib.enums import TA_JUSTIFY, TA_CENTER
from datetime import datetime
import os
import io
import zipfile
from pathlib import Path


# Configurações das IES
IES_CONFIG = {
    'UNIANDRADE': {
        'nome_completo': 'Centro Universitário Campos de Andrade – UNIANDRADE',
        'sigla': 'UNIANDRADE',
        'logo': 'logos/logo uni.png'
    },
    'UNIB': {
        'nome_completo': 'Universidade Ibirapuera - UNIB',
        'sigla': 'UNIB',
        'logo': 'logos/logo unib.png'
    },
    'UNISMG': {
        'nome_completo': 'Centro Universitário Santa Maria da Glória - UNISMG',
        'sigla': 'UNISMG',
        'logo': 'logos/logo smg.png'
    }
}


def formatar_cpf(cpf):
    """Formata CPF para o padrão XXX.XXX.XXX-XX"""
    cpf = str(cpf).replace('.', '').replace('-', '').replace(' ', '')
    if len(cpf) == 11:
        return f"{cpf[:3]}.{cpf[3:6]}.{cpf[6:9]}-{cpf[9:]}"
    return cpf


def formatar_data_extenso():
    """Retorna a data atual por extenso em português"""
    data_atual = datetime.now()
    meses = {
        1: 'janeiro', 2: 'fevereiro', 3: 'março', 4: 'abril',
        5: 'maio', 6: 'junho', 7: 'julho', 8: 'agosto',
        9: 'setembro', 10: 'outubro', 11: 'novembro', 12: 'dezembro'
    }
    return f"{data_atual.day} de {meses[data_atual.month]} de {data_atual.year}"


def gerar_termo_pdf_bytes(aluno_data, ies):
    """
    Gera um termo em PDF e retorna os bytes do arquivo
    
    Args:
        aluno_data: Dicionário com os dados do aluno
        ies: Código da IES (UNIANDRADE, UNIB ou UNISMG)
    
    Returns:
        tuple: (bytes do PDF, nome do arquivo)
    """
    # Validar IES
    if ies not in IES_CONFIG:
        raise ValueError(f"IES '{ies}' não é válida. Use: UNIANDRADE, UNIB ou UNISMG")
    
    ies_info = IES_CONFIG[ies]
    
    # Nome do arquivo PDF
    nome_arquivo = f"{aluno_data['NOME'].replace(' ', '_')}_{ies}_termo.pdf"
    
    # Criar buffer de memória para o PDF
    buffer = io.BytesIO()
    
    # Criar documento PDF
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=2*cm,
        leftMargin=2*cm,
        topMargin=2*cm,
        bottomMargin=2*cm
    )
    
    # Estilos
    styles = getSampleStyleSheet()
    
    # Estilo para título
    style_titulo = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=14,
        textColor='black',
        spaceAfter=30,
        alignment=TA_CENTER,
        fontName='Helvetica-Bold'
    )
    
    # Estilo para texto justificado
    style_texto = ParagraphStyle(
        'Justify',
        parent=styles['BodyText'],
        fontSize=11,
        alignment=TA_JUSTIFY,
        spaceAfter=12,
        leading=16,
        fontName='Helvetica'
    )
    
    # Estilo para assinatura (centralizado)
    style_assinatura = ParagraphStyle(
        'Assinatura',
        parent=styles['BodyText'],
        fontSize=11,
        alignment=TA_CENTER,
        spaceAfter=6,
        leading=14,
        fontName='Helvetica'
    )
    
    # Estilo para data (alinhado à direita)
    style_data = ParagraphStyle(
        'Data',
        parent=styles['BodyText'],
        fontSize=11,
        alignment=2,  # RIGHT alignment
        spaceAfter=6,
        leading=14,
        fontName='Helvetica'
    )
    
    # Conteúdo do documento
    story = []
    
    # Adicionar logo se existir
    logo_path = ies_info.get('logo')
    if logo_path and os.path.exists(logo_path):
        try:
            # Carregar logo e ajustar tamanho mantendo proporção
            logo = Image(logo_path)
            
            # Obter proporção original da imagem
            aspect_ratio = logo.imageWidth / logo.imageHeight
            
            # Definir tamanho máximo
            max_width = 8*cm
            max_height = 3.5*cm
            
            # Calcular dimensões mantendo proporção
            if aspect_ratio > (max_width / max_height):
                # Imagem mais larga - limitar pela largura
                logo.drawWidth = max_width
                logo.drawHeight = max_width / aspect_ratio
            else:
                # Imagem mais alta - limitar pela altura
                logo.drawHeight = max_height
                logo.drawWidth = max_height * aspect_ratio
            
            # Adicionar logo centralizado
            logo.hAlign = 'CENTER'
            story.append(logo)
            story.append(Spacer(1, 0.5*cm))
        except Exception as e:
            # Se houver erro ao carregar logo, continua sem ele
            pass
    
    # Título
    titulo = Paragraph("TERMO DE RESPONSABILIDADE DE ENTREGA DE DOCUMENTOS", style_titulo)
    story.append(titulo)
    story.append(Spacer(1, 1*cm))
    
    # Formatar CPF e data
    cpf_formatado = formatar_cpf(aluno_data['CPF'])
    data_extenso = formatar_data_extenso()
    
    # Primeiro parágrafo
    texto1 = f"""
    Eu, <b>{aluno_data['NOME']}</b>, brasileiro(a), inscrito(a) no Cadastro de Pessoas 
    Físicas CPF/MF sob n. <b>{cpf_formatado}</b>, residente e domiciliado(a) na rua 
    <b>{aluno_data['RUA']}</b>, <b>{aluno_data['BAIRRO']}</b>, <b>{aluno_data['CIDADE']}</b>, 
    <b>{aluno_data['UF']}</b>, declaro que entreguei total ou parcialmente todos os documentos 
    necessários para conclusão da minha matrícula.
    """
    paragrafo1 = Paragraph(texto1, style_texto)
    story.append(paragrafo1)
    
    # Segundo parágrafo
    texto2 = """
    Os documentos entregues serão validados pela secretaria, que poderá solicitar a 
    complementação ou nova entrega dos mesmos.
    """
    paragrafo2 = Paragraph(texto2, style_texto)
    story.append(paragrafo2)
    
    # Terceiro parágrafo (varia conforme IES)
    if ies == 'UNIANDRADE':
        texto3 = f"""
        Declaro ainda sob as penas da lei e para os devidos fins de direito, que assumo 
        integral responsabilidade pela apresentação dos comprovantes de conclusão do ensino 
        médio até 01(um) dia útil antes do início das aulas do curso de <b>{aluno_data['CURSO']}</b>, 
        a ser ministrado pelo {ies_info['nome_completo']}.
        """
    else:
        texto3 = f"""
        Declaro sob as penas da lei e para os devidos fins de direito, que assumo 
        integral responsabilidade pela apresentação dos comprovantes de conclusão do ensino 
        médio até 01(um) dia útil antes do início das aulas do curso de <b>{aluno_data['CURSO']}</b>, 
        a ser ministrado pela {ies_info['nome_completo']}.
        """
    
    paragrafo3 = Paragraph(texto3, style_texto)
    story.append(paragrafo3)
    
    # Quarto parágrafo
    texto4 = f"""
    Declaro ainda, que tenho pleno conhecimento que a ausência de apresentação do comprovante 
    de conclusão do ensino médio até o prazo acima mencionado, acarretará o imediato 
    cancelamento de minha matrícula, sem direito a restituição de qualquer mensalidade ou 
    taxa anteriormente paga, isentando a {ies_info['sigla']} de qualquer responsabilidade ou 
    obrigação relativa ao descumprimento por mim ocasionado. E, por ser a expressão da verdade, 
    firmo o presente.
    """
    paragrafo4 = Paragraph(texto4, style_texto)
    story.append(paragrafo4)
    story.append(Spacer(1, 2*cm))
    
    # Data e local (alinhado à direita)
    local_data = Paragraph(f"{aluno_data['CIDADE']}, {data_extenso}.", style_data)
    story.append(local_data)
    story.append(Spacer(1, 1.5*cm))
    
    # Linha de assinatura (centralizada)
    linha_assinatura = Paragraph("_" * 50, style_assinatura)
    story.append(linha_assinatura)
    
    assinatura = Paragraph("Assinatura do Aluno (a)", style_assinatura)
    story.append(assinatura)
    
    # Gerar PDF
    doc.build(story)
    
    # Obter bytes do PDF
    pdf_bytes = buffer.getvalue()
    buffer.close()
    
    return pdf_bytes, nome_arquivo


def criar_zip_termos(df, ies_padrao=None):
    """
    Cria um arquivo ZIP com todos os termos gerados
    
    Args:
        df: DataFrame com os dados dos alunos
        ies_padrao: IES padrão caso não exista coluna IES
    
    Returns:
        bytes do arquivo ZIP
    """
    # Criar buffer para o ZIP
    zip_buffer = io.BytesIO()
    
    # Criar arquivo ZIP
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        total = len(df)
        sucesso = 0
        erros = []
        
        # Progress bar do Streamlit
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        for index, row in df.iterrows():
            try:
                aluno_data = row.to_dict()
                
                # Definir IES para este aluno
                if 'IES' in df.columns:
                    ies = str(row['IES']).strip().upper()
                else:
                    ies = ies_padrao
                
                # Validar IES
                if ies not in IES_CONFIG:
                    raise ValueError(f"IES '{ies}' inválida")
                
                # Gerar PDF
                pdf_bytes, nome_arquivo = gerar_termo_pdf_bytes(aluno_data, ies)
                
                # Adicionar ao ZIP
                zip_file.writestr(nome_arquivo, pdf_bytes)
                
                sucesso += 1
                status_text.text(f"✓ {nome_arquivo}")
                
            except Exception as e:
                erro_msg = f"Erro na linha {index+1} ({row.get('NOME', 'Nome não encontrado')}): {str(e)}"
                erros.append(erro_msg)
                status_text.text(f"✗ {erro_msg}")
            
            # Atualizar progress bar
            progress_bar.progress((index + 1) / total)
        
        progress_bar.empty()
        status_text.empty()
        
        return zip_buffer.getvalue(), sucesso, erros


def main():
    """Aplicação Streamlit"""
    
    # Configuração da página
    st.set_page_config(
        page_title="Gerador de Termos PDF",
        page_icon="📄",
        layout="wide"
    )
    
    # Título
    st.title("📄 Gerador de Termos em PDF - Processamento em Lote")
    st.markdown("---")
    
    # Informações
    st.info("🎓 **Instituições suportadas:** UNIANDRADE | UNIB | UNISMG")
    
    # Upload do arquivo
    st.subheader("1️⃣ Fazer Upload da Planilha")
    uploaded_file = st.file_uploader(
        "Escolha um arquivo CSV ou Excel",
        type=['csv', 'xlsx', 'xls'],
        help="A planilha deve conter as colunas: NOME, CPF, RUA, BAIRRO, CIDADE, UF, CURSO (e opcionalmente IES)"
    )
    
    if uploaded_file is not None:
        try:
            # Ler arquivo
            if uploaded_file.name.endswith('.csv'):
                df = pd.read_csv(uploaded_file)
            else:
                df = pd.read_excel(uploaded_file)
            
            # ✅ PADRONIZAR COLUNAS (SOLUÇÃO DO KEYERROR)
            df.columns = df.columns.str.strip().str.upper()
            
            # Validar colunas necessárias
            colunas_necessarias = ['NOME', 'CPF', 'RUA', 'BAIRRO', 'CIDADE', 'UF', 'CURSO']
            colunas_faltando = [col for col in colunas_necessarias if col not in df.columns]
            
            if colunas_faltando:
                st.error(f"❌ Colunas faltando na planilha: **{', '.join(colunas_faltando)}**")
                st.info("**Colunas necessárias:** NOME, CPF, RUA, BAIRRO, CIDADE, UF, CURSO, IES (opcional)")
                return
            
            # Verificar se tem coluna IES
            tem_coluna_ies = 'IES' in df.columns
            
            # Mostrar preview
            st.success(f"✅ Arquivo carregado com sucesso! **{len(df)} aluno(s)** encontrado(s).")
            
            with st.expander("📋 Visualizar dados da planilha"):
                st.dataframe(df.head(10))
            
            st.markdown("---")
            
            # Seleção de IES (se não houver coluna IES)
            ies_padrao = None
            
            if not tem_coluna_ies:
                st.subheader("2️⃣ Selecionar Instituição (IES)")
                st.warning("⚠️ A planilha não possui coluna 'IES'. Selecione a instituição para todos os alunos:")
                
                opcao_ies = st.radio(
                    "Escolha a IES:",
                    options=[
                        "UNIANDRADE - Centro Universitário Campos de Andrade",
                        "UNIB - Universidade Ibirapuera",
                        "UNISMG - Centro Universitário Santa Maria da Glória"
                    ],
                    index=0
                )
                
                # Mapear opção para código da IES
                if "UNIANDRADE" in opcao_ies:
                    ies_padrao = "UNIANDRADE"
                elif "UNIB" in opcao_ies:
                    ies_padrao = "UNIB"
                else:
                    ies_padrao = "UNISMG"
                
                st.info(f"✓ IES selecionada: **{ies_padrao}**")
            else:
                st.success("✅ Planilha contém coluna 'IES' - usando IES individual para cada aluno")
                
                # Mostrar distribuição de IES
                with st.expander("📊 Distribuição por IES"):
                    ies_counts = df['IES'].value_counts()
                    st.bar_chart(ies_counts)
            
            st.markdown("---")
            
            # Botão para gerar termos
            st.subheader("3️⃣ Gerar Termos em PDF")
            
            if st.button("🚀 Gerar PDFs e Baixar ZIP", type="primary", use_container_width=True):
                with st.spinner("⏳ Gerando termos... Por favor, aguarde."):
                    try:
                        # Criar ZIP com todos os termos
                        zip_bytes, sucesso, erros = criar_zip_termos(df, ies_padrao)
                        
                        # Resumo
                        st.markdown("---")
                        st.subheader("📊 Resumo da Geração")
                        
                        col1, col2, col3 = st.columns(3)
                        with col1:
                            st.metric("Total de Alunos", len(df))
                        with col2:
                            st.metric("✅ Sucesso", sucesso)
                        with col3:
                            st.metric("❌ Erros", len(erros))
                        
                        # Mostrar erros, se houver
                        if erros:
                            with st.expander("⚠️ Ver detalhes dos erros"):
                                for erro in erros:
                                    st.error(erro)
                        
                        # Botão de download do ZIP
                        if sucesso > 0:
                            st.markdown("---")
                            st.success("🎉 **Termos gerados com sucesso!**")
                            
                            # Nome do arquivo ZIP
                            data_hoje = datetime.now().strftime("%Y%m%d_%H%M%S")
                            nome_zip = f"termos_{data_hoje}.zip"
                            
                            st.download_button(
                                label="📥 Baixar ZIP com todos os PDFs",
                                data=zip_bytes,
                                file_name=nome_zip,
                                mime="application/zip",
                                use_container_width=True
                            )
                        
                    except Exception as e:
                        st.error(f"❌ Erro ao gerar termos: {str(e)}")
        
        except Exception as e:
            st.error(f"❌ Erro ao ler arquivo: {str(e)}")
    
    else:
        # Instruções quando não há arquivo
        st.markdown("---")
        st.subheader("📝 Instruções")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("""
            **Formato da Planilha:**
            
            Sua planilha deve conter as seguintes colunas:
            - `NOME` - Nome completo do aluno
            - `CPF` - CPF do aluno
            - `RUA` - Endereço (rua/avenida)
            - `BAIRRO` - Bairro
            - `CIDADE` - Cidade
            - `UF` - Estado (sigla)
            - `CURSO` - Nome do curso
            - `IES` *(opcional)* - UNIANDRADE, UNIB ou UNISMG
            
            **Nota:** Os nomes das colunas não precisam estar em maiúsculas - o sistema padroniza automaticamente!
            """)
        
        with col2:
            st.markdown("""
            **Sobre as IES:**
            
            - **UNIANDRADE** - Centro Universitário Campos de Andrade
            - **UNIB** - Universidade Ibirapuera
            - **UNISMG** - Centro Universitário Santa Maria da Glória
            
            Se a planilha não tiver coluna `IES`, você poderá selecionar uma instituição padrão para todos os alunos.
            
            **Resultado:** Um arquivo ZIP contendo todos os PDFs gerados com os termos de cada aluno.
            """)
        
        # Exemplo de planilha
        st.markdown("---")
        st.subheader("💡 Exemplo de Planilha")
        
        exemplo_df = pd.DataFrame({
            'NOME': ['João da Silva', 'Maria Santos'],
            'CPF': ['12345678901', '98765432100'],
            'RUA': ['Rua das Flores, 123', 'Av. Brasil, 456'],
            'BAIRRO': ['Centro', 'Jardins'],
            'CIDADE': ['São Paulo', 'Curitiba'],
            'UF': ['SP', 'PR'],
            'CURSO': ['Engenharia Civil', 'Administração'],
            'IES': ['UNIB', 'UNIANDRADE']
        })
        
        st.dataframe(exemplo_df)
    
    # Footer
    st.markdown("---")
    st.markdown(
        "<div style='text-align: center; color: gray;'>"
        "Desenvolvido para geração automática de termos | "
        "Suporte: UNIANDRADE, UNIB, UNISMG"
        "</div>",
        unsafe_allow_html=True
    )


if __name__ == "__main__":
    main()
