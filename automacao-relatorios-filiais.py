import os
import smtplib
import pandas as pd
from email.message import EmailMessage
from datetime import datetime

# Configurações do servidor SMTP
SMTP_SERVER = 'smtp.seuservidor.com'
SMTP_PORT = 587
EMAIL_USER = 'seu.email@dominio.com'
EMAIL_PASS = 'sua_senha'

# Definindo as pastas de entrada e saída
PASTA_ENTRADA_FATURADOS = 'C:/relatorios/entrada_faturados'
PASTA_ENTRADA_INVENTARIO = 'C:/relatorios/entrada_inventario'
PASTA_SAIDA_FATURADOS = 'C:/relatorios/saida_faturados/'
PASTA_SAIDA_INVENTARIO = 'C:/relatorios/saida_inventario/'

# Garantindo que as pastas de saída existam
os.makedirs(PASTA_SAIDA_FATURADOS, exist_ok=True)
os.makedirs(PASTA_SAIDA_INVENTARIO, exist_ok=True)

# Definir colunas a serem removidas (aplicável ao faturamento)
columns_to_remove_faturados = [
    'Localidade', 'Cliente Final', 'Número do Pedido', 'Forma Pagamento', 
    'Banco', 'Valor Total', 'Parcelas', 
    '% Comissão', 'Valor Comissão', '% Desconto', 
    'Desconto Total', 'Itens Opcionais'
]

# Mapeamento de filiais (genérico para futuros ajustes)
email_map = {
    "Filial_01": {"to": [], "cc": []},
    "Filial_02": {"to": [], "cc": []},
    "Filial_03": {"to": [], "cc": []},
    "Filial_04": {"to": [], "cc": []},
    "Filial_05": {"to": [], "cc": []},
    "Filial_06": {"to": [], "cc": []},
    "Filial_07": {"to": [], "cc": []},
    "Filial_08": {"to": [], "cc": []},
    "Filial_09": {"to": [], "cc": []}
}

def enviar_email(destinatarios, copia, arquivos, filial):
    msg = EmailMessage()
    msg['Subject'] = f'Relatório {filial} - Inventário e Faturamento'
    msg['From'] = EMAIL_USER
    msg['To'] = ', '.join(destinatarios)
    if copia:
        msg['Cc'] = ', '.join(copia)

    msg.set_content(
        f"Bom dia,\n\n"
        f"Segue anexo o relatório de inventário e faturamento referente à filial {filial}.\n\n"
        f"Atenciosamente,\nEquipe de Automação"
    )

    for arquivo_path in arquivos:
        with open(arquivo_path, 'rb') as f:
            msg.add_attachment(f.read(), maintype='application', subtype='xlsx', filename=os.path.basename(arquivo_path))

    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(EMAIL_USER, EMAIL_PASS)
            server.send_message(msg)
            print(f"E-mail enviado para {filial} com sucesso.")
    except Exception as e:
        print(f"Erro ao enviar o e-mail para {filial}: {e}")

def processar_planilha(arquivo_entrada, tipo):
    try:
        df = pd.read_excel(arquivo_entrada)
    except Exception as e:
        print(f"Erro ao ler o arquivo {arquivo_entrada}: {e}")
        return

    for coluna in df.select_dtypes(include=['datetime']).columns:
        df[coluna] = pd.to_datetime(df[coluna], errors='coerce')

    if tipo == 'faturados':
        df = df.drop(columns=[col for col in columns_to_remove_faturados if col in df.columns])
        coluna_filial = 'Entrega Local'
        pasta_saida = PASTA_SAIDA_FATURADOS
        sufixo = "FATURADOS"
    else:
        coluna_filial = 'Venda Local'
        pasta_saida = PASTA_SAIDA_INVENTARIO
        sufixo = "INVENTARIO"

    df_conjunto = df[df[coluna_filial].isin(['Filial_01', 'Filial_02'])]
    output_file_path_conjunto = f"{pasta_saida}Filial_01_02 {sufixo}.xlsx"

    if not df_conjunto.empty:
        try:
            with pd.ExcelWriter(output_file_path_conjunto, engine='xlsxwriter') as writer:
                df_conjunto.to_excel(writer, index=False, sheet_name='Dados')
            print(f"Arquivo conjunto Filial_01_02 salvo em: {output_file_path_conjunto}")
        except Exception as e:
            print(f"Erro ao salvar o arquivo {output_file_path_conjunto}: {e}")

    for filial in df[coluna_filial].unique():
        if filial not in ['Filial_01', 'Filial_02']:
            df_filtrado = df[df[coluna_filial] == filial]
            output_file_path = f"{pasta_saida}{filial} {sufixo}.xlsx"

            try:
                with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer:
                    df_filtrado.to_excel(writer, index=False, sheet_name='Dados')
                print(f"Arquivo individual salvo em: {output_file_path}")
            except Exception as e:
                print(f"Erro ao salvar o arquivo {output_file_path}: {e}")

def confirmar_envio_email():
    for filial in email_map.keys():
        pasta_faturados = PASTA_SAIDA_FATURADOS
        pasta_inventario = PASTA_SAIDA_INVENTARIO

        arquivo_faturados = os.path.join(pasta_faturados, f"{filial} FATURADOS.xlsx")
        arquivo_inventario = os.path.join(pasta_inventario, f"{filial} INVENTARIO.xlsx")

        arquivos = []
        if os.path.exists(arquivo_faturados):
            arquivos.append(arquivo_faturados)
        if os.path.exists(arquivo_inventario):
            arquivos.append(arquivo_inventario)

        if arquivos:
            destinatarios = email_map[filial]["to"]
            copia = email_map[filial]["cc"]
            enviar_email(destinatarios, copia, arquivos, filial)

def main():
    for arquivo in os.listdir(PASTA_ENTRADA_FATURADOS):
        if arquivo.endswith('.xlsx'):
            caminho_entrada = os.path.join(PASTA_ENTRADA_FATURADOS, arquivo)
            processar_planilha(caminho_entrada, tipo='faturados')

    for arquivo in os.listdir(PASTA_ENTRADA_INVENTARIO):
        if arquivo.endswith('.xlsx'):
            caminho_entrada = os.path.join(PASTA_ENTRADA_INVENTARIO, arquivo)
            processar_planilha(caminho_entrada, tipo='inventario')

    confirmar_envio_email()

if __name__ == "__main__":
    main()
