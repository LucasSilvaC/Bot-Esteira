import pandas as pd
from openpyxl import load_workbook
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

# Contadores para cada situação em cada esteira
contagem_estoque = {
    'esteira1': {'baixo': 0, 'medio': 0, 'alto': 0},
    'esteira2': {'baixo': 0, 'medio': 0, 'alto': 0},
    'esteira3': {'baixo': 0, 'medio': 0, 'alto': 0}
}

# Função para processar cada esteira
def processar_esteira(esteira, numero_esteira):
    # Contadores de cada nível de estoque
    contagem_local = {'baixo': 0, 'medio': 0, 'alto': 0}
    print(f"Valores Esteira {numero_esteira}:")
    
    for valor in esteira:
        if valor is None:
            continue
        if valor == 1:
            print(f"Esteira {numero_esteira}, Estoque baixo, nível crítico")
            contagem_local['baixo'] += 1
        elif valor == 2:
            print(f"Esteira {numero_esteira},{valor} Estoque médio, planejar")
            contagem_local['medio'] += 1
        elif valor == 3:
            print(f"Esteira {numero_esteira},{valor} Estoque cheio, sem necessidade de planejamento")
            contagem_local['alto'] += 1
        else:
            print(f"Esteira {numero_esteira},{valor} Valor fora de rotação da esteira")
    
    # Atualizando os totais das esteiras
    for nivel in contagem_local:
        contagem_estoque[f'esteira{numero_esteira}'][nivel] += contagem_local[nivel]

# Carregar a planilha existente
df = load_workbook('Esp8266_Receiver.xlsx')
planilha = df.active

# Extraindo os valores das colunas das esteiras
esteira1 = [celula.value for celula in planilha['C']]
esteira2 = [celula.value for celula in planilha['D']]
esteira3 = [celula.value for celula in planilha['E']]

# Processando cada esteira
processar_esteira(esteira1, 1)
processar_esteira(esteira2, 2)
processar_esteira(esteira3, 3)

# Exibindo os totais
print("\nTotais de estoque por esteira:")
for esteira in contagem_estoque:
    print(f"{esteira}:")
    for nivel in contagem_estoque[esteira]:
        quantidade = contagem_estoque[esteira][nivel]
        print(f"  {nivel}: {quantidade}")

# Preparando os dados para o novo relatório
dados_relatorio = {
    "Esteira": ['Esteira 1', 'Esteira 2', 'Esteira 3'],
    "Estoque Baixo": [contagem_estoque['esteira1']['baixo'], 
                      contagem_estoque['esteira2']['baixo'], 
                      contagem_estoque['esteira3']['baixo']],
    "Estoque Médio": [contagem_estoque['esteira1']['medio'], 
                      contagem_estoque['esteira2']['medio'], 
                      contagem_estoque['esteira3']['medio']],
    "Estoque Alto": [contagem_estoque['esteira1']['alto'], 
                     contagem_estoque['esteira2']['alto'], 
                     contagem_estoque['esteira3']['alto']]
}

# Criando o DataFrame do relatório
df_relatorio = pd.DataFrame(dados_relatorio)

# Salvando o relatório em uma planilha Excel
relatorio_filename = "relatorio.xlsx"
df_relatorio.to_excel(relatorio_filename, index=False)

print("Planilha 'relatorio.xlsx' criada com sucesso.")

# Configuração do servidor SMTP
server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
server.login("lcjade76@gmail.com", "tzor lsyo ecst dfkz")

# Preparando a mensagem
msg = MIMEMultipart()
msg['From'] = "lcjade76@gmail.com"
msg['To'] = "burocraciaplays12@gmail.com"
msg['Subject'] = "Relatório de Estoque das Esteiras"

# Corpo do e-mail
body = "Olá, \nSegue em anexo o relatório de estoque das esteiras."
msg.attach(MIMEText(body, 'plain', 'utf-8'))

# Adicionando o arquivo Excel como anexo
with open(relatorio_filename, "rb") as attachment:
    part = MIMEApplication(attachment.read(), Name=relatorio_filename)
    part['Content-Disposition'] = f'attachment; filename="{relatorio_filename}"'
    msg.attach(part)

# Enviando o e-mail
server.sendmail("lcjade76@gmail.com", "burocraciaplays12@gmail.com", msg.as_string())

# Fechando a conexão com o servidor SMTP
server.quit()

print("E-mail enviado com sucesso.")