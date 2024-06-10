# Gerais
bSendSFTP = False
bSendEmail = True # Envia e-mail 
bValidaQuery = False 
bTelemetria = False
bForce = False # Executa de Qualquer jeito
# Planilha
sNomePlanilha = "Apuração de Ponto.xlsx"
sNomePlanilhaEquipe = "Apuração de Ponto Equipe.xlsx"
iDiasBloqueio = 0

# Diretórios

sDiretorioQueries = '.\\queries\\'
sDiretorioOrigem = '.\\arqs\\'
sLogDirectory = '.\\logs\\'
sDiretorioValidator = '.\\validator\\'
sExcelBase = '.\\ExcelEmails\\'


# Oracle

username = 'DWSFTP'
password = '1NT3GR4@5F7P'
host = 'ndw.redeoba.com.br'
port = 1521
service = 'dw'
dsn = 'ndw.redeoba.com.br/dw'
encoding = 'UTF-8'
dialet = 'oracle'
sql_driver = 'cx_oracle'

# Oracle Telemetria

usernameTelemetria = 'Telemetria'
passwordTelemetria = '9#B7I5q!Dq'
hostTelemetria = 'tlmprd.redeoba.com.br'
serviceTelemetria = 'tlmprd'
dsnTelemetria = 'tlmprd.redeoba.com.br/tlmprd'
codEntradaTelemetria = 20
sObjetivoTelemetria = 'INCONSISTÊNCIA PONTO RH'

# # Oracle Telemetria Homologaão

# usernameTelemetria = 'Telemetria'
# passwordTelemetria = '9#B7I5q!Dq'
# hostTelemetria = 'tlmhom.redeoba.com.br'
# serviceTelemetria = 'tlmhom'
# dsnTelemetria = 'tlmhom.redeoba.com.br/tlmhom'
# codEntradaTelemetria = 20
# sObjetivoTelemetria = 'INCONSISTÊNCIA PONTO RH'

# SFTP

sHOST = 'sftp.redeoba.com.br'
sUSER = ''
sPASSWORD = ''
iPORT = 6591
iPerRetencao = 15
sDiretorioDefault = './dados/'
# SMTP

sSMTP = 'smtp.office365.com'
iSMTPport = 587
sCertfile='office365.cer'
sSMTPUserName = 'itsmadm@redeoba.com.br'
sSMTPUserPWD = 'SM@0b@$!t04972'
sMailSender = 'itsmadm@redeoba.com.br'
sMailDestino =''
sMailSender = 'itsmadm@redeoba.com.br'

sMailSubject = "Apontamento de Horas Geral"
sDirEmails = '\\\\SRV-VDAPROCESSO\\excelemails\\ListaDistribuição.xlsx'
# sDirEmails = '.\\ExcelEmails\\ListaDistribuição.xlsx'
sEmailTeste = "" # Deixe em branco para enviar ao destinatário correto


# Execução - True - Não valida manda tudo, False - Valida o período


dDataHojeType = 'P'
dDataHojeDays = 8
dDataHoje = '' #'2022_08_24'

# Log
sLogFile = "Ocorrencias "

# Time to wait em seconds
sWaitTime = 600