from pyzabbix import ZabbixAPI
from datetime import datetime
import os
from termcolor import colored
import getpass


os.system('cls')


def banner():
    print(colored('''
 .----------------.  .----------------.  .----------------.  .----------------.  .----------------.  .----------------. 
| .--------------. || .--------------. || .--------------. || .--------------. || .--------------. || .--------------. |
| |   ________   | || |      __      | || |   ______     | || |   ______     | || |     _____    | || |  ____  ____  | |
| |  |  __   _|  | || |     /  \     | || |  |_   _ \    | || |  |_   _ \    | || |    |_   _|   | || | |_  _||_  _| | |
| |  |_/  / /    | || |    / /\ \    | || |    | |_) |   | || |    | |_) |   | || |      | |     | || |   \ \  / /   | |
| |     .'.' _   | || |   / ____ \   | || |    |  __'.   | || |    |  __'.   | || |      | |     | || |    > `' <    | |
| |   _/ /__/ |  | || | _/ /    \ \_ | || |   _| |__) |  | || |   _| |__) |  | || |     _| |_    | || |  _/ /'`\ \_  | |
| |  |________|  | || ||____|  |____|| || |  |_______/   | || |  |_______/   | || |    |_____|   | || | |____||____| | |
| |              | || |              | || |              | || |              | || |              | || |              | |
| '--------------' || '--------------' || '--------------' || '--------------' || '--------------' || '--------------' |
 '----------------'  '----------------'  '----------------'  '----------------'  '----------------'  '----------------' 
    ''', 'red', attrs=['bold']))
    print("")


def banner_vader():
    print("")
    print("           _.-'~~~~~~`-._           ")
    print("          /      ||      \          ")
    print("         /       ||       \         ")
    print("        |        ||        |        ")
    print("        | _______||_______ |        ")
    print("        |/ ----- \/ ----- \|        ")
    print("       /  (     )  (     )  \       ")
    print("      / \  ----- () -----  / \      ")
    print("     /   \      /||\      /   \     ")
    print("    /     \    /||||\    /     \    ")
    print("   /       \  /||||||\  /       \   ")
    print("  /_        \o========o/        _\  ")
    print("    `--...__|`-._  _.-'|__...--'    ")
    print("            |    `'    |            ")
    print("")


def info():
    print(colored('- Bem-vindo ao ZABBIX Consulta de Dados Recentes!\n'
                  '- A extração é realizada por ItemID\n'
                  '- Desenvolvido por Aldori Junior\n'
                  '- Dúvidas/Sugestões envie e-mail para aldorisantosjunior@hotmail.com', 'red', attrs=['bold']))

banner()
banner_vader()
info()

print(colored('\nExemplo de URL do Zabbix: http://zabbixserver.example.com', 'yellow', attrs=['bold']))
server = input('\nInforme a URL de conexão com o seu Zabbix: ')
username = input('\nInforme o usuário: ')
password = getpass.getpass('\nInforme a senha: ')

#server = 'http://zabbixserver.com.br'
#username = 'user'
#password = 'pass'
ZABBIX_SERVER = server
TIMEOUT = None
zapi = ZabbixAPI(ZABBIX_SERVER, timeout=TIMEOUT)
zapi.login(username, password)
print("Conectado no Zabbix API. Versao: %s" % zapi.api_version())

os.system('cls')
banner()
info()

print(colored("--- Escolha o Grupo ---", 'yellow', attrs=['bold']))
print(colored('\nID       |   Grupo de Host', 'yellow', attrs=['bold']))
print('--------- ---------------------------------')

group_id = zapi.hostgroup.get(output='extend', sortfield='groupid')
for a in group_id:
    print('ID: {}   |   Nome: {}'.format(a['groupid'],  a['name']))

hostgroups = input('\nInforme o cliente (Group) ID: ')

os.system('cls')
banner()
info()

print(colored("--- Escolha o Host ---", 'yellow', attrs=['bold']))
print(colored('\nID       |   Host', 'yellow', attrs=['bold']))
print('--------- ---------------------------------')

host_id = zapi.host.get(output='extend', groupids=hostgroups)
for a in host_id:
    print('ID: {}   |   Nome: {}'.format(a['hostid'],  a['host']))

hostid = input('\nInforme o HostID: ')

os.system('cls')
banner()
info()

print(colored("--- Escolha o Item ---", 'yellow', attrs=['bold']))
print(colored('\nID       |   Host', 'yellow', attrs=['bold']))
print('--------- ---------------------------------')

item_id = zapi.item.get(output='extend', hostids=hostid)
for a in item_id:
    print('ID: {}   |   Nome: {}'.format(a['itemid'],  a['name']))

itemid = input('\nInforme o ItemID: ')

print(colored("Infome o período de consulta", 'yellow', attrs=['bold']))
print(colored('O formato de data deve ser DD/MM/YYYY HH:MM:SS', 'yellow', attrs=['bold']))
print('')

data_inicio = input('Informe a data de início: ')
data_fim = input('Informe a data de fim: ')

inicio = int(datetime.strptime(data_inicio, '%d/%m/%Y %H:%M:%S').timestamp()) #Converte da data humana para TimeStamp
fim = int(datetime.strptime(data_fim, '%d/%m/%Y %H:%M:%S').timestamp()) #Converte da data humana para TimeStamp

# Get history entre as datas específicadas
history = zapi.history.get(itemids=itemid, time_from=inicio, time_till=fim, output='extend', limit=500)

# Print
for point in history:
    print("{0}: {1}".format(datetime.fromtimestamp(int(point['clock'])).strftime("%Y-%m-%d %X"), point['value']))

print("fim")