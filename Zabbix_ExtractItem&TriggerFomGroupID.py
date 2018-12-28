import os
from pyzabbix import ZabbixAPI
from termcolor import colored
from xlwt import Workbook
import getpass
import pyprind


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
    print(colored('- Bem-vindo ao ZABBIX Extração de Itens e Triggers!\n'
                  '- A extração é realizada por Grupo de Hosts\n'
                  '- Desenvolvido por Aldori Junior\n'
                  '- Dúvidas/Sugestões envie e-mail para aldorisantosjunior@hotmail.com', 'red', attrs=['bold']))

banner()
banner_vader()
info()

print(colored('\nExemplo de URL do Zabbix: http://zabbixserver.example.com', 'yellow', attrs=['bold']))
server = input('\nInforme a URL de conexão com o seu Zabbix: ')
username = input('\nInforme o usuário: ')
password = getpass.getpass('\nInforme a senha: ')

#server = 'http://zabbixserver.com.br
#server = 'http://192.168.0.1
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

print(colored("--- Escolha uma opção do menu ---", 'yellow', attrs=['bold']))
print(colored('\nID       |   Grupo de Host', 'yellow', attrs=['bold']))
print('--------- ---------------------------------')

group_id = zapi.hostgroup.get(output='extend', sortfield='groupid')
for a in group_id:
    print('ID: {}   |   Nome: {}'.format(a['groupid'],  a['name']))

hostgroups = input('\nInforme o cliente (Group) ID: ')
path = input('\nDiretório de exportação, exemplo: C:\documentos\ \nInforme o local onde deseja salvaro o arquivo: ')
group = zapi.hostgroup.get(output='extend', groupids=hostgroups)
for point in group:
    GroupName = point[u'name']
    nomearquivo = (GroupName + ".xls")
    os.system('cls')
    banner()
    info()

print("")
print('Extraindo itens e triggers do GroupID', hostgroups, 'para o arquivo', nomearquivo, "...")
print('')


wb = Workbook()
sheet1 = wb.add_sheet('Sheet 1')
sheet1.write(0, 0, 'HostID')
sheet1.write(0, 1, 'HostName')
sheet1.write(0, 2, 'ItemID')
sheet1.write(0, 3, 'Chave')
sheet1.write(0, 4, 'Nome do item')
sheet1.write(0, 5, 'Status do Item')
sheet1.write(0, 6, 'TriggerID')
sheet1.write(0, 7, 'Descrição da Trigger')
sheet1.write(0, 8, 'Expressão da Trigger')
sheet1.write(0, 9, 'Severidade')
sheet1.write(0, 10, 'Status da Trigger')
row = 1

host = zapi.host.get(output=['hostid', 'host'], groupids=hostgroups)
for x in pyprind.prog_percent(host):
    hostid=x[u'hostid']
    hostname=x[u'host']
    for y in hostid:
        item = zapi.item.get(output=['key_','name','status','itemid'], hostids=hostid, selectTriggers=['extended'], selectApplications=['extended'])
    for w in item:
        key = w[u'key_']
        NameItem = w[u'name']
        StatusItem = w[u'status']
        ItemId = w[u'itemid']
        trigger = w[u'triggers']
        application = w['applications']
        if trigger:
            for b in trigger:
                triggerid = b[u'triggerid']
                for c in triggerid:
                    triggerfilter = zapi.trigger.get(output=['description', 'expression', 'status', 'priority'], triggerids=triggerid, itemids=ItemId, expandExpression=True, expandDescription=True)
                for a in triggerfilter:
                    descricao = a[u'description']
                    expressao = a[u'expression']
                    statusTrigger = a[u'status']
                    severidade = a[u'priority']
                    sheet1.write(row, 0, hostid)
                    sheet1.write(row, 1, hostname)
                    sheet1.write(row, 2, ItemId)
                    sheet1.write(row, 3, key)
                    sheet1.write(row, 4, NameItem)
                    sheet1.write(row, 5, StatusItem)
                    sheet1.write(row, 6, triggerid)
                    sheet1.write(row, 7, descricao)
                    sheet1.write(row, 8, expressao)
                    sheet1.write(row, 9, severidade)
                    sheet1.write(row, 10, statusTrigger)
                    row += 1
                    wb.save(path+nomearquivo)
        else:
            sheet1.write(row, 0, hostid)
            sheet1.write(row, 1, hostname)
            sheet1.write(row, 2, ItemId)
            sheet1.write(row, 3, key)
            sheet1.write(row, 4, NameItem)
            sheet1.write(row, 5, StatusItem)
            row += 1
            wb.save(path+nomearquivo)