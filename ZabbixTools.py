from xlwt import Workbook
import pyprind
from pyzabbix import ZabbixAPI
from datetime import datetime
#import datetime
import time
import os
import sys
import getpass
import telnetlib
import socket
from termcolor import colored
from progressbar import ProgressBar, Percentage, ReverseBar, ETA, Timer, RotatingMarker

os.system('cls')


def banner():
    print(colored('''
     ______________________________________________________________
    |                                                              |
    |    ########      ##      ######   ######   ##  ###    ###    |
    |         ##     ##  ##    ##   ##  ##   ##  ##   ##    ##     | 
    |        ##     ##    ##   ##   ##  ##   ##  ##    ##  ##      |
    |       ##     ##########  #######  #######  ##      ##        | 
    |      ##      ##      ##  ##   ##  ##   ##  ##    ##  ##      |
    |     ##       ##      ##  ##   ##  ##   ##  ##   ##    ##     |
    |    ########  ##      ##  ######   ######   ##  ###    ###    |
    |______________________________________________________________|
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
    print(colored(' - Bem-vindo ao ZABBIX Tools!\n'
                  ' - Desenvolvido por Aldori Junior\n'
                  ' - Dúvidas/Sugestões envie e-mail para aldorisantosjunior@hotmail.com', 'red', attrs=['bold']))


def typeacess():
    global server

    URL = "ZabbixServer.com.br"  ## Informe aqui a URL do seu Zabbix!
    TEMPOTIMEOUT = 3

    try:
        socket.gethostbyname('ZabbixServer.com.br')  # Testa se a URL é válida/publicada
        try:
            port80 = telnetlib.Telnet(URL, 10051, TEMPOTIMEOUT)
            server = "http://ZabbixServer.com.br"
            infologin(server)
        except socket.timeout:
            port443 = telnetlib.Telnet(URL, 443, TEMPOTIMEOUT)
            server = "https://ZabbixServer.com.br"
            infologin(server)


    except:
        os.system('cls')
        banner()
        print(" Não foi possível conectar com o ZabbixServer!")
        exit(0)


def infologin(server):
    global username, password, TIMEOUT, ZABBIX_SERVER
    username = input('\n Informe o usuário: ')
    password = getpass.getpass(' Informe a senha: ')
    ZABBIX_SERVER = server
    TIMEOUT = None
    connect(username, password, TIMEOUT, ZABBIX_SERVER)


def connect(username, password, TIMEOUT, ZABBIX_SERVER):
    global zapi
    try:
        zapi = ZabbixAPI(server=ZABBIX_SERVER, timeout=TIMEOUT)
        zapi.login(username, password)
        print("")
        print(" Conectado no Zabbix API. Versao: %s" % zapi.api_version())
    except:
        os.system('cls')
        banner()
        print(colored('    Não foi possível conectar ao Zabbix Server.', 'yellow', attrs=['bold']))
        print(u"\n    Verifique se a URL " + colored(ZABBIX_SERVER, 'red', attrs=['bold']) + u" está disponível.")
        print("\n    Verifique se o usuário e senha estão corretos, ou se está conectado a alguma VPN.")
        print("")
        typeacess()


def menugroup():
    os.system('cls')
    banner()
    info()
    ConsultaGroupID()


def menu_opcao():
    print("")
    print(colored(" --- Relatórios ---", 'yellow', attrs=['bold']))
    print("")
    #
    # Itens 1,2,3,4,6 e 7 Extraídos do script "ZabbixTunner" de Janssen Lima - janssenreislima@gmail.com
    #
    print(" [01] - Relatório de itens do sistema")
    print(" [02] - Listar itens não suportados")
    print(" [03] - Listar itens não suportados por GroupID")
    print(" [04] - Desabilitar itens não suportados")
    # print(" [04] - Relatório da média de coleta dos itens (por tipo)")
    print(" [05] - Diagnóstico do Zabbix")
    print(" [06] - Relatório de Agentes Zabbix desatualizados")
    print(" [07] - Relatório de Triggers por tempo de alarme e estado  #Em avaliação!")
    print("")
    print(colored(" ---  Extração de Dados ---", 'yellow', attrs=['bold']))
    print("")
    print(" [08] - Extração de Itens e Triggers dos hosts do GroupID")
    print(" [09] - Extração de Itens fora de template dos hosts do GroupID")
    print(" [10] - Extração de Cenários WEB do GroupID")
    print(" [11] - Extração de Itens e Triggers de Templates")
    print(" [12] - Extração de Itens e Triggers de Templates com Severidade Desastre e Alta")
    print(" [13] - Extração de Templates vinculados aos Hosts por GroupID")
    print("")
    print(colored(" --- Períodos de Manutenção ---", 'yellow', attrs=['bold']))
    print("")
    print(" [14] - Consultar Períodos de manutenção")
    print(" [15] - Criar períodos de manutenção")
    print(" [16] - Deletar períodos de manutenção")
    print("")
#    print(colored(" --- Ações ---", 'yellow', attrs=['bold']))
#    print("")
#    print(" [17] - Adicionar dependência em triggers")
#    print("")
    print(colored(" --- APOIO ---", 'yellow', attrs=['bold']))
    print("")
    print(" [99] - Consulta de GroupID")
    print(" [0 ] - Sair")
    print("")
    opcao = input(" [+] - Selecione a opção desejada: ")

    if opcao == '1' or opcao == '01':
        dadosItens()
    elif opcao == '2' or opcao == '02':
        listagemItensNaoSuportados()
    elif opcao == '3' or opcao == '03':
        listagemItensNaoSuportadosGroupID()
    elif opcao == '4' or opcao == '04':
        desabilitaItensNaoSuportados()
    elif opcao == '5' or opcao == '05':
        diagnosticoAmbiente()
    elif opcao == '6' or opcao == '06':
        agentesDesatualizados()
    elif opcao == '7' or opcao == '07':
        menu_relack()
    elif opcao == '8' or opcao == '08':
        ExtracItemAndTrigger()
    elif opcao == '9' or opcao == '09':
        ExtractItemForaTemplate()
    elif opcao == '10':
        ExtractCenarioWEB()
    elif opcao == '11':
        ExtractItemAndTriggerTemplate()
    elif opcao == '12':
        ExtractItemAndTriggerTemplateSeveridade()
    elif opcao == '13':
        ExtractTemplateFromHostGroup()
    elif opcao == '14':
        ConsultaManutencao()
    elif opcao == '15':
        CriaManutencao()
    elif opcao == '16':
        DeletaManutencao()
    elif opcao == '17':
        AddDependenciaTrigger()
    elif opcao == '99':
        menugroup()
    elif opcao == '0':
        sys.exit()
    else:
        menu_opcao()


def ConsultaGroupID():
    print("")
    GrupoFilter = input(" Digite o nome do Grupo de Hosts ou pressione ENTER para consultar os grupos disponíveis: ")
    print("")

    if not GrupoFilter:
        GetGroup = zapi.hostgroup.get(
            output='extend',
            sortfield='name'
        )

        for point in GetGroup:
            if 'Template' not in point['name'] and 'TEMPLATE' not in point['name'] and 'machines' not in point['name'] and 'Discovered' not in point['name'] and 'Hypervisors' not in point['name']:
                print(' ID: {}   |   Nome: {}'.format(point['groupid'], point['name']))
        print("")
        input(" Pressione ENTER para voltar ao menu de opções!")
        os.system('cls')
        banner()
        info()
        menu_opcao()

    elif GrupoFilter:
        GetGroup = zapi.hostgroup.get(
            output='extend',
            sortfield='name'
        )

        for point in GetGroup:
            if GrupoFilter in point['name']:
                print(' ID: {}   |   Nome: {}'.format(point['groupid'], point['name']))
        print("")
        input(" Pressione ENTER para voltar ao menu de opções!")
        os.system('cls')
        banner()
        info()
        menu_opcao()

    else:
        print(" Não foi identificado nenhum Grupo de Hosts com o nome informado!")
        input(" Pressione ENTER para voltar ao menu de opção!")
        os.system('cls')
        banner()
        info()
        menu_opcao()


def ExtractItemAndTriggerTemplate():
    wb = Workbook()
    sheet1 = wb.add_sheet('Sheet 1')
    sheet1.write(0, 0, 'Nome do Template')
    sheet1.write(0, 1, 'Nome do Item')
    sheet1.write(0, 2, 'Chave')
    sheet1.write(0, 3, 'Status')
    sheet1.write(0, 4, 'Descrição da Trigger')
    sheet1.write(0, 5, 'Expressão da Trigger')
    sheet1.write(0, 6, 'Severidade')
    sheet1.write(0, 7, 'Status da Trigger')
    row = 1

    print("")
    TemplateFilter = input(" Digite o nome do template ou pressione ENTER para consultar os templates disponíveis: ")
    print("")

    if not TemplateFilter:
        GetTemplate = zapi.template.get(
            output='extend',
            filter='extended',
            sortfield='name'
        )

        for point in GetTemplate:
            print(' ID: {}   |   Nome: {}'.format(point['templateid'], point['host']))
        print("")
        templateid = int(input(" Informe o ID do Template: "))
    elif TemplateFilter:
        GetTemplate = zapi.template.get(
            output='extend',
            filter='extended',
            sortfield='name'
        )

        for point in GetTemplate:
            if TemplateFilter in point['host']:
                print(' ID: {}   |   Nome: {}'.format(point['templateid'], point['host']))
        print("")
        templateid = int(input(" Informe o ID do Template: "))
    else:
        print(" Não foi identificado nenhum template com o nome informado!")
        input(" Pressione ENTER para voltar ao menu de opção!")
        os.system('cls')
        banner()
        info()
        menu_opcao()

    saida = input(" Informe o nome do arquivo de saída: ")
    nomearquivo = saida + '.xls'
    GetTemplateData = zapi.template.get(
        output=['templateid', 'name', 'hostids'],
        templateids=templateid
    )

    for x in GetTemplateData:
        templateid = x[u'templateid']
        templatename = x[u'name']
        for y in templateid:
            item = zapi.item.get(
                output=['key_', 'name', 'status', 'itemid'],
                hostids=templateid
            )

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
                        triggerfilter = zapi.trigger.get(
                            output=['description', 'expression', 'status', 'priority'],
                            triggerids=triggerid,
                            itemids=ItemId,
                            expandExpression=True,
                            expandDescription=True
                        )

                    for a in triggerfilter:
                        descricao = a[u'description']
                        expressao = a[u'expression']
                        statusTrigger = a[u'status']
                        severidade = a[u'priority']
                        sheet1.write(row, 0, templatename)
                        sheet1.write(row, 1, NameItem)
                        sheet1.write(row, 2, key)
                        sheet1.write(row, 3, StatusItem)
                        sheet1.write(row, 4, descricao)
                        sheet1.write(row, 5, expressao)
                        sheet1.write(row, 6, severidade)
                        sheet1.write(row, 7, statusTrigger)
                        row += 1
                        wb.save(nomearquivo)
            else:
                sheet1.write(row, 0, templatename)
                sheet1.write(row, 1, NameItem)
                sheet1.write(row, 2, key)
                sheet1.write(row, 3, StatusItem)
                row += 1
                wb.save(nomearquivo)

    print("")
    input(" Exportação de itens do Template" + templatename + " finalizado, pressione ENTER para voltar ao Menu Opção!")

    os.system('cls')
    banner()
    info()
    menu_opcao()


def ExtractItemAndTriggerTemplateSeveridade():
    wb = Workbook()
    sheet1 = wb.add_sheet('Sheet 1')
    sheet1.write(0, 0, 'Nome do Template')
    sheet1.write(0, 1, 'Nome do Item')
    sheet1.write(0, 2, 'Chave')
    sheet1.write(0, 3, 'Status')
    sheet1.write(0, 4, 'Descrição da Trigger')
    sheet1.write(0, 5, 'Expressão da Trigger')
    sheet1.write(0, 6, 'Severidade')
    sheet1.write(0, 7, 'Status da Trigger')
    row = 1

    saida = input(" Informe o nome do arquivo de saída: ")
    nomearquivo = saida + '.xls'
    GetTemplateData = zapi.template.get(
        output=['templateid', 'name', 'hostids']
        #templateids=templateid
    )

    for x in GetTemplateData:
        templateid = x[u'templateid']
        templatename = x[u'name']
        for y in templateid:
            item = zapi.item.get(
                output=['key_', 'name', 'status', 'itemid'],
                hostids=templateid,
                selectTriggers=['extended'],
                selectApplications=['extended']
            )

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
                        triggerfilter = zapi.trigger.get(
                            output=['description', 'expression', 'status', 'priority'],
                            triggerids=triggerid,
                            itemids=ItemId,
                            expandExpression=True,
                            expandDescription=True
                        )

                    for a in triggerfilter:
                        descricao = a[u'description']
                        expressao = a[u'expression']
                        statusTrigger = a[u'status']
                        severidade = a[u'priority']
                        if severidade == '5' or severidade == '4':
                            sheet1.write(row, 0, templatename)
                            sheet1.write(row, 1, NameItem)
                            sheet1.write(row, 2, key)
                            sheet1.write(row, 3, StatusItem)
                            sheet1.write(row, 4, descricao)
                            sheet1.write(row, 5, expressao)
                            sheet1.write(row, 6, severidade)
                            sheet1.write(row, 7, statusTrigger)
                            row += 1
                            wb.save(nomearquivo)

    print("")
    input(" Exportação de itens do Template" + templatename + " finalizado, pressione ENTER para voltar ao Menu Opção!")

    os.system('cls')
    banner()
    info()
    menu_opcao()


def ExtracItemAndTrigger():

    print("")
    GrupoFilter = input(" Digite o nome do Grupo de Hosts ou pressione ENTER para consultar os grupos disponíveis: ")
    print("")

    if not GrupoFilter:
        GetGroup = zapi.hostgroup.get(
            output='extend',
            sortfield='name'
        )

        for point in GetGroup:
            if 'Template' not in point['name'] and 'TEMPLATE' not in point['name'] and 'machines' not in point['name'] and 'Discovered' not in point['name'] and 'Hypervisors' not in point['name']:
                print(' ID: {}   |   Nome: {}'.format(point['groupid'], point['name']))
        print("")
        hostgroups = int(input(" Informe o ID do Grupo: "))
        arquivosaida = input(" Informe o nome do arquivo de saída: ")
        nomearquivo = arquivosaida + ".xls"

    elif GrupoFilter:
        GetGroup = zapi.hostgroup.get(
            output='extend',
            sortfield='name'
        )

        for point in GetGroup:
            if GrupoFilter in point['name']:
                print(' ID: {}   |   Nome: {}'.format(point['groupid'], point['name']))
        print("")
        hostgroups = int(input(" Informe o ID do Grupo: "))
        arquivosaida = input(" Informe o nome do arquivo de saída: ")
        nomearquivo = arquivosaida + ".xls"

    else:
        print(" Não foi identificado nenhum Grupo de Hosts com o nome informado!")
        input(" Pressione ENTER para voltar ao menu de opção!")
        os.system('cls')
        banner()
        info()
        menu_opcao()

    print("")
    print(' Extraindo itens e triggers do GroupID', hostgroups, 'para o arquivo', nomearquivo, '...')
    print("")

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

    host = zapi.host.get(
        output=['hostid', 'host'],
        groupids=hostgroups
    )

    for x in pyprind.prog_percent(host):
        hostid = x[u'hostid']
        hostname = x[u'host']
        for y in hostid:
            item = zapi.item.get(
                output=['key_', 'name', 'status', 'itemid'],
                hostids=hostid,
                selectTriggers=['extended'],
                selectApplications=['extended']
            )

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
                        triggerfilter = zapi.trigger.get(
                            output=['description', 'expression', 'status', 'priority'],
                            triggerids=triggerid,
                            itemids=ItemId,
                            expandExpression=True,
                            expandDescription=True
                        )

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
                        wb.save(nomearquivo)
            else:
                sheet1.write(row, 0, hostid)
                sheet1.write(row, 1, hostname)
                sheet1.write(row, 2, ItemId)
                sheet1.write(row, 3, key)
                sheet1.write(row, 4, NameItem)
                sheet1.write(row, 5, StatusItem)
                row += 1
                wb.save(nomearquivo)
    print("")
    print(' Extração de itens e triggers do GroupID', hostgroups, 'finalizado!')
    input(" Pressione ENTER para voltar ao Menu de Grupos!")
    os.system('cls')
    banner()
    info()
    menu_opcao()


def ExtractTemplateFromHostGroup():

    print("")
    GrupoFilter = input(" Digite o nome do Grupo de Hosts ou pressione ENTER para consultar os grupos disponíveis: ")
    print("")

    if not GrupoFilter:
        GetGroup = zapi.hostgroup.get(
            output='extend',
            sortfield='name'
        )

        for point in GetGroup:
            if 'Template' not in point['name'] and 'TEMPLATE' not in point['name'] and 'machines' not in point['name'] and 'Discovered' not in point['name'] and 'Hypervisors' not in point['name']:
                print(' ID: {}   |   Nome: {}'.format(point['groupid'], point['name']))
        print("")
        hostgroups = int(input(" Informe o ID do Grupo: "))
        arquivosaida = input(" Informe o nome do arquivo de saída: ")
        nomearquivo = arquivosaida + ".xls"

    elif GrupoFilter:
        GetGroup = zapi.hostgroup.get(
            output='extend',
            sortfield='name'
        )

        for point in GetGroup:
            if GrupoFilter in point['name']:
                print(' ID: {}   |   Nome: {}'.format(point['groupid'], point['name']))
        print("")
        hostgroups = int(input(" Informe o ID do Grupo: "))
        arquivosaida = input(" Informe o nome do arquivo de saída: ")
        nomearquivo = arquivosaida + ".xls"

    else:
        print(" Não foi identificado nenhum Grupo de Hosts com o nome informado!")
        input(" Pressione ENTER para voltar ao menu de opção!")
        os.system('cls')
        banner()
        info()
        menu_opcao()

    print("")
    print(' Extraindo Templates vinculados aos hosts do GroupID', hostgroups, 'para o arquivo', nomearquivo, '...')
    print("")

    wb = Workbook()
    sheet1 = wb.add_sheet('Sheet 1')
    sheet1.write(0, 0, 'Host')
    sheet1.write(0, 1, 'Descrição')
    sheet1.write(0, 2, 'Template')
    row = 1

    host_get = zapi.host.get(output="extend", selectGroups="extend",
                             selectParentTemplates="extend", groupids=hostgroups) # Coleta de informações dos Hosts do Grupo Especificado na variável HostGroups
    for x in host_get:
        tmpl = x[u'parentTemplates']
        host = x[u'host']
        descricao = x['description']
        if tmpl:
            for y in tmpl:
                tname = y[u'name']
                sheet1.write(row, 0, host)
                sheet1.write(row, 1, descricao)
                sheet1.write(row, 2, tname)
                row += 1
                wb.save(nomearquivo)
        else:
            sheet1.write(row, 0, host)
            sheet1.write(row, 1, descricao)
            row += 1
            wb.save(nomearquivo)
    os.system('cls')
    print('\nFinalizado a exportação dos Templates vinculados aos Hosts do Grupo de Host', hostgroups, 'para o arquivo', nomearquivo, '!')
    print("")
    input(" Pressione ENTER para voltar ao Menu de Grupos!")
    os.system('cls')
    banner()
    info()
    menu_opcao()


def ExtractItemForaTemplate():

    print("")
    GrupoFilter = input(" Digite o nome do Grupo de Hosts ou pressione ENTER para consultar os grupos disponíveis: ")
    print("")

    if not GrupoFilter:
        GetGroup = zapi.hostgroup.get(
            output='extend',
            sortfield='name'
        )

        for point in GetGroup:
            if 'Template' not in point['name'] and 'TEMPLATE' not in point['name'] and 'machines' not in point['name'] and 'Discovered' not in point['name'] and 'Hypervisors' not in point['name']:
                print(' ID: {}   |   Nome: {}'.format(point['groupid'], point['name']))
        print("")
        hostgroups = int(input(" Informe o ID do Grupo: "))
        arquivosaida = input(" Informe o nome do arquivo de saída: ")
        nomearquivo = arquivosaida + ".xls"

    elif GrupoFilter:
        GetGroup = zapi.hostgroup.get(
            output='extend',
            sortfield='name'
        )

        for point in GetGroup:
            if GrupoFilter in point['name']:
                print(' ID: {}   |   Nome: {}'.format(point['groupid'], point['name']))
        print("")
        hostgroups = int(input("Informe o ID do Grupo: "))
        arquivosaida = input("Informe o nome do arquivo de saída: ")
        nomearquivo = arquivosaida + ".xls"

    else:
        print(" Não foi identificado nenhum Grupo de Hosts com o nome informado!")
        input(" Pressione ENTER para voltar ao menu de opção!")
        os.system('cls')
        banner()
        info()
        menu_opcao()


    print("")
    print(' Extraindo itens fora de template do GroupID', hostgroups, 'para o arquivo', nomearquivo, '...')
    print("")

    wb = Workbook()
    sheet1 = wb.add_sheet('Sheet 1')
    sheet1.write(0, 0, 'HostID')
    sheet1.write(0, 1, 'HostName')
    sheet1.write(0, 2, 'ItemID')
    sheet1.write(0, 3, 'Nome do Item')
    sheet1.write(0, 4, 'Chave')
    sheet1.write(0, 5, 'Status')
    row = 1

    host = zapi.host.get(
        output=['hostid', 'host'],
        groupids=hostgroups
    )

    for x in pyprind.prog_percent(host):
        hostid = x[u'hostid']
        hostname = x[u'host']
        for y in hostid:
            item = zapi.item.get(
                output=['key_', 'name', 'status', 'itemid'],
                hostids=hostid,
                inherited=False
            )

        for w in item:
            key = w[u'key_']
            NameItem = w[u'name']
            StatusItem = w[u'status']
            ItemId = w[u'itemid']
            sheet1.write(row, 0, hostid)
            sheet1.write(row, 1, hostname)
            sheet1.write(row, 2, ItemId)
            sheet1.write(row, 3, NameItem)
            sheet1.write(row, 4, key)
            sheet1.write(row, 5, StatusItem)
            row += 1
            wb.save(nomearquivo)
    print("")
    print(' Extração de itens fora de template do GroupID', hostgroups, 'finalizado!')
    input(" Pressione ENTER para voltar ao Menu de Grupos!")
    os.system('cls')
    banner()
    info()
    menu_opcao()


def ExtractCenarioWEB():

    print("")
    GrupoFilter = input(" Digite o nome do Grupo de Hosts ou pressione ENTER para consultar os grupos disponíveis: ")
    print("")

    if not GrupoFilter:
        GetGroup = zapi.hostgroup.get(
            output='extend',
            sortfield='name'
        )

        for point in GetGroup:
            if 'Template' not in point['name'] and 'TEMPLATE' not in point['name'] and 'machines' not in point['name'] and 'Discovered' not in point['name'] and 'Hypervisors' not in point['name']:
                print(' ID: {}   |   Nome: {}'.format(point['groupid'], point['name']))
        print("")
        hostgroups = int(input(" Informe o ID do Grupo: "))
        arquivosaida = input(" Informe o nome do arquivo de saída: ")
        nomearquivo = arquivosaida + ".xls"

    elif GrupoFilter:
        GetGroup = zapi.hostgroup.get(
            output='extend',
            sortfield='name'
        )

        for point in GetGroup:
            if GrupoFilter in point['name']:
                print(' ID: {}   |   Nome: {}'.format(point['groupid'], point['name']))
        print("")
        hostgroups = int(input(" Informe o ID do Grupo: "))
        arquivosaida = input(" Informe o nome do arquivo de saída: ")
        nomearquivo = arquivosaida + ".xls"

    else:
        print(" Não foi identificado nenhum Grupo de Hosts com o nome informado!")
        input(" Pressione ENTER para voltar ao menu de opção!")
        os.system('cls')
        banner()
        info()
        menu_opcao()


    print("")
    print(' Extraindo cenários WEB do GroupID', hostgroups, 'para o arquivo', nomearquivo, '...')
    print("")

    wb = Workbook()
    sheet1 = wb.add_sheet('Sheet 1')
    sheet1.write(0, 0, 'HostID')
    sheet1.write(0, 1, 'HostName')
    sheet1.write(0, 2, 'Nome visível')
    sheet1.write(0, 3, 'IP')
    sheet1.write(0, 4, 'Cenário ID')
    sheet1.write(0, 5, 'Nome do Cenário')
    sheet1.write(0, 6, 'Status')
    sheet1.write(0, 7, 'Intervalo atualização (em seg)')
    sheet1.write(0, 8, 'Tentativas')
    sheet1.write(0, 9, 'Proxy')
    sheet1.write(0, 10, 'Usa Autenticação?')
    sheet1.write(0, 11, 'Usuário')
    sheet1.write(0, 12, 'Senha')
    sheet1.write(0, 13, 'Etapa ID')
    sheet1.write(0, 14, 'Nome da Etapa')
    sheet1.write(0, 15, 'URL')
    sheet1.write(0, 16, 'TimeOut')
    sheet1.write(0, 17, 'StatusCode')
    row = 1

    hostids = zapi.host.get(
        output='extend',
        selectInterfaces='extend',
        groupids=hostgroups,
        sortfield='name'
    )

    for a in pyprind.prog_percent(hostids):
        hostid = a[u'hostid']
        hostname = a[u'host']
        nomevisivel = a[u'name']
        interface = a[u'interfaces']
        for b in interface:
            ip = b[u'ip']
        for b in hostid:
            GetStepID = zapi.httptest.get(
                outout='httptestid',
                hostids=hostid
            )

        for c in GetStepID:
            StepId = c[u'httptestid']
            for d in StepId:
                GetCenarioWeb = zapi.httptest.get(
                    outout='extend',
                    selectSteps='extend',
                    expandName=True,
                    expandStepName=True,
                    httptestids=StepId
                )

            for e in GetCenarioWeb:
                CenarioID = e[u'httptestid']
                NomeCenario = e[u'name']
                Etapas = e[u'steps']
                Intervalo = e[u'delay']
                StatusCenario = e[u'status']
                UsaAutenticacao = e[u'authentication']
                Usuario = e[u'http_user']
                Senha = e[u'http_password']
                Tentativas = e[u'retries']
                Proxy = e[u'http_proxy']
                Agente = e[u'agent']
                for f in Etapas:
                    EtapaID = f[u'httpstepid']
                    NomeEtapa = f[u'name']
                    URLEtapa = f[u'url']
                    TimeoutEtapa = f[u'timeout']
                    StatusCode = f[u'status_codes']
                    sheet1.write(row, 0, hostid)
                    sheet1.write(row, 1, hostname)
                    sheet1.write(row, 2, nomevisivel)
                    sheet1.write(row, 3, ip)
                    sheet1.write(row, 4, CenarioID)
                    sheet1.write(row, 5, NomeCenario)
                    sheet1.write(row, 6, StatusCenario)
                    sheet1.write(row, 7, Intervalo)
                    sheet1.write(row, 8, Tentativas)
                    sheet1.write(row, 9, Proxy)
                    sheet1.write(row, 10, UsaAutenticacao)
                    sheet1.write(row, 11, Usuario)
                    sheet1.write(row, 12, Senha)
                    sheet1.write(row, 13, EtapaID)
                    sheet1.write(row, 14, NomeEtapa)
                    sheet1.write(row, 15, URLEtapa)
                    sheet1.write(row, 16, TimeoutEtapa)
                    sheet1.write(row, 17, StatusCode)
                    row += 1
                    wb.save(nomearquivo)
    print("")
    print(' Extração de cenários WEB do GroupID', hostgroups, 'finalizado!')
    input(" Pressione ENTER para voltar ao Menu de Grupos!")
    os.system('cls')
    banner()
    info()
    menu_opcao()


def desabilitaItensNaoSuportados():
    query = {
        "output": "extend",
        "filter": {
            "state": 1
        },
        "monitored": True
    }

    filtro = input(' Qual a busca para key_? [NULL = ENTER] ')
    if filtro.__len__() > 0:
        query['search'] = {'key_': filtro}

    limite = input(' Qual o limite de itens? [NULL = ENTER] ')
    if limite.__len__() > 0:
        try:
            query['limit'] = int(limite)
        except:
            print(' Limite invalido')
            input(" Pressione ENTER para voltar")
            menu_opcao()

    opcao = input(" Confirma operação? [s/n]")
    if opcao == 's' or opcao == 'S':
        itens = zapi.item.get(query)
        print(' Encontramos {} itens'.format(itens.__len__()))
        bar = ProgressBar(maxval=itens.__len__(),
                          widgets=[Percentage(), ReverseBar(), ETA(), RotatingMarker(), Timer()]).start()
        i = 0
        for x in itens:
            result = zapi.item.update({"itemid": x['itemid'], "status": 1})
            i += 1
            bar.update(i)
        bar.finish()
        print(" Itens desabilitados!!!")
    input(" Pressione ENTER para continuar")

    os.system('cls')
    banner()
    info()
    menu_opcao()


def agentesDesatualizados():
    itens = zapi.item.get(
        filter={"key_": "agent.version"},
        output=["lastvalue", "hostid"],
        templated=False,
        selectHosts=["host"],
        sortorder="ASC"
    )

    try:
        versaoZabbixServer = zapi.item.get(
            filter={"key_": "agent.version"},
            output=["lastvalue", "hostid"],
            hostids="10084"
        )[0]["lastvalue"]

        print(colored(' {0:6} | {1:30}'.format("Versão", "Host"), attrs=['bold']))

        for x in itens:
            if x['lastvalue'] != versaoZabbixServer and x['lastvalue'] <= versaoZabbixServer:
                print(' {0:6} | {1:30}'.format(x["lastvalue"], x["hosts"][0]["host"]))
        print("")
        input(" Pressione ENTER para continuar")
        os.system('cls')
        banner()
        info()
        menu_opcao()

    except IndexError:
        print(" Não foi possível obter a versão do agent no Zabbix Server.")
        input(" Pressione ENTER para continuar")

    os.system('cls')
    banner()
    info()
    menu_opcao()


def diagnosticoAmbiente():
    print(colored(" [+++]", 'green'), "analisando itens não númericos")
    itensNaoNumericos = zapi.item.get(
        output="extend",
        monitored=True,
        filter={"value_type": [1, 2, 4]},
        countOutput=True
    )
    print(colored(" [+++]", 'green'), "analisando itens ICMPPING com histórico acima de 7 dias")
    itensPing = zapi.item.get(
        output="extend",
        monitored=True,
        filter={"key_": "icmpping"},
    )

    contPing = 0
    for x in itensPing:
        if int(x["history"]) > 7:
            contPing += 1

    print("")
    print(colored(" Resultado do diagnóstico:", attrs=['bold']))
    print(colored(" [INFO]", 'blue'), "Quantidade de itens com chave icmpping armazenando histórico por mais de 7 dias:",
          contPing)
    print(colored(" [WARN]", 'red', attrs=['bold']), "Quantidade de itens não numéricos (ativos): ",
          itensNaoNumericos)
    print("")
    input(" Pressione ENTER para continuar")

    os.system('cls')
    banner()
    info()
    menu_opcao()


def listagemItensNaoSuportados():
    itensNaoSuportados = zapi.item.get(output=["itemid", "error", "name"],
                                       filter={"state": 1, "status": 0},
                                       monitored=True,
                                       selectHosts=["hostid", "host"],
                                       )

    if itensNaoSuportados:
        print(colored(' {0:10} | {1:90} | {2:180} | {3:20}'.format("Item", "Nome", "Error", "Host"), attrs=['bold']))

        for x in itensNaoSuportados:
            print(
                u'{0:10} | {1:90} | {2:180} | {3:20}'.format(x["itemid"], x["name"], x["error"], x["hosts"][0]["host"]))
        print("")
    else:
        print(" Não há dados a exibir")
        print("")
    input(" Pressione ENTER para continuar")

    os.system('cls')
    banner()
    info()
    menu_opcao()



def listagemItensNaoSuportadosGroupID():

    print("")
    GrupoFilter = input(" Digite o nome do Grupo de Hosts ou pressione ENTER para consultar os grupos disponíveis: ")
    print("")

    if not GrupoFilter:
        GetGroup = zapi.hostgroup.get(
            output='extend',
            sortfield='name'
        )

        for point in GetGroup:
            if 'Template' not in point['name'] and 'TEMPLATE' not in point['name'] and 'machines' not in point['name'] and 'Discovered' not in point['name'] and 'Hypervisors' not in point['name']:
                print(' ID: {}   |   Nome: {}'.format(point['groupid'], point['name']))
        print("")
        hostgroups = int(input(" Informe o ID do Grupo: "))
        #arquivosaida = input(" Informe o nome do arquivo de saída: ")
        #nomearquivo = arquivosaida + ".xls"

    elif GrupoFilter:
        GetGroup = zapi.hostgroup.get(
            output='extend',
            sortfield='name'
        )

        for point in GetGroup:
            if GrupoFilter in point['name']:
                print(' ID: {}   |   Nome: {}'.format(point['groupid'], point['name']))
        print("")
        hostgroups = int(input("Informe o ID do Grupo: "))
        arquivosaida = input("Informe o nome do arquivo de saída: ")
        nomearquivo = arquivosaida + ".xls"

    else:
        print(" Não foi identificado nenhum Grupo de Hosts com o nome informado!")
        input(" Pressione ENTER para voltar ao menu de opção!")
        os.system('cls')
        banner()
        info()
        menu_opcao()

    itensNaoSuportados = zapi.item.get(output=["itemid", "error", "name"],
                                       filter={"state": 1, "status": 0},
                                       monitored=True,
                                       groupids=hostgroups,
                                       selectHosts=["hostid", "host"],
                                       )

    if itensNaoSuportados:
        print(colored(' {0:10} | {1:90} | {2:180} | {3:20}'.format("Item", "Nome", "Error", "Host"), attrs=['bold']))

        for x in itensNaoSuportados:
            print(
                u'{0:10} | {1:90} | {2:180} | {3:20}'.format(x["itemid"], x["name"], x["error"], x["hosts"][0]["host"]))
        print("")
    else:
        print(" Não há dados a exibir")
        print("")
    input(" Pressione ENTER para continuar")

    os.system('cls')
    banner()
    info()
    menu_opcao()

def dadosItens():
    itensNaoSuportados = zapi.item.get(output="extend",
                                       filter={"state": 1, "status": 0},
                                       monitored=True,
                                       countOutput=True
                                       )

    totalItensHabilitados = zapi.item.get(output="extend",
                                          filter={"state": 0},
                                          monitored=True,
                                          countOutput=True
                                          )

    itensDesabilitados = zapi.item.get(output="extend",
                                       filter={"status": 1, "flags": 0},
                                       templated=False,
                                       countOutput=True
                                       )

    itensDescobertos = zapi.item.get(output="extend",
                                     selectItemDiscovery=["itemid"],
                                     selectTriggers=["triggers"],
                                     countOutput=True,
                                     monitored=True
                                     )

    itensZabbixAgent = zapi.item.get(output="extend",
                                     filter={"type": 0},
                                     templated=False,
                                     countOutput=True,
                                     monitored=True
                                     )

    itensSNMPv1 = zapi.item.get(output="extend",
                                filter={"type": 1},
                                templated=False,
                                countOutput=True,
                                monitored=True
                                )

    itensZabbixTrapper = zapi.item.get(output="extend",
                                       filter={"type": 2},
                                       templated=False,
                                       countOutput=True,
                                       monitored=True
                                       )

    itensChecagemSimples = zapi.item.get(output="extend",
                                         filter={"type": 3},
                                         templated=False,
                                         countOutput=True,
                                         monitored=True
                                         )

    itensSNMPv2 = zapi.item.get(output="extend",
                                filter={"type": 4},
                                templated=False,
                                countOutput=True,
                                monitored=True
                                )

    itensZabbixInterno = zapi.item.get(output="extend",
                                       filter={"type": 5},
                                       templated=False,
                                       countOutput=True,
                                       monitored=True
                                       )

    itensSNMPv3 = zapi.item.get(output="extend",
                                filter={"type": 6},
                                templated=False,
                                countOutput=True,
                                monitored=True
                                )

    itensZabbixAgentAtivo = zapi.item.get(output="extend",
                                          filter={"type": 7},
                                          templated=False,
                                          countOutput=True,
                                          monitored=True
                                          )

    itensZabbixAggregate = zapi.item.get(output="extend",
                                         filter={"type": 8},
                                         templated=False,
                                         countOutput=True,
                                         monitored=True
                                         )

    itensWeb = zapi.item.get(output="extend",
                             filter={"type": 9},
                             templated=False,
                             webitems=True,
                             countOutput=True,
                             monitored=True
                             )

    itensExterno = zapi.item.get(output="extend",
                                 filter={"type": 10},
                                 templated=False,
                                 countOutput=True,
                                 monitored=True
                                 )

    itensDatabase = zapi.item.get(output="extend",
                                  filter={"type": 11},
                                  templated=False,
                                  countOutput=True,
                                  monitored=True
                                  )

    itensIPMI = zapi.item.get(output="extend",
                              filter={"type": 12},
                              templated=False,
                              countOutput=True,
                              monitored=True
                              )

    itensSSH = zapi.item.get(output="extend",
                             filter={"type": 13},
                             templated=False,
                             countOutput=True,
                             monitored=True
                             )

    itensTelnet = zapi.item.get(output="extend",
                                filter={"type": 14},
                                templated=False,
                                countOutput=True,
                                monitored=True
                                )

    itensCalculado = zapi.item.get(output="extend",
                                   filter={"type": 15},
                                   templated=False,
                                   countOutput=True,
                                   monitored=True
                                   )

    itensJMX = zapi.item.get(output="extend",
                             filter={"type": 16},
                             templated=False,
                             countOutput=True,
                             monitored=True
                             )

    itensSNMPTrap = zapi.item.get(output="extend",
                                  filter={"type": 17},
                                  templated=False,
                                  countOutput=True,
                                  monitored=True
                                  )


    print("")
    print(" Relatório de itens")
    print(" =" * 18)
    print("")
    print(colored(" [INFO]", 'blue'), "Total de itens: ",
          int(totalItensHabilitados) + int(itensDesabilitados) + int(itensNaoSuportados))
    print(colored(" [INFO]", 'blue'), "Itens habilitados: ", totalItensHabilitados)
    print(colored(" [INFO]", 'blue'), "Itens desabilitados: ", itensDesabilitados)
    if itensNaoSuportados > "0":
        print(colored(" [ERRO]", 'red'), "Itens não suportados: ", itensNaoSuportados)
    else:
        print(colored(" [-OK-]", 'green'), "Itens não suportados: ", itensNaoSuportados)
    print(colored(" [INFO]",'blue'), "Itens descobertos: ", itensDescobertos)
    print("")
    print(" Itens por tipo em monitoramento")
    print(" =" * 14)
    print(colored(" [INFO]", 'blue'), "Itens Zabbix Agent (passivo): ", itensZabbixAgent)
    print(colored(" [INFO]", 'blue'), "Itens Zabbix Agent (ativo): ", itensZabbixAgentAtivo)
    print(colored(" [INFO]", 'blue'), "Itens Zabbix Trapper: ", itensZabbixTrapper)
    print(colored(" [INFO]", 'blue'), "Itens Zabbix Interno: ", itensZabbixInterno)
    print(colored(" [INFO]", 'blue'), "Itens Zabbix Agregado: ", itensZabbixAggregate)
    print(colored(" [INFO]", 'blue'), "Itens SNMPv1: ", itensSNMPv1)
    print(colored(" [INFO]", 'blue'), "Itens SNMPv2: ", itensSNMPv2)
    print(colored(" [INFO]", 'blue'), "Itens SNMPv3: ", itensSNMPv3)
    print(colored(" [INFO]", 'blue'), "Itens SNMNP Trap: ", itensSNMPTrap)
    print(colored(" [INFO]", 'blue'), "Itens JMX: ", itensJMX)
    print(colored(" [INFO]", 'blue'), "Itens IPMI: ", itensIPMI)
    print(colored(" [INFO]", 'blue'), "Itens SSH: ", itensSSH)
    print(colored(" [INFO]", 'blue'), "Itens Telnet: ", itensTelnet)
    print(colored(" [INFO]", 'blue'), "Itens Web: ", itensWeb)
    print(colored(" [INFO]", 'blue'), "Itens Checagem Simples: ", itensChecagemSimples)
    print(colored(" [INFO]", 'blue'), "Itens Calculado: ", itensCalculado)
    print(colored(" [INFO]", 'blue'), "Itens Checagem Externa: ", itensExterno)
    print(colored(" [INFO]", 'blue'), "Itens Database: ", itensDatabase)
    print("")
    input(" Pressione ENTER para continuar")

    os.system('cls')
    banner()
    info()
    menu_opcao()


def menu_relack():
    os.system('cls')
    banner()
    print(colored(" [+] - Bem-vindo ao ZABBIX TOOLS - [+]\n"
                  " [+] - Zabbix Tools faz um diagnóstico do seu ambiente e propõe melhorias na busca de um melhor desempenho - [+]\n",
                  'blue'))
    print("")
    print(colored(" --- Escolha uma opção para o relatório ---", 'yellow', attrs=['bold']))
    print("")
    print(" [1] - Relatório de triggers com Acknowledged")
    print(" [2] - Relatório de triggers com Unacknowledged")
    print(" [3] - Relatório de triggers com ACK/UNACK")
    print("")
    print(" [0] - Sair")
    print("")
    menu_opcao_relack()


def menu_opcao_relack():
    opcao = input(" [+] - Selecione uma opção[0-3]: ")

    params = "(output=['triggerid', 'lastchange', 'comments', 'description'], selectHosts=['hostid', 'host'], expandDescription=True, only_true=True, active=True)"
    if opcao == '1':
        params + "withAcknowledgedEvents=True"
        label = 'ACK'
    elif opcao == '2':
        params['withUnacknowledgedEvents'] = True
        label = 'UNACK'
    elif opcao == '3':
        label = 'ACK/UNACK'
    elif opcao == '0':
        os.system('cls')
        banner()
        info()
        menu_opcao()
    else:
        input("\n Pressione ENTER para voltar")
        menu_relack()

    # tm_till = time.mktime(datetime.now().timetuple())
    # hoje = tm_till - 60 * 60 * 144  # 1 hour
    hoje = datetime.date.today()
    try:
        tmp_trigger = int(input(" [+] - Selecione qual o tempo de alarme (dias): "))
    except Exception as e:
        input("\n Pressione ENTER para voltar")
        menu_relack()
    dt = (hoje - datetime.timedelta(days=tmp_trigger))
    conversao = int(time.mktime(dt.timetuple()))
    operador = input(" [+] - Deseja ver Triggers com mais ou menos de {0} dias [ + / - ] ? ".format(tmp_trigger))

    if operador == '+':
        params + "lastChangeTill=conversao)"
    elif operador == '-':
        params + "lastChangeSince=conversao)"
    else:
        input("\nPressione ENTER para voltar")
        menu_relack()

    rel_ack = zapi.trigger.get(params)
    for relatorio in rel_ack:
        lastchangeConverted = datetime.datetime.fromtimestamp(float(relatorio["lastchange"])).strftime('%Y-%m-%d %H:%M')
        print("")
        print(colored(" [-PROBLEM-]", 'red'), "Trigger {} com {} de {} dias".format(label, operador, tmp_trigger))
        print(" =" * 80)
        print("")
        print(colored(" [INFO]", 'blue'), "Nome da Trigger: ", relatorio["description"],
              "| HOST:" + relatorio["hosts"][0]["host"] + " | ID:" + relatorio["hosts"][0]["hostid"])
        print(colored(" [INFO]", 'blue'), "Hora de alarme: ", lastchangeConverted)
        print(colored(" [INFO]", 'blue'),
              "URL da trigger: {}/zabbix.php?action=problem.view&filter_set=1&filter_triggerids%5B%5D={}".format(server,
                                                                                                                 relatorio[
                                                                                                                     "triggerid"]))
        # print(colored("[INFO]", 'blue'), "Descrição da Trigger: ", relatorio["comments"])
        print("")

    print(colored("\n [INFO]", 'green'), "Total de {} triggers encontradas".format(rel_ack.__len__()))
    opcao = input("\n Deseja gerar relatorio em arquivo? [s/n]")
    if opcao == 's' or opcao == 'S':
        with open("relatorio.csv", "w") as arquivo:
            arquivo.write("Nome da Trigger,Hora de alarme:,URL da trigger:,Descrição da Trigger:\r\n ")
            for relatorio in rel_ack:
                arquivo.write((relatorio["description"]).encode('utf-8'))
                arquivo.write(("| HOST:" + relatorio["hosts"][0]["host"] + " | ID:" + relatorio["hosts"][0]["hostid"]))
                arquivo.write(",")
                arquivo.write(lastchangeConverted)
                arquivo.write(",")
                arquivo.write("{}/zabbix.php?action=problem.view&filter_set=1&filter_triggerids%5B%5D={}".format(server,
                                                                                                                 relatorio[
                                                                                                                     "triggerid"]))
                arquivo.write(",")
                arquivo.write(("\"" + relatorio["comments"] + "\"").encode('utf-8'))
                arquivo.write("\r\n")

        input("\n Arquivo gerado com sucesso ! Pressione ENTER para voltar")
        menu_relack()
    else:
        input("\n Pressione ENTER para voltar")
        menu_relack()


def ConsultaManutencao():

    print("")
    GrupoFilter = input(" Digite o nome do Grupo de Hosts ou pressione ENTER para consultar os grupos disponíveis: ")
    print("")

    if not GrupoFilter:
        GetGroup = zapi.hostgroup.get(
            output='extend',
            sortfield='name'
        )

        for point in GetGroup:
            if 'Template' not in point['name'] and 'TEMPLATE' not in point['name'] and 'machines' not in point['name'] and 'Discovered' not in point['name'] and 'Hypervisors' not in point['name']:
                print(' ID: {}   |   Nome: {}'.format(point['groupid'], point['name']))
        print("")
        hostgroups = int(input(" Informe o ID do Grupo: "))

    elif GrupoFilter:
        GetGroup = zapi.hostgroup.get(
            output='extend',
            sortfield='name'
        )

        for point in GetGroup:
            if GrupoFilter in point['name']:
                print(' ID: {}   |   Nome: {}'.format(point['groupid'], point['name']))
        print("")
        hostgroups = int(input("Informe o ID do Grupo: "))

    else:
        print(" Não foi identificado nenhum Grupo de Hosts com o nome informado!")
        input(" Pressione ENTER para voltar ao menu de opção!")
        os.system('cls')
        banner()
        info()
        menu_opcao()

    ConsultManut = zapi.maintenance.get(
        groupids=hostgroups,
        output='extend',
        SelectGroups='extend',
        selectTimeperiods='extend'
    )

    if ConsultManut:
        for a in ConsultManut:
            if 'Reinício' not in a['name'] and 'Reinicio' not in a['name']:
                print("\n Períodos de Manutenções:")
                print('\n Nome: {}\n Descrição: {}\n ID: {}'.format(a['name'], a['description'],
                                                                 a['maintenanceid']))
                print(' Ativo de: {}\n Ativado até: {}'.format(
                    datetime.fromtimestamp(int(a['active_since'])).strftime("%Y-%m-%d %X"),
                    datetime.fromtimestamp(int(a['active_till'])).strftime("%Y-%m-%d %X")))
                print("")
                input(" Pressione ENTER para voltar ao menu de opção!")
    else:
        print("\n Não há períodos de manutenção para o Grupo informado!")
        input(" Pressione ENTER para voltar ao menu de opção!")

    os.system('cls')
    banner()
    info()
    menu_opcao()


def CriaManutencao():

    print("")
    print(colored('''
    ##############################################################
    #                                                            #
    #  O período de manutenção iniciará ao término do cadastro!  #
    #                                                            #
    ##############################################################
    ''', 'red', attrs=['bold']))

    name = input("\n Informe um nome para o período (Ex. TJ/SP - Manutenção da base PG5XXX): ")
    duracao = int(input("\n Informe a duração da manutenção em minutos: "))
    desc = str(name)

    now = datetime.now()
    start_time = time.mktime(now.timetuple())
    period = int(duracao * 60)
    end_time = start_time + period

    print("")
    GrupoFilter = input(" Digite o nome do Grupo de Hosts ou pressione ENTER para consultar os grupos disponíveis: ")
    print("")

    if not GrupoFilter:
        GetGroup = zapi.hostgroup.get(
            output='extend',
            sortfield='name'
        )

        for point in GetGroup:
            if 'Template' not in point['name'] and 'TEMPLATE' not in point['name'] and 'machines' not in point['name'] and 'Discovered' not in point['name'] and 'Hypervisors' not in point['name']:
                print(' ID: {}   |   Nome: {}'.format(point['groupid'], point['name']))
        print("")
        hostgroups = int(input(" Informe o ID do Grupo: "))

    elif GrupoFilter:
        GetGroup = zapi.hostgroup.get(
            output='extend',
            sortfield='name'
        )

        for point in GetGroup:
            if GrupoFilter in point['name']:
                print(' ID: {}   |   Nome: {}'.format(point['groupid'], point['name']))
        print("")
        hostgroups = int(input("Informe o ID do Grupo: "))

    else:
        print(" Não foi identificado nenhum Grupo de Hosts com o nome informado!")
        input(" Pressione ENTER para voltar ao menu de opção!")
        os.system('cls')
        banner()
        info()
        menu_opcao()


    print("")
    input(" Dados inseridos, pressione ENTER para continuar")


    zapi.maintenance.create(
            {
                "groupids": [hostgroups],
                "name": name,
                "maintenance_type": 0,  # With data collection
                "active_since": str(start_time),
                "active_till": str(end_time),
                "description": desc,
                "timeperiods": [{
                    "timeperiod_type": "0",  # one time only;
                    "start_date": str(start_time),
                    "period": str(period),
                }]
            }
        )

    ConsultManutCriada = zapi.maintenance.get(
        groupids=[hostgroups],
        output='extend',
        SelectGroups='extend',
        selectTimeperiods='extend',
        filter={'name': name},
    )
    for point in ConsultManutCriada:
        print("\n Período de manutenção criado!")
        print("")
        print(
            '\n Nome: {}\n Descrição: {}\n ID: {}'.format(point['name'], point['description'], point['maintenanceid']))
        print(' Ativo de: {}\n Ativado até: {}'.format(
            datetime.fromtimestamp(int(point['active_since'])).strftime("%Y-%m-%d %X"),
            datetime.fromtimestamp(int(point['active_till'])).strftime("%Y-%m-%d %X")))
        print("")
        input(" Pressione ENTER para voltar ao menu de opção!")


    os.system('cls')
    banner()
    info()
    menu_opcao()

def DeletaManutencao():

    print("")
    GrupoFilter = input(" Digite o nome do Grupo de Hosts ou pressione ENTER para consultar os grupos disponíveis: ")
    print("")

    if not GrupoFilter:
        GetGroup = zapi.hostgroup.get(
            output='extend',
            sortfield='name'
        )

        for point in GetGroup:
            if 'Template' not in point['name'] and 'TEMPLATE' not in point['name'] and 'machines' not in point['name'] and 'Discovered' not in point['name'] and 'Hypervisors' not in point['name']:
                print(' ID: {}   |   Nome: {}'.format(point['groupid'], point['name']))
        print("")
        hostgroups = int(input(" Informe o ID do Grupo: "))

    elif GrupoFilter:
        GetGroup = zapi.hostgroup.get(
            output='extend',
            sortfield='name'
        )

        for point in GetGroup:
            if GrupoFilter in point['name']:
                print(' ID: {}   |   Nome: {}'.format(point['groupid'], point['name']))
        print("")
        hostgroups = int(input("Informe o ID do Grupo: "))

    else:
        print(" Não foi identificado nenhum Grupo de Hosts com o nome informado!")
        input(" Pressione ENTER para voltar ao menu de opção!")
        os.system('cls')
        banner()
        info()
        menu_opcao()

    ConsultManut = zapi.maintenance.get(
        groupids=hostgroups,
        output='extend',
        SelectGroups='extend',
        selectTimeperiods='extend'
    )

    if ConsultManut:
        for a in ConsultManut:
            if 'Reinício' not in a['name'] and 'Reinicio' not in a['name']:
                print("\n Períodos de Manutenções:")
                print('\n Nome: {}\n Descrição: {}\n ID: {}'.format(a['name'], a['description'],
                                                                 a['maintenanceid']))
                print(' Ativo de: {}\n Ativado até: {}'.format(
                    datetime.fromtimestamp(int(a['active_since'])).strftime("%Y-%m-%d %X"),
                    datetime.fromtimestamp(int(a['active_till'])).strftime("%Y-%m-%d %X")))
                print("")

    else:
        print("\n Não há períodos de manutenção para o Grupo informado!")
        input(" Pressione ENTER para voltar ao menu de opção!")

    print(colored(" CONFIRMAÇÃO DE EXCLUSÃO!", 'red', attrs=['bold']))
    print(colored('''
    ##############################################################
    #                                                            #
    #    O PERÍODO DE MANUTENÇÃO SERÁ EXCLUÍDO IMEDIATAMENTE!    #
    #                                                            #
    ##############################################################
    ''', 'red', attrs=['bold']))
    ManutID = input("\n Informe o ID a ser deletado: ")
    print("\n Foi informado o ID ", ManutID, ", confirme e pressione ENTER para prosseguir!")
    input("\n Pressione ENTER para continuar!")

    try:
        zapi.maintenance.delete(ManutID
                                )
        print("\n O período de manutenção ", ManutID, " foi excluído com sucesso!")

    except:
        print("\n Ocorreu um erro na exclusão do período de manutenção ", ManutID,
              ", por favor verifique se o ID informado está correto!")

    input("\n Pressione ENTER para voltar ao Menu.")
    os.system('cls')
    banner()
    info()
    menu_opcao()


def main():
    menu_opcao()


banner()
banner_vader()
info()
typeacess()
os.system('cls')
banner()
info()

main()
