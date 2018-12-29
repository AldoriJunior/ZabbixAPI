# Consumindo a API do Zabbix Server utilizando Python

No universo de números e informações gerados pelo Zabbix, desenvolvi scripts para diversas operações, como a extração de dados, manutenção e ajuste de informações no ZabbixServer através da API usando Python.

## Requisitos

- ZabbixServer 3.2 ou superior
- Python 3.5 ou superior

## Instalação
```
git clone https://github.com/AldoriJunior/ZabbixAPI
```

## Informações de acesso ao Zabbix

As informações URL, usuário e senha para acesso a API do seu Zabbix serão solicitadas na execução do script, mas podem ser inibidas e inseridas diretamente no código do arquivo.
```
server = input('\nInforme a URL de conexão com o seu Zabbix: ')
username = input('\nInforme o usuário: ')
password = getpass.getpass('\nInforme a senha: ')

#server = 'http://zabbixserver.com.br
#server = 'http://192.168.0.1
#username = 'user'
#password = 'pass'
```

## Scripts disponibilizados

Aqui listo os scripts disponibilizados e suas funções:

- Extração de itens e triggers de todos os HostID de um GroupID através de um menu que lista automaticamente todos os GroupID e Nome do seu Zabbix. Os dados são extraídos para um arquivo XLS de local a definir pelo executor.

  - Zabbix_ExtractItem&TriggerFomGroupID.py

- Consulta de Dados Recentes de Itens através dos filtros GroupID -> HostID -> ItemID -> Data de Início e Fim.

  - Zabbix_ConsultarDadosRecentes.py

## Execução dos Scripts

Para executar os scripts:
```
python script.py
```
