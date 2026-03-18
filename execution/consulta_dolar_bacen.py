import requests
from datetime import datetime, timedelta

def obter_fechamento_dolar(data_base=None):
    # Feriados bancários para 2026
    feriados_bancarios = {
        "01-01-2026",  # Confraternização Universal
        "16-02-2026",  # Carnaval
        "17-02-2026",  # Carnaval
        "18-02-2026",  # Quarta-feira de Cinzas
        "03-04-2026",  # Paixão de Cristo
        "21-04-2026",  # Tiradentes
        "01-05-2026",  # Dia do Trabalho
        "04-06-2026",  # Corpus Christi
        "07-09-2026",  # Independência do Brasil
        "12-10-2026",  # Nossa Sra. Aparecida
        "02-11-2026",  # Finados
        "15-11-2026",  # Proclamação da República
        "20-11-2026",  # Dia da Consciência Negra
        "25-12-2026",  # Natal
        "31-12-2026",  # Véspera de Ano Novo (feriado bancário)
    }

    def data_util(data):
        data_formatada = data.strftime('%m-%d-%Y')
        return (data_formatada not in feriados_bancarios and 
                data.weekday() < 5) 

    # Data atual ou base
    if data_base is None:
        data_atual_calc = datetime.now()
    else:
        # Se for string, tentamos converter tentando varios formatos
        if isinstance(data_base, str) and data_base.strip():
            parsed = False
            for fmt in ['%Y-%m-%d', '%d/%m/%Y', '%m/%d/%Y']:
                try:
                    data_atual_calc = datetime.strptime(data_base.strip(), fmt)
                    parsed = True
                    break
                except ValueError:
                    continue
            if not parsed:
                data_atual_calc = datetime.now()
        else:
            data_atual_calc = datetime.now() if isinstance(data_base, str) else data_base

    hoje = datetime.now() # Ainda precisamos de hoje para retornar no fim
    print(f"Data base para cálculo: {data_atual_calc.strftime('%Y-%m-%d')}")
    print(f"Data atual: {hoje.strftime('%Y-%m-%d')}")

    # Começar a verificar a partir do dia anterior à data base
    data = data_atual_calc - timedelta(days=1)
    print(f"Iniciando a verificação a partir de: {data.strftime('%Y-%m-%d')}")

    # Verifica as datas até encontrar um dia útil
    while True:
        print(f"Verificando data: {data.strftime('%Y-%m-%d')}")
        
        if data_util(data):
            # Verifica se existe cotação para a data
            data_formatada = data.strftime('%m-%d-%Y')
            url = f"https://olinda.bcb.gov.br/olinda/servico/PTAX/versao/v1/odata/CotacaoDolarDia(dataCotacao=@dataCotacao)?@dataCotacao='{data_formatada}'&$top=1&$format=json"
            resposta = requests.get(url)

            print(f"Consultando URL: {url}")

            if resposta.status_code == 200:
                dados = resposta.json()
                if dados['value']:
                    # Retorna a cotação encontrada
                    cotacao_venda = dados['value'][0]['cotacaoVenda']
                    print(f"Cotação de venda em {data_formatada}: R$ {cotacao_venda:.4f}")
                    return data_formatada, cotacao_venda, hoje
                else:
                    print(f"Nenhum dado encontrado para a data: {data_formatada}")
            else:
                print(f"Erro ao acessar a API: {resposta.status_code}, URL: {url}")

        # Retrocede um dia se não encontrar cotação válida
        data -= timedelta(days=1)

# Para usar este arquivo em outro script:
# data_formatada, cotacao_venda = obter_fechamento_dolar()