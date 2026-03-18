import requests

def consultar_receitaws(cnpj_number):
    # Consulta à API ReceitaWS
    url_receita = f"https://receitaws.com.br/v1/cnpj/{cnpj_number}"
    headers_receita = {"Accept": "application/json"}

    try:
        response_receita = requests.get(url_receita, headers=headers_receita)
        print(f"Status da resposta ReceitaWS: {response_receita.status_code}")
        print(f"Resposta ReceitaWS: {response_receita.text}")
        
        if response_receita.status_code == 200:
            data_receita = response_receita.json()
            
            # Extrair os dados necessários
            return {
                "nome": data_receita.get('nome', 'N/A'),
                "fantasia": data_receita.get('fantasia', 'N/A'),
                "cnpj": data_receita.get('cnpj', 'N/A'),
                "logradouro_numero": data_receita.get('logradouro', 'N/A') + ', ' + data_receita.get('numero', 'N/A'),
                "complemento": data_receita.get('complemento', 'N/A'),
                "municipio": data_receita.get('municipio', 'N/A'),
                "uf": data_receita.get('uf', 'N/A'),
                "bairro": data_receita.get('bairro', 'N/A'),
                "cep": data_receita.get('cep', 'N/A')
            }
        return None
    except Exception as e:
        print(f"Erro na consulta ReceitaWS: {str(e)}")
        return None 