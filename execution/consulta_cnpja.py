import requests

def consultar_cnpja(cnpj_number, uf):
    # Consulta à API CNPJ
    url_cnpj = f'https://api.cnpja.com/ccc?states=BR&taxId={cnpj_number}'
    headers_cnpj = {'Authorization': 'ce7383d4-2e0d-4922-b673-5489d2aaf7c1-2a03ec20-d71d-49de-a4c9-78a78a69ea7d'}

    try:
        response_cnpj = requests.get(url_cnpj, headers=headers_cnpj)
        print(f"Status da resposta CNPJA: {response_cnpj.status_code}")
        print(f"Resposta CNPJA: {response_cnpj.text}")

        resultados_ie = []
        number_ie = 'N/A'

        if response_cnpj.status_code == 200:
            data_cnpja = response_cnpj.json()
            registrations = data_cnpja.get('registrations', [])

            for reg in registrations:
                estado = reg.get('state', 'N/A')
                number = reg.get('number', 'N/A')
                resultados_ie.append({'estado': estado, 'numero': number})

                # Se encontrar uma inscrição do mesmo estado, usa ela
                if estado == uf:
                    number_ie = number
                    break

            # Se não encontrou inscrição do mesmo estado, usa a primeira disponível
            if number_ie == 'N/A' and resultados_ie:
                number_ie = resultados_ie[0].get('numero', 'N/A')

        return {
            "resultados_ie": resultados_ie,
            "quantidade_total": len(resultados_ie),
            "number_ie": number_ie
        }
    except Exception as e:
        print(f"Erro na consulta CNPJA: {str(e)}")
        return {
            "resultados_ie": [],
            "quantidade_total": 0,
            "number_ie": 'N/A'
        } 