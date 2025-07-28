# utils/gera_token.py
import requests
import json

def obter_token(code: str, silent_mode=True, base: str = "homologacao"):
    """
    Obtém um token de acesso usando o código de autorização.
    
    Args:
        code (str): Código de autorização obtido do Acesso Cidadão
        silent_mode (bool): Se True, reduz as mensagens de log para o mínimo necessário
            
    Returns:
        dict: Dicionário contendo access_token, token_type e expires_in ou
        str: Mensagem de erro "Não foi possível obter o token!"
    """

    def log(*args):
        if not silent_mode:
            print(*args)

    log("Iniciando processo de obtenção do token...")

    #base = "acesso_cidadao" só é usada para obter os papeis dos membros, o token é de aplicação não de usuário
    if base == "producao":
        base64 = "YTM1NzZmMjMtODBkOC00YzQ3LWI5ZWMtZjEyMmFlMTZlMzRlOk83UHVAR3VlaGtlQm9VNXVFcGYxTypAKnIzZipCaw=="   #produção
    elif base == "homologacao":
        base64 = "ZTE0NjMwNDAtNzdjYi00ZThhLWFlNzgtN2Q5YjZlY2Q0YjZjOnlCQHN3NHZDQk14OUJRUypjNFJyMFlQTmpDRmRweA=="   #Homologação
    elif base == "acesso_cidadao":
        base64 = "MjU3ZmRjOTgtMDZmOC00ZDQ5LTkyNzUtNWEyMGE1OGE2MGUxOlVRMnZzcU82bXVCJGloeFJZb05jeXgxWW5keVF4TA=="   #Acesso Cidadão
    else:
        raise FileNotFoundError("A base de dados informada não é válida!")

    # Configurações fixas
    url = "https://acessocidadao.es.gov.br/is/connect/token"
    
    # Headers da requisição
    headers = {
        'Authorization': f'Basic {base64}',
        'Content-Type': 'application/x-www-form-urlencoded'
    }
    
    # Dados do formulário
    if base == "acesso_cidadao":
        data = {
            'grant_type': 'client_credentials',
            #'scope': 'openid profile ApiAcessoCidadao api-acessocidadao-base api-acessocidadao-cpf api-acessocidadao-servidores'
            'scope': 'ApiAcessoCidadao api-acessocidadao-base api-acessocidadao-cpf api-acessocidadao-servidores'
        }
    else:
        data = {
            'grant_type': 'authorization_code',
            'code': code,
            'redirect_uri': 'https://detran.es.gov.br/'
        }   
    
    try:
        log("Enviando requisição para obter token...")
        print ("request: ", url, " ", headers, " ", data)
        response = requests.post(url, headers=headers, data=data)
        
        # Verifica se a requisição foi bem sucedida
        response.raise_for_status()
        
        # Parse do JSON de resposta
        token_data = response.json()
        
        # Verifica se todos os campos necessários estão presentes
        required_fields = ['access_token', 'token_type', 'expires_in']
        if not all(field in token_data for field in required_fields):
            raise ValueError("Resposta da API não contém todos os campos necessários")
        
        # Extrai apenas os campos necessários
        result = {
            'access_token': token_data['access_token'],
            'token_type': token_data['token_type'],
            'expires_in': token_data['expires_in']
        }
        
        log("Token obtido com sucesso!", result)
        print("Token obtido com sucesso!", result)
        return result
        
    except requests.exceptions.RequestException as e:
        log(f"Erro na requisição HTTP: {str(e)}")
        print(f"Erro na requisição HTTP: {str(e)}")
        return "Não foi possível obter o token!"
        
    except ValueError as e:
        log(f"Erro ao processar resposta: {str(e)}")
        print(f"Erro ao processar resposta: {str(e)}")
        return "Não foi possível obter o token!"
        
    except Exception as e:
        log(f"Erro inesperado: {str(e)}")
        print(f"Erro inesperado: {str(e)}")
        return "Não foi possível obter o token!"