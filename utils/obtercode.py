# utils/obtercode.py
import subprocess
import time
import urllib.parse
import os
import re

def obter_code(silent_mode=True, base: str = "homologacao"):
    """
    Obtém um código de autorização OAuth do Acesso Cidadão ES automaticamente.
    
    Args:
        silent_mode (bool): Se True, reduz as mensagens de log para o mínimo necessário
            
    Returns:
        str: Código de autorização ou None se não for possível obter
    """
    # Função para log condicional
    def log(message):
        if not silent_mode:
            print(message)
    
    log("Iniciando o processo de obtenção do código...")
    
    # Verificando se as bibliotecas necessárias estão instaladas
    try:
        import pyautogui
        import pyperclip
    except ImportError:
        log("Instalando dependências necessárias...")
        subprocess.check_call(["pip", "install", "pyautogui", "pyperclip"])
        import pyautogui
        import pyperclip
    
    # client_id = "e1463040-77cb-4e8a-ae78-7d9b6ecd4b6c" #Base de homologação
    # client_id = "a3576f23-80d8-4c47-b9ec-f122ae16e34e" #Base de Produção
    #base = "acesso_cidadao"
    if base == "producao":
        client_id = "a3576f23-80d8-4c47-b9ec-f122ae16e34e" 
    elif base == "homologacao":
        client_id = "e1463040-77cb-4e8a-ae78-7d9b6ecd4b6c" 
    elif base == "treinamento":
        client_id = "f39b8f52-26e7-43fa-9b2b-0fa576547767" 
    elif base == "acesso_cidadao":   #as credenciais do acesso cidadão só existem para a base de produção
        client_id = "257fdc98-06f8-4d49-9275-5a20a58a60e1" 
    else:
        raise FileNotFoundError("A base de dados informada não é válida!")
    
    if base == "acesso_cidadao":
        scope = "openid profile ApiAcessoCidadao api-acessocidadao-base api-acessocidadao-cpf api-acessocidadao-servidores"    #escopo de acesso
    else:
        scope = "openid profile api-sigades-documento api-sigades-processo api-sigades-consultar"    #escopo de acesso

        # Homologação
    redirect_uri = "https://detran.es.gov.br/"         #url para redirecionamento
    nonce = "1000"                                     #numero aleatorio
    #scope = "openid profile api-sigades-documento api-sigades-processo api-sigades-consultar"    #escopo de acesso
    auth_url = "https://acessocidadao.es.gov.br/is/connect/authorize?response_type=code%20id_token&client_id=" + client_id + "&scope=" + scope + "&redirect_uri=" + redirect_uri + "&nonce=" + nonce

    # URL de login/autorização
    #auth_url = "https://acessocidadao.es.gov.br/is/connect/authorize?response_type=code%20id_token&client_id=e1463040-77cb-4e8a-ae78-7d9b6ecd4b6c&scope=openid%20profile%20api-sigades-documento%20api-sigades-processo%20api-sigades-consultar&redirect_uri=https://detran.es.gov.br/&nonce=1000"
    
    chrome_process = None
    janela_autorizada = None
    
    try:
        # Caminho para o Chrome
        chrome_path = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
        if not os.path.exists(chrome_path):
            chrome_path = r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
            if not os.path.exists(chrome_path):
                raise FileNotFoundError("Chrome não encontrado nos locais padrão")
        
        # Abrir Chrome em nova aba (sem --new-window)
        chrome_args = [
            chrome_path,
            f"--window-size=900,800",
            auth_url
        ]
        log(f"Abrindo Chrome em nova aba: {chrome_path}")
        print(f'[DEBUG] Comando para abrir Chrome: {chrome_args}')
        chrome_process = subprocess.Popen(chrome_args)
        print('[DEBUG] Chrome chamado, aguardando 2 segundos para abrir...')
        time.sleep(2)
        print('[DEBUG] Tentando capturar janelas do Chrome após abertura...')
        try:
            import pyautogui
            max_tentativas = 30  # aumenta o tempo de tentativa
            encontrou_code = False
            for tentativa in range(max_tentativas):
                janelas_chrome = pyautogui.getWindowsWithTitle("Chrome")
                if not janelas_chrome:
                    time.sleep(1)
                    continue
                for janela in janelas_chrome:
                    if janela is None:
                        continue
                    try:
                        janela.activate()
                        time.sleep(0.5)
                        pyautogui.hotkey('ctrl', 'l')
                        time.sleep(0.3)
                        pyautogui.hotkey('ctrl', 'c')
                        time.sleep(0.3)
                        url_inicial = pyperclip.paste()
                        print(f'[DEBUG] URL capturada da janela Chrome: {url_inicial}')
                        if "detran.es.gov.br" in url_inicial and "code=" in url_inicial:
                            log("\n✅ Redirecionamento detectado em uma das janelas do Chrome!")
                            janela_autorizada = janela  # Guarda a referência da janela
                            # Extrair o código da URL
                            parsed_url = urllib.parse.urlparse(url_inicial)
                            if "#" in url_inicial:
                                fragment = parsed_url.fragment
                                params = urllib.parse.parse_qs(fragment)
                                code = params.get("code", [None])[0]
                            else:
                                query = parsed_url.query
                                params = urllib.parse.parse_qs(query)
                                code = params.get("code", [None])[0]
                            if not code:
                                match = re.search(r'code=([^&#]+)', url_inicial)
                                if match:
                                    code = match.group(1)
                            if code:
                                log(f"\n🔑 Código de acesso obtido: {code}")
                                print("Código de acesso obtido com sucesso: ", code)
                                encontrou_code = True
                                return code
                    except Exception as e:
                        print(f'[ERRO] Falha ao ativar/capturar URL da janela Chrome: {e}')
                time.sleep(1)
            if not encontrou_code:
                log("(Aviso) Não foi possível garantir o foco/captura automática em nenhuma aba do Chrome. Por favor, clique manualmente na janela do Chrome e faça o login/redirecionamento.")
        except Exception as e:
            print(f'[ERRO] Falha ao capturar janelas do Chrome: {e}')
        
        log("\n📱 Chrome aberto com a página de autenticação.")
        log("⚠️ Por favor, faça o login manualmente e complete qualquer desafio do Cloudflare.")
        log("✅ O sistema capturará automaticamente o código após o redirecionamento.")
        
        # Função para verificar se o navegador foi redirecionado para o Detran e capturar a URL
        def detectar_redirecionamento_e_capturar_url():
            # Tempo total de espera: 5 minutos (300 segundos)
            tempo_total = 300
            intervalo_verificacao = 1  # Verificar a cada 1 segundo
            tempo_inicial = time.time()
            
            log("\n⏳ Monitorando redirecionamento automaticamente...")
            nonlocal janela_autorizada
            
            app_window = next(iter(pyautogui.getWindowsWithTitle("Gera Relato")), None)
            
            try:
                while (time.time() - tempo_inicial) < tempo_total:
                    # A cada 30 segundos, mostrar mensagem de espera
                    elapsed = int(time.time() - tempo_inicial)
                    if elapsed % 30 == 0 and elapsed > 0:
                        log(f"⏳ Ainda aguardando redirecionamento... ({elapsed}s)")
                    
                    try:
                        janelas_chrome = pyautogui.getWindowsWithTitle("Chrome")
                        if not janelas_chrome:
                            time.sleep(intervalo_verificacao)
                            continue

                        for janela in janelas_chrome:
                            if not janela.exists: continue

                            # Tenta focar na janela do Chrome
                            if janela.isMinimized:
                                janela.restore()
                            
                            janela.activate()
                            time.sleep(0.3) # Dar tempo para o SO reagir

                            # Se o foco falhou e a janela da app existe, minimiza a app e tenta de novo
                            if not janela.isActive and app_window and app_window.exists and not app_window.isMinimized:
                                log("[AVISO] Janela do Chrome não conseguiu foco. Minimizando app principal...")
                                app_window.minimize()
                                time.sleep(0.3)
                                janela.activate()
                                time.sleep(0.2)

                            # Confirma que a janela do Chrome está realmente ativa antes de interagir
                            if janela.isActive:
                                pyautogui.hotkey('ctrl', 'l')
                                time.sleep(0.3)
                                pyautogui.hotkey('ctrl', 'c')
                                time.sleep(0.3)
                                url = pyperclip.paste()
                                
                                if "detran.es.gov.br" in url and "code=" in url:
                                    log(f"\n✅ Redirecionamento detectado!")
                                    janela_autorizada = janela
                                    return url # Sucesso, sai da função
                                
                                if "detran.es.gov.br" in url and "code=" not in url:
                                    log("\n⚠️ Redirecionado para o Detran, mas sem código na URL.")
                                    janela_autorizada = janela
                    except Exception as e:
                        if not silent_mode:
                            print(f"Erro ao verificar redirecionamento: {e}")
                    
                    time.sleep(intervalo_verificacao)
            finally:
                # Garante que a janela da aplicação seja restaurada
                if app_window and app_window.exists:
                    if app_window.isMinimized:
                        app_window.restore()
                    app_window.activate()
                    log("Janela da aplicação restaurada.")

            return None
        
        # Detecta redirecionamento e captura URL
        url_final = detectar_redirecionamento_e_capturar_url()
        
        # Se não detectou automaticamente, solicita manualmente
        if not url_final:
            log("\n⚠️ Não foi possível detectar o redirecionamento automaticamente.")
            url_final = input("URL: ")
        
        if not url_final:
            log("❌ Nenhuma URL fornecida!")
            return None
        
        # Verificar se a URL contém o código
        if "detran.es.gov.br" not in url_final:
            log("❌ A URL fornecida não parece ser do site do Detran!")
            return None
        
        if "code=" not in url_final:
            log("❌ A URL fornecida não contém o código de acesso!")
            return None
        
        # Extrair o código da URL fornecida
        parsed_url = urllib.parse.urlparse(url_final)
        
        # Tenta extrair do fragmento primeiro
        if "#" in url_final:
            fragment = parsed_url.fragment
            params = urllib.parse.parse_qs(fragment)
            code = params.get("code", [None])[0]
        else:
            # Se não estiver no fragmento, tenta na query string
            query = parsed_url.query
            params = urllib.parse.parse_qs(query)
            code = params.get("code", [None])[0]
        
        # Se não conseguiu extrair com o parse padrão, tenta regex
        if not code:
            match = re.search(r'code=([^&#]+)', url_final)
            if match:
                code = match.group(1)
        
        if code:
            log(f"\n🔑 Código de acesso obtido com sucesso: {code}")
            print("Código de acesso obtido com sucesso: ", code)
            return code
        else:
            # Última tentativa - solicitação manual do código
            codigo_manual = input("Código: ")
            if codigo_manual:
                return codigo_manual.strip()
            return None
            
    except Exception as e:
        if not silent_mode:
            print(f"Erro ao abrir Chrome: {e}")
        
        # Solicitar URL manualmente
        log("\n👉 Por favor, copie e cole a URL completa após o redirecionamento:")
        url_final = input("URL: ")
        
        if not url_final or "code=" not in url_final:
            log("❌ URL inválida ou sem código de acesso!")
            return None
        
        # Extrair código da URL
        match = re.search(r'code=([^&]+)', url_final)
        if match:
            code = match.group(1)
            log(f"\n🔑 Código de acesso obtido com sucesso: {code}")
            print("Código de acesso obtido com sucesso: ", code)
            return code
        else:
            log("❌ Não foi possível extrair o código da URL!")
            return None
    
    finally:
        # Fecha a aba específica do Chrome que foi usada para o redirecionamento
        try:
            if janela_autorizada:
                log("Fechando a aba de autenticação do Chrome...")
                janela_autorizada.activate()
                time.sleep(0.5)
                pyautogui.hotkey('ctrl', 'w')
                log("Aba do Chrome fechada com sucesso!")
            elif chrome_process:
                # Fallback para o método antigo se a janela não foi capturada mas o processo sim
                log("Fechando o processo do Chrome iniciado pelo script...")
                chrome_process.terminate()
                time.sleep(1)
                if chrome_process.poll() is None:
                    chrome_process.kill()
        except Exception as e:
            if not silent_mode:
                print(f"Erro ao tentar fechar a aba/janela do Chrome: {e}")
