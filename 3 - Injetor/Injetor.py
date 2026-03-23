import time
import re
import threading
import tkinter as tk
from tkinter import scrolledtext, filedialog
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys

class AutomacaoMasterERP:
    def __init__(self, root):
        self.root = root
        self.root.title("Automator ERP v3.0 - Pro Edition")
        self.root.geometry("800x800")
        
        self.user_erp = tk.StringVar()
        self.pass_erp = tk.StringVar()
        self.caminho_txt = tk.StringVar()
        self.inicio_loop = tk.StringVar(value="0")
        self.fim_loop = tk.StringVar(value="50")
        self.url_base = "https://linkEmpresa/login"
        
        self.setup_ui()

    def setup_ui(self):
        tk.Label(self.root, text="MOTOR DE INJEÇÃO - SGF/SGE", font=("Consolas", 14, "bold"), fg="#1abc9c").pack(pady=10)
        
        form = tk.Frame(self.root, padx=20)
        form.pack(fill="x")

        f_login = tk.LabelFrame(form, text=" Autenticação ", padx=10, pady=5)
        f_login.grid(row=0, column=0, padx=5, sticky="nsew")
        tk.Label(f_login, text="User:").grid(row=0, column=0, sticky="w")
        tk.Entry(f_login, textvariable=self.user_erp, width=15).grid(row=0, column=1)
        tk.Label(f_login, text="Pass:").grid(row=1, column=0, sticky="w")
        tk.Entry(f_login, textvariable=self.pass_erp, width=15, show="*").grid(row=1, column=1)

        f_loop = tk.LabelFrame(form, text=" Parâmetros de Lote ", padx=10, pady=5)
        f_loop.grid(row=0, column=1, padx=5, sticky="nsew")
        tk.Label(f_loop, text="Index Inicial:").grid(row=0, column=0, sticky="w")
        tk.Entry(f_loop, textvariable=self.inicio_loop, width=10).grid(row=0, column=1)
        tk.Label(f_loop, text="Qtd Itens:").grid(row=1, column=0, sticky="w")
        tk.Entry(f_loop, textvariable=self.fim_loop, width=10).grid(row=1, column=1)

        tk.Button(self.root, text="📂 CARREGAR EXTRACAO_UNIFICADA.TXT", 
                  command=self.selecionar_txt, bg="#34495e", fg="white", font=("Arial", 9, "bold")).pack(pady=15)
        
        self.lbl_contador = tk.Label(self.root, text="Aguardando Início...", font=("Arial", 12, "bold"), fg="#f39c12")
        self.lbl_contador.pack(pady=5)
        
        self.log_area = scrolledtext.ScrolledText(self.root, height=18, bg="#1e1e1e", fg="#00ff00", font=("Consolas", 10))
        self.log_area.pack(padx=20, pady=5, fill="both")

        tk.Button(self.root, text="🚀 EXECUTAR DEPLOY NO ERP", bg="#27ae60", fg="white", 
                  font=("Arial", 11, "bold"), height=2, command=self.start).pack(pady=15)

    def log(self, msg):
        self.log_area.insert(tk.END, f"[{time.strftime('%H:%M:%S')}] {msg}\n")
        self.log_area.see(tk.END)

    def selecionar_txt(self):
        p = filedialog.askopenfilename(filetypes=[("Arquivos de Texto", "*.txt")])
        if p: 
            self.caminho_txt.set(p)
            self.log(f"Arquivo mapeado: {p}")

    def classificar_produto(self, nome):
        nome = nome.upper()
        regras = {
            "ABS": "PROTECAO FEMININA", "ABSORVENTE": "PROTECAO FEMININA",
            "AGUA PERF": "AROMATIZANTES", "AROMATIZ": "AROMATIZANTES",
            "BATOM": "MAQUIAGEM", "BASE": "MAQUIAGEM", "MAKE": "MAQUIAGEM",
            "UNHA": "UNHAS", "ESM": "UNHAS", "ALICATE": "UNHAS",
            "SHAMP": "CABELOS", "CONDIC": "CABELOS",
            "MAMADEIRA": "LINHA INFANTIL", "CHUP": "LINHA INFANTIL",
            "AGUA MIN": "CONVENIENCIA", "CHOC": "CONVENIENCIA",
            "PERF": "PERFUMES", "ILUM": "MAQUIAGEM", "PILHA":"CONVENIENCIA",
            "LUVA": "LINHA HOSPITALAR"
        }
        for chave, categoria in regras.items():
            if chave in nome: return categoria
        return "vazio"

    def extrair_dados(self):
        itens = []
        try:
            encoding = 'utf-8'
            try:
                with open(self.caminho_txt.get(), 'r', encoding=encoding) as f:
                    linhas = f.readlines()
            except UnicodeDecodeError:
                encoding = 'latin-1'
                with open(self.caminho_txt.get(), 'r', encoding=encoding) as f:
                    linhas = f.readlines()

            padrao = re.compile(r"C[oó]d:\s*(\d+)\s*\|\s*Ref:\s*(.*?)\s*\|\s*Conc:\s*(.*)", re.IGNORECASE)
            for linha in linhas:
                match = padrao.search(linha)
                if match:
                    codigo, nome, conc_original = match.groups()
                    conc_final = conc_original.strip() if conc_original.lower() != "vazio" else self.classificar_produto(nome)
                    if conc_final != "vazio":
                        itens.append({"cod": codigo.strip(), "conc": conc_final, "nome": nome.strip()})
            return itens
        except Exception as e:
            self.log(f"ERRO LEITURA: {e}")
            return []

    def interagir_campo(self, driver, wait, element_id, valor, press_enter=False):
        """Helper para garantir que o campo seja preenchido corretamente"""
        campo = wait.until(EC.element_to_be_clickable((By.ID, element_id)))
        driver.execute_script("arguments[0].focus();", campo)
        driver.execute_script("arguments[0].value = '';", campo) 
        campo.click()
        campo.send_keys(Keys.CONTROL + "a")
        campo.send_keys(Keys.BACKSPACE)
        campo.send_keys(valor)
        if press_enter:
            campo.send_keys(Keys.ENTER)
        return campo

    def tratar_alerta(self, driver):
        try:
            WebDriverWait(driver, 2).until(EC.alert_is_present())
            alerta = driver.switch_to.alert
            texto = alerta.text
            alerta.accept()
            self.log(f"Alerta do Sistema: {texto}")
            return True
        except: return False

    def main_loop(self):
        dados = self.extrair_dados()
        if not dados:
            self.log("ERRO: Nenhum dado válido encontrado.")
            return

        idx = int(self.inicio_loop.get())
        qtd = int(self.fim_loop.get())
        lote = dados[idx : idx + qtd]
        
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service)
        wait = WebDriverWait(driver, 12)

        try:
            self.log("Acessando ERP via DDNS...")
            driver.get(f"{self.url_base}/Login.pod")
            wait.until(EC.presence_of_element_located((By.ID, "id_cod_usuario"))).send_keys(self.user_erp.get())
            driver.find_element(By.ID, "nom_senha").send_keys(self.pass_erp.get())
            driver.find_element(By.ID, "login").click()
            time.sleep(3)

            for i, item in enumerate(lote, start=1):
                try:
                    self.lbl_contador.config(text=f"Processando: {i} / {len(lote)}")
                    
                    driver.get(f"{self.url_base}/Cad_0020.pod")
                    time.sleep(1.5)
                    self.tratar_alerta(driver)

                    self.log(f"-> Item {i}: Cod {item['cod']}")
                    self.interagir_campo(driver, wait, "cod_redbarraEntrada", item['cod'], True)
                    time.sleep(1.5)
                    
                    if self.tratar_alerta(driver): continue

                    campo_conc = self.interagir_campo(driver, wait, "cod_concentracaoEntrada", item['conc'])
                    time.sleep(1)
                    campo_conc.send_keys(Keys.ENTER)
                    time.sleep(1)

                    driver.find_element(By.TAG_NAME, "body").send_keys(Keys.F2)
                    self.log(f"   [OK] {item['conc']} salvo.")
                    time.sleep(1)

                except Exception as e:
                    self.log(f"   [ERRO] Falha no item {item['cod']}")
                    continue

        except Exception as e:
            self.log(f"CRÍTICO: {e}")
        finally:
            self.log("Sessão finalizada.")
            self.lbl_contador.config(text="FIM DO PROCESSO", fg="#2ecc71")

    def start(self):
        if not self.caminho_txt.get():
            self.log("ERRO: Selecione o arquivo primeiro!")
            return
        threading.Thread(target=self.main_loop, daemon=True).start()

if __name__ == "__main__":
    root = tk.Tk()
    app = AutomacaoMasterERP(root)
    root.mainloop()
