import pandas as pd
import time
import os
import threading
import tkinter as tk
from tkinter import ttk, scrolledtext, filedialog, messagebox
from concurrent.futures import ThreadPoolExecutor
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys

ARQUIVO_RESULTADO = 'extração_dados_multi.txt'

class AutomationGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Extração MULTI-THREAD ERP - v1.2")
        self.root.geometry("850x950")
        
        self.rodando = False
        self.contador_global = 0
        self.lock = threading.Lock() 
        
        self.user_erp = tk.StringVar()
        self.pass_erp = tk.StringVar()
        self.caminho_excel_var = tk.StringVar()
        self.coluna_alvo = tk.StringVar(value="C")
        self.linha_inicio = tk.IntVar(value=16)
        
        self.chk_concentracao = tk.BooleanVar(value=True)
        self.chk_tipo = tk.BooleanVar(value=True)
        self.chk_preco = tk.BooleanVar(value=True)
        self.chk_usar_filtro = tk.BooleanVar(value=True)
        self.filtro_texto = tk.StringVar(value="tipo de produtos")
        self.num_threads_var = tk.IntVar(value=2)

        tk.Label(root, text="Painel de Controle de Automação", font=("Arial", 14, "bold")).pack(pady=10)
        
        login_frame = tk.LabelFrame(root, text=" Credenciais do ERP ", padx=10, pady=10)
        login_frame.pack(pady=5, fill="x", padx=20)
        tk.Label(login_frame, text="Usuário:").grid(row=0, column=0, sticky="w")
        tk.Entry(login_frame, textvariable=self.user_erp, width=30).grid(row=0, column=1, padx=5, pady=2)
        tk.Label(login_frame, text="Senha:").grid(row=1, column=0, sticky="w")
        tk.Entry(login_frame, textvariable=self.pass_erp, width=30, show="*").grid(row=1, column=1, padx=5, pady=2)

        file_frame = tk.LabelFrame(root, text=" Configuração da Planilha Excel ", padx=10, pady=10)
        file_frame.pack(pady=5, fill="x", padx=20)
        row1 = tk.Frame(file_frame); row1.pack(fill="x")
        tk.Entry(row1, textvariable=self.caminho_excel_var, width=60).pack(side="left", padx=5)
        tk.Button(row1, text="Buscar Arquivo", command=self.selecionar_arquivo).pack(side="left", padx=5)

        row2 = tk.Frame(file_frame); row2.pack(fill="x", pady=5)
        tk.Label(row2, text="Letra Coluna:").pack(side="left", padx=5)
        tk.Entry(row2, textvariable=self.coluna_alvo, width=5).pack(side="left", padx=5)
        tk.Label(row2, text="Linha Início:").pack(side="left", padx=5)
        tk.Spinbox(row2, from_=1, to=100000, textvariable=self.linha_inicio, width=8).pack(side="left", padx=5)

        opts_frame = tk.LabelFrame(root, text=" Campos para Extração ", padx=10, pady=10)
        opts_frame.pack(pady=5, fill="x", padx=20)
        tk.Checkbutton(opts_frame, text="Concentração", variable=self.chk_concentracao).pack(side="left", padx=15)
        tk.Checkbutton(opts_frame, text="Tipo de Produto", variable=self.chk_tipo).pack(side="left", padx=15)
        tk.Checkbutton(opts_frame, text="Preço de Venda", variable=self.chk_preco).pack(side="left", padx=15)

        perf_frame = tk.LabelFrame(root, text=" Configuração de Performance ", padx=10, pady=5)
        perf_frame.pack(pady=5, fill="x", padx=20)
        tk.Label(perf_frame, text="Threads (Navegadores):").pack(side="left", padx=5)
        tk.Spinbox(perf_frame, from_=1, to=15, textvariable=self.num_threads_var, width=5).pack(side="left", padx=5)

        self.label_contador = tk.Label(root, text="Itens Processados: 0", font=("Consolas", 11, "bold"), fg="blue")
        self.label_contador.pack(pady=5)
        self.progress = ttk.Progressbar(root, orient="horizontal", length=750, mode="determinate")
        self.progress.pack(pady=5)
        self.log_area = scrolledtext.ScrolledText(root, width=100, height=12, font=("Consolas", 9), bg="#1e1e1e", fg="#00ff00")
        self.log_area.pack(pady=10)

        btn_frame = tk.Frame(root)
        btn_frame.pack(pady=10)
        self.btn_start = tk.Button(btn_frame, text="INICIAR", command=self.start_process, bg="green", fg="white", width=20, font=("Arial", 10, "bold"))
        self.btn_start.grid(row=0, column=0, padx=10)
        self.btn_stop = tk.Button(btn_frame, text="PARAR", command=self.stop_process, bg="red", fg="white", width=20, font=("Arial", 10, "bold"), state=tk.DISABLED)
        self.btn_stop.grid(row=0, column=1, padx=10)

    def selecionar_arquivo(self):
        caminho = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")])
        if caminho: self.caminho_excel_var.set(caminho)

    def write_log(self, text):
        self.log_area.insert(tk.END, f"{time.strftime('%H:%M:%S')} > {text}\n")
        self.log_area.see(tk.END)
        self.root.update_idletasks()

    def stop_process(self):
        self.rodando = False
        self.btn_stop.config(state=tk.DISABLED)

    def start_process(self):
        if not self.caminho_excel_var.get() or not self.user_erp.get() or not self.pass_erp.get():
            messagebox.showwarning("Erro", "Preencha Credenciais e selecione o Arquivo!")
            return
        self.rodando = True
        self.btn_start.config(state=tk.DISABLED)
        self.btn_stop.config(state=tk.NORMAL)
        threading.Thread(target=self.gerenciar_threads, daemon=True).start()

    def gerenciar_threads(self):
        try:
            caminho = self.caminho_excel_var.get()
            threads = self.num_threads_var.get()
            user = self.user_erp.get()
            pwd = self.pass_erp.get()
            letra_col = self.coluna_alvo.get().upper().strip()
            linha_idx = self.linha_inicio.get() - 1
            col_idx = ord(letra_col) - ord('A')

            df = pd.read_excel(caminho)
            col_c = df.iloc[linha_idx:, col_idx].astype(str).replace('nan', '').str.strip().tolist()
            col_d = df.iloc[linha_idx:, col_idx + 1].astype(str).replace('nan', '').str.strip().tolist()
            dados = list(zip(col_c, col_d))
            
            self.progress["maximum"] = len(dados)
            self.contador_global = 0
            
            chunk_size = (len(dados) // threads) + (1 if len(dados) % threads != 0 else 0)
            fatias = [dados[i:i + chunk_size] for i in range(0, len(dados), chunk_size)]

            with ThreadPoolExecutor(max_workers=threads) as executor:
                executor.map(lambda f: self.worker(f, user, pwd), fatias)
        except Exception as e:
            self.write_log(f"ERRO: {e}")
        finally:
            self.write_log("--- PROCESSO FINALIZADO ---")
            self.btn_start.config(state=tk.NORMAL)
            self.btn_stop.config(state=tk.DISABLED)
            self.rodando = False

    def worker(self, lista_tarefas, user, pwd):
        options = webdriver.ChromeOptions()
        options.add_argument("--headless=new")
        options.add_argument("--blink-settings=imagesEnabled=false")
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
        wait = WebDriverWait(driver, 12)

        try:
            driver.get("http://redemariano21476.ddns.com.br:4647/sgfpod1/Login.pod")
            wait.until(EC.presence_of_element_located((By.ID, "id_cod_usuario"))).send_keys(user)
            driver.find_element(By.ID, "nom_senha").send_keys(pwd)
            driver.find_element(By.ID, "login").click()
            time.sleep(5)
            driver.get("http://redemariano21476.ddns.com.br:4647/sgfpod1/Cad_0020.pod")

            for item_c, item_d in lista_tarefas:
                if not self.rodando: break
                if not item_c or item_c == '': continue
                
                try:
                    entrada = wait.until(EC.element_to_be_clickable((By.ID, "cod_redbarraEntrada")))
                    entrada.click()
                    entrada.send_keys(Keys.CONTROL + "a", Keys.BACKSPACE)
                    entrada.send_keys(item_c, Keys.ENTER)
                    time.sleep(0.6) 

                    linha = [f"Cód: {item_c}", f"Ref: {item_d}"]
                    
                    campos = [
                        ('Conc', 'produtoConcentracao_concentracao_nom_concentracao', self.chk_concentracao),
                        ('Tipo', 'cadtipro_tipo_nom_tipo', self.chk_tipo),
                        ('Preço', 'vlr_venda', self.chk_preco)
                    ]

                    for label, element_id, active in campos:
                        if active.get():
                            try:
                                val = driver.find_element(By.ID, element_id).get_attribute('value')
                                texto = val.strip() if (val and val.strip()) else "vazio"
                                linha.append(f"{label}: {texto}")
                            except:
                                linha.append(f"{label}: vazio")

                    resultado = " | ".join(linha)
                    with self.lock:
                        with open(ARQUIVO_RESULTADO, "a", encoding="utf-8") as f:
                            f.write(resultado + "\n")
                        self.contador_global += 1
                        self.label_contador.config(text=f"Itens Processados: {self.contador_global}")
                        self.progress["value"] = self.contador_global
                except: continue
        finally:
            driver.quit()

if __name__ == "__main__":
    root = tk.Tk()
    app = AutomationGUI(root)
    root.mainloop()
