import pandas as pd
import time
import threading
import os
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

class AutomationGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Extração PRO - Todos os Itens")
        self.root.geometry("950x600")
        
        self.rodando = False
        self.pausado = threading.Event()
        self.pausado.set() 
        self.contador_global = 0
        
        self.lock_contador = threading.Lock()
        self.lock_arquivo = threading.Lock()
        
        self.user_erp = tk.StringVar()
        self.pass_erp = tk.StringVar()
        self.caminho_excel = tk.StringVar()
        self.pasta_destino = tk.StringVar(value=os.getcwd())
        self.num_threads = tk.IntVar(value=4)
        
        self.chk_concentracao = tk.BooleanVar(value=True)
        self.chk_tipo = tk.BooleanVar(value=True)
        self.chk_preco = tk.BooleanVar(value=True)

        self.setup_ui()

    def setup_ui(self):
        tk.Label(self.root, text="Automação de Extração ERP - Coleta Completa", font=("Arial", 14, "bold")).pack(pady=10)
        
        # Frame de Login
        f_login = tk.LabelFrame(self.root, text=" Credenciais ", padx=10, pady=5)
        f_login.pack(fill="x", padx=20, pady=5)
        tk.Label(f_login, text="Usuário:").grid(row=0, column=0)
        tk.Entry(f_login, textvariable=self.user_erp, width=20).grid(row=0, column=1, padx=5)
        tk.Label(f_login, text="Senha:").grid(row=0, column=2)
        tk.Entry(f_login, textvariable=self.pass_erp, width=20, show="*").grid(row=0, column=3, padx=5)

        # Configurações de Arquivo
        f_files = tk.LabelFrame(self.root, text=" Configurações de Entrada/Saída ", padx=10, pady=5)
        f_files.pack(fill="x", padx=20, pady=5)
        tk.Button(f_files, text="Selecionar Excel", command=self.selecionar_excel).grid(row=0, column=0, pady=5, sticky="w")
        tk.Label(f_files, textvariable=self.caminho_excel, fg="blue", wraplength=500).grid(row=0, column=1, padx=10, sticky="w")
        tk.Button(f_files, text="Pasta de Destino", command=self.selecionar_pasta).grid(row=1, column=0, pady=5, sticky="w")
        tk.Label(f_files, textvariable=self.pasta_destino, fg="green").grid(row=1, column=1, padx=10, sticky="w")
        
        tk.Label(f_files, text="Nº de Threads:").grid(row=2, column=0, pady=5, sticky="w")
        tk.Spinbox(f_files, from_=1, to=20, textvariable=self.num_threads, width=5).grid(row=2, column=1, sticky="w", padx=10)

        # Opções de Extração
        f_opts = tk.LabelFrame(self.root, text=" Campos de Interesse ", padx=10, pady=10)
        f_opts.pack(fill="x", padx=20, pady=5)
        tk.Checkbutton(f_opts, text="Equivalente", variable=self.chk_concentracao).pack(side="left", padx=20)
        tk.Checkbutton(f_opts, text="Tipo Produto", variable=self.chk_tipo).pack(side="left", padx=20)
        tk.Checkbutton(f_opts, text="Preço Venda", variable=self.chk_preco).pack(side="left", padx=20)

        # Feedback Visual
        self.progress = ttk.Progressbar(self.root, length=700, mode="determinate")
        self.progress.pack(pady=10)
        self.label_cont = tk.Label(self.root, text="Itens Processados: 0", font=("Arial", 10, "bold"))
        self.label_cont.pack()
        
        self.log_area = scrolledtext.ScrolledText(self.root, height=12, bg="#1e1e1e", fg="#d4d4d4", font=("Consolas", 9))
        self.log_area.pack(padx=20, pady=10, fill="both", expand=True)

        # Botões de Controle
        btn_f = tk.Frame(self.root)
        btn_f.pack(pady=10)
        self.btn_start = tk.Button(btn_f, text="INICIAR", bg="#28a745", fg="white", width=15, command=self.start)
        self.btn_start.grid(row=0, column=0, padx=5)
        self.btn_pause = tk.Button(btn_f, text="PAUSAR", bg="#ffc107", width=15, command=self.toggle_pause, state="disabled")
        self.btn_pause.grid(row=0, column=1, padx=5)
        self.btn_stop = tk.Button(btn_f, text="PARAR", bg="#dc3545", fg="white", width=15, command=self.stop, state="disabled")
        self.btn_stop.grid(row=0, column=2, padx=5)

    def log(self, msg):
        self.log_area.insert(tk.END, f"[{time.strftime('%H:%M:%S')}] {msg}\n")
        self.log_area.see(tk.END)

    def selecionar_excel(self):
        p = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if p: self.caminho_excel.set(p)

    def selecionar_pasta(self):
        p = filedialog.askdirectory()
        if p: self.pasta_destino.set(p)

    def toggle_pause(self):
        if self.pausado.is_set():
            self.pausado.clear()
            self.btn_pause.config(text="RETOMAR", bg="#007bff", fg="white")
        else:
            self.pausado.set()
            self.btn_pause.config(text="PAUSAR", bg="#ffc107", fg="black")

    def stop(self):
        self.rodando = False
        self.pausado.set()
        self.log("Solicitando parada do processo...")

    def update_ui_counters(self, valor):
        self.label_cont.config(text=f"Processados: {valor}")
        self.progress["value"] = valor

    def start(self):
        if not self.caminho_excel.get() or not self.user_erp.get():
            messagebox.showwarning("Atenção", "Preencha o login e selecione o arquivo Excel.")
            return
        
        self.rodando = True
        self.contador_global = 0
        self.btn_start.config(state="disabled")
        self.btn_pause.config(state="normal")
        self.btn_stop.config(state="normal")
        
        self.caminho_final_txt = os.path.join(self.pasta_destino.get(), "extracao_completa.txt")
        with open(self.caminho_final_txt, "w", encoding="utf-8") as f:
            f.write(f"--- RELATÓRIO COMPLETO: {time.strftime('%d/%m/%Y %H:%M:%S')} ---\n\n")
        
        threading.Thread(target=self.motor_principal, daemon=True).start()

    def motor_principal(self):
        try:
            self.log("--- INICIANDO DIAGNÓSTICO DO EXCEL ---")
            caminho = self.caminho_excel.get()
            
            if not os.path.exists(caminho):
                self.log(f"ERRO: Arquivo não encontrado em: {caminho}")
                return

            # Carrega o Excel
            df = pd.read_excel(caminho)
            total_linhas_arquivo = len(df)
            self.log(f"Sucesso: Arquivo lido. Total de linhas no arquivo: {total_linhas_arquivo}")
            self.log(f"Colunas detectadas: {list(df.columns)}")

            # Define o intervalo (10.000 a 21.000)
            inicio, fim = 10000, 21000
            
            # Ajuste de segurança: se o arquivo for menor que o fim, ele pega até o último
            dados = df.iloc[inicio:min(fim, total_linhas_arquivo)].values.tolist()
            total_para_processar = len(dados)

            if total_para_processar == 0:
                self.log(f"ALERTA: A fatia [{inicio}:{fim}] retornou ZERO itens.")
                self.log("Verifique se o seu Excel tem mais de 10.000 linhas.")
                return

            self.log(f"Pronto para processar {total_para_processar} itens.")
            self.root.after(0, lambda: self.progress.config(maximum=total_para_processar))

            qtd_threads = self.num_threads.get()
            fatia = (total_para_processar // qtd_threads) + 1
            chunks = [dados[i:i + fatia] for i in range(0, total_para_processar, fatia)]

            with ThreadPoolExecutor(max_workers=qtd_threads) as executor:
                for i, chunk in enumerate(chunks):
                    if len(chunk) > 0:
                        executor.submit(self.worker, chunk, i+1)

        except Exception as e:
            self.log(f"ERRO CRÍTICO AO LER EXCEL: {str(e)}")
        finally:
            self.log("Motor de agendamento finalizado. Aguardando threads...")

    def worker(self, lista, worker_id):
        self.log(f"Thread {worker_id}: Iniciando com {len(lista)} itens.")
        
        options = webdriver.ChromeOptions()
        options.add_argument("--headless=new") 
        options.add_argument("--blink-settings=imagesEnabled=false")
        
        driver = None
        try:
            driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
            wait = WebDriverWait(driver, 20)

            self.log(f"Thread {worker_id}: Acessando página de login...")
            driver.get("https://linkEmpresa/login")
            
            user_field = wait.until(EC.presence_of_element_located((By.ID, "id_cod_usuario")))
            user_field.send_keys(self.user_erp.get())
            driver.find_element(By.ID, "nom_senha").send_keys(self.pass_erp.get())
            btn_login = wait.until(EC.presence_of_element_located((By.ID, "login")))
            driver.execute_script("arguments[0].click();", btn_login)
            
            self.log(f"Thread {worker_id}: Login enviado. Aguardando carregamento...")
            time.sleep(3)
            
            driver.get("https://linkEmpresa/empresa")
            self.log(f"Thread {worker_id}: Página de extração carregada.")

            for item in lista:
                if not self.rodando: break
                self.pausado.wait()

                try:
                    cod = str(item[2]).strip()
                    ref = str(item[3]).strip()
                    
                    if cod == 'nan' or not cod:
                        continue

                    driver.execute_script("document.getElementById('cadequiv_equivalente_nom_equivalente').value = '';")
                    
                    input_busca = wait.until(EC.element_to_be_clickable((By.ID, "cod_redbarraEntrada")))
                    input_busca.click()
                    input_busca.send_keys(Keys.CONTROL + "a", Keys.BACKSPACE)
                    input_busca.send_keys(cod, Keys.ENTER)
                    
                    time.sleep(1.2)

                    res_conc = driver.execute_script("return document.getElementById('cadequiv_equivalente_nom_equivalente').value;")
                    res_conc = str(res_conc).strip() if res_conc else "vazio"

                    with self.lock_arquivo:
                        with open(self.caminho_final_txt, "a", encoding="utf-8") as f:
                            f.write(f"Cód: {cod} | Ref: {ref} | Conc: {res_conc}\n")

                    with self.lock_contador:
                        self.contador_global += 1
                        val = self.contador_global
                        self.root.after(0, lambda v=val: self.update_ui_counters(v))

                except Exception as e_item:
                    self.log(f"Thread {worker_id}: Erro no item {item[2]}: {str(e_item)[:50]}")
                    continue

        except Exception as e_geral:
            self.log(f"Thread {worker_id}: ERRO FATAL: {str(e_geral)}")
        finally:
            if driver:
                driver.quit()
            self.log(f"Thread {worker_id}: Finalizada.")

if __name__ == "__main__":
    root = tk.Tk()
    app = AutomationGUI(root)
    root.mainloop()
