import os
import re
import time
import threading
import random
import json
import subprocess
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import filedialog
from tkinter.scrolledtext import ScrolledText
from tkinter.simpledialog import askstring
import pystray
from PIL import Image, ImageDraw
import win32com.client

class LogViewerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Predadores Votemap Patch")
        self.root.geometry("1000x900")

        self.style = ttk.Style()
        self.style.theme_use("darkly")

        self.pasta_raiz = None
        self.pasta_atual = None
        self.caminho_log_atual = None
        self.arquivo_json = None
        self.arquivo_json_votemap = None
        self._stop = False
        self._paused = False
        self._interromper_leitura = False
        self.log_thread = None
        self.nome_servico = None
        self.file_log_handle = None

        top_frame = ttk.Frame(root)
        top_frame.pack(pady=10, padx=10, fill='x')

        self.selecionar_btn = ttk.Button(top_frame, text="Selecionar Pasta de Logs", command=self.selecionar_pasta, bootstyle=PRIMARY)
        self.selecionar_btn.pack(side='left')

        self.json_btn = ttk.Button(top_frame, text="Selecionar JSON do servidor", command=self.selecionar_arquivo_json_servidor, bootstyle=INFO)
        self.json_btn.pack(side='left', padx=5)

        self.json_vm_btn = ttk.Button(top_frame, text="Selecionar JSON do Votemap", command=self.selecionar_arquivo_json_votemap, bootstyle=INFO)
        self.json_vm_btn.pack(side='left', padx=5)

        self.servico_var = ttk.StringVar(value="")
        self.servico_btn = ttk.Button(top_frame, text="Selecionar Serviço", command=self.selecionar_servico, bootstyle=SECONDARY)
        self.servico_btn.pack(side='left', padx=5)
        self.servico_label = ttk.Label(top_frame, textvariable=self.servico_var, foreground="orange")
        self.servico_label.pack(side='left', padx=(5, 10))

        ttk.Label(top_frame, text="Filtro:").pack(side='left', padx=(15, 5))
        self.filtro_var = ttk.StringVar()
        self.filtro_entry = ttk.Entry(top_frame, textvariable=self.filtro_var, width=30)
        self.filtro_entry.pack(side='left')

        self.pausar_btn = ttk.Button(top_frame, text="⏸️ Pausar", command=self.toggle_pausa, bootstyle=WARNING)
        self.pausar_btn.pack(side='right', padx=5)

        self.limpar_btn = ttk.Button(top_frame, text="♻️ Limpar Tela", command=self.limpar_tela, bootstyle=SECONDARY)
        self.limpar_btn.pack(side='right', padx=5)

        ttk.Label(top_frame, text="Tema:").pack(side='right', padx=(10, 5))
        self.tema_var = ttk.StringVar(value='darkly')
        self.tema_menu = ttk.Combobox(top_frame, textvariable=self.tema_var, values=self.style.theme_names(), width=15, state='readonly')
        self.tema_menu.pack(side='right')
        self.tema_menu.bind("<<ComboboxSelected>>", self.trocar_tema)

        self.log_label = ttk.Label(root, text="LOG AO VIVO DO SERVIDOR", foreground="red")
        self.log_label.pack()
        self.text_area = ScrolledText(root, wrap='word', height=20)
        self.text_area.pack(padx=10, fill='both', expand=True)
        self.text_area.configure(state='disabled')

        self.json_title = ttk.Label(root, text="JSON CONFIG SERVIDOR", foreground="red")
        self.json_title.pack()
        self.json_text_area = ScrolledText(root, wrap='word', height=10)
        self.json_text_area.pack(padx=10, pady=(0, 10), fill='both')
        self.json_text_area.configure(state='disabled')

        self.json_vm_title = ttk.Label(root, text="JSON CONFIG VOTEMAP", foreground="red")
        self.json_vm_title.pack()
        self.json_vm_text_area = ScrolledText(root, wrap='word', height=10)
        self.json_vm_text_area.pack(padx=10, pady=(0, 10), fill='both')
        self.json_vm_text_area.configure(state='disabled')

    def selecionar_pasta(self):
        pasta = filedialog.askdirectory(title="Selecione a pasta raiz dos logs")
        if pasta:
            self.pasta_raiz = pasta
            self.append_text(f">>> Pasta selecionada: {self.pasta_raiz}\n")
            threading.Thread(target=self.monitorar_log, daemon=True).start()

    def selecionar_arquivo_json_servidor(self):
        caminho = filedialog.askopenfilename(filetypes=[("Arquivos JSON", "*.json")])
        if caminho:
            self.arquivo_json = caminho
            with open(caminho, 'r', encoding='utf-8') as f:
                conteudo = f.read()
            self.exibir_json(self.json_text_area, conteudo)

    def selecionar_arquivo_json_votemap(self):
        caminho = filedialog.askopenfilename(filetypes=[("Arquivos JSON", "*.json")])
        if caminho:
            self.arquivo_json_votemap = caminho
            with open(caminho, 'r', encoding='utf-8') as f:
                conteudo = f.read()
            self.exibir_json(self.json_vm_text_area, conteudo)

    def selecionar_servico(self):
        def executar():
            try:
                wmi = win32com.client.GetObject('winmgmts:')
                services = wmi.InstancesOf('Win32_Service')
                nomes = sorted([s.Name for s in services])

                selecao = ttk.Toplevel(self.root)
                selecao.title("Selecionar Serviço")
                selecao.geometry("400x120")

                ttk.Label(selecao, text="Escolha um serviço para reiniciar:", font=('Arial', 10)).pack(pady=10)
                combo = ttk.Combobox(selecao, values=nomes, state="readonly")
                combo.pack(pady=5)
                combo.set("Selecione...")

                def confirmar():
                    escolha = combo.get()
                    if escolha and escolha != "Selecione...":
                        self.nome_servico = escolha
                        self.servico_var.set(f"Serviço selecionado: {escolha}")
                        selecao.destroy()

                ttk.Button(selecao, text="Confirmar", command=confirmar, bootstyle=PRIMARY).pack(pady=5)
            except Exception as e:
                self.append_text(f"Erro ao listar serviços: {e}\n")

        threading.Thread(target=executar, daemon=True).start()

    def exibir_json(self, area, conteudo):
        area.configure(state='normal')
        area.delete('1.0', 'end')
        area.insert('end', conteudo)
        area.configure(state='disabled')

    def monitorar_log(self):
        while not self._stop:
            nova_pasta = self.obter_subpasta_mais_recente()
            novo_log = os.path.join(nova_pasta, 'console.log') if nova_pasta else None
            if nova_pasta and nova_pasta != self.pasta_atual and os.path.exists(novo_log):
                self.pasta_atual = nova_pasta
                self.caminho_log_atual = novo_log
                self.log_label.config(text=f"LOG AO VIVO DO SERVIDOR  |  {self.caminho_log_atual}")
                self.append_text(f"\n>>> Nova pasta identificada: {self.pasta_atual}\n")

                if self.log_thread and self.log_thread.is_alive():
                    self._interromper_leitura = True
                    time.sleep(0.5)
                    self._interromper_leitura = False

                if self.file_log_handle:
                    self.file_log_handle.close()

                self.file_log_handle = open(self.caminho_log_atual, 'r', encoding='utf-8')
                self.file_log_handle.seek(0, os.SEEK_END)
                self.log_thread = threading.Thread(target=self.acompanhar_log, daemon=True)
                self.log_thread.start()
            time.sleep(20)

    def obter_subpasta_mais_recente(self):
        subpastas = [os.path.join(self.pasta_raiz, nome) for nome in os.listdir(self.pasta_raiz)]
        subpastas = [p for p in subpastas if os.path.isdir(p)]
        if not subpastas:
            return None
        return max(subpastas, key=os.path.getmtime)

    def acompanhar_log(self):
        aguardando_winner = False
        try:
            while not self._interromper_leitura and not self._stop:
                if self._paused:
                    time.sleep(0.5)
                    continue

                linha = self.file_log_handle.readline()
                if linha:
                    filtro = self.filtro_var.get().strip().lower()
                    if filtro == "" or filtro in linha.lower():
                        self.append_text(linha)

                    if ".EndVote()" in linha:
                        aguardando_winner = True

                    if aguardando_winner and "Winner: [" in linha:
                        match = re.search(r"Winner: \[(\d+)\]", linha)
                        if match:
                            indice = int(match.group(1))
                            try:
                                with open(self.arquivo_json_votemap, 'r', encoding='utf-8') as f_vm:
                                    votemap_data = json.load(f_vm)
                                lista_mapas = votemap_data.get("list", [])
                                if indice == 0:
                                    if len(lista_mapas) <= 1:
                                        raise ValueError("Não há mapas suficientes para seleção aleatória.")
                                    indice = random.randint(1, len(lista_mapas) - 1)
                                novo_scenario = lista_mapas[indice]
                                with open(self.arquivo_json, 'r', encoding='utf-8') as f_srv:
                                    servidor_data = json.load(f_srv)
                                servidor_data["game"]["scenarioId"] = novo_scenario
                                with open(self.arquivo_json, 'w', encoding='utf-8') as f_srv_out:
                                    json.dump(servidor_data, f_srv_out, indent=4)
                                self.exibir_json(self.json_text_area, json.dumps(servidor_data, indent=4))

                                if self.nome_servico:
                                    subprocess.run(["sc", "stop", self.nome_servico], shell=True)
                                    time.sleep(3)
                                    subprocess.run(["sc", "start", self.nome_servico], shell=True)
                                    time.sleep(15)
                                    with open(self.arquivo_json, 'r', encoding='utf-8') as f_srv:
                                        servidor_data = json.load(f_srv)
                                    servidor_data["game"]["scenarioId"] = "{B88CC33A14B71FDC}Missions/V30_MapVoting_Mission.conf"
                                    with open(self.arquivo_json, 'w', encoding='utf-8') as f_srv:
                                        json.dump(servidor_data, f_srv, indent=4)
                                    self.exibir_json(self.json_text_area, json.dumps(servidor_data, indent=4))

                            except Exception as e:
                                self.append_text(f"Erro ao atualizar JSON do servidor: {e}\n")
                        aguardando_winner = False
                else:
                    time.sleep(0.5)
        except Exception as e:
            self.append_text(f"Erro ao acompanhar log: {e}\n")

    def append_text(self, texto):
        self.text_area.configure(state='normal')
        self.text_area.insert('end', texto)
        self.text_area.yview_moveto(1.0)
        self.text_area.configure(state='disabled')

    def limpar_tela(self):
        self.text_area.configure(state='normal')
        self.text_area.delete('1.0', 'end')
        self.text_area.configure(state='disabled')

    def toggle_pausa(self):
        self._paused = not self._paused
        if self._paused:
            self.pausar_btn.config(text="▶️ Retomar", bootstyle=SUCCESS)
        else:
            self.pausar_btn.config(text="⏸️ Pausar", bootstyle=WARNING)

    def trocar_tema(self, event=None):
        tema = self.tema_var.get()
        self.style.theme_use(tema)

    def on_close(self):
        self._stop = True
        self._interromper_leitura = True
        if self.file_log_handle:
            self.file_log_handle.close()
        self.root.destroy()


def main():
    root = ttk.Window(themename="darkly")
    app = LogViewerApp(root)
    root.protocol("WM_DELETE_WINDOW", app.on_close)
    root.mainloop()

if __name__ == '__main__':
    main()