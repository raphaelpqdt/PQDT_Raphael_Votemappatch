import os
import re
import time
import threading
import random
import json
import subprocess
import logging
import platform
import tkinter as tk
import sys  # Para sys._MEIPASS

import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap.tooltip import ToolTip
from ttkbootstrap.dialogs import Messagebox
from tkinter import filedialog
from tkinter.scrolledtext import ScrolledText

import pystray
from PIL import Image, ImageDraw
import win32com.client
import pythoncom  # <--- ADICIONADO PARA INICIALIZAÇÃO COM

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(module)s.%(funcName)s - %(message)s',
    filename='votemap_patch.log',
    filemode='a'
)


def resource_path(relative_path):
    """ Obtém o caminho absoluto para o recurso, funciona para dev e para PyInstaller """
    try:
        # PyInstaller cria uma pasta temp e armazena o caminho em _MEIPASS
        base_path = sys._MEIPASS
    except AttributeError:
        # sys._MEIPASS não existe, então estamos rodando em modo de desenvolvimento
        # Use o diretório do script atual
        base_path = os.path.abspath(os.path.dirname(__file__))
    return os.path.join(base_path, relative_path)


# --- CAMINHO DO ÍCONE ---
ICON_FILENAME = "pred.ico"  # Nome do arquivo de ícone na mesma pasta do script
ICON_PATH = resource_path(ICON_FILENAME)


class LogViewerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Predadores Votemap Patch")
        self.root.geometry("1000x900")

        self.app_icon_image = None  # Para manter referência ao PhotoImage, se usado
        self.set_application_icon()  # Define o ícone da janela

        self.config_file = "votemap_config.json"
        self.config = self.load_config()

        self.style = ttk.Style()
        self.style.theme_use(self.config.get("theme", "darkly"))

        self.pasta_raiz = self.config.get("log_folder", None)
        self.pasta_atual = None
        self.caminho_log_atual = None
        self.arquivo_json = self.config.get("server_json", None)
        self.arquivo_json_votemap = self.config.get("votemap_json", None)
        self.nome_servico = self.config.get("service_name", None)

        self._stop_event = threading.Event()
        self._paused = False
        self.log_tail_thread = None
        self.log_monitor_thread = None
        self.file_log_handle = None
        self.status_label_var = ttk.StringVar(value="Aguardando configuração...")

        self.create_menu()
        self.create_ui()
        self.create_status_bar()
        self.initialize_from_config()

        self.root.after(100, self.atualizar_log_sistema_periodicamente)

        if self.pasta_raiz:
            self.start_log_monitoring()

        logging.info("Aplicação iniciada.")

    def set_application_icon(self):
        """Define o ícone da aplicação para a janela principal."""
        global ICON_PATH
        try:
            if not os.path.exists(ICON_PATH):
                logging.warning(f"Arquivo de ícone '{ICON_PATH}' (resolvido de '{ICON_FILENAME}') não encontrado.")
                return

            if platform.system() == "Windows":
                self.root.iconbitmap(ICON_PATH)
                logging.info(f"Ícone da aplicação (Windows) definido a partir de: {ICON_PATH}")
            else:
                try:
                    img_pil = Image.open(ICON_PATH)
                    img_pil_rgba = img_pil.convert("RGBA")
                    self.app_icon_image = tk.PhotoImage(data=img_pil_rgba.tobytes("raw", "RGBA"))
                    self.root.iconphoto(True, self.app_icon_image)
                    logging.info(f"Ícone da aplicação (não-Windows, via Pillow) definido a partir de: {ICON_PATH}")
                except ImportError:
                    logging.warning(
                        "Pillow não está instalado. Tentando PhotoImage diretamente para o ícone (pode falhar para .ico).")
                    try:
                        self.app_icon_image = tk.PhotoImage(file=ICON_PATH)
                        self.root.iconphoto(True, self.app_icon_image)
                        logging.info(
                            f"Ícone da aplicação (não-Windows, PhotoImage direto) definido a partir de: {ICON_PATH}")
                    except tk.TclError as e_tk:
                        logging.error(
                            f"Erro TclError ao definir ícone em {platform.system()} com '{ICON_PATH}': {e_tk}")
                except Exception as e_pil:
                    logging.warning(
                        f"Falha ao carregar '{ICON_PATH}' com Pillow para {platform.system()}: {e_pil}. PhotoImage pode não suportar .ico diretamente.")
        except tk.TclError as e:
            logging.error(
                f"Erro TclError ao definir o ícone da aplicação: {e}. Certifique-se que o arquivo de imagem é válido.",
                exc_info=False)
        except Exception as e:
            logging.error(f"Erro geral ao definir o ícone da aplicação: {e}", exc_info=True)

    def _create_tray_image(self):
        """Cria uma imagem para o ícone da bandeja do sistema usando o ICON_PATH."""
        global ICON_PATH
        try:
            if os.path.exists(ICON_PATH):
                logging.info(f"Carregando ícone da bandeja de: {ICON_PATH}")
                return Image.open(ICON_PATH)
            else:
                logging.warning(f"Arquivo de ícone da bandeja '{ICON_PATH}' não encontrado. Desenhando um padrão.")
        except ImportError:
            logging.warning(
                "Pillow (PIL) não está instalado. Não é possível carregar o ícone da bandeja do arquivo. Desenhando padrão.")
        except Exception as e:
            logging.error(f"Erro ao carregar ícone da bandeja de '{ICON_PATH}': {e}. Desenhando um padrão.")

        width, height = 64, 64
        image = Image.new('RGBA', (width, height), (0, 0, 0, 0))
        draw = ImageDraw.Draw(image)
        draw.line((18, 15, 18, 49), fill='deepskyblue', width=7)
        draw.arc((18, 15, 40, 32), 270, 90, fill='deepskyblue', width=7)
        draw.arc((18, 32, 40, 49), 90, 270, fill='deepskyblue', width=7)
        return image

    def start_log_monitoring(self):
        if self.log_monitor_thread and self.log_monitor_thread.is_alive():
            logging.warning("Tentativa de iniciar monitoramento de log já em execução.")
            return
        self._stop_event.clear()
        self.log_monitor_thread = threading.Thread(target=self.monitorar_log_continuamente, daemon=True)
        self.log_monitor_thread.start()
        logging.info("Monitoramento de logs iniciado.")

    def stop_log_monitoring(self):
        self._stop_event.set()
        if self.log_tail_thread and self.log_tail_thread.is_alive():
            self.log_tail_thread.join(timeout=1.0)
        if self.log_monitor_thread and self.log_monitor_thread.is_alive():
            self.log_monitor_thread.join(timeout=1.0)
        if self.file_log_handle:
            try:
                self.file_log_handle.close()
                self.file_log_handle = None
            except Exception as e:
                logging.error(f"Erro ao fechar handle do arquivo de log: {e}", exc_info=True)
        logging.info("Monitoramento de logs parado.")

    def atualizar_log_sistema_periodicamente(self):
        try:
            if not self.notebook.winfo_exists(): return
            aba_atual_index = self.notebook.index(self.notebook.select())
            nome_aba_atual = self.notebook.tab(aba_atual_index, "text")
            if nome_aba_atual == "Log do Sistema":
                if os.path.exists('votemap_patch.log'):
                    with open('votemap_patch.log', 'r', encoding='utf-8') as f:
                        conteudo = f.read()
                    self.system_log_text_area.configure(state='normal')
                    pos_atual_scroll = self.system_log_text_area.yview()[1]
                    self.system_log_text_area.delete('1.0', 'end')
                    self.system_log_text_area.insert('end', conteudo)
                    if pos_atual_scroll == 1.0: self.system_log_text_area.yview_moveto(1.0)
                    self.system_log_text_area.configure(state='disabled')
                else:
                    self.system_log_text_area.configure(state='normal')
                    self.system_log_text_area.delete('1.0', 'end')
                    self.system_log_text_area.insert('end', "Arquivo 'votemap_patch.log' não encontrado.")
                    self.system_log_text_area.configure(state='disabled')
        except Exception as e:
            if not hasattr(self, "_system_log_update_error_count") or self._system_log_update_error_count < 5:
                logging.error(f"Erro ao atualizar log do sistema na GUI: {e}", exc_info=False)
                self._system_log_update_error_count = getattr(self, "_system_log_update_error_count", 0) + 1
        if not self._stop_event.is_set():
            if self.root.winfo_exists():
                self.root.after(3000, self.atualizar_log_sistema_periodicamente)

    def create_menu(self):
        menubar = ttk.Menu(self.root)
        self.root.config(menu=menubar)
        file_menu = ttk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Arquivo", menu=file_menu)
        file_menu.add_command(label="Salvar Configuração", command=self.save_config)
        file_menu.add_command(label="Carregar Configuração", command=self.load_config_dialog)
        file_menu.add_separator()
        file_menu.add_command(label="Sair", command=self.on_close)
        tools_menu = ttk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Ferramentas", menu=tools_menu)
        tools_menu.add_command(label="Exportar Logs do App", command=self.export_display_logs)
        tools_menu.add_command(label="Verificar Configurações", command=self.validate_configs)
        help_menu = ttk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Ajuda", menu=help_menu)
        help_menu.add_command(label="Sobre", command=self.show_about)

    def create_ui(self):
        top_frame = ttk.Frame(self.root)
        top_frame.pack(pady=10, padx=10, fill='x')
        self.selecionar_btn = ttk.Button(top_frame, text="Selecionar Pasta de Logs", command=self.selecionar_pasta,
                                         bootstyle=PRIMARY)
        self.selecionar_btn.pack(side='left')
        ToolTip(self.selecionar_btn, text="Seleciona a pasta raiz onde os logs do servidor são armazenados.")
        self.json_btn = ttk.Button(top_frame, text="JSON do Servidor", command=self.selecionar_arquivo_json_servidor,
                                   bootstyle=INFO)
        self.json_btn.pack(side='left', padx=5)
        ToolTip(self.json_btn, text="Seleciona o arquivo JSON de configuração principal do servidor.")
        self.json_vm_btn = ttk.Button(top_frame, text="JSON do Votemap", command=self.selecionar_arquivo_json_votemap,
                                      bootstyle=INFO)
        self.json_vm_btn.pack(side='left', padx=5)
        ToolTip(self.json_vm_btn, text="Seleciona o arquivo JSON de configuração do Votemap.")
        self.servico_btn = ttk.Button(top_frame, text="Selecionar Serviço", command=self.selecionar_servico,
                                      bootstyle=SECONDARY)
        self.servico_btn.pack(side='left', padx=5)
        ToolTip(self.servico_btn, text="Seleciona o serviço do Windows associado ao servidor do jogo.")
        self.servico_var = ttk.StringVar(value="Nenhum serviço selecionado")
        if self.nome_servico: self.servico_var.set(f"Serviço: {self.nome_servico}")
        self.servico_label = ttk.Label(top_frame, textvariable=self.servico_var, foreground="orange")
        self.servico_label.pack(side='left', padx=(0, 10))
        self.refresh_btn = ttk.Button(top_frame, text="Atualizar JSONs", command=self.forcar_refresh_json,
                                      bootstyle=SUCCESS)
        self.refresh_btn.pack(side='left', padx=5)
        ToolTip(self.refresh_btn, text="Recarrega e exibe o conteúdo dos arquivos JSON selecionados.")
        ttk.Label(top_frame, text="Filtro:").pack(side='left', padx=(15, 5))
        self.filtro_var = ttk.StringVar(value=self.config.get("filter", ""))
        self.filtro_entry = ttk.Entry(top_frame, textvariable=self.filtro_var, width=30)
        self.filtro_entry.pack(side='left')
        ToolTip(self.filtro_entry,
                text="Filtra as linhas de log exibidas (case-insensitive). Deixe em branco para nenhum filtro.")
        self.pausar_btn = ttk.Button(top_frame, text="⏸️ Pausar", command=self.toggle_pausa, bootstyle=WARNING)
        self.pausar_btn.pack(side='right', padx=5)
        ToolTip(self.pausar_btn, text="Pausa ou retoma o acompanhamento ao vivo dos logs.")
        self.limpar_btn = ttk.Button(top_frame, text="♻️ Limpar Tela", command=self.limpar_tela, bootstyle=SECONDARY)
        self.limpar_btn.pack(side='right', padx=5)
        ToolTip(self.limpar_btn, text="Limpa a área de exibição de logs do servidor.")
        ttk.Label(top_frame, text="Tema:").pack(side='right', padx=(10, 5))
        self.tema_var = ttk.StringVar(value=self.config.get("theme", "darkly"))
        self.tema_menu = ttk.Combobox(top_frame, textvariable=self.tema_var, values=self.style.theme_names(), width=15,
                                      state='readonly')
        self.tema_menu.pack(side='right')
        self.tema_menu.bind("<<ComboboxSelected>>", self.trocar_tema)
        ToolTip(self.tema_menu, text="Muda o tema visual da aplicação.")
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill='both', expand=True, padx=10, pady=10)
        log_frame = ttk.Frame(self.notebook)
        self.notebook.add(log_frame, text="Logs do Servidor")
        self.log_label = ttk.Label(log_frame, text="LOG AO VIVO DO SERVIDOR", foreground="red")
        self.log_label.pack(pady=(5, 0))
        self.text_area = ScrolledText(log_frame, wrap='word', height=20, state='disabled')
        self.text_area.pack(fill='both', expand=True)
        server_json_frame = ttk.Frame(self.notebook)
        self.notebook.add(server_json_frame, text="Config. Servidor")
        self.json_title = ttk.Label(server_json_frame, text="JSON CONFIG SERVIDOR", foreground="red")
        self.json_title.pack(pady=(5, 0))
        self.json_text_area = ScrolledText(server_json_frame, wrap='word', height=20, state='disabled')
        self.json_text_area.pack(fill='both', expand=True)
        votemap_json_frame = ttk.Frame(self.notebook)
        self.notebook.add(votemap_json_frame, text="Config. Votemap")
        self.json_vm_title = ttk.Label(votemap_json_frame, text="JSON CONFIG VOTEMAP", foreground="red")
        self.json_vm_title.pack(pady=(5, 0))
        self.json_vm_text_area = ScrolledText(votemap_json_frame, wrap='word', height=20, state='disabled')
        self.json_vm_text_area.pack(fill='both', expand=True)
        log_sistema_frame = ttk.Frame(self.notebook)
        self.notebook.add(log_sistema_frame, text="Log do Sistema")
        self.system_log_text_area = ScrolledText(log_sistema_frame, wrap='word', height=20, state='disabled')
        self.system_log_text_area.pack(fill='both', expand=True)
        settings_frame = ttk.Frame(self.notebook)
        self.notebook.add(settings_frame, text="Configurações App")
        settings_inner_frame = ttk.Frame(settings_frame, padding=20)
        settings_inner_frame.pack(fill='both', expand=True)
        self.auto_restart_var = ttk.BooleanVar(value=self.config.get("auto_restart", True))
        auto_restart_check = ttk.Checkbutton(settings_inner_frame,
                                             text="Reiniciar servidor automaticamente após troca de mapa",
                                             variable=self.auto_restart_var)
        auto_restart_check.grid(row=0, column=0, sticky='w', padx=5, pady=5, columnspan=2)
        ToolTip(auto_restart_check, "Se marcado, o servidor será reiniciado após uma votação de mapa bem-sucedida.")
        ttk.Label(settings_inner_frame, text="Padrão de detecção de voto (RegEx):").grid(row=1, column=0, sticky='w',
                                                                                         padx=5, pady=(15, 0))
        self.vote_pattern_var = ttk.StringVar(value=self.config.get("vote_pattern", r"\.EndVote\(\)"))
        vote_pattern_entry = ttk.Entry(settings_inner_frame, textvariable=self.vote_pattern_var, width=50)
        vote_pattern_entry.grid(row=2, column=0, sticky='ew', padx=5, pady=5, columnspan=2)
        ToolTip(vote_pattern_entry, "Expressão regular para detectar o fim de uma votação no log.")
        ttk.Label(settings_inner_frame, text="Padrão de detecção de vencedor (RegEx):").grid(row=3, column=0,
                                                                                             sticky='w', padx=5,
                                                                                             pady=(15, 0))
        self.winner_pattern_var = ttk.StringVar(value=self.config.get("winner_pattern", r"Winner: \[(\d+)\]"))
        winner_pattern_entry = ttk.Entry(settings_inner_frame, textvariable=self.winner_pattern_var, width=50)
        winner_pattern_entry.grid(row=4, column=0, sticky='ew', padx=5, pady=5, columnspan=2)
        ToolTip(winner_pattern_entry,
                "Expressão regular para capturar o índice do mapa vencedor (o primeiro grupo de captura é usado).")
        ttk.Label(settings_inner_frame, text="Missão padrão de votemap:").grid(row=5, column=0, sticky='w', padx=5,
                                                                               pady=(15, 0))
        self.default_mission_var = ttk.StringVar(
            value=self.config.get("default_mission", "{B88CC33A14B71FDC}Missions/V30_MapVoting_Mission.conf"))
        default_mission_entry = ttk.Entry(settings_inner_frame, textvariable=self.default_mission_var, width=70)
        default_mission_entry.grid(row=6, column=0, sticky='ew', padx=5, pady=5, columnspan=2)
        ToolTip(default_mission_entry,
                "ID do cenário/missão a ser carregado após um reinício para iniciar uma nova votação.")
        ttk.Label(settings_inner_frame, text="Delay ao parar serviço (s):").grid(row=7, column=0, sticky='w', padx=5,
                                                                                 pady=(15, 0))
        self.stop_delay_var = ttk.IntVar(value=self.config.get("stop_delay", 3))
        stop_delay_spinbox = ttk.Spinbox(settings_inner_frame, from_=1, to=60, textvariable=self.stop_delay_var,
                                         width=7)
        stop_delay_spinbox.grid(row=8, column=0, sticky='w', padx=5, pady=5)
        ToolTip(stop_delay_spinbox, "Tempo em segundos para aguardar após enviar o comando de parada ao serviço.")
        ttk.Label(settings_inner_frame, text="Delay ao iniciar serviço (s):").grid(row=7, column=1, sticky='w', padx=5,
                                                                                   pady=(15, 0))
        self.start_delay_var = ttk.IntVar(value=self.config.get("start_delay", 15))
        start_delay_spinbox = ttk.Spinbox(settings_inner_frame, from_=5, to=120, textvariable=self.start_delay_var,
                                          width=7)
        start_delay_spinbox.grid(row=8, column=1, sticky='w', padx=5, pady=5)
        ToolTip(start_delay_spinbox,
                "Tempo em segundos para aguardar o servidor iniciar completamente após o comando de início.")
        ttk.Button(settings_inner_frame, text="Salvar Configurações da App", command=self.save_config,
                   bootstyle=SUCCESS).grid(row=9, column=0, columnspan=2, pady=20)

    def forcar_refresh_json(self):
        refreshed_server = False
        refreshed_votemap = False
        try:
            if self.arquivo_json and os.path.exists(self.arquivo_json):
                with open(self.arquivo_json, 'r', encoding='utf-8') as f:
                    conteudo = f.read()
                self.exibir_json(self.json_text_area, conteudo)
                self.append_text_gui(f"JSON do servidor '{os.path.basename(self.arquivo_json)}' recarregado.\n")
                refreshed_server = True
            elif self.arquivo_json:
                self.append_text_gui(f"Arquivo JSON do servidor '{self.arquivo_json}' não encontrado.\n")
                self.exibir_json(self.json_text_area, "Arquivo não encontrado.")
            else:
                self.append_text_gui("Caminho do JSON do servidor não configurado.\n")
            if self.arquivo_json_votemap and os.path.exists(self.arquivo_json_votemap):
                with open(self.arquivo_json_votemap, 'r', encoding='utf-8') as f:
                    conteudo = f.read()
                self.exibir_json(self.json_vm_text_area, conteudo)
                self.append_text_gui(f"JSON do votemap '{os.path.basename(self.arquivo_json_votemap)}' recarregado.\n")
                refreshed_votemap = True
            elif self.arquivo_json_votemap:
                self.append_text_gui(f"Arquivo JSON do votemap '{self.arquivo_json_votemap}' não encontrado.\n")
                self.exibir_json(self.json_vm_text_area, "Arquivo não encontrado.")
            else:
                self.append_text_gui("Caminho do JSON do votemap não configurado.\n")
            if refreshed_server or refreshed_votemap:
                self.status_label_var.set("Arquivos JSON recarregados.")
            else:
                self.status_label_var.set("Nenhum arquivo JSON para recarregar ou arquivos não encontrados.")
        except FileNotFoundError as e:
            self.append_text_gui(f"Erro ao recarregar JSONs: Arquivo não encontrado - {e.filename}\n")
            self.status_label_var.set(f"Erro ao recarregar: {e.strerror}")
            logging.error(f"Erro ao recarregar JSONs (FileNotFoundError): {e}", exc_info=True)
        except Exception as e:
            self.append_text_gui(f"Erro ao recarregar JSONs: {e}\n")
            self.status_label_var.set("Erro ao recarregar JSONs.")
            logging.error(f"Erro ao recarregar JSONs: {e}", exc_info=True)

    def create_status_bar(self):
        status_bar = ttk.Frame(self.root)
        status_bar.pack(side='bottom', fill='x')
        ttk.Separator(status_bar).pack(side='top', fill='x')
        self.status_label = ttk.Label(status_bar, textvariable=self.status_label_var, relief='sunken', anchor='w')
        self.status_label.pack(side='left', fill='x', expand=True, padx=5, pady=2)

    def initialize_from_config(self):
        if self.pasta_raiz: self.append_text_gui(f">>> Pasta de logs configurada: {self.pasta_raiz}\n")
        self.forcar_refresh_json()
        if self.nome_servico: self.servico_var.set(f"Serviço: {self.nome_servico}")
        self.status_label_var.set("Configuração carregada. Pronto.")

    def load_config(self):
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f: config = json.load(f)
                logging.info(f"Configuração carregada de {self.config_file}")
                return config
            logging.info(f"Arquivo de configuração {self.config_file} não encontrado. Usando padrões.")
            return {}
        except json.JSONDecodeError as e:
            logging.error(f"Erro ao decodificar JSON em {self.config_file}: {e}", exc_info=True)
            Messagebox.show_error(
                f"Erro ao carregar configuração de '{self.config_file}':\n{e}\nUsando configuração padrão.",
                "Erro de Configuração", parent=self.root)
            return {}
        except Exception as e:
            logging.error(f"Erro desconhecido ao carregar configuração de {self.config_file}: {e}", exc_info=True)
            Messagebox.show_error(
                f"Erro desconhecido ao carregar '{self.config_file}':\n{e}\nUsando configuração padrão.",
                "Erro de Configuração", parent=self.root)
            return {}

    def load_config_dialog(self):
        caminho = filedialog.askopenfilename(defaultextension=".json",
                                             filetypes=[("Arquivos JSON", "*.json"), ("Todos os arquivos", "*.*")],
                                             title="Selecionar arquivo de configuração para carregar")
        if caminho:
            try:
                with open(caminho, 'r', encoding='utf-8') as f:
                    self.config = json.load(f)
                self.pasta_raiz = self.config.get("log_folder", self.pasta_raiz)
                self.arquivo_json = self.config.get("server_json", self.arquivo_json)
                self.arquivo_json_votemap = self.config.get("votemap_json", self.arquivo_json_votemap)
                self.nome_servico = self.config.get("service_name", self.nome_servico)
                self.tema_var.set(self.config.get("theme", self.tema_var.get()))
                self.filtro_var.set(self.config.get("filter", self.filtro_var.get()))
                self.auto_restart_var.set(self.config.get("auto_restart", self.auto_restart_var.get()))
                self.vote_pattern_var.set(self.config.get("vote_pattern", self.vote_pattern_var.get()))
                self.winner_pattern_var.set(self.config.get("winner_pattern", self.winner_pattern_var.get()))
                self.default_mission_var.set(self.config.get("default_mission", self.default_mission_var.get()))
                self.stop_delay_var.set(self.config.get("stop_delay", self.stop_delay_var.get()))
                self.start_delay_var.set(self.config.get("start_delay", self.start_delay_var.get()))
                self.style.theme_use(self.tema_var.get())
                self.initialize_from_config()
                self.status_label_var.set(f"Configuração carregada de {os.path.basename(caminho)}")
                logging.info(f"Configuração carregada de {caminho}")
                Messagebox.show_info("Configuração Carregada", f"Configuração carregada com sucesso de:\n{caminho}",
                                     parent=self.root)
                if self.pasta_raiz:
                    self.stop_log_monitoring()
                    self.start_log_monitoring()
            except json.JSONDecodeError as e:
                logging.error(f"Erro ao decodificar JSON em {caminho}: {e}", exc_info=True)
                Messagebox.show_error(f"Falha ao carregar configuração de '{caminho}':\nFormato JSON inválido.\n{e}",
                                      "Erro de Configuração", parent=self.root)
            except Exception as e:
                logging.error(f"Erro ao carregar configuração de {caminho}: {e}", exc_info=True)
                Messagebox.show_error(f"Falha ao carregar configuração de '{caminho}':\n{e}", "Erro de Configuração",
                                      parent=self.root)

    def save_config(self):
        self.config = {
            "log_folder": self.pasta_raiz, "server_json": self.arquivo_json, "votemap_json": self.arquivo_json_votemap,
            "service_name": self.nome_servico, "theme": self.tema_var.get(), "filter": self.filtro_var.get(),
            "auto_restart": self.auto_restart_var.get(), "vote_pattern": self.vote_pattern_var.get(),
            "winner_pattern": self.winner_pattern_var.get(), "default_mission": self.default_mission_var.get(),
            "stop_delay": self.stop_delay_var.get(), "start_delay": self.start_delay_var.get()
        }
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=4)
            self.status_label_var.set("Configuração salva com sucesso!")
            logging.info(f"Configuração salva em {self.config_file}")
        except IOError as e:
            self.status_label_var.set(f"Erro de E/S ao salvar configuração: {e.strerror}")
            logging.error(f"Erro de E/S ao salvar configuração: {e}", exc_info=True)
            Messagebox.show_error(
                f"Não foi possível salvar o arquivo de configuração:\n{self.config_file}\n\n{e.strerror}",
                "Erro ao Salvar", parent=self.root)
        except Exception as e:
            self.status_label_var.set(f"Erro desconhecido ao salvar configuração: {e}")
            logging.error(f"Erro desconhecido ao salvar configuração: {e}", exc_info=True)
            Messagebox.show_error(f"Ocorreu um erro ao salvar a configuração:\n{e}", "Erro ao Salvar", parent=self.root)

    def selecionar_pasta(self):
        pasta_selecionada = filedialog.askdirectory(title="Selecione a pasta raiz dos logs do servidor")
        if pasta_selecionada:
            if self.pasta_raiz != pasta_selecionada:
                self.pasta_raiz = pasta_selecionada
                self.append_text_gui(f">>> Nova pasta de logs selecionada: {self.pasta_raiz}\n")
                self.status_label_var.set(f"Pasta de logs: {os.path.basename(self.pasta_raiz)}")
                logging.info(f"Pasta de logs alterada para: {self.pasta_raiz}")
                self.stop_log_monitoring()
                self.pasta_atual = None
                self.caminho_log_atual = None
                self.start_log_monitoring()
            else:
                self.append_text_gui(f">>> Pasta de logs já selecionada: {self.pasta_raiz}\n")

    def _selecionar_arquivo_json(self, tipo_json):
        title_map = {"servidor": "Selecionar JSON de Configuração do Servidor",
                     "votemap": "Selecionar JSON de Configuração do Votemap"}
        caminho = filedialog.askopenfilename(filetypes=[("Arquivos JSON", "*.json"), ("Todos os arquivos", "*.*")],
                                             title=title_map.get(tipo_json, "Selecionar Arquivo JSON"))
        if caminho:
            try:
                with open(caminho, 'r', encoding='utf-8') as f:
                    conteudo = f.read()
                json.loads(conteudo)
                if tipo_json == "servidor":
                    self.arquivo_json = caminho
                    self.exibir_json(self.json_text_area, conteudo)
                elif tipo_json == "votemap":
                    self.arquivo_json_votemap = caminho
                    self.exibir_json(self.json_vm_text_area, conteudo)
                msg = f"JSON de {tipo_json} carregado: {os.path.basename(caminho)}"
                self.status_label_var.set(msg)
                self.append_text_gui(f">>> {msg}\n")
                logging.info(f"Arquivo JSON de {tipo_json} selecionado: {caminho}")
            except FileNotFoundError:
                err_msg = f"Erro: Arquivo JSON de {tipo_json} não encontrado em '{caminho}'."
                self.status_label_var.set(err_msg);
                logging.error(err_msg)
                Messagebox.show_error(err_msg, "Arquivo não encontrado", parent=self.root)
            except json.JSONDecodeError:
                err_msg = f"Erro: Arquivo JSON de {tipo_json} ('{os.path.basename(caminho)}') não é um JSON válido."
                self.status_label_var.set(err_msg);
                logging.error(err_msg)
                Messagebox.show_error(err_msg, "JSON Inválido", parent=self.root)
            except Exception as e:
                err_msg = f"Erro ao carregar JSON de {tipo_json} '{os.path.basename(caminho)}': {e}"
                self.status_label_var.set(err_msg);
                logging.error(err_msg, exc_info=True)
                Messagebox.show_error(err_msg, "Erro de Leitura", parent=self.root)

    def selecionar_arquivo_json_servidor(self):
        self._selecionar_arquivo_json("servidor")

    def selecionar_arquivo_json_votemap(self):
        self._selecionar_arquivo_json("votemap")

    def _show_progress_dialog(self, title, message):
        progress_win = ttk.Toplevel(self.root);
        progress_win.title(title)
        progress_win.geometry("300x100");
        progress_win.resizable(False, False)
        progress_win.transient(self.root);
        progress_win.grab_set()
        ttk.Label(progress_win, text=message, bootstyle=PRIMARY).pack(pady=10)
        pb = ttk.Progressbar(progress_win, mode='indeterminate', length=280);
        pb.pack(pady=10)
        pb.start(10)
        progress_win.update_idletasks()
        width, height = progress_win.winfo_width(), progress_win.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        progress_win.geometry(f'{width}x{height}+{x}+{y}')
        return progress_win, pb

    def selecionar_servico(self):
        progress_win, _ = self._show_progress_dialog("Serviços", "Carregando lista de serviços...")
        self.root.update_idletasks()
        threading.Thread(target=self._obter_servicos_worker, args=(progress_win,), daemon=True).start()

    def _obter_servicos_worker(self, progress_win):
        pythoncom.CoInitialize()
        try:
            wmi = win32com.client.GetObject('winmgmts:')
            services_raw = wmi.InstancesOf('Win32_Service')

            nomes_servicos_temp = []
            logging.info(f"Total de serviços brutos encontrados: {len(services_raw)}")  # Log inicial

            for i, s in enumerate(services_raw):
                # Log inicial para cada serviço
                # logging.debug(f"Processando serviço bruto #{i}: {getattr(s, 'Name', 'N/A')}")

                # Filtro MUITO simplificado para teste
                if hasattr(s, 'Name') and s.Name and hasattr(s, 'AcceptStop') and s.AcceptStop:
                    nomes_servicos_temp.append(s.Name)
                    # logging.debug(f"  -> Adicionado (filtro simples): {s.Name}")
                # else:
                # if not (hasattr(s, 'Name') and s.Name):
                #     logging.debug(f"  -> Rejeitado (filtro simples): Sem nome ou nome vazio. ({getattr(s, 'Name', 'N/A')})")
                # elif not (hasattr(s, 'AcceptStop') and s.AcceptStop):
                #     logging.debug(f"  -> Rejeitado (filtro simples): Não aceita parada. ({s.Name}, AcceptStop: {getattr(s, 'AcceptStop', 'N/A')})")

            logging.info(f"Serviços após filtro simples: {len(nomes_servicos_temp)}")
            nomes_servicos = sorted(nomes_servicos_temp)
            self.root.after(0, self._mostrar_dialogo_selecao_servico, nomes_servicos, progress_win)

        except Exception as e:
            # ... (código de tratamento de erro permanece o mesmo)
            logging.error(f"Erro ao listar serviços WMI: {e}", exc_info=True)
            error_message = str(e)
            if hasattr(e, 'args') and isinstance(e.args, tuple) and len(e.args) > 0:
                error_code = e.args[0]
                error_text = e.args[1] if len(e.args) > 1 else "Descrição não disponível"
                detailed_description = ""
                if len(e.args) > 2 and e.args[2] and isinstance(e.args[2], tuple) and len(e.args[2]) > 2 and e.args[2][
                    2]:
                    detailed_description = e.args[2][2]
                error_message = f"Código: {error_code}\nErro: {error_text}"
                if detailed_description:
                    error_message += f"\nDetalhes: {detailed_description}"
            self.root.after(0, self._handle_erro_listar_servicos, error_message, progress_win)
        finally:
            pythoncom.CoUninitialize()

    def _handle_erro_listar_servicos(self, error_message, progress_win):
        if progress_win and progress_win.winfo_exists():
            progress_win.destroy()
        Messagebox.show_error(f"Erro ao obter lista de serviços:\n{error_message}", "Erro WMI", parent=self.root)

    def _mostrar_dialogo_selecao_servico(self, nomes_servicos, progress_win):
        if progress_win and progress_win.winfo_exists():
            progress_win.destroy()

        if not nomes_servicos:
            Messagebox.show_warning("Nenhum serviço gerenciável encontrado.",
                                    "Seleção de Serviço", parent=self.root)
            return

        dialog = ttk.Toplevel(self.root)
        dialog.title("Selecionar Serviço do Jogo")
        dialog.geometry("500x400")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.protocol("WM_DELETE_WINDOW", dialog.destroy)

        ttk.Label(dialog, text="Escolha o serviço do servidor do jogo:", font="-size 10").pack(pady=(10, 5))

        search_frame = ttk.Frame(dialog)
        search_frame.pack(fill='x', padx=10)
        ttk.Label(search_frame, text="Buscar:").pack(side='left')
        search_var = ttk.StringVar()
        search_entry = ttk.Entry(search_frame, textvariable=search_var)
        search_entry.pack(side='left', fill='x', expand=True, padx=5)

        list_frame = ttk.Frame(dialog)
        list_frame.pack(fill='both', expand=True, padx=10, pady=5)
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side='right', fill='y')

        listbox = ttk.Treeview(list_frame, columns=("name",), show="headings", selectmode="browse")
        listbox.heading("name", text="Nome do Serviço")
        listbox.column("name", width=450)
        listbox.pack(side='left', fill='both', expand=True)
        listbox.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=listbox.yview)

        def _populate_listbox(query=""):
            for item in listbox.get_children():
                listbox.delete(item)
            filter_query = query.lower()
            for name in nomes_servicos:
                if not filter_query or filter_query in name.lower():
                    listbox.insert("", "end", values=(name,))

        def _on_search_key_release(event=None):
            _populate_listbox(search_var.get())

        search_entry.bind("<KeyRelease>", _on_search_key_release)
        _populate_listbox()

        def on_confirm():
            selection = listbox.selection()
            if selection:
                selected_item = listbox.item(selection[0])
                service_name = selected_item["values"][0]
                self.nome_servico = service_name
                self.servico_var.set(f"Serviço: {service_name}")
                self.status_label_var.set(f"Serviço selecionado: {service_name}")
                logging.info(f"Serviço selecionado: {service_name}")
                dialog.destroy()
            else:
                Messagebox.show_warning("Nenhum serviço selecionado.", parent=dialog)

        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text="Confirmar", command=on_confirm, bootstyle=SUCCESS).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="Cancelar", command=dialog.destroy, bootstyle=DANGER).pack(side='left', padx=5)

        dialog.update_idletasks()
        ws = dialog.winfo_screenwidth()
        hs = dialog.winfo_screenheight()
        w = dialog.winfo_width()
        h = dialog.winfo_height()
        x = (ws / 2) - (w / 2)
        y = (hs / 2) - (h / 2)
        dialog.geometry(f'+{int(x)}+{int(y)}')
        search_entry.focus_set()
        dialog.wait_window()

    def exibir_json(self, text_area_widget, conteudo_json):
        try:
            dados_formatados = json.dumps(json.loads(conteudo_json), indent=4, ensure_ascii=False)
        except json.JSONDecodeError:
            dados_formatados = conteudo_json
        except TypeError:
            dados_formatados = "Conteúdo JSON inválido ou não fornecido."
        text_area_widget.configure(state='normal');
        text_area_widget.delete('1.0', 'end')
        text_area_widget.insert('end', dados_formatados);
        text_area_widget.configure(state='disabled')

    def monitorar_log_continuamente(self):
        self.status_label_var.set(
            f"Monitorando pasta: {os.path.basename(self.pasta_raiz) if self.pasta_raiz else 'N/A'}")
        logging.info(f"Iniciando monitoramento contínuo de: {self.pasta_raiz}")
        while not self._stop_event.is_set():
            if not self.pasta_raiz or not os.path.isdir(self.pasta_raiz):
                if self.pasta_raiz: logging.warning(
                    f"Pasta de logs '{self.pasta_raiz}' não encontrada ou não é um diretório.")
                self.status_label_var.set("Pasta de logs não configurada ou inválida.")
                self._stop_event.wait(10);
                continue
            try:
                nova_pasta_logs = self.obter_subpasta_mais_recente()
                if nova_pasta_logs:
                    novo_arquivo_log = os.path.join(nova_pasta_logs, 'console.log')
                    if os.path.exists(novo_arquivo_log) and novo_arquivo_log != self.caminho_log_atual:
                        logging.info(f"Nova pasta de log detectada: {nova_pasta_logs}")
                        self.append_text_gui(f"\n>>> Novo arquivo de log detectado: {novo_arquivo_log}\n")
                        if self.log_tail_thread and self.log_tail_thread.is_alive(): self.log_tail_thread.join(
                            timeout=0.5)
                        if self.file_log_handle: self.file_log_handle.close(); self.file_log_handle = None
                        self.pasta_atual = nova_pasta_logs;
                        self.caminho_log_atual = novo_arquivo_log
                        self.log_label.config(text=f"LOG AO VIVO: {self.caminho_log_atual}")
                        self.status_label_var.set(f"Monitorando: {os.path.basename(self.caminho_log_atual)}")
                        try:
                            self.file_log_handle = open(self.caminho_log_atual, 'r', encoding='utf-8', errors='replace')
                            self.file_log_handle.seek(0, os.SEEK_END)
                            self.log_tail_thread = threading.Thread(target=self.acompanhar_log_do_arquivo, daemon=True)
                            self.log_tail_thread.start()
                        except FileNotFoundError:
                            logging.error(f"Arquivo de log {self.caminho_log_atual} não encontrado ao tentar abrir.")
                            self.append_text_gui(f"Erro: Arquivo {self.caminho_log_atual} desapareceu.\n");
                            self.caminho_log_atual = None
                        except Exception as e:
                            logging.error(f"Erro ao abrir ou iniciar acompanhamento de {self.caminho_log_atual}: {e}",
                                          exc_info=True)
                            self.append_text_gui(f"Erro ao abrir {self.caminho_log_atual}: {e}\n");
                            self.caminho_log_atual = None
                    elif not os.path.exists(novo_arquivo_log) and self.caminho_log_atual == novo_arquivo_log:
                        logging.warning(f"Arquivo de log {self.caminho_log_atual} não existe mais.")
                        self.append_text_gui(f"Aviso: Arquivo de log {self.caminho_log_atual} não encontrado.\n");
                        self.caminho_log_atual = None
                        if self.file_log_handle: self.file_log_handle.close(); self.file_log_handle = None
            except Exception as e:
                logging.error(f"Erro no loop de monitoramento de logs: {e}", exc_info=True)
                self.append_text_gui(f"Erro ao monitorar logs: {e}\n")
            self._stop_event.wait(10)
        logging.info("Thread de monitoramento de log contínuo encerrada.")

    def obter_subpasta_mais_recente(self):
        if not self.pasta_raiz or not os.path.isdir(self.pasta_raiz): return None
        try:
            subpastas = [os.path.join(self.pasta_raiz, nome) for nome in os.listdir(self.pasta_raiz) if
                         os.path.isdir(os.path.join(self.pasta_raiz, nome))]
            if not subpastas: return None
            return max(subpastas, key=os.path.getmtime)
        except FileNotFoundError:
            logging.warning(
                f"Pasta raiz '{self.pasta_raiz}' não encontrada ao buscar subpastas.");
            self.pasta_raiz = None;
            return None
        except PermissionError:
            logging.error(
                f"Permissão negada ao acessar '{self.pasta_raiz}' para buscar subpastas.");
            self.pasta_raiz = None;
            return None
        except Exception as e:
            logging.error(f"Erro ao obter subpasta mais recente em '{self.pasta_raiz}': {e}",
                          exc_info=True);
            return None

    def acompanhar_log_do_arquivo(self):
        if not self.file_log_handle: logging.error("Tentativa de acompanhar log sem um file_log_handle válido."); return
        logging.info(f"Iniciando acompanhamento de: {self.caminho_log_atual}")
        aguardando_winner = False;
        vote_pattern_re = None;
        winner_pattern_re = None
        try:
            vote_pattern_str = self.vote_pattern_var.get()
            if vote_pattern_str: vote_pattern_re = re.compile(vote_pattern_str)
            winner_pattern_str = self.winner_pattern_var.get()
            if winner_pattern_str: winner_pattern_re = re.compile(winner_pattern_str)
        except re.error as e:
            logging.error(f"Erro de RegEx nos padrões de votação/vencedor: {e}", exc_info=True)
            self.append_text_gui(f"ERRO DE REGEX: Verifique os padrões nas configurações: {e}\n")
            if not vote_pattern_re or not winner_pattern_re: self.status_label_var.set(
                "Erro de RegEx! Verifique Configurações."); return
        while not self._stop_event.is_set():
            if self._paused: self._stop_event.wait(0.5); continue
            try:
                linha = self.file_log_handle.readline()
                if linha:
                    filtro = self.filtro_var.get().strip().lower()
                    if not filtro or filtro in linha.lower(): self.append_text_gui(linha)
                    if vote_pattern_re and vote_pattern_re.search(linha):
                        aguardando_winner = True;
                        logging.info("Padrão de fim de votação detectado. Aguardando vencedor.")
                        self.status_label_var.set("Fim da votação detectado. Aguardando vencedor...")
                    if aguardando_winner and winner_pattern_re:
                        match = winner_pattern_re.search(linha)
                        if match:
                            try:
                                indice_str = match.group(1);
                                indice = int(indice_str)
                                logging.info(f"Padrão de vencedor detectado: índice {indice}")
                                self.status_label_var.set(f"Vencedor: índice {indice}. Processando...")
                                self.root.after(0, self.processar_troca_mapa, indice);
                                aguardando_winner = False
                            except IndexError:
                                logging.error(
                                    f"Padrão de vencedor '{winner_pattern_str}' casou, mas não possui grupo de captura para o índice.")
                                self.append_text_gui(
                                    f"ERRO: Padrão de vencedor '{winner_pattern_str}' não tem grupo de captura.\n");
                                aguardando_winner = False
                            except ValueError:
                                logging.error(
                                    f"Padrão de vencedor '{winner_pattern_str}' capturou '{indice_str}', que não é um número de índice válido.")
                                self.append_text_gui(f"ERRO: Vencedor capturado '{indice_str}' não é um número.\n");
                                aguardando_winner = False
                else:
                    self._stop_event.wait(0.2)
            except UnicodeDecodeError as ude:
                logging.warning(
                    f"Erro de decodificação Unicode ao ler log {self.caminho_log_atual}: {ude}. Linha ignorada.")
            except Exception as e:
                if not self._stop_event.is_set():
                    logging.error(f"Erro ao acompanhar log {self.caminho_log_atual}: {e}", exc_info=True)
                    self.append_text_gui(f"Erro ao ler log: {e}\n")
                    self.status_label_var.set("Erro na leitura do log. Verifique o Log do Sistema.");
                    break
        logging.info(f"Acompanhamento de {self.caminho_log_atual} encerrado.")
        if self.file_log_handle:
            try:
                self.file_log_handle.close();
                self.file_log_handle = None
            except:
                pass

    def processar_troca_mapa(self, indice_vencedor):
        logging.info(f"Processando troca de mapa para o índice: {indice_vencedor}")
        if not self.arquivo_json or not self.arquivo_json_votemap:
            msg = "Arquivos JSON de servidor ou votemap não configurados para troca de mapa."
            self.append_text_gui(f"ERRO: {msg}\n");
            logging.error(msg);
            self.status_label_var.set("Erro: JSONs não configurados.");
            return
        try:
            with open(self.arquivo_json_votemap, 'r', encoding='utf-8') as f_vm:
                votemap_data = json.load(f_vm)
        except FileNotFoundError:
            msg = f"Arquivo votemap.json ('{self.arquivo_json_votemap}') não encontrado."
            self.append_text_gui(f"ERRO: {msg}\n");
            logging.error(msg);
            self.status_label_var.set("Erro: votemap.json não encontrado.");
            return
        except json.JSONDecodeError as e:
            msg = f"Erro ao decodificar votemap.json: {e}"
            self.append_text_gui(f"ERRO: {msg}\n");
            logging.error(msg);
            self.status_label_var.set("Erro: votemap.json inválido.");
            return
        map_list = votemap_data.get("list", [])
        if not map_list:
            msg = "Lista de mapas vazia ou não encontrada no votemap.json."
            self.append_text_gui(f"AVISO: {msg}\n");
            logging.warning(msg);
            self.status_label_var.set("Aviso: Lista de mapas vazia.");
            return
        novo_scenario_id = None
        if indice_vencedor == 0:
            if len(map_list) > 1:
                indice_selecionado_random = random.randint(1, len(map_list) - 1)
                novo_scenario_id = map_list[indice_selecionado_random]
                self.append_text_gui(
                    f"Voto aleatório: selecionado mapa '{novo_scenario_id}' (índice {indice_selecionado_random}).\n")
                logging.info(f"Seleção aleatória: {novo_scenario_id} (índice {indice_selecionado_random})")
            else:
                msg = "Voto aleatório, mas não há mapas suficientes para escolher."
                self.append_text_gui(f"AVISO: {msg}\n");
                logging.warning(msg);
                self.status_label_var.set("Aviso: Poucos mapas para aleatório.");
                return
        elif 0 < indice_vencedor < len(map_list):
            novo_scenario_id = map_list[indice_vencedor]
            self.append_text_gui(f"Mapa vencedor: '{novo_scenario_id}' (índice {indice_vencedor}).\n")
            logging.info(f"Mapa vencedor selecionado: {novo_scenario_id}")
        else:
            msg = f"Índice do mapa vencedor ({indice_vencedor}) inválido para a lista de mapas (tamanho {len(map_list)})."
            self.append_text_gui(f"ERRO: {msg}\n");
            logging.error(msg);
            self.status_label_var.set("Erro: Índice de mapa inválido.");
            return
        try:
            with open(self.arquivo_json, 'r+', encoding='utf-8') as f_srv:
                server_data = json.load(f_srv);
                server_data["game"]["scenarioId"] = novo_scenario_id
                f_srv.seek(0);
                json.dump(server_data, f_srv, indent=4);
                f_srv.truncate()
            self.exibir_json(self.json_text_area, json.dumps(server_data, indent=4))
            self.append_text_gui(f"JSON do servidor atualizado para o mapa: {novo_scenario_id}\n")
            logging.info(f"JSON do servidor atualizado com scenarioId: {novo_scenario_id}")
            if self.auto_restart_var.get() and self.nome_servico:
                self.append_text_gui("Iniciando reinício automático do servidor...\n")
                threading.Thread(target=self.reiniciar_servidor_com_progresso, args=(novo_scenario_id,),
                                 daemon=True).start()
            else:
                self.status_label_var.set(
                    f"Mapa alterado para: {os.path.basename(str(novo_scenario_id))}. Reinício manual necessário.")
                logging.info("Reinício automático desabilitado ou serviço não configurado.")
        except FileNotFoundError:
            msg = f"Arquivo de config. do servidor ('{self.arquivo_json}') não encontrado."
            self.append_text_gui(f"ERRO: {msg}\n");
            logging.error(msg);
            self.status_label_var.set("Erro: server.json não encontrado.")
        except (KeyError, TypeError) as e:
            msg = f"Estrutura do JSON do servidor inválida (game -> scenarioId não encontrado): {e}"
            self.append_text_gui(f"ERRO: {msg}\n");
            logging.error(msg);
            self.status_label_var.set("Erro: Estrutura server.json inválida.")
        except json.JSONDecodeError as e:
            msg = f"Erro ao decodificar JSON do servidor: {e}"
            self.append_text_gui(f"ERRO: {msg}\n");
            logging.error(msg);
            self.status_label_var.set("Erro: server.json inválido.")
        except Exception as e:
            msg = f"Erro inesperado ao processar troca de mapa: {e}"
            self.append_text_gui(f"ERRO: {msg}\n");
            logging.error(msg, exc_info=True);
            self.status_label_var.set("Erro inesperado na troca de mapa.")

    def verificar_status_servico(self, nome_servico):
        if not nome_servico: return "NOT_FOUND"
        try:
            startupinfo = None
            if platform.system() == "Windows":
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                startupinfo.wShowWindow = subprocess.SW_HIDE

            result = subprocess.run(['sc', 'query', nome_servico], capture_output=True, text=True, check=False,
                                    startupinfo=startupinfo, encoding='latin-1')

            service_not_found_errors = [
                "failed 1060", "falha 1060", "지정된 서비스를 설치된 서비스로 찾을 수 없습니다."
            ]
            output_lower = result.stdout.lower() + result.stderr.lower()
            for err_string in service_not_found_errors:
                if err_string in output_lower:
                    return "NOT_FOUND"

            if "state" not in output_lower:
                logging.warning(
                    f"Saída inesperada do 'sc query {nome_servico}': STDOUT='{result.stdout}' STDERR='{result.stderr}'")
                return "ERROR"

            if "running" in output_lower: return "RUNNING"
            if "stopped" in output_lower: return "STOPPED"
            if "start_pending" in output_lower: return "START_PENDING"
            if "stop_pending" in output_lower: return "STOP_PENDING"
            return "UNKNOWN"
        except FileNotFoundError:
            logging.error("'sc.exe' não encontrado. Verifique se o System32 está no PATH.", exc_info=True)
            Messagebox.show_error("Comando 'sc.exe' não encontrado. Verifique as configurações do sistema.",
                                  "Erro de Comando", parent=self.root);
            return "ERROR"
        except Exception as e:
            logging.error(f"Erro ao verificar status do serviço '{nome_servico}': {e}", exc_info=True);
            return "ERROR"

    def reiniciar_servidor_com_progresso(self, novo_scenario_id_para_log):
        progress_win, pb = self._show_progress_dialog("Reiniciando Servidor", f"Reiniciando {self.nome_servico}...")
        self.root.update_idletasks()
        success = self._reiniciar_servidor_logica(novo_scenario_id_para_log)
        if progress_win.winfo_exists():  # Verificar se a janela ainda existe
            pb.stop()
            progress_win.destroy()
        if success: Messagebox.show_info("Servidor Reiniciado",
                                         f"O serviço {self.nome_servico} foi reiniciado com sucesso.", parent=self.root)

    def _reiniciar_servidor_logica(self, novo_scenario_id_para_log):
        if not self.nome_servico:
            self.append_text_gui("ERRO: Nome do serviço não configurado para reinício.\n");
            logging.error("Tentativa de reiniciar servidor sem nome de serviço configurado.")
            self.status_label_var.set("Erro: Serviço não configurado.");
            return False

        stop_delay = self.stop_delay_var.get();
        start_delay = self.start_delay_var.get();
        default_votemap_mission = self.default_mission_var.get()

        startupinfo = None
        if platform.system() == "Windows":
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = subprocess.SW_HIDE

        try:
            self.status_label_var.set(f"Parando serviço {self.nome_servico}...");
            self.append_text_gui(f"Parando serviço '{self.nome_servico}'...\n")
            logging.info(f"Tentando parar o serviço: {self.nome_servico}")
            status = self.verificar_status_servico(self.nome_servico)
            if status == "RUNNING" or status == "START_PENDING":
                subprocess.run(["sc", "stop", self.nome_servico], check=True, shell=False,
                               startupinfo=startupinfo)
                self.append_text_gui(f"Comando de parada enviado. Aguardando {stop_delay}s...\n");
                time.sleep(stop_delay)
                status_after_stop = self.verificar_status_servico(self.nome_servico)
                if status_after_stop != "STOPPED":
                    logging.warning(f"Serviço {self.nome_servico} não parou como esperado. Status: {status_after_stop}")
                    self.append_text_gui(
                        f"AVISO: Serviço '{self.nome_servico}' pode não ter parado. Status: {status_after_stop}\n")
            elif status == "STOPPED":
                self.append_text_gui(f"Serviço '{self.nome_servico}' já estava parado.\n");
                logging.info(
                    f"Serviço {self.nome_servico} já estava parado.")
            elif status == "NOT_FOUND":
                self.append_text_gui(f"ERRO: Serviço '{self.nome_servico}' não encontrado.\n");
                logging.error(f"Serviço {self.nome_servico} não encontrado para parada.")
                self.status_label_var.set(f"Erro: Serviço '{self.nome_servico}' não existe.");
                return False
            else:
                self.append_text_gui(
                    f"ERRO: Não foi possível determinar o estado do serviço '{self.nome_servico}' ou estado inesperado: {status}.\n")
                logging.error(f"Estado do serviço {self.nome_servico} desconhecido ou erro: {status}")
                self.status_label_var.set(f"Erro: Estado de '{self.nome_servico}' desconhecido.");
                return False

            self.status_label_var.set(f"Iniciando serviço {self.nome_servico}...");
            self.append_text_gui(f"Iniciando serviço '{self.nome_servico}'...\n")
            logging.info(f"Tentando iniciar o serviço: {self.nome_servico}")
            subprocess.run(["sc", "start", self.nome_servico], check=True, shell=False,
                           startupinfo=startupinfo)
            self.append_text_gui(
                f"Comando de início enviado. Aguardando {start_delay}s para o servidor estabilizar...\n")
            self.status_label_var.set(f"Aguardando {self.nome_servico} iniciar ({start_delay}s)...");
            time.sleep(start_delay)
            status_after_start = self.verificar_status_servico(self.nome_servico)
            if status_after_start != "RUNNING":
                logging.error(f"Serviço {self.nome_servico} falhou ao iniciar. Status: {status_after_start}")
                self.append_text_gui(
                    f"ERRO: Serviço '{self.nome_servico}' falhou ao iniciar ou está demorando muito. Status: {status_after_start}\n")
                self.status_label_var.set(f"Erro: {self.nome_servico} não iniciou. Status: {status_after_start}");
                return False
            logging.info(f"Serviço {self.nome_servico} iniciado com sucesso.")

            self.append_text_gui("Restaurando JSON do servidor para o mapa de votação padrão...\n")
            if not self.arquivo_json or not os.path.exists(self.arquivo_json):
                msg = f"Arquivo JSON do servidor ({self.arquivo_json}) não encontrado para restaurar votemap."
                self.append_text_gui(f"ERRO: {msg}\n");
                logging.error(msg);
                self.status_label_var.set("Erro: server.json não encontrado para reset.");
                return False
            with open(self.arquivo_json, 'r+', encoding='utf-8') as f_srv:
                server_data = json.load(f_srv);
                server_data["game"]["scenarioId"] = default_votemap_mission
                f_srv.seek(0);
                json.dump(server_data, f_srv, indent=4);
                f_srv.truncate()
            self.root.after(0, self.exibir_json, self.json_text_area, json.dumps(server_data, indent=4))
            self.append_text_gui(f"JSON do servidor restaurado para votemap: {default_votemap_mission}\n")
            logging.info(f"JSON do servidor restaurado para scenarioId de votemap: {default_votemap_mission}")
            self.status_label_var.set(
                f"Servidor reiniciado. Mapa anterior: {os.path.basename(str(novo_scenario_id_para_log))}. Próximo: Votação.")
            return True
        except subprocess.CalledProcessError as e:
            err_output = ""
            if e.stderr:
                try:
                    err_output = e.stderr.decode('latin-1', errors='replace')
                except:
                    err_output = str(e.stderr)
            elif e.stdout:
                try:
                    err_output = e.stdout.decode('latin-1', errors='replace')
                except:
                    err_output = str(e.stdout)
            else:
                err_output = "Nenhuma saída de erro detalhada."

            err_msg = f"Erro ao executar comando 'sc' para '{self.nome_servico}': {err_output.strip()}"
            self.append_text_gui(f"ERRO: {err_msg}\n");
            logging.error(err_msg, exc_info=True)
            self.status_label_var.set(f"Erro ao gerenciar serviço: {e.cmd}");
            return False
        except FileNotFoundError:
            Messagebox.show_error("Comando 'sc.exe' não encontrado. Verifique o PATH.", "Erro de Comando",
                                  parent=self.root);
            logging.error("Comando 'sc.exe' não encontrado.")
            self.status_label_var.set("Erro: sc.exe não encontrado.");
            return False
        except (json.JSONDecodeError, KeyError, TypeError) as e:
            err_msg = f"Erro ao manipular JSON do servidor durante o reinício: {e}"
            self.append_text_gui(f"ERRO: {err_msg}\n");
            logging.error(err_msg, exc_info=True)
            self.status_label_var.set("Erro: Falha ao atualizar JSON do servidor.");
            return False
        except Exception as e:
            err_msg = f"Erro inesperado ao reiniciar o servidor: {e}"
            self.append_text_gui(f"ERRO: {err_msg}\n");
            logging.error(err_msg, exc_info=True)
            self.status_label_var.set("Erro inesperado no reinício do servidor.");
            return False

    def append_text_gui(self, texto):
        if self.root.winfo_exists() and self.text_area.winfo_exists():
            self.text_area.configure(state='normal');
            self.text_area.insert('end', texto)
            self.text_area.yview_moveto(1.0);
            self.text_area.configure(state='disabled')

    def limpar_tela(self):
        self.text_area.configure(state='normal');
        self.text_area.delete('1.0', 'end');
        self.text_area.configure(state='disabled')
        self.status_label_var.set("Tela de logs do servidor limpa.");
        logging.info("Tela de logs do servidor limpa pelo usuário.")

    def toggle_pausa(self):
        self._paused = not self._paused
        if self._paused:
            self.pausar_btn.config(text="▶️ Retomar", bootstyle=SUCCESS);
            self.status_label_var.set("Monitoramento de logs pausado.")
            logging.info("Monitoramento de logs pausado.")
        else:
            self.pausar_btn.config(text="⏸️ Pausar", bootstyle=WARNING);
            self.status_label_var.set("Monitoramento de logs retomado.")
            logging.info("Monitoramento de logs retomado.")

    def trocar_tema(self, event=None):
        novo_tema = self.tema_var.get()
        try:
            self.style.theme_use(novo_tema);
            logging.info(f"Tema alterado para: {novo_tema}")
            self.status_label_var.set(f"Tema alterado para '{novo_tema}'.")
        except Exception as e:
            logging.error(f"Erro ao tentar trocar para o tema '{novo_tema}': {e}", exc_info=True)
            Messagebox.show_error(f"Não foi possível aplicar o tema '{novo_tema}'.\n{e}", "Erro de Tema",
                                  parent=self.root)
            try:
                self.style.theme_use("litera");
                self.tema_var.set("litera")
            except:
                pass

    def export_display_logs(self):
        caminho_arquivo = filedialog.asksaveasfilename(defaultextension=".txt",
                                                       filetypes=[("Arquivos de Texto", "*.txt"),
                                                                  ("Todos os arquivos", "*.*")],
                                                       title="Exportar Logs Exibidos na Tela")
        if caminho_arquivo:
            try:
                with open(caminho_arquivo, 'w', encoding='utf-8') as f:
                    f.write(self.text_area.get('1.0', 'end-1c'))
                self.status_label_var.set(f"Logs da tela exportados para: {os.path.basename(caminho_arquivo)}")
                logging.info(f"Logs da tela exportados para: {caminho_arquivo}")
                Messagebox.show_info("Exportação Concluída",
                                     f"Logs da tela foram exportados com sucesso para:\n{caminho_arquivo}",
                                     parent=self.root)
            except Exception as e:
                self.status_label_var.set(f"Erro ao exportar logs da tela: {e}")
                logging.error(f"Erro ao exportar logs da tela para {caminho_arquivo}: {e}", exc_info=True)
                Messagebox.show_error(f"Falha ao exportar logs da tela:\n{e}", "Erro de Exportação", parent=self.root)

    def validate_configs(self):
        problemas = []
        if not self.pasta_raiz or not os.path.isdir(self.pasta_raiz): problemas.append(
            "- Pasta de logs do servidor não configurada ou inválida.")
        if not self.arquivo_json or not os.path.exists(self.arquivo_json): problemas.append(
            "- Arquivo JSON de configuração do servidor não configurado ou não encontrado.")
        if not self.arquivo_json_votemap or not os.path.exists(self.arquivo_json_votemap): problemas.append(
            "- Arquivo JSON de configuração do Votemap não configurado ou não encontrado.")
        if self.auto_restart_var.get() and not self.nome_servico: problemas.append(
            "- Reinício automático habilitado, mas nenhum serviço do servidor selecionado.")
        if not self.default_mission_var.get(): problemas.append(
            "- Missão padrão de votemap não definida (necessária para resetar após reinício).")
        try:
            if self.vote_pattern_var.get(): re.compile(self.vote_pattern_var.get())
        except re.error:
            problemas.append("- Padrão de detecção de voto (RegEx) é inválido.")
        try:
            if self.winner_pattern_var.get(): re.compile(self.winner_pattern_var.get())
        except re.error:
            problemas.append("- Padrão de detecção de vencedor (RegEx) é inválido.")
        if problemas:
            Messagebox.show_warning("Validação de Configurações",
                                    "Os seguintes problemas de configuração foram encontrados:\n\n" + "\n".join(
                                        problemas), parent=self.root)
        else:
            Messagebox.show_info("Validação de Configurações",
                                 "Todas as configurações essenciais parecem estar corretas!", parent=self.root)
        logging.info(f"Validação de configurações: {len(problemas)} problemas encontrados.")

    def show_about(self):
        about_win = ttk.Toplevel(self.root);
        about_win.title("Sobre Predadores Votemap Patch");
        about_win.geometry("450x350")
        about_win.resizable(False, False);
        about_win.transient(self.root);
        about_win.grab_set()
        frame = ttk.Frame(about_win, padding=20);
        frame.pack(fill='both', expand=True)
        ttk.Label(frame, text="Predadores Votemap Patch", font="-size 16 -weight bold").pack(pady=(0, 10))
        ttk.Label(frame, text="Versão 2.2 (Ícone Integrado)", font="-size 10").pack()
        ttk.Separator(frame).pack(fill='x', pady=10)
        desc = ("Ferramenta para monitorar logs de servidores de jogos,\n"
                "detectar votações de mapa e automatizar a troca\n"
                "de mapas e reinício do servidor.\n\n"
                "Principais funcionalidades:\n"
                "- Monitoramento de logs em tempo real\n"
                "- Detecção de votação e vencedor\n"
                "- Atualização automática de JSON de configuração\n"
                "- Reinício automático de serviço (Windows)\n"
                "- Interface personalizável com temas")
        ttk.Label(frame, text=desc, justify='left').pack(pady=10)
        ttk.Separator(frame).pack(fill='x', pady=10)
        ttk.Label(frame, text="Desenvolvido para a comunidade Predadores").pack()
        ttk.Button(frame, text="Fechar", command=about_win.destroy, bootstyle=PRIMARY).pack(pady=(15, 0))
        about_win.update_idletasks()
        x = (self.root.winfo_screenwidth() // 2) - (about_win.winfo_width() // 2)
        y = (self.root.winfo_screenheight() // 2) - (about_win.winfo_height() // 2)
        about_win.geometry(f'+{x}+{y}')
        about_win.wait_window()

    def setup_tray_icon(self):
        try:
            image = self._create_tray_image()
            if image is None:
                logging.error("Não foi possível criar a imagem para o ícone da bandeja.")
                return

            menu = pystray.Menu(
                pystray.MenuItem('Mostrar Predadores Votemap', self.show_from_tray, default=True),
                pystray.MenuItem('Sair', self.on_close_from_tray)
            )
            self.tray_icon = pystray.Icon("predadores_votemap_patch", image, "Predadores Votemap Patch", menu)
            threading.Thread(target=self.tray_icon.run, daemon=True).start()
            logging.info("Ícone da bandeja do sistema configurado e iniciado.")
        except Exception as e:
            logging.error(f"Falha ao criar ícone da bandeja: {e}", exc_info=True)

    def show_from_tray(self):
        self.root.after(0, self.root.deiconify)
        self.root.after(10, self.root.lift)
        self.root.after(20, self.root.focus_force)

    def minimize_to_tray(self, event=None):
        if event and event.widget == self.root and self.root.state() == 'iconic':
            if hasattr(self, 'tray_icon') and self.tray_icon.visible:
                self.root.withdraw()
                logging.info("Aplicação minimizada para a bandeja.")

    def on_close_from_tray(self):
        logging.info("Fechando aplicação a partir do ícone da bandeja...")
        self.status_label_var.set("Encerrando...")
        self.root.update_idletasks()
        self.stop_log_monitoring()
        if hasattr(self, 'tray_icon') and self.tray_icon:
            try:
                self.tray_icon.stop()
            except Exception as e:
                logging.error(f"Erro ao parar ícone da bandeja: {e}", exc_info=True)
        try:
            self.save_config()
        except Exception as e:
            logging.error(f"Erro ao salvar configuração ao sair: {e}", exc_info=True)

        self._stop_event.set()
        logging.info("Aplicação encerrada (via bandeja).")
        self.root.destroy()

    def on_close(self):
        logging.info("Iniciando processo de fechamento da aplicação (via janela)...")
        if Messagebox.okcancel("Confirmar Saída", "Deseja realmente sair do Predadores Votemap Patch?",
                               parent=self.root) == "OK":
            self.status_label_var.set("Encerrando...")
            self.root.update_idletasks()
            self.stop_log_monitoring()
            if hasattr(self, 'tray_icon') and self.tray_icon:
                try:
                    self.tray_icon.stop()
                except Exception as e:
                    logging.error(f"Erro ao parar ícone da bandeja: {e}", exc_info=True)
            try:
                self.save_config()
            except Exception as e:
                logging.error(f"Erro ao salvar configuração ao sair: {e}", exc_info=True)

            self._stop_event.set()
            logging.info("Aplicação encerrada (via janela).")
            self.root.destroy()
        else:
            logging.info("Saída cancelada pelo usuário.")


def main():
    root = ttk.Window(themename="darkly")
    app = LogViewerApp(root)
    root.protocol("WM_DELETE_WINDOW", app.on_close)
    root.bind("<Unmap>", app.minimize_to_tray)
    app.setup_tray_icon()
    root.mainloop()


if __name__ == '__main__':
    def handle_thread_exception(args):
        logging.error(f"Exceção não capturada na thread {args.thread}:",
                      exc_info=(args.exc_type, args.exc_value, args.exc_traceback))
        import traceback
        traceback.print_exception(args.exc_type, args.exc_value, args.exc_traceback)


    threading.excepthook = handle_thread_exception

    main()