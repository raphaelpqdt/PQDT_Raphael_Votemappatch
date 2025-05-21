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
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(threadName)s - %(module)s.%(funcName)s - %(message)s',
    filename='votemap_patch.log',
    filemode='a'
)


def resource_path(relative_path):
    """ Obtém o caminho absoluto para o recurso, funciona para dev e para PyInstaller """
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(os.path.dirname(__file__))
    return os.path.join(base_path, relative_path)


ICON_FILENAME = "pred.ico"
ICON_PATH = resource_path(ICON_FILENAME)


class LogViewerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Predadores Votemap Patch")
        self.root.geometry("1000x900")

        self.app_icon_image = None
        self.set_application_icon()

        self.config_file = "votemap_config.json"
        self.config = self.load_config()

        self.style = ttk.Style()
        self.style.theme_use(self.config.get("theme", "darkly"))

        self.pasta_raiz = self.config.get("log_folder", None)
        self.pasta_atual = None
        self.caminho_log_atual = None  # Caminho do arquivo que a LogTailThread ATUALMENTE designada deve monitorar
        self.arquivo_json = self.config.get("server_json", None)
        self.arquivo_json_votemap = self.config.get("votemap_json", None)
        self.nome_servico = self.config.get("service_name", None)

        self._stop_event = threading.Event()
        self._paused = False
        self.log_tail_thread = None
        self.log_monitor_thread = None
        self.file_log_handle = None  # O ÚNICO handle de arquivo de log aberto, gerenciado por LogMonitorThread
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
        self._stop_event.clear()  # Limpar evento de parada para a nova sessão de monitoramento
        self.log_monitor_thread = threading.Thread(target=self.monitorar_log_continuamente, daemon=True,
                                                   name="LogMonitorThread")
        self.log_monitor_thread.start()
        logging.info("Monitoramento de logs iniciado.")

    def stop_log_monitoring(self):
        # Esta função agora é mais sobre sinalizar e limpar recursos,
        # as threads devem se encarregar de sair graciosamente.
        current_thread_name = threading.current_thread().name
        logging.debug(f"[{current_thread_name}] Chamada para stop_log_monitoring.")
        self._stop_event.set()  # Sinaliza para TODAS as threads que usam este evento para pararem

        # Aguarda a LogTailThread, se existir e estiver viva
        if self.log_tail_thread and self.log_tail_thread.is_alive():
            logging.debug(
                f"[{current_thread_name}] Aguardando LogTailThread ({self.log_tail_thread.name}) finalizar...")
            self.log_tail_thread.join(timeout=2.0)
            if self.log_tail_thread.is_alive():
                logging.warning(
                    f"[{current_thread_name}] LogTailThread ({self.log_tail_thread.name}) não finalizou no tempo esperado.")

        # Aguarda a LogMonitorThread, se existir, estiver viva E NÃO FOR a thread atual
        if self.log_monitor_thread and self.log_monitor_thread.is_alive() and self.log_monitor_thread != threading.current_thread():
            logging.debug(
                f"[{current_thread_name}] Aguardando LogMonitorThread ({self.log_monitor_thread.name}) finalizar...")
            self.log_monitor_thread.join(timeout=2.0)
            if self.log_monitor_thread.is_alive():
                logging.warning(
                    f"[{current_thread_name}] LogMonitorThread ({self.log_monitor_thread.name}) não finalizou no tempo esperado.")

        # Fecha o handle do arquivo se ele ainda estiver aberto
        if self.file_log_handle:
            try:
                handle_name_for_log = getattr(self.file_log_handle, 'name', 'N/A')
                logging.debug(
                    f"[{current_thread_name}] Fechando file_log_handle em stop_log_monitoring para: {handle_name_for_log}")
                self.file_log_handle.close()
            except Exception as e:
                logging.error(
                    f"[{current_thread_name}] Erro ao fechar handle do arquivo de log ({handle_name_for_log}) em stop_log_monitoring: {e}",
                    exc_info=True)
            finally:
                self.file_log_handle = None  # Garante que seja None

        self.caminho_log_atual = None  # Reseta o caminho do log que estava sendo monitorado
        logging.info(f"[{current_thread_name}] stop_log_monitoring completado.")

    def atualizar_log_sistema_periodicamente(self):
        try:
            if not self.root.winfo_exists() or not self.notebook.winfo_exists(): return
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
                logging.error(f"Erro ao atualizar log do sistema na GUI: {e}",
                              exc_info=False)  # Manter exc_info=False para não poluir muito
                self._system_log_update_error_count = getattr(self, "_system_log_update_error_count", 0) + 1
        if not self._stop_event.is_set():  # Continuar apenas se não estivermos parando
            if self.root.winfo_exists():
                self.root.after(3000, self.atualizar_log_sistema_periodicamente)

    # --- Funções da UI e Configuração (sem alterações profundas, apenas chamadas de GUI via self.root.after se necessário) ---
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
                self.set_status_from_thread("Arquivos JSON recarregados.")
            else:
                self.set_status_from_thread("Nenhum arquivo JSON para recarregar ou arquivos não encontrados.")
        except FileNotFoundError as e:
            self.append_text_gui(f"Erro ao recarregar JSONs: Arquivo não encontrado - {e.filename}\n")
            self.set_status_from_thread(f"Erro ao recarregar: {e.strerror}")
            logging.error(f"Erro ao recarregar JSONs (FileNotFoundError): {e}", exc_info=True)
        except Exception as e:
            self.append_text_gui(f"Erro ao recarregar JSONs: {e}\n")
            self.set_status_from_thread("Erro ao recarregar JSONs.")
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
        self.set_status_from_thread("Configuração carregada. Pronto.")

    def load_config(self):
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f: config = json.load(f)
                logging.info(f"Configuração carregada de {self.config_file}")
                return config
            logging.info(f"Arquivo de configuração {self.config_file} não encontrado. Usando padrões.")
            return {}  # Retorna um dicionário vazio como padrão
        except json.JSONDecodeError as e:
            logging.error(f"Erro ao decodificar JSON em {self.config_file}: {e}", exc_info=True)
            self.show_messagebox_from_thread("error", "Erro de Configuração",
                                             f"Erro ao carregar configuração de '{self.config_file}':\n{e}\nUsando configuração padrão.")
            return {}
        except Exception as e:
            logging.error(f"Erro desconhecido ao carregar configuração de {self.config_file}: {e}", exc_info=True)
            self.show_messagebox_from_thread("error", "Erro de Configuração",
                                             f"Erro desconhecido ao carregar '{self.config_file}':\n{e}\nUsando configuração padrão.")
            return {}

    def load_config_dialog(self):
        caminho = filedialog.askopenfilename(defaultextension=".json",
                                             filetypes=[("Arquivos JSON", "*.json"), ("Todos os arquivos", "*.*")],
                                             title="Selecionar arquivo de configuração para carregar")
        if caminho:
            try:
                with open(caminho, 'r', encoding='utf-8') as f:
                    loaded_config = json.load(f)  # Carregar em uma var temporária

                # Atualizar atributos da classe com base no arquivo carregado
                self.config = loaded_config  # Atualiza o self.config principal
                self.pasta_raiz = self.config.get("log_folder", self.pasta_raiz)
                self.arquivo_json = self.config.get("server_json", self.arquivo_json)
                self.arquivo_json_votemap = self.config.get("votemap_json", self.arquivo_json_votemap)
                self.nome_servico = self.config.get("service_name", self.nome_servico)

                # Atualizar variáveis da UI
                self.tema_var.set(self.config.get("theme", self.tema_var.get()))
                self.filtro_var.set(self.config.get("filter", self.filtro_var.get()))
                self.auto_restart_var.set(self.config.get("auto_restart", self.auto_restart_var.get()))
                self.vote_pattern_var.set(self.config.get("vote_pattern", self.vote_pattern_var.get()))
                self.winner_pattern_var.set(self.config.get("winner_pattern", self.winner_pattern_var.get()))
                self.default_mission_var.set(self.config.get("default_mission", self.default_mission_var.get()))
                self.stop_delay_var.set(self.config.get("stop_delay", self.stop_delay_var.get()))
                self.start_delay_var.set(self.config.get("start_delay", self.start_delay_var.get()))

                self.style.theme_use(self.tema_var.get())  # Aplicar tema
                self.initialize_from_config()  # Reaplicar configurações na UI e lógica

                self.set_status_from_thread(f"Configuração carregada de {os.path.basename(caminho)}")
                logging.info(f"Configuração carregada de {caminho}")
                self.show_messagebox_from_thread("info", "Configuração Carregada",
                                                 f"Configuração carregada com sucesso de:\n{caminho}")

                if self.pasta_raiz:
                    self.stop_log_monitoring()
                    self.start_log_monitoring()
            except json.JSONDecodeError as e:
                logging.error(f"Erro ao decodificar JSON em {caminho}: {e}", exc_info=True)
                self.show_messagebox_from_thread("error", "Erro de Configuração",
                                                 f"Falha ao carregar configuração de '{caminho}':\nFormato JSON inválido.\n{e}")
            except Exception as e:
                logging.error(f"Erro ao carregar configuração de {caminho}: {e}", exc_info=True)
                self.show_messagebox_from_thread("error", "Erro de Configuração",
                                                 f"Falha ao carregar configuração de '{caminho}':\n{e}")

    def save_config(self):
        # Atualiza o dicionário self.config com os valores atuais da UI/lógica ANTES de salvar
        self.config["log_folder"] = self.pasta_raiz
        self.config["server_json"] = self.arquivo_json
        self.config["votemap_json"] = self.arquivo_json_votemap
        self.config["service_name"] = self.nome_servico
        self.config["theme"] = self.tema_var.get()
        self.config["filter"] = self.filtro_var.get()
        self.config["auto_restart"] = self.auto_restart_var.get()
        self.config["vote_pattern"] = self.vote_pattern_var.get()
        self.config["winner_pattern"] = self.winner_pattern_var.get()
        self.config["default_mission"] = self.default_mission_var.get()
        self.config["stop_delay"] = self.stop_delay_var.get()
        self.config["start_delay"] = self.start_delay_var.get()

        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=4)  # Salva o self.config atualizado
            self.set_status_from_thread("Configuração salva com sucesso!")
            logging.info(f"Configuração salva em {self.config_file}")
        except IOError as e:
            self.set_status_from_thread(f"Erro de E/S ao salvar configuração: {e.strerror}")
            logging.error(f"Erro de E/S ao salvar configuração: {e}", exc_info=True)
            self.show_messagebox_from_thread("error", "Erro ao Salvar",
                                             f"Não foi possível salvar o arquivo de configuração:\n{self.config_file}\n\n{e.strerror}")
        except Exception as e:
            self.set_status_from_thread(f"Erro desconhecido ao salvar configuração: {e}")
            logging.error(f"Erro desconhecido ao salvar configuração: {e}", exc_info=True)
            self.show_messagebox_from_thread("error", "Erro ao Salvar",
                                             f"Ocorreu um erro ao salvar a configuração:\n{e}")

    def selecionar_pasta(self):
        pasta_selecionada = filedialog.askdirectory(title="Selecione a pasta raiz dos logs do servidor")
        if pasta_selecionada:
            if self.pasta_raiz != pasta_selecionada:
                logging.info(f"Pasta de logs alterada de '{self.pasta_raiz}' para '{pasta_selecionada}'")
                self.stop_log_monitoring()  # Parar monitoramento antigo
                self.pasta_raiz = pasta_selecionada  # Atualizar a pasta raiz
                self.append_text_gui(f">>> Nova pasta de logs selecionada: {self.pasta_raiz}\n")
                self.set_status_from_thread(f"Pasta de logs: {os.path.basename(self.pasta_raiz)}")
                # As variáveis de estado (caminho_log_atual, file_log_handle) serão resetadas por stop_log_monitoring
                # e reconfiguradas por start_log_monitoring.
                self.start_log_monitoring()  # Iniciar monitoramento na nova pasta
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
                json.loads(conteudo)  # Apenas para validar o JSON
                if tipo_json == "servidor":
                    self.arquivo_json = caminho
                    self.exibir_json(self.json_text_area, conteudo)
                elif tipo_json == "votemap":
                    self.arquivo_json_votemap = caminho
                    self.exibir_json(self.json_vm_text_area, conteudo)
                msg = f"JSON de {tipo_json} carregado: {os.path.basename(caminho)}"
                self.set_status_from_thread(msg)
                self.append_text_gui(f">>> {msg}\n")
                logging.info(f"Arquivo JSON de {tipo_json} selecionado: {caminho}")
            except FileNotFoundError:
                err_msg = f"Erro: Arquivo JSON de {tipo_json} não encontrado em '{caminho}'."
                self.set_status_from_thread(err_msg);
                logging.error(err_msg)
                self.show_messagebox_from_thread("error", "Arquivo não encontrado", err_msg)
            except json.JSONDecodeError:
                err_msg = f"Erro: Arquivo JSON de {tipo_json} ('{os.path.basename(caminho)}') não é um JSON válido."
                self.set_status_from_thread(err_msg);
                logging.error(err_msg)
                self.show_messagebox_from_thread("error", "JSON Inválido", err_msg)
            except Exception as e:
                err_msg = f"Erro ao carregar JSON de {tipo_json} '{os.path.basename(caminho)}': {e}"
                self.set_status_from_thread(err_msg);
                logging.error(err_msg, exc_info=True)
                self.show_messagebox_from_thread("error", "Erro de Leitura", err_msg)

    def selecionar_arquivo_json_servidor(self):
        self._selecionar_arquivo_json("servidor")

    def selecionar_arquivo_json_votemap(self):
        self._selecionar_arquivo_json("votemap")

    def _show_progress_dialog(self, title, message):
        progress_win = ttk.Toplevel(self.root)
        progress_win.title(title)
        progress_win.geometry("300x100")
        progress_win.resizable(False, False)
        progress_win.transient(self.root)
        progress_win.grab_set()
        ttk.Label(progress_win, text=message, bootstyle=PRIMARY).pack(pady=10)
        pb = ttk.Progressbar(progress_win, mode='indeterminate', length=280)
        pb.pack(pady=10)
        pb.start(10)
        progress_win.update_idletasks()
        width = progress_win.winfo_width()
        height = progress_win.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        progress_win.geometry(f'{width}x{height}+{x}+{y}')
        return progress_win, pb

    def selecionar_servico(self):
        progress_win, _ = self._show_progress_dialog("Serviços", "Carregando lista de serviços...")
        if self.root.winfo_exists(): self.root.update_idletasks()  # Garante que a dialog apareça
        threading.Thread(target=self._obter_servicos_worker, args=(progress_win,), daemon=True,
                         name="ServicoWMIThread").start()

    def _obter_servicos_worker(self, progress_win):
        pythoncom.CoInitialize()
        try:
            wmi = win32com.client.GetObject('winmgmts:')
            services_raw = wmi.InstancesOf('Win32_Service')
            logging.info(f"Total de serviços brutos encontrados: {len(services_raw)}")
            nomes_servicos_temp = [s.Name for s in services_raw if
                                   hasattr(s, 'Name') and s.Name and hasattr(s, 'AcceptStop') and s.AcceptStop]
            logging.info(f"Serviços após filtro simples: {len(nomes_servicos_temp)}")
            nomes_servicos = sorted(nomes_servicos_temp)

            if self.root.winfo_exists():  # Checar se a root existe antes de agendar
                self.root.after(0, self._mostrar_dialogo_selecao_servico, nomes_servicos, progress_win)
        except Exception as e:
            logging.error(f"Erro ao listar serviços WMI: {e}", exc_info=True)
            error_message = str(e)
            if hasattr(e, 'args') and isinstance(e.args, tuple) and len(e.args) > 0:  # Melhorar msg de erro
                error_code = e.args[0];
                error_text = e.args[1] if len(e.args) > 1 else "N/A"
                detailed = e.args[2][2] if len(e.args) > 2 and e.args[2] and isinstance(e.args[2], tuple) and len(
                    e.args[2]) > 2 and e.args[2][2] else ""
                error_message = f"Código: {error_code}\nErro: {error_text}\nDetalhes: {detailed}"

            if self.root.winfo_exists():
                self.root.after(0, self._handle_erro_listar_servicos, error_message, progress_win)
        finally:
            pythoncom.CoUninitialize()

    def _handle_erro_listar_servicos(self, error_message, progress_win):
        if progress_win and progress_win.winfo_exists():
            progress_win.destroy()
        if self.root.winfo_exists():  # Checar antes de mostrar messagebox
            Messagebox.show_error(f"Erro ao obter lista de serviços:\n{error_message}", "Erro WMI", parent=self.root)

    def _mostrar_dialogo_selecao_servico(self, nomes_servicos, progress_win):
        if progress_win and progress_win.winfo_exists():
            progress_win.destroy()

        if not nomes_servicos:
            if self.root.winfo_exists():
                Messagebox.show_warning("Nenhum serviço gerenciável encontrado.", "Seleção de Serviço",
                                        parent=self.root)
            return

        dialog = ttk.Toplevel(self.root)
        dialog.title("Selecionar Serviço do Jogo")
        dialog.geometry("500x400")
        dialog.transient(self.root);
        dialog.grab_set()
        dialog.protocol("WM_DELETE_WINDOW", dialog.destroy)

        ttk.Label(dialog, text="Escolha o serviço do servidor do jogo:", font="-size 10").pack(pady=(10, 5))
        search_frame = ttk.Frame(dialog);
        search_frame.pack(fill='x', padx=10)
        ttk.Label(search_frame, text="Buscar:").pack(side='left')
        search_var = ttk.StringVar()
        search_entry = ttk.Entry(search_frame, textvariable=search_var);
        search_entry.pack(side='left', fill='x', expand=True, padx=5)

        list_frame = ttk.Frame(dialog);
        list_frame.pack(fill='both', expand=True, padx=10, pady=5)
        scrollbar = ttk.Scrollbar(list_frame);
        scrollbar.pack(side='right', fill='y')

        listbox = ttk.Treeview(list_frame, columns=("name",), show="headings", selectmode="browse")
        listbox.heading("name", text="Nome do Serviço")
        listbox.column("name", width=450)  # Ajustar largura se necessário
        listbox.pack(side='left', fill='both', expand=True)
        listbox.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=listbox.yview)

        def _populate_listbox(query=""):
            for item in listbox.get_children(): listbox.delete(item)
            filter_query = query.lower()
            for name in nomes_servicos:
                if not filter_query or filter_query in name.lower():
                    listbox.insert("", "end", values=(name,))

        search_entry.bind("<KeyRelease>", lambda e: _populate_listbox(search_var.get()))
        _populate_listbox()

        def on_confirm():
            selection = listbox.selection()
            if selection:
                selected_item = listbox.item(selection[0])
                service_name = selected_item["values"][0]
                self.nome_servico = service_name
                self.servico_var.set(f"Serviço: {service_name}")
                self.set_status_from_thread(f"Serviço selecionado: {service_name}")
                logging.info(f"Serviço selecionado: {service_name}")
                dialog.destroy()
            else:
                if dialog.winfo_exists():  # Checar se dialog ainda existe
                    Messagebox.show_warning("Nenhum serviço selecionado.", parent=dialog)

        btn_frame = ttk.Frame(dialog);
        btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text="Confirmar", command=on_confirm, bootstyle=SUCCESS).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="Cancelar", command=dialog.destroy, bootstyle=DANGER).pack(side='left', padx=5)

        dialog.update_idletasks()
        # Centralizar dialog
        ws, hs = dialog.winfo_screenwidth(), dialog.winfo_screenheight()
        w, h = dialog.winfo_width(), dialog.winfo_height()
        x, y = (ws / 2) - (w / 2), (hs / 2) - (h / 2)
        dialog.geometry(f'+{int(x)}+{int(y)}')
        search_entry.focus_set()
        dialog.wait_window()

    def exibir_json(self, text_area_widget, conteudo_json):
        try:
            # Tenta formatar se for uma string JSON válida, senão exibe como está
            dados_formatados = json.dumps(json.loads(conteudo_json), indent=4, ensure_ascii=False)
        except (json.JSONDecodeError, TypeError):  # Se não for JSON ou tipo inválido
            dados_formatados = str(conteudo_json)  # Converte para string para garantir

        if self.root.winfo_exists() and text_area_widget.winfo_exists():
            text_area_widget.configure(state='normal')
            text_area_widget.delete('1.0', 'end')
            text_area_widget.insert('end', dados_formatados)
            text_area_widget.configure(state='disabled')

    # --- LÓGICA CENTRAL DE MONITORAMENTO E PROCESSAMENTO ---
    def monitorar_log_continuamente(self):
        thread_name = threading.current_thread().name
        self.set_status_from_thread(
            f"Monitorando pasta: {os.path.basename(self.pasta_raiz) if self.pasta_raiz else 'N/A'}")
        logging.info(f"[{thread_name}] Iniciando monitoramento contínuo de: {self.pasta_raiz}")

        while not self._stop_event.is_set():
            if not self.pasta_raiz or not os.path.isdir(self.pasta_raiz):
                if self.pasta_raiz:
                    logging.warning(
                        f"[{thread_name}] Pasta de logs '{self.pasta_raiz}' não encontrada ou não é um diretório.")
                self.set_status_from_thread("Pasta de logs não configurada ou inválida.")
                if self._stop_event.wait(10): break  # Espera e checa se deve parar
                continue

            try:
                nova_pasta_logs = self.obter_subpasta_mais_recente()
                if not nova_pasta_logs:  # Nenhuma subpasta encontrada
                    if self._stop_event.wait(5): break  # Intervalo menor se não há pastas
                    continue

                novo_arquivo_log_path = os.path.join(nova_pasta_logs, 'console.log')

                # Condição para trocar de arquivo de log
                if os.path.exists(novo_arquivo_log_path) and novo_arquivo_log_path != self.caminho_log_atual:
                    logging.info(
                        f"[{thread_name}] Novo arquivo de log detectado: {novo_arquivo_log_path} (anterior: {self.caminho_log_atual})")
                    self.append_text_gui(f"\n>>> Novo arquivo de log detectado: {novo_arquivo_log_path}\n")

                    # 1. Parar a thread de acompanhamento antiga, se existir
                    if self.log_tail_thread and self.log_tail_thread.is_alive():
                        logging.debug(
                            f"[{thread_name}] Parando LogTailThread ({self.log_tail_thread.name}) para o arquivo antigo '{self.caminho_log_atual}'...")
                        # A LogTailThread também verifica self._stop_event, mas o join garante
                        self.log_tail_thread.join(timeout=1.5)
                        if self.log_tail_thread.is_alive():
                            logging.warning(
                                f"[{thread_name}] LogTailThread ({self.log_tail_thread.name}) antiga não parou a tempo.")

                    # 2. Fechar o handle do arquivo antigo, se existir
                    if self.file_log_handle:
                        try:
                            handle_name_old = getattr(self.file_log_handle, 'name', 'N/A')
                            logging.debug(f"[{thread_name}] Fechando file_log_handle antigo para: {handle_name_old}")
                            self.file_log_handle.close()
                        except Exception as e_close:
                            logging.error(
                                f"[{thread_name}] Erro ao fechar handle antigo ({handle_name_old}): {e_close}",
                                exc_info=True)
                        finally:
                            self.file_log_handle = None  # Crucial: zerar antes de abrir o novo

                    # 3. Atualizar o caminho do log atual e a pasta atual
                    self.caminho_log_atual = novo_arquivo_log_path
                    self.pasta_atual = nova_pasta_logs

                    # 4. Tentar abrir o novo arquivo de log
                    novo_fh_temp = None  # Handle temporário
                    try:
                        logging.debug(f"[{thread_name}] Tentando abrir novo arquivo de log: {self.caminho_log_atual}")
                        novo_fh_temp = open(self.caminho_log_atual, 'r', encoding='utf-8', errors='replace')
                        novo_fh_temp.seek(0, os.SEEK_END)  # Ir para o fim do novo arquivo

                        # 5. Atribuir o novo handle ao membro da classe SOMENTE SE a abertura foi bem-sucedida
                        self.file_log_handle = novo_fh_temp

                        # Atualizar UI (labels, status) - via self.root.after para segurança
                        if self.root.winfo_exists():
                            self.root.after(0, lambda p=self.caminho_log_atual: self.log_label.config(
                                text=f"LOG AO VIVO: {p}"))
                            self.root.after(0, lambda p=self.caminho_log_atual: self.status_label_var.set(
                                f"Monitorando: {os.path.basename(p)}"))

                        # 6. Iniciar nova thread de acompanhamento para o novo arquivo
                        logging.info(
                            f"[{thread_name}] Novo arquivo de log {self.caminho_log_atual} aberto com sucesso. Iniciando nova LogTailThread.")
                        self.log_tail_thread = threading.Thread(
                            target=self.acompanhar_log_do_arquivo,
                            args=(self.caminho_log_atual,),  # Passa o caminho que esta thread DEVE monitorar
                            daemon=True,
                            name=f"LogTailThread-{os.path.basename(self.caminho_log_atual)}"  # Nome mais descritivo
                        )
                        self.log_tail_thread.start()

                    except FileNotFoundError:
                        logging.error(
                            f"[{thread_name}] Arquivo de log {self.caminho_log_atual} não encontrado ao tentar abrir.")
                        if novo_fh_temp: novo_fh_temp.close()  # Limpar se parcialmente aberto
                        self.file_log_handle = None  # Garantir que está None se a abertura falhou
                        self.caminho_log_atual = None  # Resetar caminho se falhou
                    except Exception as e_open:
                        logging.error(
                            f"[{thread_name}] Erro ao abrir ou iniciar acompanhamento de {self.caminho_log_atual}: {e_open}",
                            exc_info=True)
                        if novo_fh_temp: novo_fh_temp.close()
                        self.file_log_handle = None
                        self.caminho_log_atual = None

                # Caso o arquivo monitorado atualmente desapareça
                elif self.caminho_log_atual and not os.path.exists(self.caminho_log_atual):
                    logging.warning(
                        f"[{thread_name}] Arquivo de log monitorado {self.caminho_log_atual} não existe mais.")
                    self.append_text_gui(f"Aviso: Arquivo de log {self.caminho_log_atual} não encontrado.\n")
                    if self.log_tail_thread and self.log_tail_thread.is_alive():
                        self.log_tail_thread.join(timeout=1.0)  # Esperar a thread morrer
                    if self.file_log_handle:
                        try:
                            self.file_log_handle.close()
                        except:
                            pass
                    self.file_log_handle = None
                    self.caminho_log_atual = None  # Força a redetecção na próxima iteração do loop

            except Exception as e_monitor_loop:
                logging.error(f"[{thread_name}] Erro no loop principal de monitoramento de logs: {e_monitor_loop}",
                              exc_info=True)
                self.append_text_gui(f"Erro crítico ao monitorar logs: {e_monitor_loop}\n")

            if self._stop_event.wait(5):  # Intervalo de verificação (5 segundos)
                break  # Sair do loop se o evento de parada for acionado

        logging.info(f"[{thread_name}] Thread de monitoramento de log contínuo ({thread_name}) encerrada.")

    def obter_subpasta_mais_recente(self):
        if not self.pasta_raiz or not os.path.isdir(self.pasta_raiz): return None
        try:
            subpastas = [os.path.join(self.pasta_raiz, nome) for nome in os.listdir(self.pasta_raiz) if
                         os.path.isdir(os.path.join(self.pasta_raiz, nome))]
            if not subpastas: return None
            return max(subpastas, key=os.path.getmtime)  # Retorna o caminho completo
        except FileNotFoundError:  # Pode acontecer se a pasta_raiz for removida
            logging.warning(f"Pasta raiz '{self.pasta_raiz}' não encontrada ao buscar subpastas. Resetando pasta_raiz.")
            self.pasta_raiz = None  # Para evitar loops de erro
            return None
        except PermissionError:
            logging.error(
                f"Permissão negada ao acessar '{self.pasta_raiz}' para buscar subpastas. Resetando pasta_raiz.")
            self.pasta_raiz = None
            return None
        except Exception as e:  # Outros erros
            logging.error(f"Erro ao obter subpasta mais recente em '{self.pasta_raiz}': {e}", exc_info=True)
            return None

    def acompanhar_log_do_arquivo(self, caminho_log_designado_para_esta_thread):
        thread_name = threading.current_thread().name
        logging.info(f"[{thread_name}] Tentando iniciar acompanhamento para: {caminho_log_designado_para_esta_thread}")

        if self._stop_event.is_set():
            logging.info(
                f"[{thread_name}] _stop_event já está setado no início. Encerrando para {caminho_log_designado_para_esta_thread}.")
            return

        # Verificações cruciais no início da thread:
        if not self.file_log_handle or self.file_log_handle.closed:
            logging.error(
                f"[{thread_name}] ERRO CRÍTICO: file_log_handle está NULO ou FECHADO no início do acompanhamento para '{caminho_log_designado_para_esta_thread}'. Esta thread não pode prosseguir.")
            return

        try:
            # Garante que o self.file_log_handle (que é global para a classe)
            # é de fato o arquivo que esta thread específica foi designada para monitorar.
            handle_real_path_norm = os.path.normpath(self.file_log_handle.name)
            caminho_designado_norm = os.path.normpath(caminho_log_designado_para_esta_thread)

            if handle_real_path_norm != caminho_designado_norm:
                logging.warning(
                    f"[{thread_name}] DESCOMPASSO DE HANDLE NO INÍCIO! Thread designada para '{caminho_designado_norm}' mas self.file_log_handle atualmente aponta para '{handle_real_path_norm}'. Encerrando esta thread.")
                return
        except AttributeError:  # Caso self.file_log_handle seja None ou não tenha .name
            logging.error(
                f"[{thread_name}] ERRO CRÍTICO: self.file_log_handle é None ou inválido no início para '{caminho_log_designado_para_esta_thread}'. Encerrando.")
            return
        except Exception as e_check_init_handle:
            logging.error(
                f"[{thread_name}] Exceção na verificação inicial do handle para '{caminho_log_designado_para_esta_thread}': {e_check_init_handle}. Encerrando.")
            return

        logging.info(f"[{thread_name}] Iniciando acompanhamento EFETIVO de: {caminho_log_designado_para_esta_thread}")
        aguardando_winner = False
        vote_pattern_re, winner_pattern_re = None, None
        try:
            vote_pattern_str = self.vote_pattern_var.get()
            if vote_pattern_str: vote_pattern_re = re.compile(vote_pattern_str)
            winner_pattern_str = self.winner_pattern_var.get()
            if winner_pattern_str: winner_pattern_re = re.compile(winner_pattern_str)
        except re.error as e_re:
            logging.error(
                f"[{thread_name}] Erro de RegEx nos padrões para '{caminho_log_designado_para_esta_thread}': {e_re}",
                exc_info=True)
            self.append_text_gui(f"ERRO DE REGEX: Verifique os padrões: {e_re}\n")
            self.set_status_from_thread("Erro de RegEx! Verifique Configurações.")
            return  # Não continuar se os padrões são inválidos

        logging.info(
            f"[{thread_name}] Padrões para '{caminho_log_designado_para_esta_thread}': FimVoto='{vote_pattern_str}', Vencedor='{winner_pattern_str}'")
        logging.debug(
            f"[{thread_name}] Estado inicial de aguardando_winner para '{caminho_log_designado_para_esta_thread}': {aguardando_winner}")

        while not self._stop_event.is_set():
            if self._paused:
                if self._stop_event.wait(0.5): break
                continue

            # Verificações críticas DENTRO do loop para garantir consistência
            if not self.file_log_handle or self.file_log_handle.closed:
                logging.warning(
                    f"[{thread_name}] file_log_handle NULO ou FECHADO DENTRO DO LOOP para '{caminho_log_designado_para_esta_thread}'. Encerrando thread.")
                break
            try:
                current_handle_path_norm = os.path.normpath(self.file_log_handle.name)
                caminho_designado_norm_loop = os.path.normpath(caminho_log_designado_para_esta_thread)
                if current_handle_path_norm != caminho_designado_norm_loop:
                    logging.warning(
                        f"[{thread_name}] MUDANÇA DE HANDLE DETECTADA DURANTE O LOOP! Esta thread é para '{caminho_designado_norm_loop}', mas self.file_log_handle agora é '{current_handle_path_norm}'. Encerrando esta instância da thread.")
                    break
            except AttributeError:  # self.file_log_handle pode ter se tornado None
                logging.warning(
                    f"[{thread_name}] self.file_log_handle tornou-se None ou sem 'name' DENTRO DO LOOP para '{caminho_log_designado_para_esta_thread}'. Encerrando.")
                break
            except Exception as e_check_loop_attr_consistency:  # Outros erros na verificação
                logging.error(
                    f"[{thread_name}] Erro ao verificar consistência do handle no loop para '{caminho_log_designado_para_esta_thread}': {e_check_loop_attr_consistency}. Encerrando.")
                break

            try:
                linha = self.file_log_handle.readline()
                if linha:
                    linha_strip = linha.strip()
                    filtro_atual = self.filtro_var.get().strip().lower()  # Obter filtro atual
                    if not filtro_atual or filtro_atual in linha.lower():
                        self.append_text_gui(linha)

                    logging.debug(
                        f"[{thread_name}] LIDO de '{caminho_log_designado_para_esta_thread}': repr='{repr(linha)}', strip='{linha_strip}', aguardando_winner={aguardando_winner}")

                    if vote_pattern_re and vote_pattern_re.search(linha):
                        if not aguardando_winner:
                            logging.info(
                                f"[{thread_name}] Padrão de FIM DE VOTAÇÃO detectado em '{caminho_log_designado_para_esta_thread}'. Linha: '{linha_strip}'. Definindo aguardando_winner = True.")
                        else:  # Já estava aguardando winner, EndVote apareceu de novo?
                            logging.warning(
                                f"[{thread_name}] Padrão de FIM DE VOTAÇÃO detectado NOVAMENTE em '{caminho_log_designado_para_esta_thread}' enquanto aguardando_winner já era True. Linha: '{linha_strip}'.")
                        aguardando_winner = True
                        self.set_status_from_thread("Fim da votação detectado. Aguardando vencedor...")

                    if winner_pattern_re:
                        if aguardando_winner:  # Somente processa winner se está aguardando
                            logging.debug(
                                f"[{thread_name}] AGUARDANDO WINNER é TRUE para '{caminho_log_designado_para_esta_thread}'. Testando linha para Winner: repr='{repr(linha)}', strip='{linha_strip}'")
                            match = winner_pattern_re.search(linha)
                            if match:
                                try:
                                    indice_str = match.group(1)
                                    indice = int(indice_str)
                                    logging.info(
                                        f"[{thread_name}] Padrão de VENCEDOR detectado (aguardando_winner=True) em '{caminho_log_designado_para_esta_thread}'. Índice: {indice}. Linha: '{linha_strip}'")
                                    self.set_status_from_thread(f"Vencedor: índice {indice}. Processando...")

                                    # self.root.after para interagir com a GUI de forma segura
                                    if self.root.winfo_exists():
                                        self.root.after(0, self.processar_troca_mapa, indice)

                                    logging.debug(
                                        f"[{thread_name}] Winner processado para '{caminho_log_designado_para_esta_thread}', RESETANDO aguardando_winner para False.")
                                    aguardando_winner = False  # Resetar após processar o vencedor
                                except IndexError:
                                    logging.error(
                                        f"[{thread_name}] Padrão de vencedor '{winner_pattern_str}' casou em '{linha_strip}' para '{caminho_log_designado_para_esta_thread}', mas falta grupo de captura (group 1).")
                                    self.append_text_gui(
                                        f"ERRO: Padrão de vencedor '{winner_pattern_str}' não tem grupo de captura.\n")
                                    aguardando_winner = False  # Resetar mesmo em erro
                                except ValueError:
                                    logging.error(
                                        f"[{thread_name}] Padrão de vencedor '{winner_pattern_str}' capturou '{indice_str}' em '{linha_strip}' para '{caminho_log_designado_para_esta_thread}', que não é um número de índice válido.")
                                    self.append_text_gui(f"ERRO: Vencedor capturado '{indice_str}' não é um número.\n")
                                    aguardando_winner = False  # Resetar mesmo em erro
                                except Exception as e_proc_winner:  # Outros erros ao processar
                                    logging.error(
                                        f"[{thread_name}] Erro inesperado ao processar vencedor para '{caminho_log_designado_para_esta_thread}': {e_proc_winner}",
                                        exc_info=True)
                                    aguardando_winner = False  # Resetar
                            # else: # Se aguardando_winner é True mas não houve match do Winner nesta linha específica
                            #    logging.debug(f"[{thread_name}] AGUARDANDO WINNER é TRUE para '{caminho_log_designado_para_esta_thread}', mas padrão Winner NÃO CASOU na linha: strip='{linha_strip}'")

                        elif winner_pattern_re.search(
                                linha):  # Padrão do Winner casou, mas não estávamos esperando por ele
                            logging.info(
                                f"[{thread_name}] Padrão de vencedor APARECEU na linha '{linha_strip}' em '{caminho_log_designado_para_esta_thread}', MAS aguardando_winner era FALSO. Nenhum processamento de vencedor para esta linha.")
                else:  # Nenhuma linha nova
                    if self._stop_event.wait(0.2): break  # Esperar um pouco e checar se deve parar

            except UnicodeDecodeError as ude:
                logging.warning(
                    f"[{thread_name}] Erro de decodificação Unicode ao ler log {caminho_log_designado_para_esta_thread}: {ude}. Linha ignorada.")
            except ValueError as ve:  # Especificamente para file.readline() em arquivo fechado
                if "I/O operation on closed file" in str(ve).lower():
                    logging.warning(
                        f"[{thread_name}] Tentativa de I/O em arquivo fechado ({caminho_log_designado_para_esta_thread}). Encerrando thread.")
                    break  # Sair do loop se o arquivo foi fechado externamente
                else:  # Outro ValueError
                    logging.error(
                        f"[{thread_name}] Erro de ValueError ao acompanhar log {caminho_log_designado_para_esta_thread}: {ve}",
                        exc_info=True)
                    break  # Sair em outros ValueErrors também, pode ser problemático
            except Exception as e_tail_loop:
                if not self._stop_event.is_set():  # Não logar erro se o evento de parada foi acionado globalmente
                    logging.error(
                        f"[{thread_name}] Erro Inesperado ao acompanhar log {caminho_log_designado_para_esta_thread}: {e_tail_loop}",
                        exc_info=True)
                    self.append_text_gui(f"Erro ao ler log: {e_tail_loop}\n")
                    self.set_status_from_thread("Erro na leitura do log. Verifique o Log do Sistema.")
                break  # Sair do loop de acompanhamento do arquivo em caso de erro grave

        logging.info(
            f"[{thread_name}] Acompanhamento de '{caminho_log_designado_para_esta_thread}' encerrado. Estado final de aguardando_winner: {aguardando_winner}")
        # O file_log_handle é gerenciado pela LogMonitorThread; esta thread não o fecha diretamente.

    def processar_troca_mapa(self, indice_vencedor):
        logging.info(f"Processando troca de mapa para o índice: {indice_vencedor}")
        if not self.arquivo_json or not self.arquivo_json_votemap:
            msg = "Arquivos JSON de servidor ou votemap não configurados para troca de mapa."
            self.append_text_gui(f"ERRO: {msg}\n");
            logging.error(msg)
            self.set_status_from_thread("Erro: JSONs não configurados.")
            return

        try:
            with open(self.arquivo_json_votemap, 'r', encoding='utf-8') as f_vm:
                votemap_data = json.load(f_vm)
        except FileNotFoundError:
            msg = f"Arquivo votemap.json ('{self.arquivo_json_votemap}') não encontrado."
            self.append_text_gui(f"ERRO: {msg}\n");
            logging.error(msg)
            self.set_status_from_thread("Erro: votemap.json não encontrado.")
            return
        except json.JSONDecodeError as e:
            msg = f"Erro ao decodificar votemap.json: {e}"
            self.append_text_gui(f"ERRO: {msg}\n");
            logging.error(msg)
            self.set_status_from_thread("Erro: votemap.json inválido.")
            return

        map_list = votemap_data.get("list", [])
        if not map_list:
            msg = "Lista de mapas vazia ou não encontrada no votemap.json."
            self.append_text_gui(f"AVISO: {msg}\n");
            logging.warning(msg)
            self.set_status_from_thread("Aviso: Lista de mapas vazia.")
            return

        novo_scenario_id = None
        if indice_vencedor == 0:  # Voto aleatório
            if len(map_list) > 1:  # Precisa de pelo menos 2 mapas (random + o próprio 'random' na pos 0)
                # O índice 0 na lista é o "random", então escolhemos de 1 até o fim da lista
                indice_selecionado_random = random.randint(1, len(map_list) - 1)
                novo_scenario_id = map_list[indice_selecionado_random]
                self.append_text_gui(
                    f"Voto aleatório: selecionado mapa '{novo_scenario_id}' (índice real na lista: {indice_selecionado_random}).\n")
                logging.info(f"Seleção aleatória: {novo_scenario_id} (índice {indice_selecionado_random})")
            else:
                msg = "Voto aleatório, mas não há mapas suficientes para escolher (além da opção 'random')."
                self.append_text_gui(f"AVISO: {msg}\n");
                logging.warning(msg)
                self.set_status_from_thread("Aviso: Poucos mapas para aleatório.")
                return
        elif 0 < indice_vencedor < len(map_list):  # Voto em um mapa específico
            novo_scenario_id = map_list[indice_vencedor]
            self.append_text_gui(f"Mapa vencedor: '{novo_scenario_id}' (índice {indice_vencedor}).\n")
            logging.info(f"Mapa vencedor selecionado: {novo_scenario_id}")
        else:
            msg = f"Índice do mapa vencedor ({indice_vencedor}) inválido para a lista de mapas (tamanho {len(map_list)})."
            self.append_text_gui(f"ERRO: {msg}\n");
            logging.error(msg)
            self.set_status_from_thread("Erro: Índice de mapa inválido.")
            return

        try:
            with open(self.arquivo_json, 'r+', encoding='utf-8') as f_srv:
                server_data = json.load(f_srv)
                server_data["game"]["scenarioId"] = novo_scenario_id
                f_srv.seek(0);
                json.dump(server_data, f_srv, indent=4);
                f_srv.truncate()

            # Atualizar a GUI com o novo JSON do servidor via self.root.after
            if self.root.winfo_exists():
                self.root.after(0, self.exibir_json, self.json_text_area, json.dumps(server_data, indent=4))

            self.append_text_gui(f"JSON do servidor atualizado para o mapa: {novo_scenario_id}\n")
            logging.info(f"JSON do servidor atualizado com scenarioId: {novo_scenario_id}")

            if self.auto_restart_var.get() and self.nome_servico:
                self.append_text_gui("Iniciando reinício automático do servidor...\n")
                threading.Thread(target=self.reiniciar_servidor_com_progresso, args=(novo_scenario_id,),
                                 daemon=True, name="ServidorRestartThread").start()
            else:
                self.set_status_from_thread(
                    f"Mapa alterado para: {os.path.basename(str(novo_scenario_id))}. Reinício manual necessário.")
                logging.info("Reinício automático desabilitado ou serviço não configurado.")
        except FileNotFoundError:
            msg = f"Arquivo de config. do servidor ('{self.arquivo_json}') não encontrado."
            self.append_text_gui(f"ERRO: {msg}\n");
            logging.error(msg)
            self.set_status_from_thread("Erro: server.json não encontrado.")
        except (KeyError, TypeError) as e:
            msg = f"Estrutura do JSON do servidor inválida (game -> scenarioId não encontrado): {e}"
            self.append_text_gui(f"ERRO: {msg}\n");
            logging.error(msg)
            self.set_status_from_thread("Erro: Estrutura server.json inválida.")
        except json.JSONDecodeError as e:
            msg = f"Erro ao decodificar JSON do servidor: {e}"
            self.append_text_gui(f"ERRO: {msg}\n");
            logging.error(msg)
            self.set_status_from_thread("Erro: server.json inválido.")
        except Exception as e:
            msg = f"Erro inesperado ao processar troca de mapa: {e}"
            self.append_text_gui(f"ERRO: {msg}\n");
            logging.error(msg, exc_info=True)
            self.set_status_from_thread("Erro inesperado na troca de mapa.")

    def verificar_status_servico(self, nome_servico):
        if not nome_servico: return "NOT_FOUND"
        try:
            startupinfo = None
            if platform.system() == "Windows":
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                startupinfo.wShowWindow = subprocess.SW_HIDE

            result = subprocess.run(['sc', 'query', nome_servico], capture_output=True, text=True, check=False,
                                    startupinfo=startupinfo,
                                    encoding='latin-1')  # Tentar 'cp850' ou 'cp1252' se latin-1 falhar

            service_not_found_errors = ["failed 1060", "falha 1060",
                                        "지정된 서비스를 설치된 서비스로 찾을 수 없습니다."]  # Adicionar mais se necessário
            output_lower = result.stdout.lower() + result.stderr.lower()  # Combinar stdout e stderr

            for err_string in service_not_found_errors:
                if err_string in output_lower:
                    logging.warning(f"Serviço '{nome_servico}' não encontrado via 'sc query'. Output: {output_lower}")
                    return "NOT_FOUND"

            if "state" not in output_lower:  # Checagem básica se a saída é esperada
                logging.warning(
                    f"Saída inesperada do 'sc query {nome_servico}': STDOUT='{result.stdout}' STDERR='{result.stderr}'")
                return "ERROR"  # Ou UNKNOWN

            if "running" in output_lower: return "RUNNING"
            if "stopped" in output_lower: return "STOPPED"
            if "start_pending" in output_lower: return "START_PENDING"
            if "stop_pending" in output_lower: return "STOP_PENDING"
            return "UNKNOWN"  # Se nenhum estado conhecido for encontrado
        except FileNotFoundError:
            logging.error("'sc.exe' não encontrado. Verifique se o System32 está no PATH.", exc_info=True)
            self.show_messagebox_from_thread("error", "Erro de Comando",
                                             "Comando 'sc.exe' não encontrado. Verifique as configurações do sistema.")
            return "ERROR"
        except Exception as e:
            logging.error(f"Erro ao verificar status do serviço '{nome_servico}': {e}", exc_info=True)
            return "ERROR"

    def reiniciar_servidor_com_progresso(self, novo_scenario_id_para_log):
        # Esta função roda em uma thread, então interações com GUI devem usar self.root.after
        progress_win, pb = None, None  # Inicializar
        if self.root.winfo_exists():
            # Não podemos criar Toplevel diretamente de outra thread.
            # Uma solução mais complexa seria usar uma queue para comunicar com a thread principal para criar a dialog.
            # Por simplicidade aqui, vamos apenas logar o progresso.
            # Se uma dialog de progresso for essencial, a lógica precisa ser reestruturada.
            logging.info(f"Iniciando processo de reinício do servidor {self.nome_servico} em background.")
            self.set_status_from_thread(f"Reiniciando {self.nome_servico}...")

        success = self._reiniciar_servidor_logica(novo_scenario_id_para_log)

        # Destruir a dialog de progresso (se fosse criada na thread principal)
        # if progress_win and progress_win.winfo_exists():
        #     if pb: pb.stop()
        #     progress_win.destroy()

        if success:
            self.show_messagebox_from_thread("info", "Servidor Reiniciado",
                                             f"O serviço {self.nome_servico} foi reiniciado com sucesso.")
        else:
            self.show_messagebox_from_thread("error", "Falha no Reinício",
                                             f"Ocorreu um erro ao reiniciar o serviço {self.nome_servico}.")

    def _reiniciar_servidor_logica(self, novo_scenario_id_para_log):
        # Esta função é chamada por uma thread. Atualizações de GUI (status_label_var, append_text_gui)
        # devem ser feitas via self.root.after.
        if not self.nome_servico:
            self.append_text_gui_threadsafe("ERRO: Nome do serviço não configurado para reinício.\n")
            logging.error("Tentativa de reiniciar servidor sem nome de serviço configurado.")
            self.set_status_from_thread("Erro: Serviço não configurado.")
            return False

        stop_delay = self.stop_delay_var.get()
        start_delay = self.start_delay_var.get()
        default_votemap_mission = self.default_mission_var.get()

        startupinfo = None
        if platform.system() == "Windows":
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = subprocess.SW_HIDE
        try:
            self.set_status_from_thread(f"Parando serviço {self.nome_servico}...")
            self.append_text_gui_threadsafe(f"Parando serviço '{self.nome_servico}'...\n")
            logging.info(f"Tentando parar o serviço: {self.nome_servico}")
            status = self.verificar_status_servico(self.nome_servico)

            if status == "RUNNING" or status == "START_PENDING":
                subprocess.run(["sc", "stop", self.nome_servico], check=True, shell=False, startupinfo=startupinfo)
                self.append_text_gui_threadsafe(f"Comando de parada enviado. Aguardando {stop_delay}s...\n")
                time.sleep(stop_delay)  # Bloqueia esta thread, não a GUI
                status_after_stop = self.verificar_status_servico(self.nome_servico)
                if status_after_stop != "STOPPED":
                    logging.warning(f"Serviço {self.nome_servico} não parou como esperado. Status: {status_after_stop}")
                    self.append_text_gui_threadsafe(
                        f"AVISO: Serviço '{self.nome_servico}' pode não ter parado. Status: {status_after_stop}\n")
            elif status == "STOPPED":
                self.append_text_gui_threadsafe(f"Serviço '{self.nome_servico}' já estava parado.\n")
                logging.info(f"Serviço {self.nome_servico} já estava parado.")
            elif status == "NOT_FOUND":
                self.append_text_gui_threadsafe(f"ERRO: Serviço '{self.nome_servico}' não encontrado.\n")
                logging.error(f"Serviço {self.nome_servico} não encontrado para parada.")
                self.set_status_from_thread(f"Erro: Serviço '{self.nome_servico}' não existe.")
                return False
            else:  # UNKNOWN ou ERROR
                self.append_text_gui_threadsafe(
                    f"ERRO: Não foi possível determinar o estado do serviço '{self.nome_servico}' ou estado inesperado: {status}.\n")
                logging.error(f"Estado do serviço {self.nome_servico} desconhecido ou erro: {status}")
                self.set_status_from_thread(f"Erro: Estado de '{self.nome_servico}' desconhecido.")
                return False

            self.set_status_from_thread(f"Iniciando serviço {self.nome_servico}...")
            self.append_text_gui_threadsafe(f"Iniciando serviço '{self.nome_servico}'...\n")
            logging.info(f"Tentando iniciar o serviço: {self.nome_servico}")
            subprocess.run(["sc", "start", self.nome_servico], check=True, shell=False, startupinfo=startupinfo)
            self.append_text_gui_threadsafe(
                f"Comando de início enviado. Aguardando {start_delay}s para o servidor estabilizar...\n")
            self.set_status_from_thread(f"Aguardando {self.nome_servico} iniciar ({start_delay}s)...")
            time.sleep(start_delay)  # Bloqueia esta thread

            status_after_start = self.verificar_status_servico(self.nome_servico)
            if status_after_start != "RUNNING":
                logging.error(f"Serviço {self.nome_servico} falhou ao iniciar. Status: {status_after_start}")
                self.append_text_gui_threadsafe(
                    f"ERRO: Serviço '{self.nome_servico}' falhou ao iniciar ou está demorando muito. Status: {status_after_start}\n")
                self.set_status_from_thread(f"Erro: {self.nome_servico} não iniciou. Status: {status_after_start}")
                return False
            logging.info(f"Serviço {self.nome_servico} iniciado com sucesso.")

            self.append_text_gui_threadsafe("Restaurando JSON do servidor para o mapa de votação padrão...\n")
            if not self.arquivo_json or not os.path.exists(self.arquivo_json):
                msg = f"Arquivo JSON do servidor ({self.arquivo_json}) não encontrado para restaurar votemap."
                self.append_text_gui_threadsafe(f"ERRO: {msg}\n");
                logging.error(msg)
                self.set_status_from_thread("Erro: server.json não encontrado para reset.")
                return False

            with open(self.arquivo_json, 'r+', encoding='utf-8') as f_srv:
                server_data = json.load(f_srv)
                server_data["game"]["scenarioId"] = default_votemap_mission
                f_srv.seek(0);
                json.dump(server_data, f_srv, indent=4);
                f_srv.truncate()

            # Atualizar a GUI com o JSON resetado
            if self.root.winfo_exists():
                self.root.after(0, self.exibir_json, self.json_text_area, json.dumps(server_data, indent=4))

            self.append_text_gui_threadsafe(f"JSON do servidor restaurado para votemap: {default_votemap_mission}\n")
            logging.info(f"JSON do servidor restaurado para scenarioId de votemap: {default_votemap_mission}")
            self.set_status_from_thread(
                f"Servidor reiniciado. Mapa anterior: {os.path.basename(str(novo_scenario_id_para_log))}. Próximo: Votação.")
            return True

        except subprocess.CalledProcessError as e:
            err_output = e.stderr.decode('latin-1', errors='replace') if e.stderr else (
                e.stdout.decode('latin-1', errors='replace') if e.stdout else "Nenhuma saída de erro detalhada.")
            err_msg = f"Erro ao executar comando 'sc' para '{self.nome_servico}': {err_output.strip()}"
            self.append_text_gui_threadsafe(f"ERRO: {err_msg}\n");
            logging.error(err_msg, exc_info=True)
            self.set_status_from_thread(f"Erro ao gerenciar serviço: {e.cmd}")
            return False
        except FileNotFoundError:  # Para sc.exe
            self.show_messagebox_from_thread("error", "Erro de Comando",
                                             "Comando 'sc.exe' não encontrado. Verifique o PATH.")
            logging.error("Comando 'sc.exe' não encontrado.")
            self.set_status_from_thread("Erro: sc.exe não encontrado.")
            return False
        except (json.JSONDecodeError, KeyError, TypeError) as e:
            err_msg = f"Erro ao manipular JSON do servidor durante o reinício: {e}"
            self.append_text_gui_threadsafe(f"ERRO: {err_msg}\n");
            logging.error(err_msg, exc_info=True)
            self.set_status_from_thread("Erro: Falha ao atualizar JSON do servidor.")
            return False
        except Exception as e:
            err_msg = f"Erro inesperado ao reiniciar o servidor: {e}"
            self.append_text_gui_threadsafe(f"ERRO: {err_msg}\n");
            logging.error(err_msg, exc_info=True)
            self.set_status_from_thread("Erro inesperado no reinício do servidor.")
            return False

    # Funções auxiliares para interagir com a GUI de forma thread-safe
    def set_status_from_thread(self, message):
        if self.root.winfo_exists():
            self.root.after(0, lambda: self.status_label_var.set(message))

    def append_text_gui_threadsafe(self, texto):
        if self.root.winfo_exists() and hasattr(self, 'text_area') and self.text_area.winfo_exists():
            self.root.after(0, self._append_text_gui_actual, texto)

    def _append_text_gui_actual(self, texto):
        # Esta função é chamada pela thread da GUI via self.root.after
        if self.text_area.winfo_exists():  # Dupla checagem
            self.text_area.configure(state='normal')
            self.text_area.insert('end', texto)
            self.text_area.yview_moveto(1.0)
            self.text_area.configure(state='disabled')

    def show_messagebox_from_thread(self, boxtype, title, message):
        if self.root.winfo_exists():
            if boxtype == "info":
                self.root.after(0, lambda t=title, m=message: Messagebox.show_info(m, t, parent=self.root))
            elif boxtype == "error":
                self.root.after(0, lambda t=title, m=message: Messagebox.show_error(m, t, parent=self.root))
            elif boxtype == "warning":
                self.root.after(0, lambda t=title, m=message: Messagebox.show_warning(m, t, parent=self.root))
            # Adicionar outros tipos se necessário

    # Funções da UI (sem alterações de lógica, apenas chamadas de GUI)
    def append_text_gui(self, texto):  # Esta é chamada pela thread principal ou de forma segura
        if self.root.winfo_exists() and hasattr(self, 'text_area') and self.text_area.winfo_exists():
            # Se chamada da thread principal, pode ser direta. Se de outra, deveria usar append_text_gui_threadsafe
            # Para simplificar, vamos assumir que se esta for chamada diretamente, é da thread da GUI
            # ou o chamador sabe o que está fazendo. Para chamadas de threads, usar a versão _threadsafe.
            try:
                self.text_area.configure(state='normal')
                self.text_area.insert('end', texto)
                self.text_area.yview_moveto(1.0)
                self.text_area.configure(state='disabled')
            except tk.TclError as e:
                logging.warning(f"TclError em append_text_gui (provavelmente GUI fechando): {e}")

    def limpar_tela(self):
        self.text_area.configure(state='normal');
        self.text_area.delete('1.0', 'end');
        self.text_area.configure(state='disabled')
        self.status_label_var.set("Tela de logs do servidor limpa.");  # OK, da thread da GUI
        logging.info("Tela de logs do servidor limpa pelo usuário.")

    def toggle_pausa(self):
        self._paused = not self._paused
        if self._paused:
            self.pausar_btn.config(text="▶️ Retomar", bootstyle=SUCCESS)
            self.status_label_var.set("Monitoramento de logs pausado.")  # OK
            logging.info("Monitoramento de logs pausado.")
        else:
            self.pausar_btn.config(text="⏸️ Pausar", bootstyle=WARNING)
            self.status_label_var.set("Monitoramento de logs retomado.")  # OK
            logging.info("Monitoramento de logs retomado.")

    def trocar_tema(self, event=None):
        novo_tema = self.tema_var.get()
        try:
            self.style.theme_use(novo_tema)
            logging.info(f"Tema alterado para: {novo_tema}")
            self.status_label_var.set(f"Tema alterado para '{novo_tema}'.")  # OK
        except Exception as e:
            logging.error(f"Erro ao tentar trocar para o tema '{novo_tema}': {e}", exc_info=True)
            self.show_messagebox_from_thread("error", "Erro de Tema",
                                             f"Não foi possível aplicar o tema '{novo_tema}'.\n{e}")
            try:  # Tentar voltar para um tema padrão seguro
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
                self.status_label_var.set(f"Logs da tela exportados para: {os.path.basename(caminho_arquivo)}")  # OK
                logging.info(f"Logs da tela exportados para: {caminho_arquivo}")
                self.show_messagebox_from_thread("info", "Exportação Concluída",
                                                 f"Logs da tela foram exportados com sucesso para:\n{caminho_arquivo}")
            except Exception as e:
                self.status_label_var.set(f"Erro ao exportar logs da tela: {e}")  # OK
                logging.error(f"Erro ao exportar logs da tela para {caminho_arquivo}: {e}", exc_info=True)
                self.show_messagebox_from_thread("error", "Erro de Exportação", f"Falha ao exportar logs da tela:\n{e}")

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
            self.show_messagebox_from_thread("warning", "Validação de Configurações",
                                             "Os seguintes problemas de configuração foram encontrados:\n\n" + "\n".join(
                                                 problemas))
        else:
            self.show_messagebox_from_thread("info", "Validação de Configurações",
                                             "Todas as configurações essenciais parecem estar corretas!")
        logging.info(f"Validação de configurações: {len(problemas)} problemas encontrados.")

    def show_about(self):
        about_win = ttk.Toplevel(self.root);
        about_win.title("Sobre Predadores Votemap Patch");
        about_win.geometry("450x350")  # Ajustar tamanho conforme necessário
        about_win.resizable(False, False);
        about_win.transient(self.root);
        about_win.grab_set()
        frame = ttk.Frame(about_win, padding=20);
        frame.pack(fill='both', expand=True)
        ttk.Label(frame, text="Predadores Votemap Patch", font="-size 16 -weight bold").pack(pady=(0, 10))
        ttk.Label(frame, text="Versão 2.2 (Ícone Integrado)", font="-size 10").pack()  # Mantenha sua versão
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
                pystray.MenuItem('Sair', self.on_close_from_tray_menu_item)  # Nome diferente para clareza
            )
            self.tray_icon = pystray.Icon("predadores_votemap_patch", image, "Predadores Votemap Patch", menu)
            threading.Thread(target=self.tray_icon.run, daemon=True, name="TrayIconThread").start()
            logging.info("Ícone da bandeja do sistema configurado e iniciado.")
        except Exception as e:
            logging.error(f"Falha ao criar ícone da bandeja: {e}", exc_info=True)

    def show_from_tray(self):
        if self.root.winfo_exists():  # Checar se a janela ainda existe
            self.root.after(0, self.root.deiconify)  # Usar after para garantir que é da thread da GUI

    def minimize_to_tray(self, event=None):
        # Chamado quando a janela é minimizada (evento <Unmap> ou state() == 'iconic')
        if hasattr(self, 'tray_icon') and self.tray_icon.visible:
            if self.root.winfo_exists() and self.root.state() == 'iconic':  # Checar se foi realmente minimizada
                self.root.withdraw()  # Esconder a janela da barra de tarefas
                logging.info("Aplicação minimizada para a bandeja.")

    def on_close_from_tray_menu_item(self):  # Chamado pelo item de menu 'Sair' da bandeja
        logging.info("Comando 'Sair' do menu da bandeja recebido.")
        self.on_close_common_logic(initiated_by_tray=True)  # Forçar fechamento

    def on_close(self):  # Chamado pelo 'X' da janela
        logging.info("Tentativa de fechar a janela principal (WM_DELETE_WINDOW).")
        if self.root.winfo_exists():
            # Usar after para Messagebox para garantir que seja da thread da GUI, especialmente se on_close puder ser chamado de outra forma
            self.root.after(0, self._confirm_close_dialog)

    def _confirm_close_dialog(self):
        if Messagebox.okcancel("Confirmar Saída", "Deseja realmente sair do Predadores Votemap Patch?",
                               parent=self.root, alert=True) == "OK":  # alert=True para garantir que fique no topo
            self.on_close_common_logic()
        else:
            logging.info("Saída cancelada pelo usuário (via janela).")

    def on_close_common_logic(self, initiated_by_tray=False):
        logging.info(
            f"Iniciando lógica comum de fechamento (iniciado por {'bandeja' if initiated_by_tray else 'janela'}).")
        if self.root.winfo_exists():
            self.set_status_from_thread("Encerrando...")  # Usar a versão thread-safe
            if not initiated_by_tray:  # A atualização da UI pode não ser necessária se a janela já estiver fechada
                try:
                    self.root.update_idletasks()
                except tk.TclError:
                    pass  # Ignorar erro se a janela já estiver sendo destruída

        self.stop_log_monitoring()  # Parar todas as threads de monitoramento e seus recursos

        if hasattr(self, 'tray_icon') and self.tray_icon.visible:  # Checar visible
            try:
                self.tray_icon.stop()
            except Exception as e_tray:
                logging.error(f"Erro ao parar ícone da bandeja: {e_tray}", exc_info=True)

        try:
            self.save_config()  # Salvar configuração
        except Exception as e_save:
            logging.error(f"Erro ao salvar configuração ao sair: {e_save}", exc_info=True)

        logging.info(f"Aplicação encerrada (via {'bandeja' if initiated_by_tray else 'janela'}).")
        if self.root.winfo_exists():
            self.root.destroy()  # Destruir a janela principal


def main():
    # pythoncom.CoInitialize() # Descomentar se WMI for usado na thread principal ANTES de outras threads
    root = ttk.Window(themename="darkly")
    app = LogViewerApp(root)
    root.protocol("WM_DELETE_WINDOW", app.on_close)  # Chama on_close quando o 'X' da janela é clicado
    root.bind("<Unmap>", app.minimize_to_tray)  # Para minimizar para a bandeja quando a janela é minimizada
    app.setup_tray_icon()  # Configura o ícone da bandeja

    try:
        root.mainloop()
    except KeyboardInterrupt:  # Lidar com Ctrl+C no console
        logging.info("Interrupção por teclado recebida. Encerrando...")
        app.on_close_common_logic(initiated_by_tray=True)  # Forçar fechamento sem confirmação
    finally:
        # pythoncom.CoUninitialize() # Descomentar se CoInitialize foi usado
        logging.info("Aplicação finalizada (bloco finally do main).")


if __name__ == '__main__':
    def handle_thread_exception(args):
        logging.critical(f"EXCEÇÃO NÃO CAPTURADA NA THREAD {args.thread.name if args.thread else 'Desconhecida'}:",
                         # Usar args.thread.name
                         exc_info=(args.exc_type, args.exc_value, args.exc_traceback))
        import traceback
        traceback.print_exception(args.exc_type, args.exc_value, args.exc_traceback)


    threading.excepthook = handle_thread_exception
    main()