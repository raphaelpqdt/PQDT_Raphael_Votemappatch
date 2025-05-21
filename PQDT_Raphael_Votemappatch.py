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
import sys
import webbrowser

import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap.tooltip import ToolTip
from ttkbootstrap.dialogs import Messagebox
from tkinter import filedialog
from tkinter.scrolledtext import ScrolledText

import pystray
from PIL import Image, ImageDraw

try:
    import win32com.client
    import pythoncom

    PYWIN32_AVAILABLE = True
except ImportError:
    PYWIN32_AVAILABLE = False
    # O logging será configurado depois, então não podemos logar aqui ainda.
    # A mensagem de erro será mostrada na UI se necessário.

# Configure logging
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(threadName)s - %(module)s.%(funcName)s - %(message)s',
    filename='votemap_patch.log',
    filemode='a'
)


def resource_path(relative_path):
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
        self.config_changed = False
        self.loading_config = True

        self.log_folder_path_label_var = ttk.StringVar(value="Pasta Logs: Nenhuma")
        self.server_json_path_label_var = ttk.StringVar(value="JSON Servidor: Nenhum")
        self.votemap_json_path_label_var = ttk.StringVar(value="JSON Votemap: Nenhum")
        self.servico_var = ttk.StringVar(value="Nenhum serviço selecionado")

        self.config = {}
        self.tema_var = ttk.StringVar()
        self.filtro_var = ttk.StringVar()
        self.auto_restart_var = ttk.BooleanVar()
        self.vote_pattern_var = ttk.StringVar()
        self.winner_pattern_var = ttk.StringVar()
        self.default_mission_var = ttk.StringVar()
        self.stop_delay_var = ttk.IntVar()
        self.start_delay_var = ttk.IntVar()

        self.config = self.load_config()

        self.style = ttk.Style()
        default_theme = "litera"
        try:
            current_theme_from_config = self.config.get("theme", default_theme)
            self.style.theme_use(current_theme_from_config)
            logging.info(f"Tema '{current_theme_from_config}' aplicado com sucesso.")
        except tk.TclError:
            logging.warning(
                f"Tema '{self.config.get('theme')}' não encontrado ou inválido. Usando tema padrão '{default_theme}'.")
            try:
                self.style.theme_use(default_theme)
                if "theme" in self.config:
                    self.config["theme"] = default_theme
            except tk.TclError:  # Se até o default_theme falhar (improvável com ttkbootstrap)
                logging.error(f"Tema padrão '{default_theme}' também falhou. Verifique a instalação do ttkbootstrap.")
                # A aplicação pode ter problemas visuais sérios aqui.

        self.tema_var.set(self.style.theme)
        self.filtro_var.set(self.config.get("filter", ""))
        self.auto_restart_var.set(self.config.get("auto_restart", True))
        self.vote_pattern_var.set(self.config.get("vote_pattern", r"\.EndVote\(\)"))
        self.winner_pattern_var.set(self.config.get("winner_pattern", r"Winner: \[(\d+)\]"))
        self.default_mission_var.set(
            self.config.get("default_mission", "{B88CC33A14B71FDC}Missions/V30_MapVoting_Mission.conf"))
        self.stop_delay_var.set(self.config.get("stop_delay", 3))
        self.start_delay_var.set(self.config.get("start_delay", 15))

        self.tema_var.trace_add("write", self._config_value_changed)
        self.filtro_var.trace_add("write", self._config_value_changed)
        self.auto_restart_var.trace_add("write", self._config_value_changed)
        self.vote_pattern_var.trace_add("write", self._config_value_changed)
        self.winner_pattern_var.trace_add("write", self._config_value_changed)
        self.default_mission_var.trace_add("write", self._config_value_changed)
        self.stop_delay_var.trace_add("write", self._config_value_changed)
        self.start_delay_var.trace_add("write", self._config_value_changed)

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

        self.auto_scroll_log_var = ttk.BooleanVar(value=True)
        self.log_search_var = ttk.StringVar()
        self.last_search_pos = "1.0"
        self.search_log_frame_visible = False

        self.create_menu()
        self.create_ui()
        self.create_status_bar()

        self.loading_config = False
        self.initialize_from_config()
        self.config_changed = False
        self._update_save_buttons_state()

        self.root.after(100, self.atualizar_log_sistema_periodicamente)

        if self.pasta_raiz:
            self.start_log_monitoring()

        if not PYWIN32_AVAILABLE:
            self.show_messagebox_from_thread("warning", "Dependência Faltando",
                                             "pywin32 não encontrado. A funcionalidade de serviço do Windows (listar, iniciar, parar) estará desabilitada.")
            if hasattr(self, 'servico_btn') and self.servico_btn:
                self.servico_btn.config(state="disabled")
            if hasattr(self, 'refresh_servico_status_btn') and self.refresh_servico_status_btn:
                self.refresh_servico_status_btn.config(state="disabled")
            if hasattr(self, 'auto_restart_check') and self.auto_restart_check:  # Precisa do nome do checkbutton
                # Se o auto_restart_check estiver na aba de configurações, você precisaria desabilitá-lo
                # ou o label associado, e talvez o auto_restart_var.set(False)
                pass

        logging.info("Aplicação iniciada.")

    def _config_value_changed(self, *args):
        if self.loading_config:
            return
        if not self.config_changed:
            self.config_changed = True
            self.set_status_from_thread("Configurações alteradas. Não se esqueça de salvar.")
        self._update_save_buttons_state()

    def _update_save_buttons_state(self):
        state = "normal" if self.config_changed else "disabled"
        try:
            if hasattr(self, 'file_menu') and self.file_menu.winfo_exists():
                self.file_menu.entryconfigure("Salvar Configuração", state=state)
            if hasattr(self, 'save_config_button_app_settings') and self.save_config_button_app_settings.winfo_exists():
                self.save_config_button_app_settings.config(state=state)
        except tk.TclError:
            logging.debug("TclError em _update_save_buttons_state (widgets podem estar sendo destruídos).")

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
        self._stop_event.clear()
        self.log_monitor_thread = threading.Thread(target=self.monitorar_log_continuamente, daemon=True,
                                                   name="LogMonitorThread")
        self.log_monitor_thread.start()
        logging.info("Monitoramento de logs iniciado.")

    def stop_log_monitoring(self):
        current_thread_name = threading.current_thread().name
        logging.debug(f"[{current_thread_name}] Chamada para stop_log_monitoring.")
        self._stop_event.set()

        if self.log_tail_thread and self.log_tail_thread.is_alive():
            logging.debug(
                f"[{current_thread_name}] Aguardando LogTailThread ({self.log_tail_thread.name}) finalizar...")
            self.log_tail_thread.join(timeout=2.0)
            if self.log_tail_thread.is_alive():
                logging.warning(
                    f"[{current_thread_name}] LogTailThread ({self.log_tail_thread.name}) não finalizou no tempo esperado.")

        if self.log_monitor_thread and self.log_monitor_thread.is_alive() and self.log_monitor_thread != threading.current_thread():
            logging.debug(
                f"[{current_thread_name}] Aguardando LogMonitorThread ({self.log_monitor_thread.name}) finalizar...")
            self.log_monitor_thread.join(timeout=2.0)
            if self.log_monitor_thread.is_alive():
                logging.warning(
                    f"[{current_thread_name}] LogMonitorThread ({self.log_monitor_thread.name}) não finalizou no tempo esperado.")

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
                self.file_log_handle = None

        self.caminho_log_atual = None
        logging.info(f"[{current_thread_name}] stop_log_monitoring completado.")

    def atualizar_log_sistema_periodicamente(self):
        try:
            if not self.root.winfo_exists() or not hasattr(self, 'notebook') or not self.notebook.winfo_exists(): return
            if not self.notebook.tabs(): return  # Se não há abas, não há o que selecionar
            current_selection = self.notebook.select()
            if not current_selection: return  # Nenhuma aba selecionada

            aba_atual_index = self.notebook.index(current_selection)
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
        except tk.TclError as e:
            if "invalid command name" not in str(e).lower():
                logging.error(f"TclError ao atualizar log do sistema na GUI: {e}", exc_info=False)
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
        self.file_menu = ttk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Arquivo", menu=self.file_menu)
        self.file_menu.add_command(label="Salvar Configuração", command=self.save_config, state="disabled")
        self.file_menu.add_command(label="Carregar Configuração", command=self.load_config_dialog)
        self.file_menu.add_separator()
        self.file_menu.add_command(label="Sair", command=self.on_close)

        tools_menu = ttk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Ferramentas", menu=tools_menu)
        tools_menu.add_command(label="Exportar Logs do App", command=self.export_display_logs)
        tools_menu.add_command(label="Verificar Configurações", command=self.validate_configs)
        tools_menu.add_command(label="Buscar no Log (Ctrl+F)",
                               command=lambda: self.toggle_log_search_bar(force_show=True))

        help_menu = ttk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Ajuda", menu=help_menu)
        help_menu.add_command(label="Sobre", command=self.show_about)
        help_menu.add_separator()
        help_menu.add_command(label="Verificar Atualizações...", command=self.check_for_updates)

    def create_ui(self):
        outer_top_frame = ttk.Frame(self.root)
        outer_top_frame.pack(pady=10, padx=10, fill='x')

        selection_labelframe = ttk.Labelframe(outer_top_frame, text="Configuração de Caminhos e Serviço",
                                              padding=(10, 5))
        selection_labelframe.pack(side='top', fill='x', pady=(0, 5))

        path_buttons_frame = ttk.Frame(selection_labelframe)
        path_buttons_frame.pack(fill='x')

        self.selecionar_btn = ttk.Button(path_buttons_frame, text="Selecionar Pasta de Logs",
                                         command=self.selecionar_pasta, bootstyle=PRIMARY)
        self.selecionar_btn.pack(side='left', pady=2, padx=(0, 5))
        ToolTip(self.selecionar_btn, text="Seleciona a pasta raiz onde os logs do servidor são armazenados.")

        self.json_btn = ttk.Button(path_buttons_frame, text="JSON do Servidor",
                                   command=self.selecionar_arquivo_json_servidor, bootstyle=INFO)
        self.json_btn.pack(side='left', padx=5, pady=2)
        ToolTip(self.json_btn, text="Seleciona o arquivo JSON de configuração principal do servidor.")

        self.json_vm_btn = ttk.Button(path_buttons_frame, text="JSON do Votemap",
                                      command=self.selecionar_arquivo_json_votemap, bootstyle=INFO)
        self.json_vm_btn.pack(side='left', padx=5, pady=2)
        ToolTip(self.json_vm_btn, text="Seleciona o arquivo JSON de configuração do Votemap.")

        self.servico_btn = ttk.Button(path_buttons_frame, text="Selecionar Serviço", command=self.selecionar_servico,
                                      bootstyle=SECONDARY)
        self.servico_btn.pack(side='left', padx=5, pady=2)
        ToolTip(self.servico_btn, text="Seleciona o serviço do Windows associado ao servidor do jogo.")

        self.refresh_servico_status_btn = ttk.Button(path_buttons_frame, text="↻",
                                                     command=self.update_service_status_display,
                                                     bootstyle=(TOOLBUTTON, LIGHT), width=2)
        self.refresh_servico_status_btn.pack(side='left', padx=(0, 5), pady=2)
        ToolTip(self.refresh_servico_status_btn, text="Atualizar status do serviço selecionado.")

        path_labels_frame = ttk.Frame(selection_labelframe)
        path_labels_frame.pack(fill='x', pady=(5, 0))

        self.log_folder_path_label = ttk.Label(path_labels_frame, textvariable=self.log_folder_path_label_var,
                                               wraplength=250, anchor='w')
        self.log_folder_path_label.pack(side='left', padx=5, fill='x', expand=True)

        self.json_server_path_label = ttk.Label(path_labels_frame, textvariable=self.server_json_path_label_var,
                                                wraplength=250, anchor='w')
        self.json_server_path_label.pack(side='left', padx=5, fill='x', expand=True)

        self.json_votemap_path_label = ttk.Label(path_labels_frame, textvariable=self.votemap_json_path_label_var,
                                                 wraplength=250, anchor='w')
        self.json_votemap_path_label.pack(side='left', padx=5, fill='x', expand=True)

        self.servico_label_widget = ttk.Label(path_labels_frame, textvariable=self.servico_var, anchor='w')
        self.servico_label_widget.pack(side='left', padx=(5, 0), fill='x', expand=True)

        controls_labelframe = ttk.Labelframe(outer_top_frame, text="Controles e Visualização", padding=(10, 5))
        controls_labelframe.pack(side='top', fill='x')

        log_controls_subframe = ttk.Frame(controls_labelframe)
        log_controls_subframe.pack(side='left', fill='x', expand=True, padx=(0, 10))

        ttk.Label(log_controls_subframe, text="Filtro:").pack(side='left', padx=(0, 5))
        self.filtro_entry = ttk.Entry(log_controls_subframe, textvariable=self.filtro_var, width=25)
        self.filtro_entry.pack(side='left')
        ToolTip(self.filtro_entry,
                text="Filtra as linhas de log exibidas (case-insensitive). Deixe em branco para nenhum filtro.")

        self.refresh_json_btn = ttk.Button(log_controls_subframe, text="Atualizar JSONs",
                                           command=self.forcar_refresh_json, bootstyle=SUCCESS)
        self.refresh_json_btn.pack(side='left', padx=5)
        ToolTip(self.refresh_json_btn, text="Recarrega e exibe o conteúdo dos arquivos JSON selecionados.")

        self.pausar_btn = ttk.Button(log_controls_subframe, text="⏸️ Pausar", command=self.toggle_pausa,
                                     bootstyle=WARNING)
        self.pausar_btn.pack(side='left', padx=5)
        ToolTip(self.pausar_btn, text="Pausa ou retoma o acompanhamento ao vivo dos logs.")

        self.limpar_btn = ttk.Button(log_controls_subframe, text="♻️ Limpar Tela", command=self.limpar_tela,
                                     bootstyle=SECONDARY)
        self.limpar_btn.pack(side='left', padx=5)
        ToolTip(self.limpar_btn, text="Limpa a área de exibição de logs do servidor.")

        theme_subframe = ttk.Frame(controls_labelframe)
        theme_subframe.pack(side='right')

        ttk.Label(theme_subframe, text="Tema:").pack(side='left', padx=(10, 5))
        self.tema_menu = ttk.Combobox(theme_subframe, textvariable=self.tema_var, values=self.style.theme_names(),
                                      width=12, state='readonly')
        self.tema_menu.pack(side='left')
        self.tema_menu.bind("<<ComboboxSelected>>", self.trocar_tema)
        ToolTip(self.tema_menu, text="Muda o tema visual da aplicação.")

        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill='both', expand=True, padx=10, pady=(0, 10))

        log_frame = ttk.Frame(self.notebook)
        self.notebook.add(log_frame, text="Logs do Servidor")
        self.log_label = ttk.Label(log_frame, text="LOG AO VIVO DO SERVIDOR", foreground="red")
        self.log_label.pack(pady=(5, 0))

        self.search_log_frame = ttk.Frame(log_frame)
        ttk.Label(self.search_log_frame, text="Buscar:").pack(side='left', padx=(5, 2))
        self.log_search_entry = ttk.Entry(self.search_log_frame, textvariable=self.log_search_var)
        self.log_search_entry.pack(side='left', fill='x', expand=True, padx=2)
        self.log_search_entry.bind("<Return>", self.search_log_next)
        search_next_btn = ttk.Button(self.search_log_frame, text="Próximo", command=self.search_log_next,
                                     bootstyle=SECONDARY)
        search_next_btn.pack(side='left', padx=2)
        search_prev_btn = ttk.Button(self.search_log_frame, text="Anterior", command=self.search_log_prev,
                                     bootstyle=SECONDARY)
        search_prev_btn.pack(side='left', padx=2)
        close_search_btn = ttk.Button(self.search_log_frame, text="X", command=self.toggle_log_search_bar,
                                      bootstyle=(SECONDARY, DANGER), width=2)
        close_search_btn.pack(side='left', padx=(2, 5))

        self.text_area = ScrolledText(log_frame, wrap='word', height=20, state='disabled')
        self.text_area.pack(fill='both', expand=True, pady=(0, 5))
        self.text_area.bind("<Control-f>", lambda e: self.toggle_log_search_bar(force_show=True))
        self.root.bind_all("<Escape>", lambda e: self.toggle_log_search_bar(
            force_hide=True) if self.search_log_frame_visible else None)

        self.auto_scroll_check = ttk.Checkbutton(log_frame, text="Rolar Automaticamente",
                                                 variable=self.auto_scroll_log_var)  # Nomeado
        self.auto_scroll_check.pack(side='bottom', anchor='se', pady=2, padx=5)

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

        self.auto_restart_check = ttk.Checkbutton(settings_inner_frame,
                                                  text="Reiniciar servidor automaticamente após troca de mapa",
                                                  variable=self.auto_restart_var)  # Nomeado
        self.auto_restart_check.grid(row=0, column=0, sticky='w', padx=5, pady=5, columnspan=2)
        ToolTip(self.auto_restart_check,
                "Se marcado, o servidor será reiniciado após uma votação de mapa bem-sucedida.")

        ttk.Label(settings_inner_frame, text="Padrão de detecção de voto (RegEx):").grid(row=1, column=0, sticky='w',
                                                                                         padx=5, pady=(15, 0))
        vote_pattern_entry = ttk.Entry(settings_inner_frame, textvariable=self.vote_pattern_var, width=50)
        vote_pattern_entry.grid(row=2, column=0, sticky='ew', padx=5, pady=5, columnspan=2)
        ToolTip(vote_pattern_entry, "Expressão regular para detectar o fim de uma votação no log.")

        ttk.Label(settings_inner_frame, text="Padrão de detecção de vencedor (RegEx):").grid(row=3, column=0,
                                                                                             sticky='w', padx=5,
                                                                                             pady=(15, 0))
        winner_pattern_entry = ttk.Entry(settings_inner_frame, textvariable=self.winner_pattern_var, width=50)
        winner_pattern_entry.grid(row=4, column=0, sticky='ew', padx=5, pady=5, columnspan=2)
        ToolTip(winner_pattern_entry,
                "Expressão regular para capturar o índice do mapa vencedor (o primeiro grupo de captura é usado).")

        ttk.Label(settings_inner_frame, text="Missão padrão de votemap:").grid(row=5, column=0, sticky='w', padx=5,
                                                                               pady=(15, 0))
        default_mission_entry = ttk.Entry(settings_inner_frame, textvariable=self.default_mission_var, width=70)
        default_mission_entry.grid(row=6, column=0, sticky='ew', padx=5, pady=5, columnspan=2)
        ToolTip(default_mission_entry,
                "ID do cenário/missão a ser carregado após um reinício para iniciar uma nova votação.")

        ttk.Label(settings_inner_frame, text="Delay ao parar serviço (s):").grid(row=7, column=0, sticky='w', padx=5,
                                                                                 pady=(15, 0))
        stop_delay_spinbox = ttk.Spinbox(settings_inner_frame, from_=1, to=60, textvariable=self.stop_delay_var,
                                         width=7)
        stop_delay_spinbox.grid(row=8, column=0, sticky='w', padx=5, pady=5)
        ToolTip(stop_delay_spinbox, "Tempo em segundos para aguardar após enviar o comando de parada ao serviço.")

        ttk.Label(settings_inner_frame, text="Delay ao iniciar serviço (s):").grid(row=7, column=1, sticky='w', padx=5,
                                                                                   pady=(15, 0))
        start_delay_spinbox = ttk.Spinbox(settings_inner_frame, from_=5, to=120, textvariable=self.start_delay_var,
                                          width=7)
        start_delay_spinbox.grid(row=8, column=1, sticky='w', padx=5, pady=5)
        ToolTip(start_delay_spinbox,
                "Tempo em segundos para aguardar o servidor iniciar completamente após o comando de início.")

        self.save_config_button_app_settings = ttk.Button(settings_inner_frame, text="Salvar Configurações da App",
                                                          command=self.save_config, bootstyle=SUCCESS, state="disabled")
        self.save_config_button_app_settings.grid(row=9, column=0, columnspan=2, pady=20)

    def forcar_refresh_json(self):
        refreshed_server = False
        refreshed_votemap = False
        default_fg = self.style.colors.fg if hasattr(self.style, 'colors') and self.style.colors else "black"

        try:
            if self.arquivo_json:
                if os.path.exists(self.arquivo_json):
                    try:
                        with open(self.arquivo_json, 'r', encoding='utf-8') as f:
                            conteudo = f.read()
                        json.loads(conteudo)
                        self.exibir_json(self.json_text_area, conteudo)
                        self.append_text_gui(f"JSON do servidor '{os.path.basename(self.arquivo_json)}' recarregado.\n")
                        self.server_json_path_label_var.set(f"JSON Servidor: {os.path.basename(self.arquivo_json)}")
                        self.json_server_path_label.config(foreground="green")
                        refreshed_server = True
                    except json.JSONDecodeError:
                        self.exibir_json(self.json_text_area,
                                         f"ERRO: Conteúdo de '{os.path.basename(self.arquivo_json)}' não é JSON válido.")
                        self.append_text_gui(
                            f"ERRO: JSON do servidor '{os.path.basename(self.arquivo_json)}' é inválido.\n")
                        self.server_json_path_label_var.set(
                            f"JSON Servidor (INVÁLIDO): {os.path.basename(self.arquivo_json)}")
                        self.json_server_path_label.config(foreground="red")
                    except Exception as e_read:
                        self.exibir_json(self.json_text_area,
                                         f"ERRO ao ler '{os.path.basename(self.arquivo_json)}': {e_read}")
                        self.append_text_gui(
                            f"ERRO ao ler JSON do servidor '{os.path.basename(self.arquivo_json)}': {e_read}\n")
                        self.server_json_path_label_var.set(
                            f"JSON Servidor (ERRO LEITURA): {os.path.basename(self.arquivo_json)}")
                        self.json_server_path_label.config(foreground="red")
                else:
                    self.append_text_gui(f"Arquivo JSON do servidor '{self.arquivo_json}' não encontrado.\n")
                    self.exibir_json(self.json_text_area, "Arquivo não encontrado.")
                    self.server_json_path_label_var.set(
                        f"JSON Servidor (NÃO ENCONTRADO): {os.path.basename(self.arquivo_json)}")
                    self.json_server_path_label.config(foreground="orange")
            else:
                self.append_text_gui("Caminho do JSON do servidor não configurado.\n")
                self.exibir_json(self.json_text_area, "Não configurado.")
                self.server_json_path_label_var.set("JSON Servidor: Nenhum")
                self.json_server_path_label.config(foreground=default_fg)

            if self.arquivo_json_votemap:
                if os.path.exists(self.arquivo_json_votemap):
                    try:
                        with open(self.arquivo_json_votemap, 'r', encoding='utf-8') as f:
                            conteudo = f.read()
                        json.loads(conteudo)
                        self.exibir_json(self.json_vm_text_area, conteudo)
                        self.append_text_gui(
                            f"JSON do votemap '{os.path.basename(self.arquivo_json_votemap)}' recarregado.\n")
                        self.votemap_json_path_label_var.set(
                            f"JSON Votemap: {os.path.basename(self.arquivo_json_votemap)}")
                        self.json_votemap_path_label.config(foreground="green")
                        refreshed_votemap = True
                    except json.JSONDecodeError:
                        self.exibir_json(self.json_vm_text_area,
                                         f"ERRO: Conteúdo de '{os.path.basename(self.arquivo_json_votemap)}' não é JSON válido.")
                        self.append_text_gui(
                            f"ERRO: JSON do votemap '{os.path.basename(self.arquivo_json_votemap)}' é inválido.\n")
                        self.votemap_json_path_label_var.set(
                            f"JSON Votemap (INVÁLIDO): {os.path.basename(self.arquivo_json_votemap)}")
                        self.json_votemap_path_label.config(foreground="red")
                    except Exception as e_read_vm:
                        self.exibir_json(self.json_vm_text_area,
                                         f"ERRO ao ler '{os.path.basename(self.arquivo_json_votemap)}': {e_read_vm}")
                        self.append_text_gui(
                            f"ERRO ao ler JSON do votemap '{os.path.basename(self.arquivo_json_votemap)}': {e_read_vm}\n")
                        self.votemap_json_path_label_var.set(
                            f"JSON Votemap (ERRO LEITURA): {os.path.basename(self.arquivo_json_votemap)}")
                        self.json_votemap_path_label.config(foreground="red")
                else:
                    self.append_text_gui(f"Arquivo JSON do votemap '{self.arquivo_json_votemap}' não encontrado.\n")
                    self.exibir_json(self.json_vm_text_area, "Arquivo não encontrado.")
                    self.votemap_json_path_label_var.set(
                        f"JSON Votemap (NÃO ENCONTRADO): {os.path.basename(self.arquivo_json_votemap)}")
                    self.json_votemap_path_label.config(foreground="orange")
            else:
                self.append_text_gui("Caminho do JSON do votemap não configurado.\n")
                self.exibir_json(self.json_vm_text_area, "Não configurado.")
                self.votemap_json_path_label_var.set("JSON Votemap: Nenhum")
                self.json_votemap_path_label.config(foreground=default_fg)

            if refreshed_server or refreshed_votemap:
                self.set_status_from_thread("Arquivos JSON recarregados.")
            else:
                self.set_status_from_thread(
                    "Nenhum arquivo JSON para recarregar ou arquivos não encontrados/inválidos.")
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
        default_fg = self.style.colors.fg if hasattr(self.style, 'colors') and self.style.colors else "black"

        if self.pasta_raiz and os.path.isdir(self.pasta_raiz):
            self.append_text_gui(f">>> Pasta de logs configurada: {self.pasta_raiz}\n")
            self.log_folder_path_label_var.set(f"Pasta Logs: {os.path.basename(self.pasta_raiz)}")
            self.log_folder_path_label.config(foreground="green")
        elif self.pasta_raiz:
            self.log_folder_path_label_var.set(f"Pasta Logs (Inválida): {os.path.basename(self.pasta_raiz)}")
            self.log_folder_path_label.config(foreground="red")
        else:
            self.log_folder_path_label_var.set("Pasta Logs: Nenhuma")
            self.log_folder_path_label.config(foreground=default_fg)

        self.forcar_refresh_json()

        if self.nome_servico and PYWIN32_AVAILABLE:  # Só atualiza se pywin32 estiver disponível
            self.update_service_status_display()
        elif not PYWIN32_AVAILABLE:
            self.servico_var.set("Serviço: N/A (pywin32 ausente)")
            self.servico_label_widget.config(foreground="gray")
        else:
            self.servico_var.set("Nenhum serviço selecionado")
            self.servico_label_widget.config(
                foreground=default_fg if default_fg != "black" else "orange")  # Evitar laranja em tema claro

        self.set_status_from_thread("Configuração carregada. Pronto.")

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
            return {}
        except Exception as e:
            logging.error(f"Erro desconhecido ao carregar configuração de {self.config_file}: {e}", exc_info=True)
            return {}

    def load_config_dialog(self):
        caminho = filedialog.askopenfilename(defaultextension=".json",
                                             filetypes=[("Arquivos JSON", "*.json"), ("Todos os arquivos", "*.*")],
                                             title="Selecionar arquivo de configuração para carregar")
        if caminho:
            try:
                with open(caminho, 'r', encoding='utf-8') as f:
                    loaded_config = json.load(f)

                self.loading_config = True

                self.config = loaded_config
                self.pasta_raiz = self.config.get("log_folder", None)
                self.arquivo_json = self.config.get("server_json", None)
                self.arquivo_json_votemap = self.config.get("votemap_json", None)
                self.nome_servico = self.config.get("service_name", None)

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

                self.loading_config = False
                self.config_changed = False
                self._update_save_buttons_state()

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
                self.loading_config = False
            except Exception as e:
                logging.error(f"Erro ao carregar configuração de {caminho}: {e}", exc_info=True)
                self.show_messagebox_from_thread("error", "Erro de Configuração",
                                                 f"Falha ao carregar configuração de '{caminho}':\n{e}")
                self.loading_config = False

    def save_config(self):
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
                json.dump(self.config, f, indent=4)
            self.config_changed = False
            self._update_save_buttons_state()
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
        default_fg = self.style.colors.fg if hasattr(self.style, 'colors') and self.style.colors else "black"
        if pasta_selecionada:
            if self.pasta_raiz != pasta_selecionada:
                logging.info(f"Pasta de logs alterada de '{self.pasta_raiz}' para '{pasta_selecionada}'")
                self.stop_log_monitoring()
                self.pasta_raiz = pasta_selecionada
                self.append_text_gui(f">>> Nova pasta de logs selecionada: {self.pasta_raiz}\n")
                self.set_status_from_thread(f"Pasta de logs: {os.path.basename(self.pasta_raiz)}")
                self.log_folder_path_label_var.set(f"Pasta Logs: {os.path.basename(self.pasta_raiz)}")
                self.log_folder_path_label.config(foreground="green")
                self._config_value_changed()
                self.start_log_monitoring()
            else:
                self.log_folder_path_label_var.set(f"Pasta Logs: {os.path.basename(self.pasta_raiz)}")
                self.log_folder_path_label.config(foreground="green")
                self.append_text_gui(f">>> Pasta de logs já selecionada: {self.pasta_raiz}\n")
        else:
            if not self.pasta_raiz:
                self.log_folder_path_label_var.set("Pasta Logs: Nenhuma")
                self.log_folder_path_label.config(foreground=default_fg)

    def _selecionar_arquivo_json(self, tipo_json):
        title_map = {"servidor": "Selecionar JSON de Configuração do Servidor",
                     "votemap": "Selecionar JSON de Configuração do Votemap"}
        caminho = filedialog.askopenfilename(filetypes=[("Arquivos JSON", "*.json"), ("Todos os arquivos", "*.*")],
                                             title=title_map.get(tipo_json, "Selecionar Arquivo JSON"))

        target_label_var = self.server_json_path_label_var if tipo_json == "servidor" else self.votemap_json_path_label_var
        target_widget_label = self.json_server_path_label if tipo_json == "servidor" else self.json_votemap_path_label
        current_path_attr = "arquivo_json" if tipo_json == "servidor" else "arquivo_json_votemap"
        default_fg = self.style.colors.fg if hasattr(self.style, 'colors') and self.style.colors else "black"

        if caminho:
            try:
                with open(caminho, 'r', encoding='utf-8') as f:
                    conteudo = f.read()
                json.loads(conteudo)

                path_changed = getattr(self, current_path_attr) != caminho
                setattr(self, current_path_attr, caminho)

                if tipo_json == "servidor":
                    self.exibir_json(self.json_text_area, conteudo)
                elif tipo_json == "votemap":
                    self.exibir_json(self.json_vm_text_area, conteudo)

                msg = f"JSON de {tipo_json} carregado: {os.path.basename(caminho)}"
                self.set_status_from_thread(msg)
                self.append_text_gui(f">>> {msg}\n")
                target_label_var.set(f"JSON {tipo_json.capitalize()}: {os.path.basename(caminho)}")
                target_widget_label.config(foreground="green")
                logging.info(f"Arquivo JSON de {tipo_json} selecionado: {caminho}")
                if path_changed:
                    self._config_value_changed()

            except FileNotFoundError:
                err_msg = f"Erro: Arquivo JSON de {tipo_json} não encontrado em '{caminho}'."
                self.set_status_from_thread(err_msg);
                logging.error(err_msg)
                self.show_messagebox_from_thread("error", "Arquivo não encontrado", err_msg)
                target_label_var.set(f"JSON {tipo_json.capitalize()} (NÃO ENCONTRADO): {os.path.basename(caminho)}")
                target_widget_label.config(foreground="orange")
            except json.JSONDecodeError:
                err_msg = f"Erro: Arquivo JSON de {tipo_json} ('{os.path.basename(caminho)}') não é um JSON válido."
                self.set_status_from_thread(err_msg);
                logging.error(err_msg)
                self.show_messagebox_from_thread("error", "JSON Inválido", err_msg)
                target_label_var.set(f"JSON {tipo_json.capitalize()} (INVÁLIDO): {os.path.basename(caminho)}")
                target_widget_label.config(foreground="red")
            except Exception as e:
                err_msg = f"Erro ao carregar JSON de {tipo_json} '{os.path.basename(caminho)}': {e}"
                self.set_status_from_thread(err_msg);
                logging.error(err_msg, exc_info=True)
                self.show_messagebox_from_thread("error", "Erro de Leitura", err_msg)
                target_label_var.set(f"JSON {tipo_json.capitalize()} (ERRO): {os.path.basename(caminho)}")
                target_widget_label.config(foreground="red")
        else:
            if not getattr(self, current_path_attr):
                target_label_var.set(f"JSON {tipo_json.capitalize()}: Nenhum")
                target_widget_label.config(foreground=default_fg)

    def selecionar_arquivo_json_servidor(self):
        self._selecionar_arquivo_json("servidor")

    def selecionar_arquivo_json_votemap(self):
        self._selecionar_arquivo_json("votemap")

    def _show_progress_dialog(self, title, message):
        progress_win = ttk.Toplevel(self.root)
        progress_win.title(str(title) if title else "Progresso")  # Garantir que title é string
        progress_win.geometry("300x100")
        progress_win.resizable(False, False)
        progress_win.transient(self.root)
        progress_win.grab_set()
        ttk.Label(progress_win, text=str(message) if message else "Carregando...", bootstyle=PRIMARY).pack(pady=10)
        pb = ttk.Progressbar(progress_win, mode='indeterminate', length=280)
        pb.pack(pady=10)
        pb.start(10)

        progress_win.update_idletasks()
        try:
            width = progress_win.winfo_width()
            height = progress_win.winfo_height()
            x = (self.root.winfo_screenwidth() // 2) - (width // 2)
            y = (self.root.winfo_screenheight() // 2) - (height // 2)
            progress_win.geometry(f'{width}x{height}+{x}+{y}')
        except tk.TclError:
            logging.warning("TclError ao tentar centralizar _show_progress_dialog, janela pode ter sido destruída.")

        return progress_win, pb

    def selecionar_servico(self):
        if not PYWIN32_AVAILABLE:
            self.show_messagebox_from_thread("error", "Funcionalidade Indisponível",
                                             "A biblioteca pywin32 é necessária para listar e gerenciar serviços do Windows.")
            return

        progress_win, _ = self._show_progress_dialog("Serviços", "Carregando lista de serviços...")
        # Garantir que progress_win é válido antes de usá-lo
        if not (progress_win and progress_win.winfo_exists()):
            logging.error("Falha ao criar a janela de progresso para selecionar serviço.")
            return  # Não continuar se a dialog de progresso falhou

        if self.root.winfo_exists(): self.root.update_idletasks()
        threading.Thread(target=self._obter_servicos_worker, args=(progress_win,), daemon=True,
                         name="ServicoWMIThread").start()

    def _obter_servicos_worker(self, progress_win):
        if not PYWIN32_AVAILABLE:
            logging.warning("_obter_servicos_worker chamado mas PYWIN32_AVAILABLE é False.")
            if progress_win and progress_win.winfo_exists():
                self.root.after(0, lambda: progress_win.destroy() if progress_win.winfo_exists() else None)
            return

        initialized_com = False
        try:
            logging.debug("Tentando inicializar COM na thread ServicoWMIThread.")
            pythoncom.CoInitialize()
            initialized_com = True
            logging.debug("COM inicializado com sucesso.")

            logging.debug("Tentando obter objeto WMI 'winmgmts:'.")
            wmi = win32com.client.GetObject('winmgmts:')
            logging.info(f"Objeto WMI obtido.")  # Removido {wmi} para evitar log muito grande

            logging.debug("Tentando obter instâncias de 'Win32_Service'.")
            services_raw = wmi.InstancesOf('Win32_Service')
            logging.info(f"Instâncias de Win32_Service obtidas. Tipo: {type(services_raw)}")

            if services_raw is None:
                logging.warning("wmi.InstancesOf('Win32_Service') retornou None.")
                nomes_servicos_temp = []
            else:
                try:
                    num_services_raw = len(services_raw)
                    logging.info(f"Total de serviços brutos encontrados: {num_services_raw}")
                except TypeError:
                    logging.warning(
                        f"Não foi possível obter o tamanho de services_raw (tipo: {type(services_raw)}). Iterando...")

                nomes_servicos_temp = []
                for i, s in enumerate(services_raw):
                    if hasattr(s, 'Name') and s.Name and hasattr(s, 'AcceptStop') and s.AcceptStop:
                        nomes_servicos_temp.append(s.Name)
                logging.info(f"Serviços após filtro (Name, AcceptStop): {len(nomes_servicos_temp)}")

            nomes_servicos = sorted(nomes_servicos_temp)

            if self.root.winfo_exists():
                self.root.after(0, self._mostrar_dialogo_selecao_servico, nomes_servicos, progress_win)
            elif progress_win and progress_win.winfo_exists():
                progress_win.destroy()

        except pythoncom.com_error as e_com:
            logging.error(f"Erro COM ao listar serviços WMI: {e_com}", exc_info=True)
            error_message = f"Erro COM ({e_com.hresult}): {e_com.strerror}"
            if hasattr(e_com, 'excepinfo') and e_com.excepinfo:  # Adicionar hasattr
                error_message += f"\nDetalhes: {e_com.excepinfo[2]}"
            if self.root.winfo_exists():
                self.root.after(0, self._handle_erro_listar_servicos, error_message, progress_win)
            elif progress_win and progress_win.winfo_exists():
                progress_win.destroy()
        except Exception as e:
            logging.error(f"Erro geral ao listar serviços WMI: {e}", exc_info=True)
            error_message = str(e)
            if self.root.winfo_exists():
                self.root.after(0, self._handle_erro_listar_servicos, error_message, progress_win)
            elif progress_win and progress_win.winfo_exists():
                progress_win.destroy()
        finally:
            if initialized_com:
                try:
                    pythoncom.CoUninitialize()
                    logging.debug("COM desinicializado com sucesso.")
                except Exception as e_uninit:
                    logging.error(f"Erro ao desinicializar COM: {e_uninit}", exc_info=True)
            # Garantir que progress_win seja fechado se a root não existir mais
            if progress_win and progress_win.winfo_exists() and not self.root.winfo_exists():
                try:
                    progress_win.destroy()
                except tk.TclError:
                    pass  # Já foi destruída
                except Exception as e_destroy:  # Outros erros ao destruir
                    logging.error(f"Erro não Tcl ao destruir progress_win no finally: {e_destroy}")

    def _handle_erro_listar_servicos(self, error_message, progress_win):
        if progress_win and progress_win.winfo_exists():
            try:
                progress_win.destroy()
            except tk.TclError:
                logging.debug(
                    "TclError ao destruir progress_win em _handle_erro_listar_servicos (provavelmente já destruída).")
            except Exception as e_destroy:
                logging.error(f"Erro não Tcl ao destruir progress_win em _handle_erro_listar_servicos: {e_destroy}")

        if self.root.winfo_exists():
            Messagebox.show_error(f"Erro ao obter lista de serviços:\n{error_message}", "Erro WMI", parent=self.root)

    def _mostrar_dialogo_selecao_servico(self, nomes_servicos, progress_win):
        if progress_win and progress_win.winfo_exists():
            try:
                progress_win.destroy()
            except tk.TclError:
                logging.debug(
                    "TclError ao destruir progress_win em _mostrar_dialogo_selecao_servico (provavelmente já destruída).")
            except Exception as e_destroy:
                logging.error(f"Erro não Tcl ao destruir progress_win em _mostrar_dialogo_selecao_servico: {e_destroy}")

        if not nomes_servicos:
            if self.root.winfo_exists():
                Messagebox.show_warning("Nenhum serviço gerenciável encontrado (ou erro ao listar).",
                                        "Seleção de Serviço", parent=self.root)
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
        listbox.column("name", width=450)
        listbox.pack(side='left', fill='both', expand=True)
        listbox.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=listbox.yview)

        def _populate_listbox(query=""):
            for item_id in listbox.get_children():
                listbox.delete(item_id)

            if not nomes_servicos:
                return

            filter_query = query.lower() if query else ""
            for name in nomes_servicos:
                if name and (not filter_query or filter_query in name.lower()):
                    listbox.insert("", "end", values=(name,))

        search_entry.bind("<KeyRelease>", lambda e: _populate_listbox(search_var.get()))
        _populate_listbox()

        def on_confirm():
            selection = listbox.selection()
            if selection:
                selected_item_id = selection[0]
                selected_item_values = listbox.item(selected_item_id, "values")
                if selected_item_values:
                    service_name = selected_item_values[0]
                    if self.nome_servico != service_name:
                        self.nome_servico = service_name
                        self._config_value_changed()
                    self.update_service_status_display()
                    self.set_status_from_thread(f"Serviço selecionado: {service_name}")
                    logging.info(f"Serviço selecionado: {service_name}")
                    dialog.destroy()
                else:
                    if dialog.winfo_exists():
                        Messagebox.show_warning("Falha ao obter o nome do serviço selecionado.", parent=dialog)
            else:
                if dialog.winfo_exists():
                    Messagebox.show_warning("Nenhum serviço selecionado.", parent=dialog)

        btn_frame = ttk.Frame(dialog);
        btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text="Confirmar", command=on_confirm, bootstyle=SUCCESS).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="Cancelar", command=dialog.destroy, bootstyle=DANGER).pack(side='left', padx=5)

        dialog.update_idletasks()
        ws, hs = dialog.winfo_screenwidth(), dialog.winfo_screenheight()
        w, h = dialog.winfo_width(), dialog.winfo_height()
        x, y = (ws / 2) - (w / 2), (hs / 2) - (h / 2)
        dialog.geometry(f'+{int(x)}+{int(y)}')
        search_entry.focus_set()
        dialog.wait_window()

    def update_service_status_display(self):
        if not PYWIN32_AVAILABLE:
            self.servico_var.set("Serviço: N/A (pywin32 ausente)")
            if hasattr(self.style, 'colors') and self.style.colors:
                self.servico_label_widget.config(foreground=self.style.colors.get('disabled', 'gray'))
            else:  # Fallback se colors não estiver disponível
                self.servico_label_widget.config(foreground="gray")
            return

        if self.nome_servico:
            current_text_base = f"Serviço: {self.nome_servico}"
            self.servico_var.set(f"{current_text_base} (Verificando...)")
            self.servico_label_widget.config(foreground="blue")
            threading.Thread(target=self._get_and_display_service_status,
                             args=(self.nome_servico, current_text_base), daemon=True,
                             name="ServiceStatusCheckThread").start()
        else:
            self.servico_var.set("Nenhum serviço selecionado")
            default_fg = "orange"
            if hasattr(self.style, 'colors') and self.style.colors:
                default_fg = self.style.colors.get('fg', 'orange')  # Tenta usar fg do tema, senão laranja
            self.servico_label_widget.config(foreground=default_fg)

    def _get_and_display_service_status(self, service_name, base_text):
        status = self.verificar_status_servico(service_name)
        status_map_colors = {
            "RUNNING": ("(Rodando)", "green"),
            "STOPPED": ("(Parado)", "red"),
            "START_PENDING": ("(Iniciando...)", "blue"),
            "STOP_PENDING": ("(Parando...)", "blue"),
            "NOT_FOUND": ("(Não encontrado!)", "orange"),
            "ERROR": ("(Erro ao verificar!)", "red"),
            "UNKNOWN": ("(Desconhecido)", "gray")
        }
        display_status_text, color = status_map_colors.get(status, ("(Status ?)", "gray"))

        if self.root.winfo_exists():
            self.root.after(0, lambda: (
                self.servico_var.set(f"{base_text} {display_status_text}"),
                self.servico_label_widget.config(foreground=color) if self.servico_label_widget.winfo_exists() else None
            ))

    def exibir_json(self, text_area_widget, conteudo_json):
        try:
            dados_formatados = json.dumps(json.loads(conteudo_json), indent=4, ensure_ascii=False)
        except (json.JSONDecodeError, TypeError):
            dados_formatados = str(conteudo_json)

        if self.root.winfo_exists() and text_area_widget.winfo_exists():
            try:
                text_area_widget.configure(state='normal')
                text_area_widget.delete('1.0', 'end')
                text_area_widget.insert('end', dados_formatados)
                text_area_widget.configure(state='disabled')
            except tk.TclError:
                logging.warning(f"TclError ao exibir JSON no widget {text_area_widget} (provavelmente GUI fechando).")

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
                if self._stop_event.wait(10): break
                continue

            try:
                nova_pasta_logs = self.obter_subpasta_mais_recente()
                if not nova_pasta_logs:
                    if self._stop_event.wait(5): break
                    continue

                novo_arquivo_log_path = os.path.join(nova_pasta_logs, 'console.log')

                if os.path.exists(novo_arquivo_log_path) and novo_arquivo_log_path != self.caminho_log_atual:
                    logging.info(
                        f"[{thread_name}] Novo arquivo de log detectado: {novo_arquivo_log_path} (anterior: {self.caminho_log_atual})")
                    self.append_text_gui(f"\n>>> Novo arquivo de log detectado: {novo_arquivo_log_path}\n")

                    if self.log_tail_thread and self.log_tail_thread.is_alive():
                        logging.debug(
                            f"[{thread_name}] Parando LogTailThread ({self.log_tail_thread.name}) para o arquivo antigo '{self.caminho_log_atual}'...")
                        self.log_tail_thread.join(timeout=1.5)
                        if self.log_tail_thread.is_alive():
                            logging.warning(
                                f"[{thread_name}] LogTailThread ({self.log_tail_thread.name}) antiga não parou a tempo.")

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
                            self.file_log_handle = None

                    self.caminho_log_atual = novo_arquivo_log_path
                    self.pasta_atual = nova_pasta_logs
                    novo_fh_temp = None
                    try:
                        logging.debug(f"[{thread_name}] Tentando abrir novo arquivo de log: {self.caminho_log_atual}")
                        novo_fh_temp = open(self.caminho_log_atual, 'r', encoding='utf-8', errors='replace')
                        novo_fh_temp.seek(0, os.SEEK_END)
                        self.file_log_handle = novo_fh_temp

                        if self.root.winfo_exists():
                            self.root.after(0, lambda p=self.caminho_log_atual: self.log_label.config(
                                text=f"LOG AO VIVO: {p}") if self.log_label.winfo_exists() else None)
                            self.root.after(0, lambda p=self.caminho_log_atual: self.status_label_var.set(
                                f"Monitorando: {os.path.basename(p)}"))

                        logging.info(
                            f"[{thread_name}] Novo arquivo de log {self.caminho_log_atual} aberto com sucesso. Iniciando nova LogTailThread.")
                        self.log_tail_thread = threading.Thread(
                            target=self.acompanhar_log_do_arquivo,
                            args=(self.caminho_log_atual,),
                            daemon=True,
                            name=f"LogTailThread-{os.path.basename(self.caminho_log_atual)}"
                        )
                        self.log_tail_thread.start()

                    except FileNotFoundError:
                        logging.error(
                            f"[{thread_name}] Arquivo de log {self.caminho_log_atual} não encontrado ao tentar abrir.")
                        if novo_fh_temp: novo_fh_temp.close()
                        self.file_log_handle = None
                        self.caminho_log_atual = None
                    except Exception as e_open:
                        logging.error(
                            f"[{thread_name}] Erro ao abrir ou iniciar acompanhamento de {self.caminho_log_atual}: {e_open}",
                            exc_info=True)
                        if novo_fh_temp: novo_fh_temp.close()
                        self.file_log_handle = None
                        self.caminho_log_atual = None

                elif self.caminho_log_atual and not os.path.exists(self.caminho_log_atual):
                    logging.warning(
                        f"[{thread_name}] Arquivo de log monitorado {self.caminho_log_atual} não existe mais.")
                    self.append_text_gui(f"Aviso: Arquivo de log {self.caminho_log_atual} não encontrado.\n")
                    if self.log_tail_thread and self.log_tail_thread.is_alive():
                        self.log_tail_thread.join(timeout=1.0)
                    if self.file_log_handle:
                        try:
                            self.file_log_handle.close()
                        except:
                            pass
                    self.file_log_handle = None
                    self.caminho_log_atual = None

            except Exception as e_monitor_loop:
                logging.error(f"[{thread_name}] Erro no loop principal de monitoramento de logs: {e_monitor_loop}",
                              exc_info=True)
                self.append_text_gui(f"Erro crítico ao monitorar logs: {e_monitor_loop}\n")

            if self._stop_event.wait(5): break
        logging.info(f"[{thread_name}] Thread de monitoramento de log contínuo ({thread_name}) encerrada.")

    def obter_subpasta_mais_recente(self):
        if not self.pasta_raiz or not os.path.isdir(self.pasta_raiz): return None
        try:
            subpastas = [os.path.join(self.pasta_raiz, nome) for nome in os.listdir(self.pasta_raiz) if
                         os.path.isdir(os.path.join(self.pasta_raiz, nome))]
            if not subpastas: return None
            return max(subpastas, key=os.path.getmtime)
        except FileNotFoundError:
            logging.warning(f"Pasta raiz '{self.pasta_raiz}' não encontrada ao buscar subpastas. Resetando pasta_raiz.")
            self.pasta_raiz = None
            return None
        except PermissionError:
            logging.error(
                f"Permissão negada ao acessar '{self.pasta_raiz}' para buscar subpastas. Resetando pasta_raiz.")
            self.pasta_raiz = None
            return None
        except Exception as e:
            logging.error(f"Erro ao obter subpasta mais recente em '{self.pasta_raiz}': {e}", exc_info=True)
            return None

    def acompanhar_log_do_arquivo(self, caminho_log_designado_para_esta_thread):
        thread_name = threading.current_thread().name
        logging.info(f"[{thread_name}] Tentando iniciar acompanhamento para: {caminho_log_designado_para_esta_thread}")

        if self._stop_event.is_set():
            logging.info(
                f"[{thread_name}] _stop_event já está setado no início. Encerrando para {caminho_log_designado_para_esta_thread}.")
            return

        if not self.file_log_handle or self.file_log_handle.closed:
            logging.error(
                f"[{thread_name}] ERRO CRÍTICO: file_log_handle está NULO ou FECHADO no início do acompanhamento para '{caminho_log_designado_para_esta_thread}'. Esta thread não pode prosseguir.")
            return

        try:
            handle_real_path_norm = os.path.normpath(self.file_log_handle.name)
            caminho_designado_norm = os.path.normpath(caminho_log_designado_para_esta_thread)
            if handle_real_path_norm != caminho_designado_norm:
                logging.warning(
                    f"[{thread_name}] DESCOMPASSO DE HANDLE NO INÍCIO! Thread designada para '{caminho_designado_norm}' mas self.file_log_handle atualmente aponta para '{handle_real_path_norm}'. Encerrando esta thread.")
                return
        except AttributeError:
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
            return

        logging.info(
            f"[{thread_name}] Padrões para '{caminho_log_designado_para_esta_thread}': FimVoto='{vote_pattern_str}', Vencedor='{winner_pattern_str}'")
        logging.debug(
            f"[{thread_name}] Estado inicial de aguardando_winner para '{caminho_log_designado_para_esta_thread}': {aguardando_winner}")

        while not self._stop_event.is_set():
            if self._paused:
                if self._stop_event.wait(0.5): break
                continue

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
            except AttributeError:
                logging.warning(
                    f"[{thread_name}] self.file_log_handle tornou-se None ou sem 'name' DENTRO DO LOOP para '{caminho_log_designado_para_esta_thread}'. Encerrando.")
                break
            except Exception as e_check_loop_attr_consistency:
                logging.error(
                    f"[{thread_name}] Erro ao verificar consistência do handle no loop para '{caminho_log_designado_para_esta_thread}': {e_check_loop_attr_consistency}. Encerrando.")
                break

            try:
                linha = self.file_log_handle.readline()
                if linha:
                    linha_strip = linha.strip()
                    filtro_atual = self.filtro_var.get().strip().lower()
                    if not filtro_atual or filtro_atual in linha.lower():
                        self.append_text_gui(linha)

                    logging.debug(
                        f"[{thread_name}] LIDO de '{caminho_log_designado_para_esta_thread}': repr='{repr(linha)}', strip='{linha_strip}', aguardando_winner={aguardando_winner}")

                    if vote_pattern_re and vote_pattern_re.search(linha):
                        if not aguardando_winner:
                            logging.info(
                                f"[{thread_name}] Padrão de FIM DE VOTAÇÃO detectado em '{caminho_log_designado_para_esta_thread}'. Linha: '{linha_strip}'. Definindo aguardando_winner = True.")
                        else:
                            logging.warning(
                                f"[{thread_name}] Padrão de FIM DE VOTAÇÃO detectado NOVAMENTE em '{caminho_log_designado_para_esta_thread}' enquanto aguardando_winner já era True. Linha: '{linha_strip}'.")
                        aguardando_winner = True
                        self.set_status_from_thread("Fim da votação detectado. Aguardando vencedor...")

                    if winner_pattern_re:
                        if aguardando_winner:
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
                                    if self.root.winfo_exists():
                                        self.root.after(0, self.processar_troca_mapa, indice)
                                    logging.debug(
                                        f"[{thread_name}] Winner processado para '{caminho_log_designado_para_esta_thread}', RESETANDO aguardando_winner para False.")
                                    aguardando_winner = False
                                except IndexError:
                                    logging.error(
                                        f"[{thread_name}] Padrão de vencedor '{winner_pattern_str}' casou em '{linha_strip}' para '{caminho_log_designado_para_esta_thread}', mas falta grupo de captura (group 1).")
                                    self.append_text_gui(
                                        f"ERRO: Padrão de vencedor '{winner_pattern_str}' não tem grupo de captura.\n")
                                    aguardando_winner = False
                                except ValueError:
                                    logging.error(
                                        f"[{thread_name}] Padrão de vencedor '{winner_pattern_str}' capturou '{indice_str}' em '{linha_strip}' para '{caminho_log_designado_para_esta_thread}', que não é um número de índice válido.")
                                    self.append_text_gui(f"ERRO: Vencedor capturado '{indice_str}' não é um número.\n")
                                    aguardando_winner = False
                                except Exception as e_proc_winner:
                                    logging.error(
                                        f"[{thread_name}] Erro inesperado ao processar vencedor para '{caminho_log_designado_para_esta_thread}': {e_proc_winner}",
                                        exc_info=True)
                                    aguardando_winner = False
                        elif winner_pattern_re.search(linha):
                            logging.info(
                                f"[{thread_name}] Padrão de vencedor APARECEU na linha '{linha_strip}' em '{caminho_log_designado_para_esta_thread}', MAS aguardando_winner era FALSO. Nenhum processamento de vencedor para esta linha.")
                else:
                    if self._stop_event.wait(0.2): break

            except UnicodeDecodeError as ude:
                logging.warning(
                    f"[{thread_name}] Erro de decodificação Unicode ao ler log {caminho_log_designado_para_esta_thread}: {ude}. Linha ignorada.")
            except ValueError as ve:
                if "i/o operation on closed file" in str(ve).lower():
                    logging.warning(
                        f"[{thread_name}] Tentativa de I/O em arquivo fechado ({caminho_log_designado_para_esta_thread}). Encerrando thread.")
                    break
                else:
                    logging.error(
                        f"[{thread_name}] Erro de ValueError ao acompanhar log {caminho_log_designado_para_esta_thread}: {ve}",
                        exc_info=True)
                    break
            except Exception as e_tail_loop:
                if not self._stop_event.is_set():
                    logging.error(
                        f"[{thread_name}] Erro Inesperado ao acompanhar log {caminho_log_designado_para_esta_thread}: {e_tail_loop}",
                        exc_info=True)
                    self.append_text_gui(f"Erro ao ler log: {e_tail_loop}\n")
                    self.set_status_from_thread("Erro na leitura do log. Verifique o Log do Sistema.")
                break
        logging.info(
            f"[{thread_name}] Acompanhamento de '{caminho_log_designado_para_esta_thread}' encerrado. Estado final de aguardando_winner: {aguardando_winner}")

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
        if indice_vencedor == 0:
            if len(map_list) > 1:
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
        elif 0 < indice_vencedor < len(map_list):
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
        if not PYWIN32_AVAILABLE: return "ERROR"
        if not nome_servico: return "NOT_FOUND"
        try:
            startupinfo = None
            if platform.system() == "Windows":
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                startupinfo.wShowWindow = subprocess.SW_HIDE

            result = subprocess.run(['sc', 'query', nome_servico], capture_output=True, text=True, check=False,
                                    startupinfo=startupinfo,
                                    encoding='latin-1', errors='replace')

            service_not_found_errors = ["failed 1060", "falha 1060",
                                        "지정된 서비스를 설치된 서비스로 찾을 수 없습니다."]
            output_lower = result.stdout.lower() + result.stderr.lower()

            for err_string in service_not_found_errors:
                if err_string in output_lower:
                    logging.warning(f"Serviço '{nome_servico}' não encontrado via 'sc query'. Output: {output_lower}")
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
            self.show_messagebox_from_thread("error", "Erro de Comando",
                                             "Comando 'sc.exe' não encontrado. Verifique as configurações do sistema.")
            return "ERROR"
        except Exception as e:
            logging.error(f"Erro ao verificar status do serviço '{nome_servico}': {e}", exc_info=True)
            return "ERROR"

    def reiniciar_servidor_com_progresso(self, novo_scenario_id_para_log):
        if not PYWIN32_AVAILABLE:
            self.show_messagebox_from_thread("error", "Funcionalidade Indisponível",
                                             "pywin32 é necessário para reiniciar serviços.")
            return

        progress_win, pb = None, None
        if self.root.winfo_exists():
            logging.info(f"Iniciando processo de reinício do servidor {self.nome_servico} em background.")
            self.set_status_from_thread(f"Reiniciando {self.nome_servico}...")

        success = self._reiniciar_servidor_logica(novo_scenario_id_para_log)

        if success:
            self.show_messagebox_from_thread("info", "Servidor Reiniciado",
                                             f"O serviço {self.nome_servico} foi reiniciado com sucesso.")
        else:
            self.show_messagebox_from_thread("error", "Falha no Reinício",
                                             f"Ocorreu um erro ao reiniciar o serviço {self.nome_servico}.")

    def _reiniciar_servidor_logica(self, novo_scenario_id_para_log):
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
                time.sleep(stop_delay)
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
            else:
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
            time.sleep(start_delay)

            status_after_start = self.verificar_status_servico(self.nome_servico)
            if status_after_start != "RUNNING":
                logging.error(f"Serviço {self.nome_servico} falhou ao iniciar. Status: {status_after_start}")
                self.append_text_gui_threadsafe(
                    f"ERRO: Serviço '{self.nome_servico}' falhou ao iniciar ou está demorando muito. Status: {status_after_start}\n")
                self.set_status_from_thread(f"Erro: {self.nome_servico} não iniciou. Status: {status_after_start}")
                return False
            logging.info(f"Serviço {self.nome_servico} iniciado com sucesso.")
            self.update_service_status_display()

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
            self.update_service_status_display()
            return False
        except FileNotFoundError:
            self.show_messagebox_from_thread("error", "Erro de Comando",
                                             "Comando 'sc.exe' não encontrado. Verifique o PATH.")
            logging.error("Comando 'sc.exe' não encontrado.")
            self.set_status_from_thread("Erro: sc.exe não encontrado.")
            self.update_service_status_display()
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
            self.update_service_status_display()
            return False

    def set_status_from_thread(self, message):
        if self.root.winfo_exists():
            self.root.after(0, lambda: self.status_label_var.set(message))

    def append_text_gui_threadsafe(self, texto):
        if self.root.winfo_exists() and hasattr(self, 'text_area') and self.text_area.winfo_exists():
            self.root.after(0, self._append_text_gui_actual, texto)

    def _append_text_gui_actual(self, texto):
        if self.text_area.winfo_exists():
            try:
                current_state = self.text_area.cget("state")
                self.text_area.configure(state='normal')
                self.text_area.insert('end', texto)
                if self.auto_scroll_log_var.get():
                    self.text_area.yview_moveto(1.0)
                self.text_area.configure(state=current_state)
            except tk.TclError:
                logging.warning(f"TclError em _append_text_gui_actual (provavelmente GUI fechando).")

    def show_messagebox_from_thread(self, boxtype, title, message):
        if self.root.winfo_exists():
            safe_title = str(title) if title is not None else "Notificação"
            safe_message = str(message) if message is not None else ""

            if boxtype == "info":
                self.root.after(0, lambda t=safe_title, m=safe_message: Messagebox.show_info(m, t,
                                                                                             parent=self.root) if self.root.winfo_exists() else None)
            elif boxtype == "error":
                self.root.after(0, lambda t=safe_title, m=safe_message: Messagebox.show_error(m, t,
                                                                                              parent=self.root) if self.root.winfo_exists() else None)
            elif boxtype == "warning":
                self.root.after(0, lambda t=safe_title, m=safe_message: Messagebox.show_warning(m, t,
                                                                                                parent=self.root) if self.root.winfo_exists() else None)

    def append_text_gui(self, texto):
        if self.root.winfo_exists() and hasattr(self, 'text_area') and self.text_area.winfo_exists():
            try:
                current_state = self.text_area.cget("state")
                self.text_area.configure(state='normal')
                self.text_area.insert('end', texto)
                if self.auto_scroll_log_var.get():
                    self.text_area.yview_moveto(1.0)
                self.text_area.configure(state=current_state)
            except tk.TclError as e:
                logging.warning(f"TclError em append_text_gui (provavelmente GUI fechando): {e}")

    def limpar_tela(self):
        if hasattr(self, 'text_area') and self.text_area.winfo_exists():
            self.text_area.configure(state='normal');
            self.text_area.delete('1.0', 'end');
            self.text_area.configure(state='disabled')
            self.status_label_var.set("Tela de logs do servidor limpa.");
            logging.info("Tela de logs do servidor limpa pelo usuário.")

    def toggle_pausa(self):
        self._paused = not self._paused
        if self._paused:
            self.pausar_btn.config(text="▶️ Retomar", bootstyle=SUCCESS)
            self.status_label_var.set("Monitoramento de logs pausado.")
            logging.info("Monitoramento de logs pausado.")
        else:
            self.pausar_btn.config(text="⏸️ Pausar", bootstyle=WARNING)
            self.status_label_var.set("Monitoramento de logs retomado.")
            logging.info("Monitoramento de logs retomado.")

    def trocar_tema(self, event=None):
        novo_tema = self.tema_var.get()
        try:
            self.style.theme_use(novo_tema)
            self.initialize_from_config()  # Re-inicializa para aplicar cores corretas ao tema
            logging.info(f"Tema alterado para: {novo_tema}")
            self.status_label_var.set(f"Tema alterado para '{novo_tema}'.")
        except Exception as e:
            logging.error(f"Erro ao tentar trocar para o tema '{novo_tema}': {e}", exc_info=True)
            self.show_messagebox_from_thread("error", "Erro de Tema",
                                             f"Não foi possível aplicar o tema '{novo_tema}'.\n{e}")
            try:
                self.style.theme_use("litera");
                self.tema_var.set("litera")
                self.initialize_from_config()  # Tentar re-inicializar com tema padrão
            except:
                pass

    def export_display_logs(self):
        caminho_arquivo = filedialog.asksaveasfilename(defaultextension=".txt",
                                                       filetypes=[("Arquivos de Texto", "*.txt"),
                                                                  ("Todos os arquivos", "*.*")],
                                                       title="Exportar Logs Exibidos na Tela")
        if caminho_arquivo:
            try:
                if hasattr(self, 'text_area') and self.text_area.winfo_exists():
                    with open(caminho_arquivo, 'w', encoding='utf-8') as f:
                        f.write(self.text_area.get('1.0', 'end-1c'))
                    self.status_label_var.set(f"Logs da tela exportados para: {os.path.basename(caminho_arquivo)}")
                    logging.info(f"Logs da tela exportados para: {caminho_arquivo}")
                    self.show_messagebox_from_thread("info", "Exportação Concluída",
                                                     f"Logs da tela foram exportados com sucesso para:\n{caminho_arquivo}")
                else:
                    self.show_messagebox_from_thread("error", "Erro de Exportação", "Área de log não encontrada.")
            except Exception as e:
                self.status_label_var.set(f"Erro ao exportar logs da tela: {e}")
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
        if self.auto_restart_var.get() and not PYWIN32_AVAILABLE: problemas.append(
            "- Reinício automático habilitado, mas pywin32 (para controle de serviço) não está disponível.")
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
        about_win.geometry("450x350")
        about_win.resizable(False, False);
        about_win.transient(self.root);
        about_win.grab_set()
        frame = ttk.Frame(about_win, padding=20);
        frame.pack(fill='both', expand=True)
        ttk.Label(frame, text="Predadores Votemap Patch", font="-size 16 -weight bold").pack(pady=(0, 10))
        ttk.Label(frame, text="Versão 2.3.1 (Fix WMI)", font="-size 10").pack()  # Atualizar versão
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

    def check_for_updates(self):
        url = "https://github.com/raphaelpqdt/PQDT_Raphael_Votemappatch/tree/main/dist"
        self.append_text_gui(f"Abrindo página de atualizações: {url}\n")
        logging.info(f"Abrindo URL para verificação de atualizações: {url}")
        try:
            webbrowser.open_new_tab(url)
            self.set_status_from_thread("Página de atualizações aberta no navegador.")
        except Exception as e:
            logging.error(f"Erro ao abrir URL de atualizações '{url}': {e}", exc_info=True)
            self.show_messagebox_from_thread("error", "Erro ao Abrir Navegador",
                                             f"Não foi possível abrir o link de atualizações:\n{url}\n\nErro: {e}")
            self.set_status_from_thread("Falha ao abrir página de atualizações.")

    def toggle_log_search_bar(self, event=None, force_hide=False, force_show=False):
        if force_hide or (self.search_log_frame_visible and not force_show):
            if self.search_log_frame.winfo_ismapped():  # Só desempacota se estiver visível
                self.search_log_frame_visible = False
                self.search_log_frame.pack_forget()
                if self.text_area.winfo_exists(): self.text_area.focus_set()
                self.text_area.tag_remove("search_match", "1.0", "end")
        elif force_show or not self.search_log_frame_visible:
            if not self.search_log_frame.winfo_ismapped():  # Só empacota se não estiver visível
                self.search_log_frame_visible = True
                self.search_log_frame.pack(fill='x', before=self.text_area, pady=(0, 5))
                if self.log_search_entry.winfo_exists(): self.log_search_entry.focus_set()
                self.log_search_entry.select_range(0, 'end')
        self.last_search_pos = "1.0"  # Resetar posição de busca ao mostrar/esconder a barra

    def _perform_log_search(self, term, start_pos, direction_forward=True, wrap=True):
        if not term or not self.text_area.winfo_exists():
            if self.text_area.winfo_exists(): self.text_area.tag_remove("search_match", "1.0", "end")
            return None

        self.text_area.tag_remove("search_match", "1.0", "end")
        count_var = tk.IntVar()

        original_state = self.text_area.cget("state")
        self.text_area.config(state="normal")

        pos = None
        if direction_forward:
            pos = self.text_area.search(term, start_pos, stopindex="end", count=count_var, nocase=True)
            if not pos and wrap and start_pos != "1.0":  # Se não achou e deve dar wrap E não começou do início
                pos = self.text_area.search(term, "1.0", stopindex=start_pos, count=count_var, nocase=True)
        else:
            pos = self.text_area.search(term, start_pos, stopindex="1.0", count=count_var, nocase=True, backwards=True)
            if not pos and wrap and start_pos != "end":  # Se não achou e deve dar wrap E não começou do fim
                pos = self.text_area.search(term, "end", stopindex=start_pos, count=count_var, nocase=True,
                                            backwards=True)

        if pos:
            end_pos = f"{pos}+{count_var.get()}c"
            self.text_area.tag_add("search_match", pos, end_pos)
            self.text_area.tag_config("search_match", background="yellow", foreground="black")
            self.text_area.see(pos)
            self.text_area.config(state=original_state)
            return end_pos if direction_forward else pos
        else:
            self.text_area.config(state=original_state)
            # Não mostrar messagebox aqui para não interromper a digitação
            self.set_status_from_thread(f"'{term}' não encontrado.")
            return None

    def search_log_next(self, event=None):
        term = self.log_search_var.get()
        if not term: return

        # Se houver uma seleção atual, começar depois dela
        current_match_ranges = self.text_area.tag_ranges("search_match")
        start_from = self.last_search_pos
        if current_match_ranges:
            start_from = current_match_ranges[1]  # Fim da seleção atual

        next_start_pos = self._perform_log_search(term, start_from, direction_forward=True)
        if next_start_pos:
            self.last_search_pos = next_start_pos
        else:  # Não encontrou mais para frente, tentar do início (wrap)
            next_start_pos_wrapped = self._perform_log_search(term, "1.0", direction_forward=True,
                                                              wrap=False)  # Wrap é implícito
            if next_start_pos_wrapped:
                self.last_search_pos = next_start_pos_wrapped
            else:  # Realmente não encontrou
                self.last_search_pos = "1.0"  # Resetar para próxima busca do início

    def search_log_prev(self, event=None):
        term = self.log_search_var.get()
        if not term: return

        current_match_ranges = self.text_area.tag_ranges("search_match")
        start_from = self.last_search_pos
        if current_match_ranges:
            start_from = current_match_ranges[0]  # Início da seleção atual

        new_match_start_pos = self._perform_log_search(term, start_from, direction_forward=False)
        if new_match_start_pos:
            self.last_search_pos = new_match_start_pos
        else:  # Não encontrou mais para trás, tentar do fim (wrap)
            new_match_start_pos_wrapped = self._perform_log_search(term, "end", direction_forward=False, wrap=False)
            if new_match_start_pos_wrapped:
                self.last_search_pos = new_match_start_pos_wrapped
            else:  # Realmente não encontrou
                self.last_search_pos = "end"  # Resetar para próxima busca do fim

    def setup_tray_icon(self):
        try:
            image = self._create_tray_image()
            if image is None:
                logging.error("Não foi possível criar a imagem para o ícone da bandeja.")
                return

            menu = pystray.Menu(
                pystray.MenuItem('Mostrar Predadores Votemap', self.show_from_tray, default=True),
                pystray.MenuItem('Sair', self.on_close_from_tray_menu_item)
            )
            self.tray_icon = pystray.Icon("predadores_votemap_patch", image, "Predadores Votemap Patch", menu)
            threading.Thread(target=self.tray_icon.run, daemon=True, name="TrayIconThread").start()
            logging.info("Ícone da bandeja do sistema configurado e iniciado.")
        except Exception as e:
            logging.error(f"Falha ao criar ícone da bandeja: {e}", exc_info=True)

    def show_from_tray(self):
        if self.root.winfo_exists():
            self.root.after(0, self.root.deiconify)

    def minimize_to_tray(self, event=None):
        # Só minimiza para bandeja se o ícone estiver visível e a janela estiver sendo minimizada
        if hasattr(self, 'tray_icon') and self.tray_icon.visible:
            if self.root.winfo_exists() and self.root.state() == 'iconic':
                self.root.withdraw()
                logging.info("Aplicação minimizada para a bandeja.")

    def on_close_from_tray_menu_item(self):
        logging.info("Comando 'Sair' do menu da bandeja recebido.")
        self.on_close_common_logic(initiated_by_tray=True)

    def on_close(self):
        logging.info("Tentativa de fechar a janela principal (WM_DELETE_WINDOW).")
        if self.root.winfo_exists():
            if self.config_changed:
                response = Messagebox.yesnocancel(
                    title="Salvar Alterações?",
                    message="Existem configurações não salvas. Deseja salvá-las antes de sair?",
                    parent=self.root,
                    alert=True
                )
                if response == "Yes":
                    self.save_config()  # Salva e então prossegue para fechar
                    self.on_close_common_logic()
                elif response == "No":
                    self.on_close_common_logic()  # Prossegue para fechar sem salvar
                # else: "Cancel", não faz nada, a janela permanece aberta
            else:  # Nenhuma alteração, apenas confirma
                if Messagebox.okcancel("Confirmar Saída", "Deseja realmente sair do Predadores Votemap Patch?",
                                       parent=self.root, alert=True) == "OK":
                    self.on_close_common_logic()
                else:
                    logging.info("Saída cancelada pelo usuário (via janela).")

    def on_close_common_logic(self, initiated_by_tray=False):
        logging.info(
            f"Iniciando lógica comum de fechamento (iniciado por {'bandeja' if initiated_by_tray else 'janela'}).")
        if self.root.winfo_exists():
            self.set_status_from_thread("Encerrando...")
            if not initiated_by_tray:
                try:
                    self.root.update_idletasks()
                except tk.TclError:
                    pass

        self.stop_log_monitoring()

        if hasattr(self, 'tray_icon') and self.tray_icon.visible:
            try:
                self.tray_icon.stop()
            except Exception as e_tray:
                logging.error(f"Erro ao parar ícone da bandeja: {e_tray}", exc_info=True)

        # A lógica de salvar config já foi tratada em on_close() se não for initiated_by_tray
        # Se for initiated_by_tray, geralmente não salvamos automaticamente a menos que seja um comportamento desejado.
        # Para manter simples, não salvamos se initiated_by_tray aqui, assumindo que on_close fez o prompt.
        # Se initiated_by_tray E config_changed, o usuário perde as alterações. Uma alternativa
        # seria mostrar uma notificação da bandeja ou um prompt rápido, mas é mais complexo.

        logging.info(f"Aplicação encerrada (via {'bandeja' if initiated_by_tray else 'janela'}).")
        if self.root.winfo_exists():
            self.root.destroy()


def main():
    # Não inicializar COM aqui, cada thread que precisar fará sua própria inicialização/desinicialização.
    root = ttk.Window()  # Deixar ttkbootstrap escolher o tema inicial ou usar o da config
    app = LogViewerApp(root)
    root.protocol("WM_DELETE_WINDOW", app.on_close)
    root.bind("<Unmap>", app.minimize_to_tray)
    app.setup_tray_icon()

    try:
        root.mainloop()
    except KeyboardInterrupt:
        logging.info("Interrupção por teclado recebida. Encerrando...")
        app.on_close_common_logic(initiated_by_tray=True)  # Força fechamento
    finally:
        logging.info("Aplicação finalizada (bloco finally do main).")


if __name__ == '__main__':
    if not PYWIN32_AVAILABLE and platform.system() == "Windows":
        # Logar mais cedo se pywin32 estiver faltando no Windows
        logging.warning("pywin32 não está instalado. Funcionalidades de serviço do Windows serão desabilitadas.")


    def handle_thread_exception(args):
        # Assegurar que args.thread não é None antes de acessar .name
        thread_name = args.thread.name if args.thread else 'Desconhecida'
        logging.critical(f"EXCEÇÃO NÃO CAPTURADA NA THREAD {thread_name}:",
                         exc_info=(args.exc_type, args.exc_value, args.exc_traceback))
        # Evitar print se estivermos em um ambiente sem console (PyInstaller --noconsole)
        if sys.stderr:
            import traceback
            traceback.print_exception(args.exc_type, args.exc_value, args.exc_traceback)


    threading.excepthook = handle_thread_exception
    main()