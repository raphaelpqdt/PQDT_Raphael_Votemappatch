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
from tkinter import simpledialog  # Para renomear
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

# Configure logging (ajustado)
logging.basicConfig(
    level=logging.INFO,  # Use INFO para produção, DEBUG para desenvolvimento
    format='%(asctime)s - %(levelname)s - [%(threadName)s] - %(module)s.%(funcName)s:%(lineno)d - %(message)s',
    filename='votemap_patch_multi.log',  # Nome de log diferente
    filemode='a',
    encoding='utf-8'
)


def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(os.path.dirname(__file__))
    return os.path.join(base_path, relative_path)


ICON_FILENAME = "pred.ico"
ICON_PATH = resource_path(ICON_FILENAME)


# ############################################################################
# # Classe ServidorTab - Representa uma aba individual de servidor
# ############################################################################
class ServidorTab(ttk.Frame):
    def __init__(self, master_notebook, app_instance, nome_servidor, config_dict=None):
        super().__init__(master_notebook)
        self.app = app_instance  # Referência à instância LogViewerApp
        self.nome = nome_servidor
        self.config_inicial = config_dict if config_dict else {}

        # --- Variáveis de Estado Específicas do Servidor ---
        self.pasta_raiz = tk.StringVar(value=self.config_inicial.get("log_folder", ""))
        self.arquivo_json = tk.StringVar(value=self.config_inicial.get("server_json", ""))
        self.arquivo_json_votemap = tk.StringVar(value=self.config_inicial.get("votemap_json", ""))
        self.nome_servico = tk.StringVar(value=self.config_inicial.get("service_name", ""))

        self.log_folder_path_label_var = tk.StringVar(value="Pasta Logs: Nenhuma")
        self.server_json_path_label_var = tk.StringVar(value="JSON Servidor: Nenhum")
        self.votemap_json_path_label_var = tk.StringVar(value="JSON Votemap: Nenhum")
        self.servico_label_var = tk.StringVar(value="Serviço: Nenhum")  # Renomeado para evitar conflito

        self.filtro_var = tk.StringVar(value=self.config_inicial.get("filter", ""))
        self.auto_restart_var = tk.BooleanVar(value=self.config_inicial.get("auto_restart", True))
        self.vote_pattern_var = tk.StringVar(value=self.config_inicial.get("vote_pattern", r"\.EndVote\(\)"))
        self.winner_pattern_var = tk.StringVar(
            value=self.config_inicial.get("winner_pattern", r"Winner: \[(\d+)\]"))
        self.default_mission_var = tk.StringVar(
            value=self.config_inicial.get("default_mission", "{B88CC33A14B71FDC}Missions/V30_MapVoting_Mission.conf"))
        self.stop_delay_var = tk.IntVar(value=self.config_inicial.get("stop_delay", 10))
        self.start_delay_var = tk.IntVar(value=self.config_inicial.get("start_delay", 30))
        self.auto_scroll_log_var = tk.BooleanVar(value=self.config_inicial.get("auto_scroll_log", True))

        self.log_search_var = tk.StringVar()
        self.last_search_pos = "1.0"
        self.search_log_frame_visible = False

        # --- Atributos de Monitoramento ---
        self._stop_event = threading.Event()
        self._paused = False
        self.log_monitor_thread = None
        self.log_tail_thread = None
        self.file_log_handle = None
        self.caminho_log_atual = None
        self.pasta_log_detectada_atual = None  # Para rastrear a pasta de log ex: 2023-10-27_10-00-00

        self._create_ui_for_tab()
        self.initialize_from_config_vars()

        # Adicionar listeners para detectar mudanças e notificar app principal
        vars_to_trace_str = [
            self.pasta_raiz, self.arquivo_json, self.arquivo_json_votemap, self.nome_servico,
            self.filtro_var, self.vote_pattern_var, self.winner_pattern_var, self.default_mission_var
        ]
        vars_to_trace_bool = [self.auto_restart_var, self.auto_scroll_log_var]
        vars_to_trace_int = [self.stop_delay_var, self.start_delay_var]

        for var in vars_to_trace_str:
            var.trace_add("write", lambda *args, v=var: self._value_changed(v.get()))
        for var in vars_to_trace_bool:
            var.trace_add("write", lambda *args, v=var: self._value_changed(v.get()))
        for var in vars_to_trace_int:
            var.trace_add("write", lambda *args, v=var: self._value_changed(v.get()))

    def _value_changed(self, new_value=None):  # new_value é para o trace
        # logging.debug(f"Tab '{self.nome}': Value changed to '{new_value}'")
        self.app.mark_config_changed()

    def get_current_config(self):
        """Retorna a configuração atual desta aba de servidor."""
        return {
            "nome": self.nome,  # O nome da aba é gerenciado por LogViewerApp, mas útil ter aqui
            "log_folder": self.pasta_raiz.get(),
            "server_json": self.arquivo_json.get(),
            "votemap_json": self.arquivo_json_votemap.get(),
            "service_name": self.nome_servico.get(),
            "filter": self.filtro_var.get(),
            "auto_restart": self.auto_restart_var.get(),
            "vote_pattern": self.vote_pattern_var.get(),
            "winner_pattern": self.winner_pattern_var.get(),
            "default_mission": self.default_mission_var.get(),
            "stop_delay": self.stop_delay_var.get(),
            "start_delay": self.start_delay_var.get(),
            "auto_scroll_log": self.auto_scroll_log_var.get(),
        }

    def _create_ui_for_tab(self):
        # Frame superior para botões de seleção e labels de caminho
        outer_top_frame = ttk.Frame(self)
        outer_top_frame.pack(pady=5, padx=5, fill='x')  # Reduzido pady/padx

        selection_labelframe = ttk.Labelframe(outer_top_frame, text="Configuração de Caminhos e Serviço",
                                              padding=(10, 5))
        selection_labelframe.pack(side='top', fill='x', pady=(0, 5))

        path_buttons_frame = ttk.Frame(selection_labelframe)
        path_buttons_frame.pack(fill='x')

        # Botões de seleção
        self.selecionar_btn = ttk.Button(path_buttons_frame, text="Pasta de Logs", command=self.selecionar_pasta,
                                         bootstyle=PRIMARY)
        self.selecionar_btn.pack(side='left', pady=2, padx=(0, 2))
        ToolTip(self.selecionar_btn,
                text="Seleciona a pasta raiz onde os logs do servidor são armazenados (ex: .../ArmaReforgerServer/profile/logs).")

        self.json_btn = ttk.Button(path_buttons_frame, text="JSON Servidor",
                                   command=self.selecionar_arquivo_json_servidor, bootstyle=INFO)
        self.json_btn.pack(side='left', padx=2, pady=2)
        ToolTip(self.json_btn,
                text="Seleciona o arquivo JSON de configuração principal do servidor (geralmente config.json).")

        self.json_vm_btn = ttk.Button(path_buttons_frame, text="JSON Votemap",
                                      command=self.selecionar_arquivo_json_votemap, bootstyle=INFO)
        self.json_vm_btn.pack(side='left', padx=2, pady=2)
        ToolTip(self.json_vm_btn, text="Seleciona o arquivo JSON de configuração do Votemap (ex: votemap.json).")

        self.servico_btn = ttk.Button(path_buttons_frame, text="Serviço Win", command=self.selecionar_servico,
                                      bootstyle=SECONDARY)
        self.servico_btn.pack(side='left', padx=2, pady=2)
        ToolTip(self.servico_btn, text="Seleciona o serviço do Windows associado ao servidor do jogo.")
        if not PYWIN32_AVAILABLE: self.servico_btn.config(state=DISABLED)

        self.refresh_servico_status_btn = ttk.Button(path_buttons_frame, text="↻",
                                                     command=self.update_service_status_display,
                                                     bootstyle=(TOOLBUTTON, LIGHT), width=2)
        self.refresh_servico_status_btn.pack(side='left', padx=(0, 2), pady=2)
        ToolTip(self.refresh_servico_status_btn, text="Atualizar status do serviço selecionado.")
        if not PYWIN32_AVAILABLE: self.refresh_servico_status_btn.config(state=DISABLED)

        # Labels de caminho (agora em duas linhas para melhor layout)
        path_labels_frame_line1 = ttk.Frame(selection_labelframe)
        path_labels_frame_line1.pack(fill='x', pady=(5, 2))
        self.log_folder_path_label = ttk.Label(path_labels_frame_line1, textvariable=self.log_folder_path_label_var,
                                               wraplength=450, anchor='w')
        self.log_folder_path_label.pack(side='left', padx=5, fill='x', expand=True)

        self.servico_label_widget = ttk.Label(path_labels_frame_line1, textvariable=self.servico_label_var, anchor='w',
                                              width=30)  # Largura fixa
        self.servico_label_widget.pack(side='left', padx=(5, 0))

        path_labels_frame_line2 = ttk.Frame(selection_labelframe)
        path_labels_frame_line2.pack(fill='x', pady=(0, 0))
        self.json_server_path_label = ttk.Label(path_labels_frame_line2, textvariable=self.server_json_path_label_var,
                                                wraplength=220, anchor='w')
        self.json_server_path_label.pack(side='left', padx=5, fill='x', expand=True)
        self.json_votemap_path_label = ttk.Label(path_labels_frame_line2, textvariable=self.votemap_json_path_label_var,
                                                 wraplength=220, anchor='w')
        self.json_votemap_path_label.pack(side='left', padx=5, fill='x', expand=True)

        # Frame para controles de log
        controls_labelframe = ttk.Labelframe(outer_top_frame, text="Controles de Log", padding=(10, 5))
        controls_labelframe.pack(side='top', fill='x', pady=(5, 0))

        log_controls_subframe = ttk.Frame(controls_labelframe)
        log_controls_subframe.pack(fill='x', expand=True)

        ttk.Label(log_controls_subframe, text="Filtro:").pack(side='left', padx=(0, 5))
        self.filtro_entry = ttk.Entry(log_controls_subframe, textvariable=self.filtro_var, width=20)
        self.filtro_entry.pack(side='left', padx=(0, 5))
        ToolTip(self.filtro_entry, text="Filtra as linhas de log exibidas (case-insensitive).")

        self.refresh_json_btn = ttk.Button(log_controls_subframe, text="Atualizar JSONs",
                                           command=self.forcar_refresh_json_display, bootstyle=SUCCESS)
        self.refresh_json_btn.pack(side='left', padx=5)
        ToolTip(self.refresh_json_btn, text="Recarrega e exibe o conteúdo dos arquivos JSON selecionados.")

        self.pausar_btn = ttk.Button(log_controls_subframe, text="⏸️ Pausar", command=self.toggle_pausa,
                                     bootstyle=WARNING)
        self.pausar_btn.pack(side='left', padx=5)
        ToolTip(self.pausar_btn, text="Pausa ou retoma o acompanhamento ao vivo dos logs.")

        self.limpar_btn = ttk.Button(log_controls_subframe, text="♻️ Limpar Log", command=self.limpar_tela_log,
                                     bootstyle=SECONDARY)
        self.limpar_btn.pack(side='left', padx=5)
        ToolTip(self.limpar_btn, text="Limpa a área de exibição de logs do servidor.")

        # Notebook interno da aba do servidor (Logs, Configs JSON, Opções Votemap)
        self.tab_notebook = ttk.Notebook(self)
        self.tab_notebook.pack(fill='both', expand=True, padx=5, pady=(5, 5))  # Reduzido pady

        # --- Aba de Logs do Servidor (dentro do tab_notebook) ---
        log_frame = ttk.Frame(self.tab_notebook)
        self.tab_notebook.add(log_frame, text="Logs do Servidor")
        self.log_label_display = ttk.Label(log_frame, text="LOG AO VIVO DO SERVIDOR", foreground="red")
        self.log_label_display.pack(pady=(5, 0))

        self.search_log_frame = ttk.Frame(log_frame)
        # ... (código da barra de busca como antes, mas usando self.search_log_frame, self.log_search_entry, etc.)
        ttk.Label(self.search_log_frame, text="Buscar:").pack(side='left', padx=(5, 2))
        self.log_search_entry = ttk.Entry(self.search_log_frame, textvariable=self.log_search_var)
        self.log_search_entry.pack(side='left', fill='x', expand=True, padx=2)
        self.log_search_entry.bind("<Return>", self._search_log_next)  # Usar métodos prefixados com _ se forem internos
        search_next_btn = ttk.Button(self.search_log_frame, text="Próximo", command=self._search_log_next,
                                     bootstyle=SECONDARY)
        search_next_btn.pack(side='left', padx=2)
        search_prev_btn = ttk.Button(self.search_log_frame, text="Anterior", command=self._search_log_prev,
                                     bootstyle=SECONDARY)
        search_prev_btn.pack(side='left', padx=2)
        close_search_btn = ttk.Button(self.search_log_frame, text="X", command=self._toggle_log_search_bar,
                                      bootstyle=(SECONDARY, DANGER), width=2)
        close_search_btn.pack(side='left', padx=(2, 5))
        # self.search_log_frame não é empacotado aqui, _toggle_log_search_bar fará isso.

        self.text_area_log = ScrolledText(log_frame, wrap='word', height=10, state='disabled')  # Altura ajustada
        self.text_area_log.pack(fill='both', expand=True, pady=(0, 5))
        self.text_area_log.bind("<Control-f>", lambda e: self._toggle_log_search_bar(force_show=True))
        # O bind Escape no root da app principal para fechar a barra de busca é feito em LogViewerApp

        self.auto_scroll_check = ttk.Checkbutton(log_frame, text="Rolar Auto.", variable=self.auto_scroll_log_var)
        self.auto_scroll_check.pack(side='left', anchor='sw', pady=2, padx=5)  # Ajustado para esquerda

        # --- Aba de Config. Servidor (JSON) ---
        server_json_frame = ttk.Frame(self.tab_notebook)
        self.tab_notebook.add(server_json_frame, text="JSON Servidor")
        self.json_server_title_label = ttk.Label(server_json_frame,
                                                 text="CONTEÚDO DO ARQUIVO JSON DE CONFIGURAÇÃO DO SERVIDOR",
                                                 foreground="blue")
        self.json_server_title_label.pack(pady=(5, 0))
        self.json_text_area_server = ScrolledText(server_json_frame, wrap='word', height=10, state='disabled')
        self.json_text_area_server.pack(fill='both', expand=True, padx=5, pady=5)

        # --- Aba de Config. Votemap (JSON) ---
        votemap_json_frame = ttk.Frame(self.tab_notebook)
        self.tab_notebook.add(votemap_json_frame, text="JSON Votemap")
        self.json_votemap_title_label = ttk.Label(votemap_json_frame,
                                                  text="CONTEÚDO DO ARQUIVO JSON DE CONFIGURAÇÃO DO VOTEMAP",
                                                  foreground="blue")
        self.json_votemap_title_label.pack(pady=(5, 0))
        self.json_text_area_votemap = ScrolledText(votemap_json_frame, wrap='word', height=10, state='disabled')
        self.json_text_area_votemap.pack(fill='both', expand=True, padx=5, pady=5)

        # --- Aba de Opções Votemap (Regex, Delays) ---
        options_votemap_frame = ttk.Frame(self.tab_notebook)
        self.tab_notebook.add(options_votemap_frame, text=f"Opções Votemap")
        options_inner_frame = ttk.Frame(options_votemap_frame, padding=15)
        options_inner_frame.pack(fill='both', expand=True)

        self.auto_restart_check = ttk.Checkbutton(options_inner_frame,
                                                  text="Reiniciar servidor automaticamente após troca de mapa",
                                                  variable=self.auto_restart_var)
        self.auto_restart_check.grid(row=0, column=0, sticky='w', padx=5, pady=5, columnspan=2)
        ToolTip(self.auto_restart_check,
                "Se marcado, o servidor será reiniciado após uma votação de mapa bem-sucedida.")

        ttk.Label(options_inner_frame, text="Padrão detecção de voto (RegEx):").grid(row=1, column=0, sticky='w',
                                                                                     padx=5, pady=(10, 0))
        vote_pattern_entry = ttk.Entry(options_inner_frame, textvariable=self.vote_pattern_var, width=60)
        vote_pattern_entry.grid(row=2, column=0, sticky='ew', padx=5, pady=2, columnspan=2)
        ToolTip(vote_pattern_entry, "Expressão regular para detectar o fim de uma votação no log.")

        ttk.Label(options_inner_frame, text="Padrão detecção de vencedor (RegEx):").grid(row=3, column=0, sticky='w',
                                                                                         padx=5, pady=(10, 0))
        winner_pattern_entry = ttk.Entry(options_inner_frame, textvariable=self.winner_pattern_var, width=60)
        winner_pattern_entry.grid(row=4, column=0, sticky='ew', padx=5, pady=2, columnspan=2)
        ToolTip(winner_pattern_entry,
                "Expressão regular para capturar o índice do mapa vencedor (o primeiro grupo de captura é usado).")

        ttk.Label(options_inner_frame, text="Missão padrão de votemap (ScenarioID):").grid(row=5, column=0, sticky='w',
                                                                                           padx=5, pady=(10, 0))
        default_mission_entry = ttk.Entry(options_inner_frame, textvariable=self.default_mission_var, width=60)
        default_mission_entry.grid(row=6, column=0, sticky='ew', padx=5, pady=2, columnspan=2)
        ToolTip(default_mission_entry,
                "ID do cenário/missão a ser carregado após um reinício para iniciar uma nova votação.")

        delay_frame = ttk.Frame(options_inner_frame)
        delay_frame.grid(row=7, column=0, columnspan=2, sticky='ew', pady=(10, 0))
        ttk.Label(delay_frame, text="Delay Parar (s):").pack(side='left', padx=5)
        stop_delay_spinbox = ttk.Spinbox(delay_frame, from_=1, to=60, textvariable=self.stop_delay_var, width=5)
        stop_delay_spinbox.pack(side='left', padx=5)
        ToolTip(stop_delay_spinbox, "Tempo (s) para aguardar após comando de parada.")

        ttk.Label(delay_frame, text="Delay Iniciar (s):").pack(side='left', padx=15)
        start_delay_spinbox = ttk.Spinbox(delay_frame, from_=5, to=180, textvariable=self.start_delay_var, width=5)
        start_delay_spinbox.pack(side='left', padx=5)
        ToolTip(start_delay_spinbox, "Tempo (s) para aguardar o servidor iniciar completamente.")
        options_inner_frame.columnconfigure(0, weight=1)  # Para expandir entries

    def initialize_from_config_vars(self):
        """Inicializa os labels e o monitoramento com base nas tk.StringVars já populadas."""
        default_fg = self.app.style.colors.fg if hasattr(self.app.style,
                                                         'colors') and self.app.style.colors else "black"
        pasta_raiz_val = self.pasta_raiz.get()
        if pasta_raiz_val and os.path.isdir(pasta_raiz_val):
            self.append_text_to_log_area(f">>> Pasta de logs configurada: {pasta_raiz_val}\n")
            self.log_folder_path_label_var.set(f"Pasta Logs: {os.path.basename(pasta_raiz_val)}")
            self.log_folder_path_label.config(foreground="green")
            self.start_log_monitoring()
        elif pasta_raiz_val:
            self.log_folder_path_label_var.set(f"Pasta Logs (INVÁLIDA): {os.path.basename(pasta_raiz_val)}")
            self.log_folder_path_label.config(foreground="red")
        else:
            self.log_folder_path_label_var.set("Pasta Logs: Nenhuma")
            self.log_folder_path_label.config(foreground=default_fg)

        self.forcar_refresh_json_display()

        nome_servico_val = self.nome_servico.get()
        if nome_servico_val and PYWIN32_AVAILABLE:
            self.update_service_status_display()
        elif not PYWIN32_AVAILABLE:
            self.servico_label_var.set("Serviço: N/A (pywin32)")
            self.servico_label_widget.config(foreground="gray")
        else:
            self.servico_label_var.set("Serviço: Nenhum")
            self.servico_label_widget.config(foreground=default_fg if default_fg != "black" else "orange")

        # logging.info(f"ServidorTab '{self.nome}' inicializado com base nas variáveis.")

    def forcar_refresh_json_display(self):
        """Atualiza a exibição dos JSONs com base nos caminhos em self.arquivo_json e self.arquivo_json_votemap."""
        # JSON do Servidor
        self._refresh_single_json_display(
            self.arquivo_json.get(),
            self.json_text_area_server,
            self.server_json_path_label_var,
            self.json_server_path_label,
            "Servidor"
        )
        # JSON do Votemap
        self._refresh_single_json_display(
            self.arquivo_json_votemap.get(),
            self.json_text_area_votemap,
            self.votemap_json_path_label_var,
            self.json_votemap_path_label,
            "Votemap"
        )
        if self.app:  # Pode ser None durante a inicialização muito cedo
            self.app.set_status_from_thread(f"JSONs para '{self.nome}' atualizados.")

    def _refresh_single_json_display(self, file_path, text_widget, label_var, label_widget, json_type_name):
        default_fg = self.app.style.colors.fg if hasattr(self.app.style,
                                                         'colors') and self.app.style.colors else "black"
        if file_path:
            if os.path.exists(file_path):
                try:
                    with open(file_path, 'r', encoding='utf-8') as f:
                        conteudo = f.read()
                    json_data = json.loads(conteudo)  # Valida
                    self._display_json_in_widget(text_widget, json_data)
                    self.append_text_to_log_area(
                        f"JSON de {json_type_name} '{os.path.basename(file_path)}' recarregado.\n")
                    label_var.set(f"JSON {json_type_name}: {os.path.basename(file_path)}")
                    label_widget.config(foreground="green")
                except json.JSONDecodeError:
                    self._display_json_in_widget(text_widget,
                                                 f"ERRO: Conteúdo de '{os.path.basename(file_path)}' não é JSON válido.")
                    self.append_text_to_log_area(
                        f"ERRO: JSON de {json_type_name} '{os.path.basename(file_path)}' é inválido.\n")
                    label_var.set(f"JSON {json_type_name} (INVÁLIDO): {os.path.basename(file_path)}")
                    label_widget.config(foreground="red")
                except Exception as e_read:
                    self._display_json_in_widget(text_widget, f"ERRO ao ler '{os.path.basename(file_path)}': {e_read}")
                    self.append_text_to_log_area(
                        f"ERRO ao ler JSON de {json_type_name} '{os.path.basename(file_path)}': {e_read}\n")
                    label_var.set(f"JSON {json_type_name} (ERRO LEITURA): {os.path.basename(file_path)}")
                    label_widget.config(foreground="red")
            else:
                self.append_text_to_log_area(f"Arquivo JSON de {json_type_name} '{file_path}' não encontrado.\n")
                self._display_json_in_widget(text_widget, "Arquivo não encontrado.")
                label_var.set(
                    f"JSON {json_type_name} (NÃO ENC.): {os.path.basename(file_path) if file_path else 'Nenhum'}")
                label_widget.config(foreground="orange")
        else:
            self.append_text_to_log_area(f"Caminho do JSON de {json_type_name} não configurado.\n")
            self._display_json_in_widget(text_widget, "Não configurado.")
            label_var.set(f"JSON {json_type_name}: Nenhum")
            label_widget.config(foreground=default_fg)

    def _display_json_in_widget(self, text_area_widget, content):
        """Exibe o conteúdo (string ou dados JSON) formatado na área de texto."""
        try:
            # Se for dict/list, formata. Se já for string (erro), usa direto.
            dados_formatados = json.dumps(content, indent=4, ensure_ascii=False) if isinstance(content,
                                                                                               (dict, list)) else str(
                content)
        except (TypeError):  # Caso content já seja uma string de erro
            dados_formatados = str(content)

        if self.winfo_exists() and text_area_widget.winfo_exists():
            try:
                text_area_widget.configure(state='normal')
                text_area_widget.delete('1.0', 'end')
                text_area_widget.insert('end', dados_formatados)
                text_area_widget.configure(state='disabled')
            except tk.TclError:
                logging.warning(
                    f"TclError ao exibir JSON no widget {text_area_widget} para '{self.nome}' (GUI fechando?).")

    def selecionar_pasta(self):
        pasta_selecionada = filedialog.askdirectory(
            title=f"Selecione a pasta de logs para '{self.nome}'",
            initialdir=self.pasta_raiz.get() or os.path.expanduser("~")
        )
        if pasta_selecionada:
            if self.pasta_raiz.get() != pasta_selecionada:
                logging.info(f"Tab '{self.nome}': Pasta de logs alterada para '{pasta_selecionada}'")
                self.stop_log_monitoring()  # Para o monitoramento antigo
                self.pasta_raiz.set(pasta_selecionada)  # Isso vai trigar _value_changed
                self.initialize_from_config_vars()  # Re-inicializa labels e começa novo monitoramento se válido
            else:
                self.append_text_to_log_area(f">>> Pasta de logs já selecionada: {pasta_selecionada}\n")
        # Se não selecionar nada, não fazemos nada, o valor antigo permanece.

    def _selecionar_arquivo_json_generico(self, tipo_json_str, tk_var_caminho, text_widget, label_var, label_widget):
        """Função genérica para selecionar arquivos JSON."""
        caminho_atual = tk_var_caminho.get()
        diretorio_inicial = os.path.dirname(caminho_atual) if caminho_atual and os.path.exists(
            os.path.dirname(caminho_atual)) else self.pasta_raiz.get() or os.path.expanduser("~")

        caminho_selecionado = filedialog.askopenfilename(
            title=f"Selecionar JSON de {tipo_json_str} para '{self.nome}'",
            filetypes=[("Arquivos JSON", "*.json"), ("Todos os arquivos", "*.*")],
            initialdir=diretorio_inicial
        )
        if caminho_selecionado:
            if tk_var_caminho.get() != caminho_selecionado:
                tk_var_caminho.set(caminho_selecionado)  # Isso vai trigar _value_changed
                # A atualização da UI do JSON (texto e label) será feita por forcar_refresh_json_display
                # que é chamado indiretamente ou diretamente após a mudança de valor.
                self.forcar_refresh_json_display()  # Força refresh imediato da UI do JSON
            else:
                self.append_text_to_log_area(
                    f">>> Arquivo JSON de {tipo_json_str} já selecionado: {caminho_selecionado}\n")

    def selecionar_arquivo_json_servidor(self):
        self._selecionar_arquivo_json_generico(
            "Configuração do Servidor",
            self.arquivo_json,
            self.json_text_area_server,
            self.server_json_path_label_var,
            self.json_server_path_label
        )

    def selecionar_arquivo_json_votemap(self):
        self._selecionar_arquivo_json_generico(
            "Configuração do Votemap",
            self.arquivo_json_votemap,
            self.json_text_area_votemap,
            self.votemap_json_path_label_var,
            self.json_votemap_path_label
        )

    def selecionar_servico(self):
        if not PYWIN32_AVAILABLE:
            self.app.show_messagebox_from_thread("error", "Funcionalidade Indisponível",
                                                 "A biblioteca pywin32 é necessária para listar e gerenciar serviços do Windows.")
            return

        # Usar o método da app principal para mostrar o diálogo de progresso
        # e obter a lista de serviços, pois isso é mais global.
        # O resultado do diálogo será então usado para atualizar ESTA aba.
        self.app.iniciar_selecao_servico_para_aba(self)

    def set_selected_service(self, service_name):
        """Chamado por LogViewerApp após o diálogo de seleção de serviço."""
        if self.nome_servico.get() != service_name:
            self.nome_servico.set(service_name)  # Isso triggará _value_changed
            self.update_service_status_display()
            self.app.set_status_from_thread(f"Serviço '{service_name}' selecionado para '{self.nome}'.")
            logging.info(f"Tab '{self.nome}': Serviço selecionado: {service_name}")

    def update_service_status_display(self):
        if not PYWIN32_AVAILABLE:
            self.servico_label_var.set("Serviço: N/A (pywin32)")
            if hasattr(self.app.style, 'colors') and self.app.style.colors:
                self.servico_label_widget.config(foreground=self.app.style.colors.get('disabled', 'gray'))
            else:
                self.servico_label_widget.config(foreground="gray")
            return

        nome_servico_val = self.nome_servico.get()
        if nome_servico_val:
            current_text_base = f"Serviço: {nome_servico_val}"
            self.servico_label_var.set(f"{current_text_base} (Verificando...)")
            self.servico_label_widget.config(foreground="blue")  # Cor temporária enquanto verifica
            threading.Thread(
                target=self._get_and_display_service_status_thread_worker,
                args=(nome_servico_val, current_text_base),
                daemon=True,
                name=f"ServiceStatusCheck-{self.nome}"
            ).start()
        else:
            self.servico_label_var.set("Serviço: Nenhum")
            default_fg = "orange"
            if hasattr(self.app.style, 'colors') and self.app.style.colors:
                default_fg = self.app.style.colors.get('fg', 'orange')
            self.servico_label_widget.config(foreground=default_fg)

    def _get_and_display_service_status_thread_worker(self, service_name_to_check, base_text_for_label):
        status = self._verificar_status_servico_win(service_name_to_check)  # Método interno
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

        if self.app.root.winfo_exists() and self.winfo_exists():  # Verifica se a aba e a app ainda existem
            self.app.root.after(0, lambda: (
                self.servico_label_var.set(f"{base_text_for_label} {display_status_text}"),
                self.servico_label_widget.config(foreground=color) if self.servico_label_widget.winfo_exists() else None
            ))

    def _verificar_status_servico_win(self, nome_servico_local):
        """Verifica o status de um serviço do Windows usando 'sc query'."""
        if not PYWIN32_AVAILABLE: return "ERROR"  # Deveria ser checado antes, mas por segurança
        if not nome_servico_local: return "NOT_FOUND"
        try:
            startupinfo = None
            if platform.system() == "Windows":
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                startupinfo.wShowWindow = subprocess.SW_HIDE

            # Tentar com 'latin-1' para decodificar a saída do 'sc', comum em PT-BR Windows
            # Se falhar, tentar com 'utf-8' como fallback.
            encodings_to_try = ['latin-1', 'utf-8', 'cp850', 'cp1252']
            output_text = None
            for enc in encodings_to_try:
                try:
                    result = subprocess.run(
                        ['sc', 'query', nome_servico_local],
                        capture_output=True, text=False,  # text=False para decodificar manualmente
                        check=False,  # Não levantar exceção para códigos de saída != 0
                        startupinfo=startupinfo
                    )
                    # Tentar decodificar stdout e stderr
                    stdout_decoded = result.stdout.decode(enc, errors='replace')
                    stderr_decoded = result.stderr.decode(enc, errors='replace')
                    output_text = stdout_decoded + stderr_decoded
                    break  # Sucesso na decodificação
                except UnicodeDecodeError:
                    logging.warning(
                        f"Tab '{self.nome}': Falha ao decodificar saída 'sc query' com {enc} para '{nome_servico_local}'. Tentando próximo.")
                except Exception as e_run:  # Outros erros do subprocess
                    logging.error(
                        f"Tab '{self.nome}': Erro ao executar 'sc query' para '{nome_servico_local}': {e_run}",
                        exc_info=True)
                    return "ERROR"

            if output_text is None:  # Se todas as tentativas de decodificação falharem
                logging.error(
                    f"Tab '{self.nome}': Não foi possível decodificar a saída de 'sc query' para '{nome_servico_local}'.")
                return "ERROR"

            output_lower = output_text.lower()

            # Erros comuns para serviço não encontrado (incluindo PT-BR)
            service_not_found_errors = [
                "failed 1060", "falha 1060", "o servi‡o especificado nÆo existe como servi‡o instalado",
                "specified service does not exist as an installed service"
            ]
            if any(err_str in output_lower for err_str in service_not_found_errors):
                logging.warning(
                    f"Tab '{self.nome}': Serviço '{nome_servico_local}' não encontrado via 'sc query'. Output: {output_text[:200]}")
                return "NOT_FOUND"

            if "state" not in output_lower:  # Se a palavra "STATE" não estiver na saída
                logging.warning(
                    f"Tab '{self.nome}': Saída inesperada de 'sc query {nome_servico_local}', sem 'STATE': {output_text[:200]}")
                return "ERROR"  # Ou UNKNOWN se preferir

            # Mapeamento de estados (incluindo PT-BR comuns)
            if "running" in output_lower or "em execu‡Æo" in output_lower: return "RUNNING"
            if "stopped" in output_lower or "parado" in output_lower: return "STOPPED"
            if "start_pending" in output_lower or "pendente deinÝcio" in output_lower: return "START_PENDING"  # "início" com acento pode variar
            if "stop_pending" in output_lower or "pendente deparada" in output_lower: return "STOP_PENDING"  # "parada" com acento pode variar

            logging.info(
                f"Tab '{self.nome}': Status desconhecido para '{nome_servico_local}' com saída: {output_text[:200]}")
            return "UNKNOWN"

        except FileNotFoundError:  # sc.exe não encontrado
            logging.error(f"Tab '{self.nome}': 'sc.exe' não encontrado. Verifique se o System32 está no PATH.",
                          exc_info=True)
            # Não mostrar messagebox daqui, pois é uma thread worker. LogViewerApp pode mostrar se for um erro persistente.
            return "ERROR"
        except Exception as e:
            logging.error(f"Tab '{self.nome}': Erro ao verificar status do serviço '{nome_servico_local}': {e}",
                          exc_info=True)
            return "ERROR"

    def start_log_monitoring(self):
        if self.log_monitor_thread and self.log_monitor_thread.is_alive():
            logging.warning(f"Tab '{self.nome}': Tentativa de iniciar monitoramento de log já em execução.")
            return
        if not self.pasta_raiz.get() or not os.path.isdir(self.pasta_raiz.get()):
            self.append_text_to_log_area(
                f"AVISO: Pasta de logs '{self.pasta_raiz.get()}' inválida. Monitoramento não iniciado.\n")
            return

        self._stop_event.clear()
        self.log_monitor_thread = threading.Thread(
            target=self.monitorar_log_continuamente_worker,
            daemon=True,
            name=f"LogMonitor-{self.nome}"
        )
        self.log_monitor_thread.start()
        logging.info(f"Tab '{self.nome}': Monitoramento de logs iniciado para pasta '{self.pasta_raiz.get()}'.")

    def stop_log_monitoring(self, from_tab_closure=False):
        """Para o monitoramento de logs desta aba."""
        thread_name = threading.current_thread().name
        logging.debug(f"Tab '{self.nome}' [{thread_name}]: Chamada para stop_log_monitoring.")
        self._stop_event.set()

        # Parar thread de acompanhamento de arquivo (log_tail_thread)
        if self.log_tail_thread and self.log_tail_thread.is_alive():
            logging.debug(
                f"Tab '{self.nome}' [{thread_name}]: Aguardando LogTailThread ({self.log_tail_thread.name}) finalizar...")
            self.log_tail_thread.join(timeout=2.0)  # Timeout para evitar bloqueio indefinido
            if self.log_tail_thread.is_alive():
                logging.warning(
                    f"Tab '{self.nome}' [{thread_name}]: LogTailThread ({self.log_tail_thread.name}) não finalizou no tempo esperado.")
        self.log_tail_thread = None

        # Parar thread principal de monitoramento de pasta (log_monitor_thread)
        if self.log_monitor_thread and self.log_monitor_thread.is_alive() and self.log_monitor_thread != threading.current_thread():
            logging.debug(
                f"Tab '{self.nome}' [{thread_name}]: Aguardando LogMonitorThread ({self.log_monitor_thread.name}) finalizar...")
            self.log_monitor_thread.join(timeout=2.0)
            if self.log_monitor_thread.is_alive():
                logging.warning(
                    f"Tab '{self.nome}' [{thread_name}]: LogMonitorThread ({self.log_monitor_thread.name}) não finalizou no tempo esperado.")
        self.log_monitor_thread = None

        # Fechar handle do arquivo de log
        if self.file_log_handle:
            try:
                handle_name_for_log = getattr(self.file_log_handle, 'name', 'N/A')
                logging.debug(
                    f"Tab '{self.nome}' [{thread_name}]: Fechando file_log_handle para: {handle_name_for_log}")
                self.file_log_handle.close()
            except Exception as e:
                logging.error(
                    f"Tab '{self.nome}' [{thread_name}]: Erro ao fechar handle do log ({handle_name_for_log}): {e}",
                    exc_info=True)
            finally:
                self.file_log_handle = None

        self.caminho_log_atual = None
        self.pasta_log_detectada_atual = None
        if not from_tab_closure:  # Só loga se não for parte do fechamento da aba/app
            logging.info(f"Tab '{self.nome}' [{thread_name}]: stop_log_monitoring completado.")

    def monitorar_log_continuamente_worker(self):
        """Thread worker para monitorar a pasta de logs e detectar novos arquivos/pastas de log."""
        thread_name = threading.current_thread().name
        pasta_raiz_monitorada = self.pasta_raiz.get()
        self.app.set_status_from_thread(
            f"'{self.nome}': Monitorando pasta: {os.path.basename(pasta_raiz_monitorada) if pasta_raiz_monitorada else 'N/A'}")
        logging.info(f"[{thread_name}] Tab '{self.nome}': Iniciando monitoramento contínuo de: {pasta_raiz_monitorada}")

        while not self._stop_event.is_set():
            if not pasta_raiz_monitorada or not os.path.isdir(pasta_raiz_monitorada):
                if pasta_raiz_monitorada:  # Se havia um caminho, mas agora é inválido
                    logging.warning(
                        f"[{thread_name}] Tab '{self.nome}': Pasta de logs '{pasta_raiz_monitorada}' não encontrada ou não é um diretório.")
                # Não precisa setar status aqui, pois a UI já deve refletir a pasta inválida
                if self._stop_event.wait(10): break  # Pausa longa se a pasta for inválida
                pasta_raiz_monitorada = self.pasta_raiz.get()  # Re-checa se o usuário corrigiu
                continue

            try:
                # Obter a subpasta de log mais recente (ex: .../logs/2023-10-27_10-00-00)
                subpasta_log_recente = self._obter_subpasta_log_mais_recente(pasta_raiz_monitorada)

                if not subpasta_log_recente:  # Nenhuma subpasta de log encontrada
                    if self.caminho_log_atual:  # Se antes estávamos monitorando algo
                        self.append_text_to_log_area(
                            f"AVISO: Nenhuma subpasta de log encontrada em '{pasta_raiz_monitorada}'. Verificando...\n")
                        self.caminho_log_atual = None  # Resetar
                        if self.log_tail_thread and self.log_tail_thread.is_alive():  # Parar tail antigo
                            self.log_tail_thread.join(timeout=1.0)
                        if self.file_log_handle: self.file_log_handle.close(); self.file_log_handle = None
                    if self._stop_event.wait(5): break  # Espera antes de checar de novo
                    continue

                # Caminho para o arquivo console.log dentro da subpasta mais recente
                novo_arquivo_log_path_potencial = os.path.join(subpasta_log_recente, 'console.log')

                # Se um novo arquivo de log foi detectado (diferente do atual) OU se não há log atual mas encontramos um
                if os.path.exists(novo_arquivo_log_path_potencial) and \
                        (novo_arquivo_log_path_potencial != self.caminho_log_atual or not self.caminho_log_atual):

                    logging.info(
                        f"[{thread_name}] Tab '{self.nome}': Novo arquivo de log detectado/mudança: '{novo_arquivo_log_path_potencial}' (anterior: '{self.caminho_log_atual}')")
                    self.append_text_to_log_area(
                        f"\n>>> Monitorando novo arquivo de log: {novo_arquivo_log_path_potencial}\n")

                    # Parar thread de acompanhamento anterior, se houver
                    if self.log_tail_thread and self.log_tail_thread.is_alive():
                        logging.debug(
                            f"[{thread_name}] Tab '{self.nome}': Parando LogTailThread antiga para '{self.caminho_log_atual}'...")
                        # Não precisa de _stop_event.set() aqui, o join deve ser suficiente ou o próprio LogTailThread deve verificar
                        self.log_tail_thread.join(timeout=1.5)
                        if self.log_tail_thread.is_alive():
                            logging.warning(
                                f"[{thread_name}] Tab '{self.nome}': LogTailThread antiga não parou a tempo.")

                    # Fechar handle do arquivo de log anterior, se houver
                    if self.file_log_handle:
                        try:
                            old_fh_name = getattr(self.file_log_handle, 'name', 'N/A')
                            logging.debug(
                                f"[{thread_name}] Tab '{self.nome}': Fechando file_log_handle antigo para: {old_fh_name}")
                            self.file_log_handle.close()
                        except Exception as e_close_old:
                            logging.error(
                                f"[{thread_name}] Tab '{self.nome}': Erro ao fechar handle antigo ({old_fh_name}): {e_close_old}",
                                exc_info=True)
                        finally:
                            self.file_log_handle = None

                    self.caminho_log_atual = novo_arquivo_log_path_potencial
                    self.pasta_log_detectada_atual = subpasta_log_recente  # Atualiza a pasta de log sendo monitorada

                    novo_fh_temp = None
                    try:
                        logging.debug(
                            f"[{thread_name}] Tab '{self.nome}': Tentando abrir novo arquivo de log: {self.caminho_log_atual}")
                        # Usar 'latin-1' como encoding padrão para logs de Arma, errors='replace' para evitar crash
                        novo_fh_temp = open(self.caminho_log_atual, 'r', encoding='latin-1', errors='replace')
                        novo_fh_temp.seek(0, os.SEEK_END)  # Ir para o fim do arquivo
                        self.file_log_handle = novo_fh_temp

                        if self.app.root.winfo_exists() and self.winfo_exists():  # Verifica se a aba e a app ainda existem
                            # Atualiza o label que mostra o arquivo de log sendo monitorado
                            log_file_display_name = os.path.join(os.path.basename(self.pasta_log_detectada_atual),
                                                                 os.path.basename(self.caminho_log_atual))
                            self.app.root.after(0, lambda p=log_file_display_name: self.log_label_display.config(
                                text=f"LOG: {p}") if self.log_label_display.winfo_exists() else None)
                            self.app.set_status_from_thread(f"'{self.nome}': Monitorando: {log_file_display_name}")

                        logging.info(
                            f"[{thread_name}] Tab '{self.nome}': Novo log {self.caminho_log_atual} aberto. Iniciando nova LogTailThread.")
                        self.log_tail_thread = threading.Thread(
                            target=self.acompanhar_log_do_arquivo_worker,
                            args=(self.caminho_log_atual,),
                            # Passa o caminho para a thread saber qual arquivo ela é responsável
                            daemon=True,
                            name=f"LogTail-{self.nome}-{os.path.basename(self.caminho_log_atual)}"
                        )
                        self.log_tail_thread.start()

                    except FileNotFoundError:
                        logging.error(
                            f"[{thread_name}] Tab '{self.nome}': Arquivo {self.caminho_log_atual} não encontrado ao tentar abrir.")
                        if novo_fh_temp: novo_fh_temp.close()
                        self.file_log_handle = None
                        self.caminho_log_atual = None  # Reset
                    except Exception as e_open_new:
                        logging.error(
                            f"[{thread_name}] Tab '{self.nome}': Erro ao abrir/acompanhar {self.caminho_log_atual}: {e_open_new}",
                            exc_info=True)
                        if novo_fh_temp: novo_fh_temp.close()
                        self.file_log_handle = None
                        self.caminho_log_atual = None  # Reset

                elif self.caminho_log_atual and not os.path.exists(self.caminho_log_atual):
                    # O arquivo que estávamos monitorando sumiu
                    logging.warning(
                        f"[{thread_name}] Tab '{self.nome}': Arquivo de log monitorado {self.caminho_log_atual} não existe mais.")
                    self.append_text_to_log_area(
                        f"AVISO: Arquivo de log {self.caminho_log_atual} não encontrado. Procurando por novo...\n")
                    if self.log_tail_thread and self.log_tail_thread.is_alive():
                        self.log_tail_thread.join(timeout=1.0)  # Tenta parar a thread antiga
                    if self.file_log_handle:
                        try:
                            self.file_log_handle.close()
                        except:
                            pass  # Ignora erros ao fechar se já estiver fechado
                    self.file_log_handle = None
                    self.caminho_log_atual = None  # Força a redetecção no próximo ciclo

            except Exception as e_monitor_loop:
                logging.error(
                    f"[{thread_name}] Tab '{self.nome}': Erro no loop principal de monitoramento: {e_monitor_loop}",
                    exc_info=True)
                self.append_text_to_log_area(
                    f"ERRO CRÍTICO AO MONITORAR LOGS: {e_monitor_loop}\nVerifique o Log do Sistema do Patch.\n")

            if self._stop_event.wait(5): break  # Intervalo de verificação da pasta de logs

        logging.info(
            f"[{thread_name}] Tab '{self.nome}': Thread de monitoramento de log contínuo ({thread_name}) encerrada.")
        # Limpeza final se a thread estiver parando
        if self.log_tail_thread and self.log_tail_thread.is_alive():
            self.log_tail_thread.join(timeout=1.0)
        if self.file_log_handle:
            try:
                self.file_log_handle.close()
            except:
                pass
        self.file_log_handle = None
        self.caminho_log_atual = None

    def _obter_subpasta_log_mais_recente(self, pasta_raiz_logs):
        """Obtém a subpasta de log mais recente dentro da pasta_raiz_logs."""
        if not pasta_raiz_logs or not os.path.isdir(pasta_raiz_logs):
            return None
        try:
            # Listar todas as entradas na pasta_raiz_logs
            entradas = os.listdir(pasta_raiz_logs)
            # Filtrar para manter apenas diretórios que parecem ser pastas de log (ex: YYYY-MM-DD_HH-MM-SS)
            # Regex simples para validar o formato esperado das pastas de log do Arma Reforger
            log_folder_pattern = re.compile(r"^logs_\d{4}-\d{2}-\d{2}_\d{2}-\d{2}-\d{2}$")

            subpastas_log_validas = []
            for nome_entrada in entradas:
                caminho_completo = os.path.join(pasta_raiz_logs, nome_entrada)
                if os.path.isdir(caminho_completo) and log_folder_pattern.match(nome_entrada):
                    subpastas_log_validas.append(caminho_completo)

            if not subpastas_log_validas:
                return None

            # Retornar a subpasta mais recente com base no tempo de modificação
            return max(subpastas_log_validas, key=os.path.getmtime)
        except FileNotFoundError:  # Se a pasta_raiz_logs sumir entre a verificação e o listdir
            logging.warning(f"Tab '{self.nome}': Pasta raiz '{pasta_raiz_logs}' não encontrada ao buscar subpastas.")
            self.pasta_raiz.set("")  # Limpa a configuração da pasta raiz se ela sumiu
            return None
        except PermissionError:
            logging.error(f"Tab '{self.nome}': Permissão negada ao acessar '{pasta_raiz_logs}' para buscar subpastas.")
            self.pasta_raiz.set("")  # Limpa a configuração
            return None
        except Exception as e:
            logging.error(f"Tab '{self.nome}': Erro ao obter subpasta mais recente em '{pasta_raiz_logs}': {e}",
                          exc_info=True)
            return None

    def acompanhar_log_do_arquivo_worker(self, caminho_log_designado_para_esta_thread):
        """Thread worker para acompanhar um arquivo de log específico (tail -f)."""
        thread_name = threading.current_thread().name
        logging.info(
            f"[{thread_name}] Tab '{self.nome}': Tentando iniciar acompanhamento para: {caminho_log_designado_para_esta_thread}")

        if self._stop_event.is_set():  # Verifica se já foi pedido para parar
            logging.info(
                f"[{thread_name}] Tab '{self.nome}': _stop_event já setado no início. Encerrando para {caminho_log_designado_para_esta_thread}.")
            return

        # Validação crítica: o file_log_handle deve estar aberto e correto
        if not self.file_log_handle or self.file_log_handle.closed:
            logging.error(
                f"[{thread_name}] Tab '{self.nome}': ERRO CRÍTICO - file_log_handle NULO ou FECHADO no início do acompanhamento para '{caminho_log_designado_para_esta_thread}'. Esta thread não pode prosseguir.")
            return
        try:
            # Compara o caminho do handle atual com o caminho que esta thread deveria monitorar
            handle_real_path_norm = os.path.normpath(self.file_log_handle.name)
            caminho_designado_norm = os.path.normpath(caminho_log_designado_para_esta_thread)
            if handle_real_path_norm != caminho_designado_norm:
                logging.warning(
                    f"[{thread_name}] Tab '{self.nome}': DESCOMPASSO DE HANDLE! Thread para '{caminho_designado_norm}' mas handle é '{handle_real_path_norm}'. Encerrando.")
                return
        except AttributeError:  # Se self.file_log_handle.name não existir
            logging.error(
                f"[{thread_name}] Tab '{self.nome}': ERRO CRÍTICO - file_log_handle inválido (sem 'name') para '{caminho_log_designado_para_esta_thread}'. Encerrando.")
            return
        except Exception as e_check_init_handle:
            logging.error(
                f"[{thread_name}] Tab '{self.nome}': Exceção na verificação inicial do handle para '{caminho_log_designado_para_esta_thread}': {e_check_init_handle}. Encerrando.")
            return

        logging.info(
            f"[{thread_name}] Tab '{self.nome}': Iniciando acompanhamento EFETIVO de: {caminho_log_designado_para_esta_thread}")
        aguardando_winner = False
        vote_pattern_re, winner_pattern_re = None, None
        try:
            # Compilar regex patterns uma vez
            vote_pattern_str = self.vote_pattern_var.get()
            if vote_pattern_str: vote_pattern_re = re.compile(vote_pattern_str)
            winner_pattern_str = self.winner_pattern_var.get()
            if winner_pattern_str: winner_pattern_re = re.compile(winner_pattern_str)
        except re.error as e_re_compile:
            logging.error(
                f"[{thread_name}] Tab '{self.nome}': Erro de RegEx nos padrões para '{caminho_log_designado_para_esta_thread}': {e_re_compile}",
                exc_info=True)
            self.append_text_to_log_area(f"ERRO DE REGEX: Verifique os padrões em 'Opções Votemap': {e_re_compile}\n")
            self.app.set_status_from_thread(f"'{self.nome}': Erro de RegEx! Verifique Configurações.")
            return  # Não continuar se os padrões são inválidos

        logging.debug(
            f"[{thread_name}] Tab '{self.nome}': Padrões para '{caminho_log_designado_para_esta_thread}': FimVoto='{vote_pattern_str}', Vencedor='{winner_pattern_str}'")
        logging.debug(
            f"[{thread_name}] Tab '{self.nome}': Estado inicial aguardando_winner={aguardando_winner} para '{caminho_log_designado_para_esta_thread}'")

        while not self._stop_event.is_set():
            if self._paused:  # Se a pausa desta aba estiver ativa
                if self._stop_event.wait(0.5): break  # Checa _stop_event periodicamente mesmo pausado
                continue

            # Validação do handle DENTRO do loop para detectar mudanças externas
            if not self.file_log_handle or self.file_log_handle.closed:
                logging.warning(
                    f"[{thread_name}] Tab '{self.nome}': file_log_handle NULO ou FECHADO DENTRO DO LOOP para '{caminho_log_designado_para_esta_thread}'. Encerrando thread.")
                break
            try:
                current_handle_path_norm = os.path.normpath(self.file_log_handle.name)
                caminho_designado_norm_loop = os.path.normpath(caminho_log_designado_para_esta_thread)
                if current_handle_path_norm != caminho_designado_norm_loop:
                    logging.warning(
                        f"[{thread_name}] Tab '{self.nome}': MUDANÇA DE HANDLE DETECTADA! Thread para '{caminho_designado_norm_loop}', mas handle é '{current_handle_path_norm}'. Encerrando esta instância.")
                    break
            except AttributeError:  # self.file_log_handle.name não existe
                logging.warning(
                    f"[{thread_name}] Tab '{self.nome}': self.file_log_handle tornou-se inválido (sem 'name') DENTRO DO LOOP para '{caminho_log_designado_para_esta_thread}'. Encerrando.")
                break
            except Exception as e_check_loop_consistency:
                logging.error(
                    f"[{thread_name}] Tab '{self.nome}': Erro ao verificar consistência do handle no loop para '{caminho_log_designado_para_esta_thread}': {e_check_loop_consistency}. Encerrando.")
                break

            try:
                linha = self.file_log_handle.readline()
                if linha:
                    # Decodificar explicitamente se necessário, embora o open já deva ter lidado com 'latin-1'
                    # linha_decoded = linha.decode('latin-1', errors='replace') if isinstance(linha, bytes) else linha
                    linha_strip = linha.strip()  # Usar a linha já decodificada pelo open()

                    # Aplicar filtro
                    filtro_atual = self.filtro_var.get().strip().lower()
                    if not filtro_atual or filtro_atual in linha.lower():  # linha original para filtro
                        self.append_text_to_log_area(linha)  # Adiciona a linha original (com \n)

                    logging.debug(
                        f"[{thread_name}] Tab '{self.nome}': LIDO de '{caminho_log_designado_para_esta_thread}': repr='{repr(linha)}', strip='{linha_strip}', aguardando_winner={aguardando_winner}")

                    # Detecção de fim de votação
                    if vote_pattern_re and vote_pattern_re.search(linha_strip):  # Usar linha_strip para regex
                        if not aguardando_winner:
                            logging.info(
                                f"[{thread_name}] Tab '{self.nome}': FIM DE VOTAÇÃO detectado em '{caminho_log_designado_para_esta_thread}'. Linha: '{linha_strip}'. Set aguardando_winner=True.")
                        else:  # Já estava aguardando, isso pode ser um log repetido ou um problema
                            logging.warning(
                                f"[{thread_name}] Tab '{self.nome}': FIM DE VOTAÇÃO detectado NOVAMENTE em '{caminho_log_designado_para_esta_thread}' (aguardando_winner já era True). Linha: '{linha_strip}'.")
                        aguardando_winner = True
                        self.app.set_status_from_thread(f"'{self.nome}': Fim da votação. Aguardando vencedor...")

                    # Detecção de vencedor (somente se estivermos aguardando um)
                    if winner_pattern_re and aguardando_winner:
                        logging.debug(
                            f"[{thread_name}] Tab '{self.nome}': AGUARDANDO WINNER é TRUE. Testando linha para Winner: '{linha_strip}'")
                        match = winner_pattern_re.search(linha_strip)  # Usar linha_strip para regex
                        if match:
                            try:
                                indice_str = match.group(1)  # Pega o primeiro grupo de captura
                                indice_vencedor = int(indice_str)
                                logging.info(
                                    f"[{thread_name}] Tab '{self.nome}': VENCEDOR detectado (aguardando_winner=True). Índice: {indice_vencedor}. Linha: '{linha_strip}'")
                                self.app.set_status_from_thread(
                                    f"'{self.nome}': Vencedor índice {indice_vencedor}. Processando...")

                                # Processar a troca de mapa na thread principal da GUI
                                if self.app.root.winfo_exists():
                                    self.app.root.after(0, self.processar_troca_mapa_logica, indice_vencedor)

                                logging.debug(
                                    f"[{thread_name}] Tab '{self.nome}': Winner processado. RESETANDO aguardando_winner para False.")
                                aguardando_winner = False  # Resetar após processar
                            except IndexError:
                                logging.error(
                                    f"[{thread_name}] Tab '{self.nome}': Padrão de vencedor '{winner_pattern_str}' casou em '{linha_strip}', mas falta grupo de captura (group 1).")
                                self.append_text_to_log_area(
                                    f"ERRO: Padrão de vencedor '{winner_pattern_str}' não tem grupo de captura (verifique Opções Votemap).\n")
                                aguardando_winner = False  # Resetar para evitar loops de erro
                            except ValueError:
                                logging.error(
                                    f"[{thread_name}] Tab '{self.nome}': Padrão vencedor capturou '{indice_str}' em '{linha_strip}', que não é um número de índice válido.")
                                self.append_text_to_log_area(
                                    f"ERRO: Vencedor capturado '{indice_str}' não é um número (verifique Opções Votemap e logs).\n")
                                aguardando_winner = False  # Resetar
                            except Exception as e_proc_winner_inesperado:
                                logging.error(
                                    f"[{thread_name}] Tab '{self.nome}': Erro inesperado ao processar vencedor: {e_proc_winner_inesperado}",
                                    exc_info=True)
                                aguardando_winner = False  # Resetar
                        # Se não deu match, continua aguardando winner na proxima linha
                    elif winner_pattern_re and winner_pattern_re.search(linha_strip) and not aguardando_winner:
                        # Encontrou padrão de vencedor, mas não estávamos esperando (ex: log antigo)
                        logging.info(
                            f"[{thread_name}] Tab '{self.nome}': Padrão de vencedor APARECEU na linha '{linha_strip}', MAS aguardando_winner era FALSO. Ignorando.")

                else:  # Linha vazia, significa fim do arquivo por enquanto
                    if self._stop_event.wait(0.2): break  # Pausa curta antes de tentar ler de novo

            except UnicodeDecodeError as ude_loop:  # Pode acontecer se o encoding mudar no meio ou caractere inválido
                logging.warning(
                    f"[{thread_name}] Tab '{self.nome}': Erro de decodificação Unicode ao ler log {caminho_log_designado_para_esta_thread}: {ude_loop}. Linha ignorada.")
                # self.append_text_to_log_area(f"AVISO: Caractere inválido no log, linha ignorada.\n")
            except ValueError as ve_loop:  # Ex: "I/O operation on closed file"
                if "closed file" in str(ve_loop).lower():
                    logging.warning(
                        f"[{thread_name}] Tab '{self.nome}': Tentativa de I/O em arquivo fechado ({caminho_log_designado_para_esta_thread}). Encerrando thread de acompanhamento.")
                    break  # Sai do loop while
                else:  # Outro ValueError
                    logging.error(
                        f"[{thread_name}] Tab '{self.nome}': Erro de ValueError ao acompanhar log {caminho_log_designado_para_esta_thread}: {ve_loop}",
                        exc_info=True)
                    break  # Sai do loop por segurança
            except Exception as e_tail_loop_inesperado:
                if not self._stop_event.is_set():  # Só loga se não for uma parada intencional
                    logging.error(
                        f"[{thread_name}] Tab '{self.nome}': Erro INESPERADO ao acompanhar log {caminho_log_designado_para_esta_thread}: {e_tail_loop_inesperado}",
                        exc_info=True)
                    self.append_text_to_log_area(
                        f"ERRO GRAVE ao ler log: {e_tail_loop_inesperado}\nVerifique o Log do Sistema do Patch.\n")
                    self.app.set_status_from_thread(f"'{self.nome}': Erro na leitura do log. Ver Log do Sistema.")
                break  # Sai do loop por segurança

        logging.info(
            f"[{thread_name}] Tab '{self.nome}': Acompanhamento de '{caminho_log_designado_para_esta_thread}' encerrado. Estado final aguardando_winner: {aguardando_winner}")
        # Não fechar self.file_log_handle aqui, a thread monitorar_log_continuamente_worker é responsável por isso.

    def processar_troca_mapa_logica(self, indice_vencedor):
        """Lógica para processar a troca de mapa, chamada pela GUI thread."""
        logging.info(f"Tab '{self.nome}': Processando troca de mapa para o índice: {indice_vencedor}")

        arquivo_json_val = self.arquivo_json.get()
        arquivo_json_votemap_val = self.arquivo_json_votemap.get()

        if not arquivo_json_val or not arquivo_json_votemap_val:
            msg = f"Arquivos JSON de servidor ({arquivo_json_val}) ou votemap ({arquivo_json_votemap_val}) não configurados para '{self.nome}'."
            self.append_text_to_log_area(f"ERRO: {msg}\n");
            logging.error(msg)
            self.app.set_status_from_thread(f"'{self.nome}': Erro - JSONs não configurados.")
            return

        try:
            with open(arquivo_json_votemap_val, 'r', encoding='utf-8') as f_vm:
                votemap_data = json.load(f_vm)
        except FileNotFoundError:
            msg = f"Arquivo votemap.json ('{arquivo_json_votemap_val}') não encontrado para '{self.nome}'."
            self.append_text_to_log_area(f"ERRO: {msg}\n");
            logging.error(msg)
            self.app.set_status_from_thread(f"'{self.nome}': Erro - votemap.json não encontrado.")
            return
        except json.JSONDecodeError as e_json_vm:
            msg = f"Erro ao decodificar votemap.json ('{arquivo_json_votemap_val}') para '{self.nome}': {e_json_vm}"
            self.append_text_to_log_area(f"ERRO: {msg}\n");
            logging.error(msg)
            self.app.set_status_from_thread(f"'{self.nome}': Erro - votemap.json inválido.")
            return

        map_list = votemap_data.get("list", [])
        if not map_list:
            msg = f"Lista de mapas ('list') vazia ou não encontrada no votemap.json ('{arquivo_json_votemap_val}') para '{self.nome}'."
            self.append_text_to_log_area(f"AVISO: {msg}\n");
            logging.warning(msg)
            self.app.set_status_from_thread(f"'{self.nome}': Aviso - Lista de mapas vazia.")
            return

        novo_scenario_id = None
        nome_mapa_log = "N/A"
        # O índice 0 no log do Arma geralmente significa "Random" ou a primeira opção da lista que pode ser "Random"
        if indice_vencedor == 0:  # Assumindo que o primeiro item (índice 0 da lista de mapas) é o "random" ou um placeholder.
            if len(map_list) > 1:  # Precisa de pelo menos mais uma opção além do "random"
                # Escolhe um mapa aleatório da lista, EXCLUINDO o primeiro item (índice 0)
                indice_selecionado_random_na_lista = random.randint(1, len(map_list) - 1)
                novo_scenario_id = map_list[indice_selecionado_random_na_lista]
                nome_mapa_log = os.path.basename(str(novo_scenario_id)).replace(".ArmaReforgerLink", "")
                self.append_text_to_log_area(
                    f"VOTO ALEATÓRIO: Selecionado mapa '{nome_mapa_log}' (índice real na lista: {indice_selecionado_random_na_lista}).\n")
                logging.info(
                    f"Tab '{self.nome}': Seleção aleatória: {novo_scenario_id} (índice da lista {indice_selecionado_random_na_lista})")
            else:
                msg = f"Voto aleatório (índice 0), mas não há mapas suficientes na lista de votemap.json (apenas {len(map_list)} item(ns)) para '{self.nome}'."
                self.append_text_to_log_area(f"AVISO: {msg}\n");
                logging.warning(msg)
                self.app.set_status_from_thread(f"'{self.nome}': Aviso - Poucos mapas para aleatório.")
                return
        elif 0 < indice_vencedor < len(
                map_list):  # Voto direto para um mapa da lista (excluindo o primeiro se for random)
            novo_scenario_id = map_list[indice_vencedor]
            nome_mapa_log = os.path.basename(str(novo_scenario_id)).replace(".ArmaReforgerLink", "")
            self.append_text_to_log_area(
                f"MAPA VENCEDOR: '{nome_mapa_log}' (índice {indice_vencedor} da lista de votemap.json).\n")
            logging.info(f"Tab '{self.nome}': Mapa vencedor selecionado: {novo_scenario_id}")
        else:  # Índice inválido
            msg = f"Índice do mapa vencedor ({indice_vencedor}) inválido para a lista de mapas (tamanho {len(map_list)}) em votemap.json para '{self.nome}'."
            self.append_text_to_log_area(f"ERRO: {msg}\n");
            logging.error(msg)
            self.app.set_status_from_thread(f"'{self.nome}': Erro - Índice de mapa inválido.")
            return

        if not novo_scenario_id:
            msg = f"Não foi possível determinar o novo scenarioId para o índice {indice_vencedor} para '{self.nome}'."
            self.append_text_to_log_area(f"ERRO: {msg}\n");
            logging.error(msg)
            return

        # Atualizar o JSON de configuração do servidor
        try:
            with open(arquivo_json_val, 'r+', encoding='utf-8') as f_srv:
                server_data = json.load(f_srv)

                # Caminho comum para scenarioId: server_data["game"]["scenarioId"]
                # Adicionar verificação se "game" existe
                if "game" not in server_data:
                    server_data["game"] = {}  # Cria a chave 'game' se não existir

                server_data["game"]["scenarioId"] = novo_scenario_id

                f_srv.seek(0)  # Voltar ao início do arquivo
                json.dump(server_data, f_srv, indent=4)
                f_srv.truncate()  # Remover qualquer conteúdo antigo restante se o novo for menor

            # Atualizar a exibição do JSON na GUI
            if self.winfo_exists() and self.json_text_area_server.winfo_exists():
                self._display_json_in_widget(self.json_text_area_server, server_data)

            self.append_text_to_log_area(
                f"JSON do servidor '{os.path.basename(arquivo_json_val)}' atualizado para o mapa: {nome_mapa_log}\n")
            logging.info(f"Tab '{self.nome}': JSON do servidor atualizado com scenarioId: {novo_scenario_id}")

            # Reiniciar o servidor se auto_restart estiver habilitado e nome_servico configurado
            if self.auto_restart_var.get() and self.nome_servico.get():
                self.append_text_to_log_area("Iniciando reinício automático do servidor...\n")
                # O reinício em si deve ser em uma thread para não bloquear a GUI
                threading.Thread(
                    target=self.reiniciar_servidor_worker,  # Método worker da aba
                    args=(novo_scenario_id,),  # Passar o scenarioId para log, se necessário
                    daemon=True,
                    name=f"ServidorRestart-{self.nome}"
                ).start()
            else:
                msg_status = f"'{self.nome}': Mapa alterado para {nome_mapa_log}. Reinício manual."
                if not self.nome_servico.get() and self.auto_restart_var.get():
                    msg_status += " (Serviço não config.)"
                self.app.set_status_from_thread(msg_status)
                logging.info(f"Tab '{self.nome}': Reinício automático desabilitado ou serviço não configurado.")

        except FileNotFoundError:
            msg = f"Arquivo de config. do servidor ('{arquivo_json_val}') não encontrado para '{self.nome}' ao tentar atualizar."
            self.append_text_to_log_area(f"ERRO: {msg}\n");
            logging.error(msg)
            self.app.set_status_from_thread(f"'{self.nome}': Erro - server.json não encontrado.")
        except (KeyError, TypeError) as e_json_key:
            msg = f"Estrutura do JSON do servidor ('{arquivo_json_val}') inválida para '{self.nome}' (game -> scenarioId não encontrado ou tipo errado): {e_json_key}"
            self.append_text_to_log_area(f"ERRO: {msg}\n");
            logging.error(msg)
            self.app.set_status_from_thread(f"'{self.nome}': Erro - Estrutura server.json inválida.")
        except json.JSONDecodeError as e_json_srv:
            msg = f"Erro ao decodificar JSON do servidor ('{arquivo_json_val}') para '{self.nome}': {e_json_srv}"
            self.append_text_to_log_area(f"ERRO: {msg}\n");
            logging.error(msg)
            self.app.set_status_from_thread(f"'{self.nome}': Erro - server.json inválido.")
        except Exception as e_proc_mapa_inesperado:
            msg = f"Erro inesperado ao processar troca de mapa para '{self.nome}': {e_proc_mapa_inesperado}"
            self.append_text_to_log_area(f"ERRO: {msg}\n");
            logging.error(msg, exc_info=True)
            self.app.set_status_from_thread(f"'{self.nome}': Erro inesperado na troca de mapa.")

    def reiniciar_servidor_worker(self, scenario_id_que_causou_restart):
        """Thread worker para reiniciar o serviço do Windows."""
        if not PYWIN32_AVAILABLE:  # Dupla checagem
            self.app.show_messagebox_from_thread("error", f"'{self.nome}': Funcionalidade Indisponível",
                                                 "pywin32 é necessário para reiniciar serviços.")
            return

        nome_servico_reiniciar = self.nome_servico.get()
        if not nome_servico_reiniciar:
            self.append_text_to_log_area_threadsafe("ERRO: Nome do serviço não configurado para reinício automático.\n")
            logging.error(f"Tab '{self.nome}': Tentativa de reiniciar servidor sem nome de serviço.")
            self.app.set_status_from_thread(f"'{self.nome}': Erro - Serviço não configurado para reinício.")
            return

        logging.info(
            f"Tab '{self.nome}': Iniciando processo de reinício do serviço '{nome_servico_reiniciar}' em background.")
        self.app.set_status_from_thread(f"'{self.nome}': Reiniciando {nome_servico_reiniciar}...")

        # Chamar a lógica de reinício que contém os comandos 'sc' e delays
        success = self._executar_logica_reinicio_servico(nome_servico_reiniciar, scenario_id_que_causou_restart)

        # Após a tentativa de reinício, mostrar mensagem e atualizar status
        if self.app.root.winfo_exists():  # Verificar se a GUI ainda existe
            if success:
                self.app.show_messagebox_from_thread("info", f"'{self.nome}': Servidor Reiniciado",
                                                     f"O serviço {nome_servico_reiniciar} foi reiniciado com sucesso.")
            else:
                self.app.show_messagebox_from_thread("error", f"'{self.nome}': Falha no Reinício",
                                                     f"Ocorreu um erro ao reiniciar o serviço {nome_servico_reiniciar}.\nVerifique os logs.")
            self.update_service_status_display()  # Atualiza o status do serviço na GUI após a tentativa

    def _executar_logica_reinicio_servico(self, nome_servico_a_gerenciar, scenario_id_anterior):
        """Contém a lógica real de parada e início do serviço usando 'sc'."""
        stop_delay_s = self.stop_delay_var.get()
        start_delay_s = self.start_delay_var.get()
        default_votemap_mission_id = self.default_mission_var.get()
        arquivo_json_servidor_path = self.arquivo_json.get()

        startupinfo = None
        if platform.system() == "Windows":
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = subprocess.SW_HIDE

        try:
            # --- Parar o Serviço ---
            self.app.set_status_from_thread(f"'{self.nome}': Parando serviço {nome_servico_a_gerenciar}...")
            self.append_text_to_log_area_threadsafe(f"Parando serviço '{nome_servico_a_gerenciar}'...\n")
            logging.info(f"Tab '{self.nome}': Tentando parar o serviço: {nome_servico_a_gerenciar}")

            status_atual = self._verificar_status_servico_win(nome_servico_a_gerenciar)
            if status_atual == "RUNNING" or status_atual == "START_PENDING":  # Se estiver rodando ou iniciando
                subprocess.run(["sc", "stop", nome_servico_a_gerenciar], check=True, shell=False,
                               startupinfo=startupinfo)
                self.append_text_to_log_area_threadsafe(f"Comando de parada enviado. Aguardando {stop_delay_s}s...\n")
                time.sleep(stop_delay_s)  # Esperar o serviço parar
                status_apos_parada = self._verificar_status_servico_win(nome_servico_a_gerenciar)
                if status_apos_parada != "STOPPED":
                    logging.warning(
                        f"Tab '{self.nome}': Serviço {nome_servico_a_gerenciar} não parou como esperado. Status: {status_apos_parada}")
                    self.append_text_to_log_area_threadsafe(
                        f"AVISO: Serviço '{nome_servico_a_gerenciar}' pode não ter parado. Status: {status_apos_parada}\n")
                else:
                    logging.info(f"Tab '{self.nome}': Serviço {nome_servico_a_gerenciar} parado com sucesso.")
            elif status_atual == "STOPPED":
                self.append_text_to_log_area_threadsafe(f"Serviço '{nome_servico_a_gerenciar}' já estava parado.\n")
                logging.info(f"Tab '{self.nome}': Serviço {nome_servico_a_gerenciar} já estava parado.")
            elif status_atual == "NOT_FOUND":
                self.append_text_to_log_area_threadsafe(
                    f"ERRO: Serviço '{nome_servico_a_gerenciar}' não encontrado para parada.\n")
                logging.error(f"Tab '{self.nome}': Serviço {nome_servico_a_gerenciar} não encontrado para parada.")
                self.app.set_status_from_thread(
                    f"'{self.nome}': Erro - Serviço '{nome_servico_a_gerenciar}' não existe.")
                return False  # Falha crítica
            else:  # Erro ou estado desconhecido
                self.append_text_to_log_area_threadsafe(
                    f"ERRO: Não foi possível determinar estado do serviço '{nome_servico_a_gerenciar}' ou estado inesperado: {status_atual}.\n")
                logging.error(
                    f"Tab '{self.nome}': Estado do serviço {nome_servico_a_gerenciar} desconhecido ou erro: {status_atual}")
                self.app.set_status_from_thread(
                    f"'{self.nome}': Erro - Estado de '{nome_servico_a_gerenciar}' desconhecido.")
                return False  # Falha

            # --- Iniciar o Serviço ---
            self.app.set_status_from_thread(f"'{self.nome}': Iniciando serviço {nome_servico_a_gerenciar}...")
            self.append_text_to_log_area_threadsafe(f"Iniciando serviço '{nome_servico_a_gerenciar}'...\n")
            logging.info(f"Tab '{self.nome}': Tentando iniciar o serviço: {nome_servico_a_gerenciar}")
            subprocess.run(["sc", "start", nome_servico_a_gerenciar], check=True, shell=False, startupinfo=startupinfo)
            self.append_text_to_log_area_threadsafe(
                f"Comando de início enviado. Aguardando {start_delay_s}s para estabilizar...\n")
            self.app.set_status_from_thread(
                f"'{self.nome}': Aguardando {nome_servico_a_gerenciar} iniciar ({start_delay_s}s)...")
            time.sleep(start_delay_s)  # Esperar o servidor iniciar

            status_apos_inicio = self._verificar_status_servico_win(nome_servico_a_gerenciar)
            if status_apos_inicio != "RUNNING":
                logging.error(
                    f"Tab '{self.nome}': Serviço {nome_servico_a_gerenciar} falhou ao iniciar. Status: {status_apos_inicio}")
                self.append_text_to_log_area_threadsafe(
                    f"ERRO: Serviço '{nome_servico_a_gerenciar}' falhou ao iniciar ou demorando. Status: {status_apos_inicio}\n")
                self.app.set_status_from_thread(
                    f"'{self.nome}': Erro - {nome_servico_a_gerenciar} não iniciou. Status: {status_apos_inicio}")
                return False  # Falha

            logging.info(f"Tab '{self.nome}': Serviço {nome_servico_a_gerenciar} iniciado com sucesso.")
            self.update_service_status_display()  # Atualiza o label na GUI via after()

            # --- Restaurar JSON do Servidor para o Mapa de Votação Padrão ---
            self.append_text_to_log_area_threadsafe("Restaurando JSON do servidor para o mapa de votação padrão...\n")
            if not default_votemap_mission_id:
                self.append_text_to_log_area_threadsafe(
                    "AVISO: Missão padrão de votemap não definida. O servidor pode não iniciar a votação.\n")
                logging.warning(f"Tab '{self.nome}': Missão padrão de votemap não definida em Opções.")
            elif not arquivo_json_servidor_path or not os.path.exists(arquivo_json_servidor_path):
                msg = f"Arquivo JSON do servidor ({arquivo_json_servidor_path}) não encontrado para restaurar votemap para '{self.nome}'."
                self.append_text_to_log_area_threadsafe(f"ERRO: {msg}\n");
                logging.error(msg)
                self.app.set_status_from_thread(f"'{self.nome}': Erro - server.json não encontrado para reset.")
                return False  # Considerar falha se não puder resetar o mapa
            else:  # Tentar restaurar
                with open(arquivo_json_servidor_path, 'r+', encoding='utf-8') as f_srv_reset:
                    server_data_reset = json.load(f_srv_reset)
                    if "game" not in server_data_reset: server_data_reset["game"] = {}
                    server_data_reset["game"]["scenarioId"] = default_votemap_mission_id
                    f_srv_reset.seek(0)
                    json.dump(server_data_reset, f_srv_reset, indent=4)
                    f_srv_reset.truncate()

                # Atualizar a exibição do JSON na GUI (se a aba ainda existir)
                if self.winfo_exists() and self.json_text_area_server.winfo_exists():
                    self.app.root.after(0, self._display_json_in_widget, self.json_text_area_server, server_data_reset)

                self.append_text_to_log_area_threadsafe(
                    f"JSON do servidor restaurado para votemap: {os.path.basename(default_votemap_mission_id)}\n")
                logging.info(
                    f"Tab '{self.nome}': JSON do servidor restaurado para scenarioId de votemap: {default_votemap_mission_id}")

            nome_mapa_anterior_log = os.path.basename(str(scenario_id_anterior)).replace(".ArmaReforgerLink", "")
            self.app.set_status_from_thread(
                f"'{self.nome}': Servidor reiniciado. Mapa anterior: {nome_mapa_anterior_log}. Próximo: Votação.")
            return True  # Sucesso no reinício

        except subprocess.CalledProcessError as e_sc:
            # Tentar decodificar a saída de erro do 'sc'
            err_output = "Nenhuma saída de erro detalhada."
            if e_sc.stderr:
                try:
                    err_output = e_sc.stderr.decode('latin-1', errors='replace')
                except:
                    pass  # Se falhar, mantém a msg padrão
            elif e_sc.stdout:
                try:
                    err_output = e_sc.stdout.decode('latin-1', errors='replace')
                except:
                    pass

            err_msg = f"Erro ao executar comando 'sc' para '{nome_servico_a_gerenciar}': {err_output.strip()}"
            self.append_text_to_log_area_threadsafe(f"ERRO: {err_msg}\n");
            logging.error(err_msg, exc_info=True)
            self.app.set_status_from_thread(f"'{self.nome}': Erro ao gerenciar serviço: {e_sc.cmd}")
            self.update_service_status_display()  # Atualiza status mesmo em erro
            return False
        except FileNotFoundError:  # sc.exe não encontrado
            self.app.show_messagebox_from_thread("error", f"'{self.nome}': Erro de Comando",
                                                 "Comando 'sc.exe' não encontrado. Verifique o PATH do sistema.")
            logging.error(f"Tab '{self.nome}': Comando 'sc.exe' não encontrado.")
            self.app.set_status_from_thread(f"'{self.nome}': Erro - sc.exe não encontrado.")
            self.update_service_status_display()
            return False
        except (json.JSONDecodeError, KeyError, TypeError) as e_json_reset:
            err_msg = f"Erro ao manipular JSON do servidor durante o reinício para '{self.nome}': {e_json_reset}"
            self.append_text_to_log_area_threadsafe(f"ERRO: {err_msg}\n");
            logging.error(err_msg, exc_info=True)
            self.app.set_status_from_thread(f"'{self.nome}': Erro - Falha ao atualizar JSON do servidor.")
            return False
        except Exception as e_reinicio_inesperado:
            err_msg = f"Erro inesperado ao reiniciar o servidor '{self.nome}': {e_reinicio_inesperado}"
            self.append_text_to_log_area_threadsafe(f"ERRO: {err_msg}\n");
            logging.error(err_msg, exc_info=True)
            self.app.set_status_from_thread(f"'{self.nome}': Erro inesperado no reinício do servidor.")
            self.update_service_status_display()
            return False

    def append_text_to_log_area(self, texto):
        """Adiciona texto à área de log da aba, garantindo que seja feito na thread da GUI se chamada de outra."""
        if not (self.winfo_exists() and self.text_area_log.winfo_exists()):
            logging.debug(f"Tab '{self.nome}': Tentativa de adicionar texto ao log, mas widget não existe.")
            return

        try:
            # Se estivermos na thread da GUI, podemos modificar diretamente.
            # Caso contrário, usamos 'after' para enfileirar a modificação.
            # No entanto, para simplificar e garantir segurança, usar 'after' sempre é mais robusto.
            # A performance é geralmente aceitável para logs.
            self.app.root.after(0, self._append_text_to_log_area_gui_thread, texto)
        except Exception as e:  # Captura genérica se a root não existir mais, etc.
            logging.warning(f"Tab '{self.nome}': Exceção ao tentar agendar append_text_to_log_area: {e}")

    def _append_text_to_log_area_gui_thread(self, texto):
        """Executado pela thread da GUI para realmente modificar o widget ScrolledText."""
        if not (self.winfo_exists() and self.text_area_log.winfo_exists()):
            return
        try:
            current_state = self.text_area_log.cget("state")
            self.text_area_log.configure(state='normal')
            self.text_area_log.insert('end', texto)
            if self.auto_scroll_log_var.get():
                self.text_area_log.yview_moveto(1.0)
            self.text_area_log.configure(state=current_state)
        except tk.TclError as e_tcl:  # Comum se a GUI estiver fechando
            logging.debug(
                f"Tab '{self.nome}': TclError em _append_text_to_log_area_gui_thread (GUI fechando?): {e_tcl}")
        except Exception as e_append:
            logging.error(f"Tab '{self.nome}': Erro em _append_text_to_log_area_gui_thread: {e_append}", exc_info=True)

    def append_text_to_log_area_threadsafe(self, texto):  # Alias para clareza
        self.append_text_to_log_area(texto)

    def limpar_tela_log(self):
        if self.text_area_log.winfo_exists():
            self.text_area_log.configure(state='normal')
            self.text_area_log.delete('1.0', 'end')
            self.text_area_log.configure(state='disabled')
            self.app.set_status_from_thread(f"Tela de logs de '{self.nome}' limpa.")
            logging.info(f"Tab '{self.nome}': Tela de logs limpa pelo usuário.")

    def toggle_pausa(self):
        self._paused = not self._paused
        if self._paused:
            self.pausar_btn.config(text="▶️ Retomar", bootstyle=SUCCESS)
            self.app.set_status_from_thread(f"Monitoramento de '{self.nome}' pausado.")
            logging.info(f"Tab '{self.nome}': Monitoramento de logs pausado.")
        else:
            self.pausar_btn.config(text="⏸️ Pausar", bootstyle=WARNING)
            self.app.set_status_from_thread(f"Monitoramento de '{self.nome}' retomado.")
            logging.info(f"Tab '{self.nome}': Monitoramento de logs retomado.")

    # --- Métodos de Busca no Log (adaptados para self.text_area_log) ---
    def _toggle_log_search_bar(self, event=None, force_hide=False, force_show=False):
        if force_hide or (self.search_log_frame_visible and not force_show):
            if self.search_log_frame.winfo_ismapped():
                self.search_log_frame_visible = False
                self.search_log_frame.pack_forget()
                if self.text_area_log.winfo_exists(): self.text_area_log.focus_set()
                self.text_area_log.tag_remove("search_match", "1.0", "end")
        elif force_show or not self.search_log_frame_visible:
            if not self.search_log_frame.winfo_ismapped():
                self.search_log_frame_visible = True
                # Empacotar search_log_frame antes de text_area_log
                # O parent de search_log_frame é log_frame (definido em _create_ui_for_tab)
                # O parent de text_area_log também é log_frame
                self.search_log_frame.pack(fill='x', before=self.text_area_log, pady=(0, 2), padx=5)
                if self.log_search_entry.winfo_exists(): self.log_search_entry.focus_set()
                self.log_search_entry.select_range(0, 'end')
        self.last_search_pos = "1.0"

    def _perform_log_search_internal(self, term, start_pos, direction_forward=True, wrap=True):
        if not term or not self.text_area_log.winfo_exists():
            if self.text_area_log.winfo_exists(): self.text_area_log.tag_remove("search_match", "1.0", "end")
            return None

        self.text_area_log.tag_remove("search_match", "1.0", "end")
        count_var = tk.IntVar()
        original_state = self.text_area_log.cget("state")
        self.text_area_log.config(state="normal")
        pos = None
        if direction_forward:
            pos = self.text_area_log.search(term, start_pos, stopindex="end", count=count_var, nocase=True)
            if not pos and wrap and start_pos != "1.0":
                pos = self.text_area_log.search(term, "1.0", stopindex=start_pos, count=count_var, nocase=True)
        else:  # Backwards
            pos = self.text_area_log.search(term, start_pos, stopindex="1.0", count=count_var, nocase=True,
                                            backwards=True)
            if not pos and wrap and start_pos != "end":
                pos = self.text_area_log.search(term, "end", stopindex=start_pos, count=count_var, nocase=True,
                                                backwards=True)

        if pos:
            end_pos = f"{pos}+{count_var.get()}c"
            self.text_area_log.tag_add("search_match", pos, end_pos)
            self.text_area_log.tag_config("search_match", background="yellow", foreground="black")  # Cores de destaque
            self.text_area_log.see(pos)
            self.text_area_log.config(state=original_state)
            return end_pos if direction_forward else pos
        else:
            self.text_area_log.config(state=original_state)
            self.app.set_status_from_thread(f"'{term}' não encontrado em '{self.nome}'.")
            return None

    def _search_log_next(self, event=None):
        term = self.log_search_var.get()
        if not term: return

        current_match_ranges = self.text_area_log.tag_ranges("search_match")
        start_from = self.last_search_pos
        if current_match_ranges:  # Se há uma seleção, começar depois dela
            start_from = current_match_ranges[1]

        next_start_pos = self._perform_log_search_internal(term, start_from, direction_forward=True)
        if next_start_pos:
            self.last_search_pos = next_start_pos
        else:  # Tentar do início (wrap)
            next_start_pos_wrapped = self._perform_log_search_internal(term, "1.0", direction_forward=True, wrap=False)
            if next_start_pos_wrapped:
                self.last_search_pos = next_start_pos_wrapped
            else:  # Realmente não encontrou
                self.last_search_pos = "1.0"  # Reset

    def _search_log_prev(self, event=None):
        term = self.log_search_var.get()
        if not term: return

        current_match_ranges = self.text_area_log.tag_ranges("search_match")
        start_from = self.last_search_pos
        if current_match_ranges:  # Se há uma seleção, começar antes dela
            start_from = current_match_ranges[0]

        new_match_start_pos = self._perform_log_search_internal(term, start_from, direction_forward=False)
        if new_match_start_pos:
            self.last_search_pos = new_match_start_pos
        else:  # Tentar do fim (wrap)
            new_match_start_pos_wrapped = self._perform_log_search_internal(term, "end", direction_forward=False,
                                                                            wrap=False)
            if new_match_start_pos_wrapped:
                self.last_search_pos = new_match_start_pos_wrapped
            else:  # Realmente não encontrou
                self.last_search_pos = "end"  # Reset


# ############################################################################
# # Classe LogViewerApp - Aplicação Principal
# ############################################################################
class LogViewerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Predadores Votemap Patch - Multi-Servidor")
        # Ajustar tamanho inicial da janela principal se necessário
        self.root.geometry("1000x750")  # Altura um pouco menor que antes

        self.style = ttk.Style()
        # Tenta carregar tema salvo, senão usa 'darkly'
        self.config_file = "votemap_config_multi.json"
        self.config = self._load_app_config_from_file()  # Método para carregar config da app

        try:
            self.style.theme_use(self.config.get("theme", "darkly"))
        except tk.TclError:
            logging.warning(f"Tema '{self.config.get('theme')}' não encontrado. Usando 'litera'.")
            self.style.theme_use("litera")  # Fallback seguro
            self.config["theme"] = "litera"

        self.servidores = []  # Lista de instâncias ServidorTab
        self.config_changed = False  # Flag para indicar se algo foi alterado
        self._app_stop_event = threading.Event()  # Evento para parar threads da app (ex: log do sistema)

        self.create_menu()

        # Notebook principal para as abas dos servidores e log do sistema
        self.main_notebook = ttk.Notebook(self.root)
        self.main_notebook.pack(fill='both', expand=True, padx=5, pady=5)
        self.main_notebook.bind("<<NotebookTabChanged>>", self.on_tab_changed)

        # Aba para Log do Sistema (do próprio Patch)
        self.system_log_frame = ttk.Frame(self.main_notebook)
        self.main_notebook.add(self.system_log_frame, text="Log do Sistema (Patch)")
        self.system_log_text_area = ScrolledText(self.system_log_frame, wrap='word', height=10, state='disabled')
        self.system_log_text_area.pack(fill='both', expand=True, padx=5, pady=5)

        # Inicializar servidores salvos (após criar a UI básica)
        self.inicializar_servidores_das_configuracoes()

        self.create_status_bar()
        self.set_application_icon()

        # Atualizar Log do Sistema periodicamente
        self._system_log_update_error_count = 0  # Para evitar spam de logs de erro
        self.atualizar_log_sistema_periodicamente()

        # Bind Escape global para fechar a barra de busca da aba ativa
        self.root.bind_all("<Escape>", self.handle_escape_key, add="+")

        if not PYWIN32_AVAILABLE and platform.system() == "Windows":
            self.show_messagebox_from_thread("warning", "pywin32 Ausente",
                                             "A biblioteca 'pywin32' não foi encontrada.\n"
                                             "Funcionalidades de gerenciamento de serviços do Windows (iniciar/parar servidor, status) estarão desabilitadas.\n"
                                             "Instale com: pip install pywin32")

    def handle_escape_key(self, event=None):
        """Fecha a barra de busca da aba de servidor atualmente ativa."""
        current_tab_widget = self.get_current_servidor_tab_widget()
        if current_tab_widget and hasattr(current_tab_widget, '_toggle_log_search_bar'):
            if current_tab_widget.search_log_frame_visible:
                current_tab_widget._toggle_log_search_bar(force_hide=True)
                return "break"  # Impede que o Escape se propague para outros binds (como fechar a janela)
        return None

    def on_tab_changed(self, event):
        """Chamado quando uma aba do notebook principal é selecionada."""
        try:
            current_tab_widget = self.get_current_servidor_tab_widget()
            if current_tab_widget:
                self.set_status_from_thread(f"Servidor '{current_tab_widget.nome}' selecionado.")
                # Se precisar fazer algo específico ao mudar de aba, como focar um widget
                # current_tab_widget.text_area_log.focus_set() # Exemplo
            elif self.main_notebook.tab(self.main_notebook.select(), "text") == "Log do Sistema (Patch)":
                self.set_status_from_thread("Visualizando Log do Sistema do Patch.")
        except tk.TclError:
            pass  # Pode acontecer se a aba estiver sendo destruída

    def get_current_servidor_tab_widget(self):
        """Retorna a instância ServidorTab da aba atualmente selecionada, se for uma."""
        try:
            selected_tab_id = self.main_notebook.select()
            if not selected_tab_id:
                return None
            # O widget associado à aba é a própria instância ServidorTab
            widget = self.main_notebook.nametowidget(selected_tab_id)
            if isinstance(widget, ServidorTab):
                return widget
        except tk.TclError:
            return None  # Aba pode não existir mais ou não ser um widget esperado
        return None

    def inicializar_servidores_das_configuracoes(self):
        servers_config_list = self.config.get("servers", [])

        if not servers_config_list:
            # Criar um servidor padrão se não houver configuração
            logging.info("Nenhuma configuração de servidor encontrada. Adicionando um servidor padrão.")
            self.adicionar_servidor_tab("Servidor 1 (Padrão)")
        else:
            for srv_conf in servers_config_list:
                nome = srv_conf.get("nome", f"Servidor {len(self.servidores) + 1}")
                self.adicionar_servidor_tab(nome, srv_conf)

        if self.servidores:  # Seleciona a primeira aba de servidor se houver alguma
            self.main_notebook.select(self.servidores[0])

    def adicionar_servidor_tab(self, nome_sugerido=None, config_servidor=None, focus_new_tab=True):
        """Adiciona uma nova aba de servidor ao notebook."""
        if nome_sugerido is None:
            nome_sugerido = f"Servidor {len(self.servidores) + 1}"

        # Garante que o nome seja único entre as abas existentes
        nomes_existentes = [s.nome for s in self.servidores]
        nome_final = nome_sugerido
        count = 1
        while nome_final in nomes_existentes:
            nome_final = f"{nome_sugerido} ({count})"
            count += 1

        servidor_tab_frame = ServidorTab(self.main_notebook, self, nome_final, config_servidor)
        self.servidores.append(servidor_tab_frame)
        self.main_notebook.add(servidor_tab_frame, text=nome_final)
        logging.info(f"Aba de servidor '{nome_final}' adicionada.")
        if focus_new_tab:
            self.main_notebook.select(servidor_tab_frame)  # Seleciona a nova aba
        self.mark_config_changed()  # Adicionar um servidor é uma mudança de config
        return servidor_tab_frame

    def remover_servidor_atual(self):
        """Remove a aba de servidor atualmente selecionada."""
        current_tab_widget = self.get_current_servidor_tab_widget()
        if not current_tab_widget:
            self.show_messagebox_from_thread("warning", "Remover Servidor",
                                             "Nenhuma aba de servidor está selecionada ou a aba selecionada não é um servidor.")
            return

        nome_servidor_removido = current_tab_widget.nome
        if Messagebox.okcancel(f"Remover '{nome_servidor_removido}'?",
                               f"Tem certeza que deseja remover a aba do servidor '{nome_servidor_removido}'?\n"
                               "Suas configurações para este servidor serão perdidas se não salvas.",
                               parent=self.root, alert=True) == "OK":

            # Parar monitoramento e limpar recursos da aba
            current_tab_widget.stop_log_monitoring(from_tab_closure=True)

            # Remover da lista e do notebook
            self.servidores.remove(current_tab_widget)
            self.main_notebook.forget(current_tab_widget)  # Remove a aba da UI
            current_tab_widget.destroy()  # Destroi o frame e seus widgets filhos

            logging.info(f"Aba de servidor '{nome_servidor_removido}' removida.")
            self.set_status_from_thread(f"Servidor '{nome_servidor_removido}' removido.")
            self.mark_config_changed()  # Remover um servidor é uma mudança

            if not self.servidores and self.main_notebook.index(
                    "end") > 1:  # Se não houver mais servidores, mas ainda houver a aba de log do sistema
                self.main_notebook.select(self.system_log_frame)  # Seleciona log do sistema
            elif self.servidores:
                self.main_notebook.select(self.servidores[0])  # Seleciona o primeiro servidor restante

    def renomear_servidor_atual(self):
        current_tab_widget = self.get_current_servidor_tab_widget()
        if not current_tab_widget:
            self.show_messagebox_from_thread("warning", "Renomear Servidor", "Nenhuma aba de servidor selecionada.")
            return

        nome_antigo = current_tab_widget.nome
        novo_nome = simpledialog.askstring("Renomear Servidor", f"Digite o novo nome para '{nome_antigo}':",
                                           initialvalue=nome_antigo, parent=self.root)

        if novo_nome and novo_nome.strip() and novo_nome != nome_antigo:
            nomes_existentes = [s.nome for s in self.servidores if s != current_tab_widget]
            if novo_nome in nomes_existentes:
                self.show_messagebox_from_thread("error", "Nome Duplicado",
                                                 f"O nome '{novo_nome}' já está em uso por outra aba de servidor.")
                return

            current_tab_widget.nome = novo_nome
            # Atualizar o texto da aba no notebook
            # Precisamos encontrar o ID da aba pelo widget para mudar o texto
            for i, tab_id in enumerate(self.main_notebook.tabs()):
                if self.main_notebook.nametowidget(tab_id) == current_tab_widget:
                    self.main_notebook.tab(tab_id, text=novo_nome)
                    break

            logging.info(f"Servidor '{nome_antigo}' renomeado para '{novo_nome}'.")
            self.set_status_from_thread(f"Servidor '{nome_antigo}' renomeado para '{novo_nome}'.")
            self.mark_config_changed()
        elif novo_nome is not None and novo_nome != nome_antigo:  # Se o usuário apagou o nome ou não mudou
            self.show_messagebox_from_thread("warning", "Nome Inválido", "O nome do servidor não pode ser vazio.")

    def mark_config_changed(self):
        """Marca que a configuração foi alterada e atualiza o estado do botão Salvar."""
        if not self.config_changed:
            self.config_changed = True
            if hasattr(self, 'file_menu'):  # Menu pode não existir durante init muito cedo
                self.file_menu.entryconfigure("Salvar Configuração", state="normal")
            if hasattr(self, 'save_config_button_main'):  # Se houver um botão de salvar principal
                self.save_config_button_main.config(state="normal")
            logging.debug("Configuração marcada como alterada.")

    def _load_app_config_from_file(self):
        """Carrega a configuração da aplicação do arquivo JSON."""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config_data = json.load(f)
                logging.info(f"Configuração da aplicação carregada de {self.config_file}")
                return config_data
            logging.info(f"Arquivo de configuração {self.config_file} não encontrado. Usando padrões.")
            return {"theme": "darkly", "servers": []}  # Config padrão
        except json.JSONDecodeError as e:
            logging.error(f"Erro ao decodificar JSON em {self.config_file}: {e}", exc_info=True)
            Messagebox.show_error(
                f"Erro ao ler arquivo de configuração:\n{self.config_file}\n\nJSON inválido: {e}\n\nUsando configurações padrão.",
                "Erro de Configuração")
            return {"theme": "darkly", "servers": []}
        except Exception as e:
            logging.error(f"Erro desconhecido ao carregar config de {self.config_file}: {e}", exc_info=True)
            Messagebox.show_error(
                f"Erro desconhecido ao carregar arquivo de configuração:\n{self.config_file}\n\n{e}\n\nUsando configurações padrão.",
                "Erro de Configuração")
            return {"theme": "darkly", "servers": []}

    def _save_app_config_to_file(self):
        """Salva a configuração de todos os servidores e configurações globais da app."""
        if not self.config_changed:
            # self.set_status_from_thread("Nenhuma alteração para salvar.")
            # logging.info("Tentativa de salvar configuração, mas nenhuma alteração detectada.")
            # return # Não salvar se nada mudou - mas o usuário pode querer salvar explicitamente
            pass

        current_app_config = {"theme": self.style.theme_use(), "servers": []}

        for servidor_tab in self.servidores:
            current_app_config["servers"].append(servidor_tab.get_current_config())

        # Outras configurações globais da app podem ser adicionadas aqui
        # Ex: current_app_config["last_window_size"] = self.root.geometry()

        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(current_app_config, f, indent=4)

            self.config_changed = False  # Resetar flag após salvar
            if hasattr(self, 'file_menu'):
                self.file_menu.entryconfigure("Salvar Configuração", state="disabled")
            if hasattr(self, 'save_config_button_main'):
                self.save_config_button_main.config(state="disabled")

            self.set_status_from_thread("Configuração salva com sucesso!")
            logging.info(f"Configuração salva em {self.config_file}")
        except IOError as e_io:
            self.set_status_from_thread(f"Erro de E/S ao salvar configuração: {e_io.strerror}")
            logging.error(f"Erro de E/S ao salvar configuração: {e_io}", exc_info=True)
            self.show_messagebox_from_thread("error", "Erro ao Salvar",
                                             f"Não foi possível salvar o arquivo de configuração:\n{self.config_file}\n\n{e_io.strerror}")
        except Exception as e_save:
            self.set_status_from_thread(f"Erro desconhecido ao salvar configuração: {e_save}")
            logging.error(f"Erro desconhecido ao salvar configuração: {e_save}", exc_info=True)
            self.show_messagebox_from_thread("error", "Erro ao Salvar",
                                             f"Ocorreu um erro ao salvar a configuração:\n{e_save}")

    def load_config_from_dialog(self):
        """Permite ao usuário carregar um arquivo de configuração diferente."""
        caminho = filedialog.askopenfilename(
            defaultextension=".json",
            filetypes=[("Arquivos JSON", "*.json"), ("Todos os arquivos", "*.*")],
            title="Selecionar arquivo de configuração para carregar",
            initialdir=os.path.dirname(self.config_file) or os.getcwd()
        )
        if not caminho:
            return  # Usuário cancelou

        try:
            with open(caminho, 'r', encoding='utf-8') as f:
                loaded_config_data = json.load(f)

            # Parar todos os monitoramentos das abas atuais e removê-las
            for srv_tab in list(self.servidores):  # Iterar sobre uma cópia para poder modificar a original
                srv_tab.stop_log_monitoring(from_tab_closure=True)
                self.main_notebook.forget(srv_tab)
                srv_tab.destroy()
            self.servidores.clear()

            # Atualizar o arquivo de configuração principal da app
            self.config_file = caminho
            self.config = loaded_config_data  # Assume que o novo arquivo tem a estrutura esperada

            # Recarregar tema
            new_theme = self.config.get("theme", "darkly")
            try:
                self.style.theme_use(new_theme)
                self.config["theme"] = new_theme  # Garante que está salvo
            except tk.TclError:
                logging.warning(f"Tema '{new_theme}' do arquivo carregado não encontrado. Usando 'litera'.")
                self.style.theme_use("litera")
                self.config["theme"] = "litera"

            # Reinicializar servidores com base na nova configuração
            self.inicializar_servidores_das_configuracoes()

            self.config_changed = False  # Configuração acabou de ser carregada, então não está "modificada"
            if hasattr(self, 'file_menu'):
                self.file_menu.entryconfigure("Salvar Configuração", state="disabled")

            self.set_status_from_thread(f"Configuração carregada de {os.path.basename(caminho)}")
            logging.info(f"Configuração carregada de {caminho}")
            self.show_messagebox_from_thread("info", "Configuração Carregada",
                                             f"Configuração carregada com sucesso de:\n{caminho}")

        except json.JSONDecodeError as e_json_load:
            logging.error(f"Erro ao decodificar JSON em {caminho}: {e_json_load}", exc_info=True)
            self.show_messagebox_from_thread("error", "Erro de Configuração",
                                             f"Falha ao carregar configuração de '{os.path.basename(caminho)}':\nFormato JSON inválido.\n{e_json_load}")
        except Exception as e_load:
            logging.error(f"Erro ao carregar configuração de {caminho}: {e_load}", exc_info=True)
            self.show_messagebox_from_thread("error", "Erro de Configuração",
                                             f"Falha ao carregar configuração de '{os.path.basename(caminho)}':\n{e_load}")

    def create_menu(self):
        menubar = ttk.Menu(self.root)
        self.root.config(menu=menubar)

        # --- Menu Arquivo ---
        self.file_menu = ttk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Arquivo", menu=self.file_menu)
        self.file_menu.add_command(label="Salvar Configuração", command=self._save_app_config_to_file,
                                   state="disabled")  # Inicialmente desabilitado
        self.file_menu.add_command(label="Carregar Configuração...", command=self.load_config_from_dialog)
        self.file_menu.add_separator()
        self.file_menu.add_command(label="Sair", command=self.on_close)

        # --- Menu Servidores ---
        server_menu = ttk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Servidores", menu=server_menu)
        server_menu.add_command(label="Adicionar Novo Servidor", command=self.adicionar_servidor_tab)
        server_menu.add_command(label="Remover Servidor Atual", command=self.remover_servidor_atual)
        server_menu.add_command(label="Renomear Servidor Atual...", command=self.renomear_servidor_atual)

        # --- Menu Ferramentas ---
        tools_menu = ttk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Ferramentas", menu=tools_menu)
        tools_menu.add_command(label="Exportar Logs da Aba Atual", command=self.export_current_tab_logs)  # Adaptado
        tools_menu.add_command(label="Validar Configs da Aba Atual",
                               command=self.validate_current_tab_configs)  # Adaptado

        theme_menu = ttk.Menu(tools_menu, tearoff=0)
        tools_menu.add_cascade(label="Mudar Tema", menu=theme_menu)
        self.theme_var = tk.StringVar(value=self.style.theme_use())  # Obtém o nome do tema atual
        for theme_name in sorted(self.style.theme_names()):
            theme_menu.add_radiobutton(label=theme_name, variable=self.theme_var, command=self.trocar_tema)

        # --- Menu Ajuda ---
        help_menu = ttk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Ajuda", menu=help_menu)
        help_menu.add_command(label="Sobre", command=self.show_about)
        help_menu.add_separator()
        help_menu.add_command(label="Verificar Atualizações...", command=self.check_for_updates)

    def trocar_tema(self, event=None):  # event é None quando chamado por Radiobutton
        novo_tema = self.theme_var.get()
        try:
            self.style.theme_use(novo_tema)
            # Re-inicializa cores dos labels de caminho e status de serviço em todas as abas
            for srv_tab in self.servidores:
                srv_tab.initialize_from_config_vars()  # Isso deve re-aplicar as cores baseadas no novo tema
            self.config["theme"] = novo_tema  # Atualiza a config para salvar
            self.mark_config_changed()
            logging.info(f"Tema alterado para: {novo_tema}")
            self.set_status_from_thread(f"Tema alterado para '{novo_tema}'.")
        except tk.TclError as e_theme_tcl:
            logging.error(f"Erro TclError ao tentar trocar para o tema '{novo_tema}': {e_theme_tcl}", exc_info=True)
            self.show_messagebox_from_thread("error", "Erro de Tema",
                                             f"Não foi possível aplicar o tema '{novo_tema}'.\n{e_theme_tcl}")
            # Tentar voltar para um tema padrão seguro se a troca falhar
            try:
                self.style.theme_use("litera")
                self.theme_var.set("litera")
                self.config["theme"] = "litera"
                for srv_tab in self.servidores: srv_tab.initialize_from_config_vars()
            except:
                pass  # Ignora erro no fallback

    def export_current_tab_logs(self):
        current_tab_widget = self.get_current_servidor_tab_widget()
        if not current_tab_widget:
            if self.main_notebook.tab(self.main_notebook.select(), "text") == "Log do Sistema (Patch)":
                self._export_text_widget_content(self.system_log_text_area, "Log do Sistema do Patch")
            else:
                self.show_messagebox_from_thread("info", "Exportar Logs",
                                                 "Selecione uma aba de servidor ou a aba 'Log do Sistema'.")
            return

        self._export_text_widget_content(current_tab_widget.text_area_log, f"Logs de '{current_tab_widget.nome}'")

    def _export_text_widget_content(self, text_widget, default_filename_part):
        caminho_arquivo = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Arquivos de Texto", "*.txt"), ("Todos os arquivos", "*.*")],
            title=f"Exportar {default_filename_part}",
            initialfile=f"{default_filename_part.replace(' ', '_')}.txt"
        )
        if caminho_arquivo:
            try:
                if text_widget.winfo_exists():
                    with open(caminho_arquivo, 'w', encoding='utf-8') as f:
                        f.write(text_widget.get('1.0', 'end-1c'))  # -1c para não incluir o newline final do widget
                    self.set_status_from_thread(
                        f"{default_filename_part} exportados para: {os.path.basename(caminho_arquivo)}")
                    logging.info(f"{default_filename_part} exportados para: {caminho_arquivo}")
                    self.show_messagebox_from_thread("info", "Exportação Concluída",
                                                     f"{default_filename_part} foram exportados com sucesso para:\n{caminho_arquivo}")
                else:
                    self.show_messagebox_from_thread("error", "Erro de Exportação",
                                                     "Área de texto não encontrada ou inválida.")
            except Exception as e_export:
                self.set_status_from_thread(f"Erro ao exportar {default_filename_part}: {e_export}")
                logging.error(f"Erro ao exportar {default_filename_part} para {caminho_arquivo}: {e_export}",
                              exc_info=True)
                self.show_messagebox_from_thread("error", "Erro de Exportação",
                                                 f"Falha ao exportar {default_filename_part}:\n{e_export}")

    def validate_current_tab_configs(self):
        current_tab_widget = self.get_current_servidor_tab_widget()
        if not current_tab_widget:
            self.show_messagebox_from_thread("info", "Validar Configurações",
                                             "Selecione uma aba de servidor para validar.")
            return

        problemas = []
        cfg = current_tab_widget.get_current_config()  # Pega a config atual da aba

        if not cfg["log_folder"] or not os.path.isdir(cfg["log_folder"]):
            problemas.append(f"- Pasta de logs '{cfg['log_folder']}' não configurada ou inválida.")
        if not cfg["server_json"] or not os.path.exists(cfg["server_json"]):
            problemas.append(f"- JSON do servidor '{cfg['server_json']}' não configurado ou não encontrado.")
        if not cfg["votemap_json"] or not os.path.exists(cfg["votemap_json"]):
            problemas.append(f"- JSON do Votemap '{cfg['votemap_json']}' não configurado ou não encontrado.")

        if cfg["auto_restart"]:
            if not cfg["service_name"]:
                problemas.append("- Reinício automático habilitado, mas nenhum serviço do Windows selecionado.")
            if not PYWIN32_AVAILABLE:
                problemas.append(
                    "- Reinício automático habilitado, mas pywin32 (para controle de serviço) não está disponível.")

        if not cfg["default_mission"]:
            problemas.append(
                "- Missão padrão de votemap (ScenarioID) não definida (necessária para resetar após reinício).")

        try:
            if cfg["vote_pattern"]: re.compile(cfg["vote_pattern"])
        except re.error as e_re_vote:
            problemas.append(f"- Padrão de detecção de voto (RegEx) é inválido: {e_re_vote}")
        try:
            if cfg["winner_pattern"]: re.compile(cfg["winner_pattern"])
        except re.error as e_re_winner:
            problemas.append(f"- Padrão de detecção de vencedor (RegEx) é inválido: {e_re_winner}")

        if problemas:
            self.show_messagebox_from_thread("warning", f"Validação de '{current_tab_widget.nome}'",
                                             "Os seguintes problemas de configuração foram encontrados:\n\n" + "\n".join(
                                                 problemas))
        else:
            self.show_messagebox_from_thread("info", f"Validação de '{current_tab_widget.nome}'",
                                             "Todas as configurações essenciais para esta aba parecem estar corretas!")
        logging.info(f"Validação de config para '{current_tab_widget.nome}': {len(problemas)} problemas encontrados.")

    def create_status_bar(self):
        status_bar_frame = ttk.Frame(self.root)
        status_bar_frame.pack(side='bottom', fill='x', pady=(0, 2), padx=2)  # Adicionado padx
        ttk.Separator(status_bar_frame, orient='horizontal').pack(side='top', fill='x')

        self.status_label_var = tk.StringVar(value="Pronto.")
        self.status_label = ttk.Label(status_bar_frame, textvariable=self.status_label_var, anchor='w')
        self.status_label.pack(side='left', fill='x', expand=True, padx=5, pady=(2, 0))  # Ajustado pady

    def set_application_icon(self):  # Mantido como antes, mas loga melhor
        global ICON_PATH, ICON_FILENAME
        try:
            if not os.path.exists(ICON_PATH):
                logging.warning(
                    f"Arquivo de ícone '{ICON_PATH}' (de '{ICON_FILENAME}') não encontrado. Ícone padrão será usado.")
                # Tentar criar um ícone padrão se o Pillow estiver disponível
                try:
                    img = Image.new('RGBA', (64, 64), (0, 0, 0, 0))  # Transparente
                    draw = ImageDraw.Draw(img)
                    draw.rectangle((10, 10, 54, 54), fill='blue', outline='white', width=2)
                    draw.text((18, 20), "PVP", fill="white",
                              font=ImageFont.truetype("arial.ttf", 20) if os.path.exists("arial.ttf") else None)
                    self.app_icon_image_pil = img  # Guardar referência
                    if platform.system() == "Windows":
                        # Para .iconbitmap, precisaríamos salvar como .ico. Mais fácil usar iconphoto com PhotoImage.
                        # No Windows, iconbitmap é preferível para .ico. Se for uma imagem gerada, PhotoImage é ok.
                        # Como ICON_PATH pode não ser .ico, vamos padronizar para iconphoto por enquanto.
                        pass  # Deixar o Tkinter usar seu ícone padrão se o arquivo .ico não existir.
                    else:  # Para Linux/MacOS
                        self.app_icon_image_tk = ImageTk.PhotoImage(self.app_icon_image_pil)
                        self.root.iconphoto(True, self.app_icon_image_tk)
                    logging.info("Ícone da aplicação definido com imagem gerada (Pillow).")

                except ImportError:
                    logging.warning("Pillow (PIL) não instalado. Não foi possível gerar ícone padrão.")
                except Exception as e_gen_icon:
                    logging.error(f"Erro ao gerar ícone padrão com Pillow: {e_gen_icon}")
                return  # Retorna se o arquivo original não existe

            # Se ICON_PATH existe
            if platform.system() == "Windows":
                self.root.iconbitmap(ICON_PATH)  # Ideal para .ico no Windows
                logging.info(f"Ícone da aplicação (Windows) definido de: {ICON_PATH}")
            else:  # Para outros sistemas (Linux, macOS)
                try:
                    img_pil = Image.open(ICON_PATH)
                    # Converter para RGBA para suportar transparência, se houver
                    img_pil_rgba = img_pil.convert("RGBA")
                    # PhotoImage do Tkinter não lida bem com todos os formatos PIL diretamente,
                    # especialmente sem o módulo ImageTk. Mas para .ico, pode funcionar.
                    # Para maior compatibilidade, usar ImageTk (requer `pip install Pillow`)
                    from PIL import ImageTk  # Importar aqui para não ser dependência dura
                    self.app_icon_image_tk = ImageTk.PhotoImage(img_pil_rgba)
                    self.root.iconphoto(True, self.app_icon_image_tk)
                    logging.info(f"Ícone da aplicação (não-Windows, via Pillow/ImageTk) definido de: {ICON_PATH}")
                except ImportError:  # Pillow ou ImageTk não disponível
                    logging.warning(
                        "Pillow (PIL) ou ImageTk não instalado. Tentando PhotoImage direto (pode falhar para .ico).")
                    try:  # Tentar com PhotoImage diretamente (menos robusto para .ico)
                        self.app_icon_image_tk = tk.PhotoImage(file=ICON_PATH)
                        self.root.iconphoto(True, self.app_icon_image_tk)
                        logging.info(f"Ícone da aplicação (não-Windows, PhotoImage direto) definido de: {ICON_PATH}")
                    except tk.TclError as e_tk_icon:
                        logging.error(
                            f"Erro TclError ao definir ícone (PhotoImage direto) com '{ICON_PATH}': {e_tk_icon}")
                except Exception as e_pil_icon:
                    logging.warning(
                        f"Falha ao carregar '{ICON_PATH}' com Pillow/ImageTk: {e_pil_icon}. PhotoImage pode não suportar .ico.")
        except tk.TclError as e_icon_main_tcl:
            logging.error(
                f"Erro TclError GERAL ao definir ícone da aplicação: {e_icon_main_tcl}. Ícone padrão do Tk será usado.")
        except Exception as e_icon_main_gen:
            logging.error(f"Erro GERAL ao definir ícone da aplicação: {e_icon_main_gen}", exc_info=True)

    def _create_tray_image(self):  # Mantido como antes
        global ICON_PATH
        try:
            if os.path.exists(ICON_PATH):
                logging.info(f"Carregando ícone da bandeja de: {ICON_PATH}")
                return Image.open(ICON_PATH)
            else:
                logging.warning(f"Arquivo de ícone da bandeja '{ICON_PATH}' não encontrado. Desenhando um padrão.")
        except ImportError:  # Pillow não instalado
            logging.warning(
                "Pillow (PIL) não está instalado. Não é possível carregar o ícone da bandeja do arquivo. Desenhando padrão.")
        except Exception as e_load_tray:
            logging.error(f"Erro ao carregar ícone da bandeja de '{ICON_PATH}': {e_load_tray}. Desenhando um padrão.")

        # Desenhar imagem padrão se o arquivo não existir ou Pillow não estiver disponível
        width, height = 64, 64
        image = Image.new('RGBA', (width, height), (0, 0, 0, 0))  # Transparente
        draw = ImageDraw.Draw(image)
        # Desenho simples como placeholder
        draw.rectangle((10, 10, width - 10, height - 10), outline='blue', fill='lightblue')
        draw.text((width // 2 - 10, height // 2 - 10), "P", fill="blue")  # Letra "P"
        return image

    def atualizar_log_sistema_periodicamente(self):
        try:
            if not self.root.winfo_exists() or not hasattr(self,
                                                           'system_log_text_area') or not self.system_log_text_area.winfo_exists():
                return

            # Só atualiza se a aba "Log do Sistema" estiver visível
            # Isso é opcional, mas pode economizar recursos se a aba não estiver em foco.
            # current_selected_tab_text = self.main_notebook.tab(self.main_notebook.select(), "text")
            # if current_selected_tab_text != "Log do Sistema (Patch)":
            #     if not self._app_stop_event.is_set() and self.root.winfo_exists():
            #         self.root.after(5000, self.atualizar_log_sistema_periodicamente) # Checa menos frequentemente se não visível
            #     return

            log_file_path = 'votemap_patch_multi.log'  # Usar o nome correto do arquivo de log
            if os.path.exists(log_file_path):
                with open(log_file_path, 'r', encoding='utf-8', errors='replace') as f:
                    # Ler apenas as últimas N linhas para performance, se o log for grande
                    # conteudo = "".join(f.readlines()[-500:]) # Ex: últimas 500 linhas
                    # Ou ler tudo se não for um problema
                    conteudo = f.read()

                self.system_log_text_area.configure(state='normal')
                # Otimização: verificar se o conteúdo mudou antes de deletar e reinserir tudo
                # current_content = self.system_log_text_area.get('1.0', 'end-1c')
                # if current_content != conteudo: # Só atualiza se houver mudança
                pos_atual_scroll_y, _ = self.system_log_text_area.yview()  # Pega a posição atual do scroll (0.0 a 1.0)

                self.system_log_text_area.delete('1.0', 'end')
                self.system_log_text_area.insert('end', conteudo)

                # Manter scroll no fundo apenas se já estava no fundo
                if pos_atual_scroll_y >= 0.99:  # Se estava perto do fim (considerando float precision)
                    self.system_log_text_area.yview_moveto(1.0)
                else:  # Caso contrário, tenta manter a posição (pode não ser perfeito com conteúdo mudando)
                    self.system_log_text_area.yview_moveto(pos_atual_scroll_y)

                self.system_log_text_area.configure(state='disabled')
            else:  # Arquivo de log não existe
                self.system_log_text_area.configure(state='normal')
                self.system_log_text_area.delete('1.0', 'end')
                self.system_log_text_area.insert('end', f"Arquivo '{log_file_path}' não encontrado.")
                self.system_log_text_area.configure(state='disabled')
        except tk.TclError as e_tcl_syslog:
            if "invalid command name" not in str(e_tcl_syslog).lower():  # Ignora erro comum ao fechar
                logging.error(f"TclError ao atualizar log do sistema na GUI: {e_tcl_syslog}", exc_info=False)
        except Exception as e_syslog_update:
            # Limitar a frequência de logs de erro para esta função para não poluir o log principal
            if not hasattr(self, "_system_log_update_error_count") or self._system_log_update_error_count < 5:
                logging.error(f"Erro ao atualizar log do sistema na GUI: {e_syslog_update}", exc_info=False)
                self._system_log_update_error_count = getattr(self, "_system_log_update_error_count", 0) + 1

        # Continuar agendando a atualização se a app não estiver parando
        if not self._app_stop_event.is_set() and self.root.winfo_exists():
            self.root.after(3000, self.atualizar_log_sistema_periodicamente)  # Intervalo de 3 segundos

    def iniciar_selecao_servico_para_aba(self, servidor_tab_instance):
        """Chamado por ServidorTab para iniciar o processo de seleção de serviço."""
        if not PYWIN32_AVAILABLE:  # Checagem redundante, mas segura
            servidor_tab_instance.app.show_messagebox_from_thread("error", "Funcionalidade Indisponível",
                                                                  "pywin32 é necessário para listar serviços.")
            return

        progress_win, _ = self._show_progress_dialog("Serviços", "Carregando lista de serviços...")
        if not (progress_win and progress_win.winfo_exists()):
            logging.error("Falha ao criar a janela de progresso para selecionar serviço.")
            return

        if self.root.winfo_exists(): self.root.update_idletasks()
        # Passar a instância da aba para o worker, para que ele saiba qual aba atualizar
        threading.Thread(
            target=self._obter_servicos_worker,
            args=(progress_win, servidor_tab_instance),  # Passa a aba
            daemon=True,
            name=f"ServicoWMI-{servidor_tab_instance.nome}"
        ).start()

    def _obter_servicos_worker(self, progress_win, servidor_tab_instance_target):  # Recebe a aba
        if not PYWIN32_AVAILABLE:  # Redundante
            logging.warning("_obter_servicos_worker chamado mas PYWIN32_AVAILABLE é False.")
            if progress_win and progress_win.winfo_exists():
                self.root.after(0, lambda: progress_win.destroy() if progress_win.winfo_exists() else None)
            return

        initialized_com = False
        try:
            logging.debug(f"Tab '{servidor_tab_instance_target.nome}': Tentando CoInitialize.")
            pythoncom.CoInitialize()
            initialized_com = True

            wmi = win32com.client.GetObject('winmgmts:')
            services_raw = wmi.InstancesOf('Win32_Service')

            nomes_servicos_temp = []
            if services_raw:
                for s in services_raw:
                    if hasattr(s, 'Name') and s.Name and hasattr(s, 'AcceptStop') and s.AcceptStop:
                        nomes_servicos_temp.append(s.Name)

            nomes_servicos_sorted = sorted(nomes_servicos_temp)

            if self.root.winfo_exists():
                self.root.after(0, self._mostrar_dialogo_selecao_servico, nomes_servicos_sorted, progress_win,
                                servidor_tab_instance_target)  # Passa a aba
            elif progress_win and progress_win.winfo_exists():
                progress_win.destroy()  # Se a root não existe mais, fecha o progresso

        except pythoncom.com_error as e_com:
            logging.error(f"Tab '{servidor_tab_instance_target.nome}': Erro COM ao listar serviços: {e_com}",
                          exc_info=True)
            error_message = f"Erro COM ({e_com.hresult}): {e_com.strerror}"
            if hasattr(e_com, 'excepinfo') and e_com.excepinfo and len(e_com.excepinfo) > 2:
                error_message += f"\nDetalhes: {e_com.excepinfo[2]}"
            if self.root.winfo_exists():
                self.root.after(0, self._handle_erro_listar_servicos, error_message, progress_win,
                                servidor_tab_instance_target.nome)
            elif progress_win and progress_win.winfo_exists():
                progress_win.destroy()
        except Exception as e_wmi:
            logging.error(f"Tab '{servidor_tab_instance_target.nome}': Erro geral ao listar serviços WMI: {e_wmi}",
                          exc_info=True)
            if self.root.winfo_exists():
                self.root.after(0, self._handle_erro_listar_servicos, str(e_wmi), progress_win,
                                servidor_tab_instance_target.nome)
            elif progress_win and progress_win.winfo_exists():
                progress_win.destroy()
        finally:
            if initialized_com:
                try:
                    pythoncom.CoUninitialize(); logging.debug(
                        f"Tab '{servidor_tab_instance_target.nome}': CoUninitialize bem-sucedido.")
                except Exception as e_uninit:
                    logging.error(f"Tab '{servidor_tab_instance_target.nome}': Erro ao CoUninitialize: {e_uninit}",
                                  exc_info=True)

            if progress_win and progress_win.winfo_exists() and not self.root.winfo_exists():
                try:
                    progress_win.destroy()
                except:
                    pass

    def _handle_erro_listar_servicos(self, error_message, progress_win, nome_tab_origem):
        if progress_win and progress_win.winfo_exists():
            try:
                progress_win.destroy()
            except:
                pass
        if self.root.winfo_exists():
            Messagebox.show_error(f"Erro ao obter lista de serviços para '{nome_tab_origem}':\n{error_message}",
                                  "Erro WMI", parent=self.root)

    def _mostrar_dialogo_selecao_servico(self, nomes_servicos, progress_win,
                                         servidor_tab_instance_target):  # Recebe a aba
        if progress_win and progress_win.winfo_exists():
            try:
                progress_win.destroy()
            except:
                pass

        if not nomes_servicos:
            if self.root.winfo_exists():
                Messagebox.show_warning(
                    f"Nenhum serviço gerenciável encontrado (ou erro ao listar) para '{servidor_tab_instance_target.nome}'.",
                    "Seleção de Serviço", parent=self.root)
            return

        dialog = ttk.Toplevel(self.root)
        dialog.title(f"Selecionar Serviço para '{servidor_tab_instance_target.nome}'")
        # ... resto da UI do diálogo como antes ...
        # No on_confirm:
        #   service_name = ...
        #   servidor_tab_instance_target.set_selected_service(service_name) # Chama o método da aba correta
        #   dialog.destroy()

        dialog.geometry("500x400");
        dialog.transient(self.root);
        dialog.grab_set()
        dialog.protocol("WM_DELETE_WINDOW", dialog.destroy)
        ttk.Label(dialog, text=f"Escolha o serviço para '{servidor_tab_instance_target.nome}':", font="-size 10").pack(
            pady=(10, 5))
        search_frame = ttk.Frame(dialog);
        search_frame.pack(fill='x', padx=10)
        ttk.Label(search_frame, text="Buscar:").pack(side='left')
        search_var = tk.StringVar()
        search_entry = ttk.Entry(search_frame, textvariable=search_var);
        search_entry.pack(side='left', fill='x', expand=True, padx=5)

        list_frame = ttk.Frame(dialog);
        list_frame.pack(fill='both', expand=True, padx=10, pady=5)
        scrollbar = ttk.Scrollbar(list_frame);
        scrollbar.pack(side='right', fill='y')
        listbox = ttk.Treeview(list_frame, columns=("name",), show="headings", selectmode="browse")
        listbox.heading("name", text="Nome do Serviço");
        listbox.column("name", width=450)
        listbox.pack(side='left', fill='both', expand=True)
        listbox.config(yscrollcommand=scrollbar.set);
        scrollbar.config(command=listbox.yview)

        initial_selection = servidor_tab_instance_target.nome_servico.get()

        def _populate_listbox(query=""):
            listbox.delete(*listbox.get_children())  # Limpa a lista
            filter_query = query.lower().strip() if query else ""
            item_to_select_id = None
            for name in nomes_servicos:
                if name and (not filter_query or filter_query in name.lower()):
                    item_id = listbox.insert("", "end", values=(name,))
                    if name == initial_selection and not query:  # Seleciona o item atual da aba, se não estiver buscando
                        item_to_select_id = item_id
            if item_to_select_id:
                listbox.selection_set(item_to_select_id)
                listbox.see(item_to_select_id)

        search_entry.bind("<KeyRelease>", lambda e: _populate_listbox(search_var.get()))
        listbox.bind("<Double-1>", lambda e: on_confirm())  # Duplo clique confirma
        _populate_listbox()  # Popula inicialmente

        def on_confirm():
            selection = listbox.selection()
            if selection:
                selected_item_id = selection[0]
                selected_item_values = listbox.item(selected_item_id, "values")
                if selected_item_values:
                    service_name = selected_item_values[0]
                    servidor_tab_instance_target.set_selected_service(service_name)
                    dialog.destroy()
                else:  # Pouco provável
                    if dialog.winfo_exists(): Messagebox.show_warning("Falha ao obter nome do serviço.", parent=dialog)
            else:
                if dialog.winfo_exists(): Messagebox.show_warning("Nenhum serviço selecionado.", parent=dialog)

        btn_frame = ttk.Frame(dialog);
        btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text="Confirmar", command=on_confirm, bootstyle=SUCCESS).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="Cancelar", command=dialog.destroy, bootstyle=DANGER).pack(side='left', padx=5)

        self.root.update_idletasks()  # Garante que as dimensões do diálogo estão corretas
        ws, hs = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        w, h = dialog.winfo_width(), dialog.winfo_height()
        if w <= 1 or h <= 1: w, h = 500, 400  # Fallback se winfo_width/height retornarem 1
        x, y = (ws / 2) - (w / 2), (hs / 2) - (h / 2)
        dialog.geometry(f'{w}x{h}+{int(x)}+{int(y)}')
        search_entry.focus_set()
        dialog.wait_window()

    def _show_progress_dialog(self, title, message):  # Mantido como antes
        progress_win = ttk.Toplevel(self.root)
        # ... (código do diálogo de progresso)
        progress_win.title(str(title) if title else "Progresso")
        progress_win.geometry("300x100");
        progress_win.resizable(False, False)
        progress_win.transient(self.root);
        progress_win.grab_set()
        ttk.Label(progress_win, text=str(message) if message else "Carregando...", bootstyle=PRIMARY).pack(pady=10)
        pb = ttk.Progressbar(progress_win, mode='indeterminate', length=280)
        pb.pack(pady=10);
        pb.start(10)
        progress_win.update_idletasks()
        try:
            width = progress_win.winfo_width();
            height = progress_win.winfo_height()
            if width <= 1 or height <= 1: width, height = 300, 100
            x = (self.root.winfo_screenwidth() // 2) - (width // 2)
            y = (self.root.winfo_screenheight() // 2) - (height // 2)
            progress_win.geometry(f'{width}x{height}+{x}+{y}')
        except tk.TclError:
            logging.warning("TclError ao tentar centralizar _show_progress_dialog (janela destruída?).")
        return progress_win, pb

    def set_status_from_thread(self, message):  # Mantido como antes
        if self.root.winfo_exists() and hasattr(self, 'status_label_var'):
            self.root.after(0, lambda: self.status_label_var.set(str(message)[:200]))  # Limita tamanho da msg

    def show_messagebox_from_thread(self, boxtype, title, message):  # Mantido como antes
        if self.root.winfo_exists():
            # ... (código do messagebox.after)
            safe_title = str(title) if title is not None else "Notificação"
            safe_message = str(message) if message is not None else ""
            parent_win = self.root  # ou a aba atual se fizer sentido e for um Toplevel

            # Garantir que a messagebox não seja muito grande
            max_msg_len = 500
            if len(safe_message) > max_msg_len:
                safe_message = safe_message[:max_msg_len] + "...\n(Mensagem truncada)"

            if boxtype == "info":
                self.root.after(0, lambda t=safe_title, m=safe_message: Messagebox.show_info(m, t,
                                                                                             parent=parent_win) if parent_win.winfo_exists() else None)
            elif boxtype == "error":
                self.root.after(0, lambda t=safe_title, m=safe_message: Messagebox.show_error(m, t,
                                                                                              parent=parent_win) if parent_win.winfo_exists() else None)
            elif boxtype == "warning":
                self.root.after(0, lambda t=safe_title, m=safe_message: Messagebox.show_warning(m, t,
                                                                                                parent=parent_win) if parent_win.winfo_exists() else None)

    def show_about(self):  # Mantido como antes, ajustar versão
        # ... (código da janela Sobre)
        about_win = ttk.Toplevel(self.root);
        about_win.title("Sobre Predadores Votemap Patch")
        about_win.geometry("450x380");
        about_win.resizable(False, False)
        about_win.transient(self.root);
        about_win.grab_set()
        frame = ttk.Frame(about_win, padding=20);
        frame.pack(fill='both', expand=True)
        ttk.Label(frame, text="Predadores Votemap Patch", font="-size 16 -weight bold").pack(pady=(0, 10))
        ttk.Label(frame, text="Versão 2.5 Multi-Servidor", font="-size 10").pack()  # Atualizar versão
        ttk.Separator(frame).pack(fill='x', pady=10)
        desc = ("Ferramenta para monitorar logs de múltiplos servidores,\n"
                "detectar votações de mapa e automatizar a troca\n"
                "de mapas e reinício do servidor (por aba).\n\n"
                "Principais funcionalidades:\n"
                "- Múltiplas abas de servidor independentes\n"
                "- Monitoramento de logs em tempo real por servidor\n"
                "- Detecção de votação e vencedor por servidor\n"
                "- Atualização de JSON de config. por servidor\n"
                "- Reinício de serviço (Windows) por servidor\n"
                "- Interface personalizável com temas")
        ttk.Label(frame, text=desc, justify='left').pack(pady=10)
        ttk.Separator(frame).pack(fill='x', pady=10)
        ttk.Label(frame, text="Desenvolvido para a comunidade Predadores").pack()
        ttk.Button(frame, text="Fechar", command=about_win.destroy, bootstyle=PRIMARY).pack(pady=(15, 0))
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() // 2) - (about_win.winfo_width() // 2)
        y = (self.root.winfo_screenheight() // 2) - (about_win.winfo_height() // 2)
        about_win.geometry(f'+{int(x)}+{int(y)}')
        about_win.wait_window()

    def check_for_updates(self):  # Mantido como antes
        # ...
        url = "https://github.com/raphaelpqdt/PQDT_Raphael_Votemappatch/releases"  # Link para releases
        current_tab = self.get_current_servidor_tab_widget()
        log_area = current_tab.text_area_log if current_tab else self.system_log_text_area  # Log na aba atual ou do sistema

        if log_area and log_area.winfo_exists():
            log_area.configure(state='normal')
            log_area.insert('end', f"Abrindo página de atualizações: {url}\n")
            log_area.configure(state='disabled')
            if hasattr(log_area, 'yview_moveto'): log_area.yview_moveto(1.0)

        logging.info(f"Abrindo URL para verificação de atualizações: {url}")
        try:
            webbrowser.open_new_tab(url)
            self.set_status_from_thread("Página de atualizações aberta no navegador.")
        except Exception as e_web:
            logging.error(f"Erro ao abrir URL de atualizações '{url}': {e_web}", exc_info=True)
            self.show_messagebox_from_thread("error", "Erro ao Abrir Navegador",
                                             f"Não foi possível abrir o link:\n{url}\n\nErro: {e_web}")
            self.set_status_from_thread("Falha ao abrir página de atualizações.")

    # --- Métodos da Bandeja (Tray Icon) ---
    def setup_tray_icon(self):  # Adaptado para self.root
        try:
            image = self._create_tray_image()
            if image is None:
                logging.error("Não foi possível criar a imagem para o ícone da bandeja.")
                return

            menu_items = [
                pystray.MenuItem('Mostrar Predadores Votemap', self.show_from_tray, default=True),
                pystray.MenuItem('Sair', self.on_close_from_tray_menu_item)
            ]
            # Adicionar opção para pausar/retomar todos os servidores?
            # menu_items.insert(1, pystray.MenuItem('Pausar Todos', self.pause_all_servers_from_tray))
            # menu_items.insert(2, pystray.MenuItem('Retomar Todos', self.resume_all_servers_from_tray))

            self.tray_icon = pystray.Icon("predadores_votemap_patch_multi", image, "Predadores Votemap Patch - Multi",
                                          pystray.Menu(*menu_items))
            threading.Thread(target=self.tray_icon.run, daemon=True, name="TrayIconThread").start()
            logging.info("Ícone da bandeja do sistema configurado e iniciado.")
        except Exception as e_tray_setup:
            logging.error(f"Falha ao criar ícone da bandeja: {e_tray_setup}", exc_info=True)
            self.tray_icon = None  # Garante que não tentaremos usá-lo se falhar

    def show_from_tray(self, icon=None, item=None):  # Adicionado icon e item para compatibilidade
        if self.root.winfo_exists():
            self.root.after(0, self.root.deiconify)
            self.root.after(100, self.root.lift)  # Traz para frente
            self.root.after(200, self.root.focus_force)  # Força foco

    def minimize_to_tray(self, event=None):
        # Só minimiza para bandeja se o ícone estiver visível e a janela estiver sendo minimizada (estado 'iconic')
        # E se o evento for do tipo Unmap (geralmente disparado ao minimizar)
        if hasattr(self, 'tray_icon') and self.tray_icon and self.tray_icon.visible:
            if self.root.winfo_exists() and self.root.state() == 'iconic':
                # Verificar se o evento é Unmap pode ser mais preciso,
                # mas root.state() == 'iconic' geralmente cobre o caso de minimização.
                self.root.withdraw()  # Esconde a janela
                logging.info("Aplicação minimizada para a bandeja.")
                # Mostrar uma notificação (opcional)
                # self.tray_icon.notify("Minimizado para a bandeja.", "Predadores Votemap Patch")

    def on_close_from_tray_menu_item(self, icon=None, item=None):  # Adicionado icon e item
        logging.info("Comando 'Sair' do menu da bandeja recebido.")
        # Para sair da bandeja, geralmente não perguntamos, apenas fechamos.
        # Se quiser perguntar, precisaria de uma forma de mostrar o diálogo da thread da bandeja.
        self.on_close_common_logic(initiated_by_tray=True, force_close=True)

    def on_close(self, event=None):
        logging.info("WM_DELETE_WINDOW recebido. Iniciando on_close.")
        if not self.root.winfo_exists():
            logging.warning("on_close chamado, mas root não existe mais.")
            return

        if self.config_changed:
            logging.debug("Configuração alterada. Salvando automaticamente antes de fechar.")
            try:
                self._save_app_config_to_file()
                logging.info("Configurações salvas automaticamente.")
            except Exception as e_auto_save:
                logging.error(f"Erro ao salvar automaticamente as configurações: {e_auto_save}", exc_info=True)
                # Você pode decidir se quer prosseguir com o fechamento mesmo se o salvamento automático falhar,
                # ou mostrar um erro e não fechar. Por simplicidade, vamos prosseguir:
                pass

                # Prosseguir para a lógica de fechamento comum, independentemente de ter havido alterações ou não.
        # Se quiser uma confirmação genérica de "Deseja sair?", pode colocar aqui.
        # if Messagebox.okcancel("Confirmar Saída", "Deseja realmente sair?", parent=self.root, alert=True) == "OK":
        #     logging.debug("Usuário confirmou saída. Chamando on_close_common_logic.")
        #     self.on_close_common_logic()
        # else:
        #     logging.info("Saída cancelada pelo usuário.")
        # Ou, para fechar diretamente sem confirmação genérica:
        logging.debug("Chamando on_close_common_logic para fechar a aplicação.")
        self.on_close_common_logic()
        logging.info("Finalizando on_close.")

    def on_close_common_logic(self, initiated_by_tray=False, force_close=False):
        logging.info(f"Iniciando lógica comum de fechamento (bandeja={initiated_by_tray}, forcar={force_close}).")

        self._app_stop_event.set()  # Sinaliza para threads da app pararem (ex: log do sistema)

        # Parar monitoramento de todas as abas de servidor
        for srv_tab in self.servidores:
            srv_tab.stop_log_monitoring(from_tab_closure=True)

        if self.root.winfo_exists():
            self.set_status_from_thread("Encerrando...")
            if not initiated_by_tray or force_close:  # Se não for da bandeja, ou se for forçado da bandeja
                try:
                    self.root.update_idletasks()  # Processa eventos pendentes antes de fechar
                except tk.TclError:
                    pass

        if hasattr(self, 'tray_icon') and self.tray_icon:
            try:
                self.tray_icon.stop()
                logging.info("Ícone da bandeja parado.")
            except Exception as e_tray_stop:
                logging.error(f"Erro ao parar ícone da bandeja: {e_tray_stop}", exc_info=True)

        # A lógica de salvar config já foi tratada em on_close() se não for initiated_by_tray.
        # Se initiated_by_tray e force_close, geralmente não salvamos, a menos que queiramos um comportamento diferente.
        # Atualmente, se fechar pela bandeja, não salva alterações pendentes.

        logging.info(f"Aplicação encerrada (via {'bandeja' if initiated_by_tray else 'janela'}).")
        if self.root.winfo_exists():
            self.root.destroy()


# --- Função Principal e Inicialização ---
def main():
    # Não inicializar COM globalmente aqui. Cada thread que precisar (WMI)
    # fará sua própria CoInitialize/CoUninitialize.

    # Criar a janela raiz antes de instanciar LogViewerApp
    # ttk.Window já lida com o theming inicial do ttkbootstrap
    root_window = ttk.Window()

    app = LogViewerApp(root_window)

    root_window.protocol("WM_DELETE_WINDOW", app.on_close)
    # O bind <Unmap> é para minimizar para a bandeja.
    # Pode ser preciso ajustar a condição para quando ele realmente esconde a janela.
    root_window.bind("<Unmap>", app.minimize_to_tray)

    app.setup_tray_icon()  # Configura o ícone da bandeja

    try:
        root_window.mainloop()
    except KeyboardInterrupt:
        logging.info("Interrupção por teclado (Ctrl+C) recebida. Encerrando...")
        app.on_close_common_logic(initiated_by_tray=True, force_close=True)  # Força fechamento
    finally:
        logging.info("Aplicação finalizada (bloco finally do main).")


if __name__ == '__main__':
    # Logar cedo se pywin32 estiver faltando no Windows
    if not PYWIN32_AVAILABLE and platform.system() == "Windows":
        # O logging já está configurado, então podemos usar
        logging.warning("pywin32 não está instalado. Funcionalidades de serviço do Windows serão desabilitadas.")


    def handle_unhandled_thread_exception(args):
        thread_name = args.thread.name if args.thread else 'ThreadDesconhecida'
        logging.critical(f"EXCEÇÃO NÃO CAPTURADA NA THREAD '{thread_name}':",
                         exc_info=(args.exc_type, args.exc_value, args.exc_traceback))
        # Opcional: Mostrar um erro para o usuário se a GUI ainda estiver de pé
        # if root_window and root_window.winfo_exists():
        #    Messagebox.show_error(f"Erro crítico na thread {thread_name}.\nConsulte o log para detalhes.", "Erro de Thread")


    threading.excepthook = handle_unhandled_thread_exception

    main()