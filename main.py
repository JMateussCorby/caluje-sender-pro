import customtkinter as ctk
import requests
import pandas as pd
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
from selenium import webdriver
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import os
import urllib.parse
import logging
from datetime import datetime, timedelta
import re
import threading
import json
import hashlib
from tkinter import ttk
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import zipfile
import io
import winreg

# Configuração global de cores
COLORS = {
    "dark": {
        "primary": "#4361EE",
        "secondary": "#7209B7",
        "accent": "#3A86FF",
        "background": "#1E1E2E",
        "surface": "#2D2D3D",
        "text": "#FFFFFF",
        "text_secondary": "#A0A0C0",
        "success": "#4CAF50",
        "error": "#F44336",
        "warning": "#FFC107",
        "chart_line1": "#4361EE",
        "chart_line2": "#7209B7"
    },
    "light": {
        "primary": "#4361EE",
        "secondary": "#7209B7",
        "accent": "#3A86FF",
        "background": "#F8F9FA",
        "surface": "#FFFFFF",
        "text": "#2B2D42",
        "text_secondary": "#4A4A6A",
        "success": "#2A9D8F",
        "error": "#D90429",
        "warning": "#F4A261",
        "chart_line1": "#4361EE",
        "chart_line2": "#7209B7",
        "sidebar_bg": "#2B2D42",
        
    }
}

class SplashScreen(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Carregando...")
        self.geometry("800x600")
        self.configure(fg_color=COLORS["dark"]["background"])
        self.overrideredirect(True)

        try:
            image = Image.open("logo.png")
            ctk_image = ctk.CTkImage(
                light_image=image,
                dark_image=image,
                size=(200, 80))
            self.logo_label = ctk.CTkLabel(self, image=ctk_image, text="")
            self.logo_label.pack(expand=True)
        except Exception as e:
            self.logo_label = ctk.CTkLabel(self, text="Caluje Sender", 
                                        font=("Roboto", 24, "bold"), 
                                        text_color=COLORS["dark"]["accent"])
            self.logo_label.pack(expand=True)

        self.progress = ctk.CTkProgressBar(self, mode="indeterminate", 
                                         progress_color=COLORS["dark"]["accent"],
                                         height=4)
        self.progress.pack(fill="x", padx=100, pady=20)
        self.progress.start()

        self.after(2500, self.destroy)

# Configuração de logging
logging.basicConfig(
    filename='whatsapp_sender.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

class WhatsAppSender:
    def __init__(self):
        self.app = ctk.CTk()
        self.app.title("Caluje Sender")
        self.app.geometry("1200x800")
        self.app.minsize(1100, 750)
        self.current_theme = "dark"
        self.api_settings = {
            'enabled': False,
            'access_token': '',
            'phone_number_id': ''
        }
        
        try:
            self.logo_image = Image.open("logo.png")
            self.logo_photo = ctk.CTkImage(
                light_image=self.logo_image,
                dark_image=self.logo_image,
                size=(200, 80))
        except Exception as e:
            print(f"Erro ao carregar logo: {str(e)}")
            self.logo_photo = None
        
        self.verificar_edgedriver()
        
        self.quantidade_contatos = ctk.StringVar(value="0 contatos")
        self.colunas_planilha = []
        self.coluna_numero = ctk.StringVar(value="")
        self.coluna_nome = ctk.StringVar(value="")
        self.nome_atendente = ctk.StringVar()
        self.atendentes = []
        self.enviando = False
        self.progresso = ctk.IntVar(value=0)
        self.enviados = 0
        self.erros = 0
        self.stats_data = []
        self.campanhas = []
        self.usuarios = []
        self.usuario_atual = None
        self.campanha_selecionada = None
        
        self.carregar_dados()
        self.criar_tela_login()
    
    def get_colors(self):
        return COLORS["dark"] if ctk.get_appearance_mode() == "Dark" else COLORS["light"]
    
    def verificar_edgedriver(self):
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, 
                             r"Software\Microsoft\Edge\BLBeacon") as key:
                edge_version = winreg.QueryValueEx(key, "version")[0]
            major_version = edge_version.split('.')[0]
            
            driver_path = os.path.join(os.getcwd(), "msedgedriver.exe")
            if os.path.exists(driver_path):
                try:
                    service = EdgeService(executable_path=driver_path)
                    driver = webdriver.Edge(service=service)
                    driver_version = driver.capabilities['browserVersion']
                    driver.quit()
                    if driver_version.startswith(major_version):
                        return
                except:
                    pass
            self.baixar_edgedriver(major_version)
        except Exception as e:
            messagebox.showwarning("Aviso", 
                                 f"Não foi possível verificar/baixar o EdgeDriver.\nErro: {str(e)}")
            raise
    
    def carregar_dados(self):
        try:
            with open('config.json', 'r') as f:
                data = json.load(f)
                self.usuarios = data.get('usuarios', [])
                self.campanhas = data.get('campanhas', [])
                self.stats_data = data.get('stats', [])
                self.atendentes = data.get('atendentes', [])
                self.api_settings = data.get('api_settings', {
                    'enabled': False,
                    'access_token': '',
                    'phone_number_id': ''
                })
        except (FileNotFoundError, json.JSONDecodeError):
            self.usuarios = []
            self.campanhas = []
            self.stats_data = []
            self.atendentes = []
    
    def salvar_dados(self):
        data = {
            'usuarios': self.usuarios,
            'campanhas': self.campanhas,
            'stats': self.stats_data,
            'atendentes': self.atendentes,
            'api_settings': self.api_settings
        }
        with open('config.json', 'w') as f:
            json.dump(data, f)
    
    def hash_senha(self, senha):
        return hashlib.sha256(senha.encode()).hexdigest()
    
    def criar_tela_login(self):
        colors = self.get_colors()
        self.login_frame = ctk.CTkFrame(self.app, fg_color=colors["background"])
        self.login_frame.pack(fill="both", expand=True)
        
        login_container = ctk.CTkFrame(self.login_frame, fg_color=colors["surface"], 
                                     corner_radius=16, width=400, height=500)
        login_container.place(relx=0.5, rely=0.5, anchor="center")
        
        if self.logo_photo:
            ctk.CTkLabel(login_container, image=self.logo_photo, text="").pack(pady=(40, 20))
        else:
            ctk.CTkLabel(login_container, text="Caluje SENDER", 
                        font=("Roboto", 20, "bold"), 
                        text_color=colors["accent"]).pack(pady=(40, 20))
        
        self.login_username = ctk.CTkEntry(login_container, placeholder_text="Usuário", 
                                         width=300, border_color=colors["primary"])
        self.login_username.pack(pady=10)
        
        self.login_password = ctk.CTkEntry(login_container, placeholder_text="Senha", show="*",
                                         width=300, border_color=colors["primary"])
        self.login_password.pack(pady=10)
        
        login_btn = ctk.CTkButton(login_container, text="Login", command=self.fazer_login,
                                fg_color=colors["primary"], hover_color=colors["secondary"],
                                corner_radius=8)
        login_btn.pack(pady=20)
        
        ctk.CTkLabel(login_container, text="OU", text_color=colors["text_secondary"]).pack()
        
        register_btn = ctk.CTkButton(login_container, text="Criar Novo Usuário", 
                                   command=self.mostrar_tela_cadastro,
                                   fg_color="transparent", border_color=colors["primary"], 
                                   border_width=1, hover_color=colors["surface"])
        register_btn.pack(pady=20)
    
    def mostrar_tela_cadastro(self):
        colors = self.get_colors()
        self.login_frame.pack_forget()
        
        self.register_frame = ctk.CTkFrame(self.app, fg_color=colors["background"])
        self.register_frame.pack(fill="both", expand=True)
        
        register_container = ctk.CTkFrame(self.register_frame, fg_color=colors["surface"],
                                        corner_radius=16, width=400, height=500)
        register_container.place(relx=0.5, rely=0.5, anchor="center")
        
        ctk.CTkLabel(register_container, text="CRIAR USUÁRIO MESTRE", 
                    font=("Roboto", 20, "bold"), 
                    text_color=colors["accent"]).pack(pady=(40, 20))
        
        self.register_username = ctk.CTkEntry(register_container, placeholder_text="Novo usuário",
                                            width=300, border_color=colors["primary"])
        self.register_username.pack(pady=10)
        
        self.register_password = ctk.CTkEntry(register_container, placeholder_text="Senha", show="*",
                                            width=300, border_color=colors["primary"])
        self.register_password.pack(pady=10)
        
        self.register_confirm = ctk.CTkEntry(register_container, placeholder_text="Confirmar senha", show="*",
                                           width=300, border_color=colors["primary"])
        self.register_confirm.pack(pady=10)
        
        register_btn = ctk.CTkButton(register_container, text="Criar Usuário", 
                                   command=self.criar_usuario,
                                   fg_color=colors["primary"], hover_color=colors["secondary"],
                                   corner_radius=8)
        register_btn.pack(pady=20)
        
        back_btn = ctk.CTkButton(register_container, text="Voltar", command=self.voltar_para_login,
                                fg_color="transparent", border_color=colors["primary"], 
                                border_width=1, hover_color=colors["surface"])
        back_btn.pack(pady=10)
    
    def voltar_para_login(self):
        self.register_frame.pack_forget()
        self.criar_tela_login()
    
    def fazer_login(self):
        username = self.login_username.get()
        password = self.login_password.get()
        
        if not username or not password:
            messagebox.showerror("Erro", "Por favor, preencha todos os campos")
            return
        
        for usuario in self.usuarios:
            if usuario['username'] == username and usuario['password'] == self.hash_senha(password):
                self.usuario_atual = usuario
                self.login_frame.pack_forget()
                self.criar_interface()
                return
        
        messagebox.showerror("Erro", "Usuário ou senha incorretos")
    
    def criar_usuario(self):
        username = self.register_username.get()
        password = self.register_password.get()
        confirm = self.register_confirm.get()
        
        if not username or not password or not confirm:
            messagebox.showerror("Erro", "Por favor, preencha todos os campos")
            return
        
        if password != confirm:
            messagebox.showerror("Erro", "As senhas não coincidem")
            return
        
        for usuario in self.usuarios:
            if usuario['username'] == username:
                messagebox.showerror("Erro", "Usuário já existe")
                return
        
        self.usuarios.append({
            'username': username,
            'password': self.hash_senha(password),
            'is_admin': True
        })
        
        self.salvar_dados()
        messagebox.showinfo("Sucesso", "Usuário criado com sucesso!")
        self.voltar_para_login()
    
    def criar_interface(self):
        colors = self.get_colors()
        self.app.grid_columnconfigure(0, weight=1)
        self.app.grid_rowconfigure(0, weight=1)
        
        self.main_frame = ctk.CTkFrame(self.app, corner_radius=0, fg_color=colors["background"])
        self.main_frame.grid(row=0, column=0, sticky="nsew")
        self.main_frame.grid_columnconfigure(1, weight=1)
        self.main_frame.grid_rowconfigure(1, weight=1)
        
        self.sidebar = ctk.CTkFrame(self.main_frame, width=200, corner_radius=0,
                                  fg_color=colors["surface"])
        self.sidebar.grid(row=0, column=0, rowspan=2, sticky="nsew")
        self.sidebar.grid_rowconfigure(6, weight=1)
        
        if self.logo_photo:
            ctk.CTkLabel(self.sidebar, image=self.logo_photo, text="").grid(row=0, column=0, padx=20, pady=(20, 40))
        else:
            ctk.CTkLabel(self.sidebar, text="Caluje Sender", 
                        font=("Roboto", 16, "bold"), 
                        text_color=colors["accent"]).grid(row=0, column=0, padx=20, pady=(20, 40))
        
        menu_buttons = [
            ("🏠 Dashboard", self.mostrar_dashboard),
            ("📤 Envio", self.mostrar_envio),
            ("📊 Campanhas", self.mostrar_campanhas),
            ("📈 Estatísticas", self.mostrar_estatisticas),
            ("⚙️ Configurações", self.mostrar_configuracoes)
        ]
        
        for i, (text, command) in enumerate(menu_buttons, start=1):
            btn = ctk.CTkButton(self.sidebar, text=text, command=command, anchor="w",
                              fg_color="transparent", hover_color=colors["background"],
                              font=("Roboto", 14), corner_radius=0)
            btn.grid(row=i, column=0, sticky="ew", padx=10, pady=5)
        
        logout_btn = ctk.CTkButton(self.sidebar, text="🔒 Sair", command=self.fazer_logout,
                                  fg_color="transparent", hover_color=colors["background"],
                                  font=("Roboto", 14), corner_radius=0)
        logout_btn.grid(row=7, column=0, sticky="ew", padx=10, pady=(20, 10))
        
        self.header_frame = ctk.CTkFrame(self.main_frame, height=60, corner_radius=0,
                                        fg_color=colors["surface"])
        self.header_frame.grid(row=0, column=1, sticky="ew")
        
        self.page_title = ctk.CTkLabel(self.header_frame, text="Dashboard", 
                                     font=("Roboto", 18, "bold"), 
                                     text_color=colors["text"])
        self.page_title.pack(side="left", padx=20)
        
        self.theme_btn = ctk.CTkButton(self.header_frame, text="🌓", width=40, height=40,
                                      command=self.alternar_tema, fg_color="transparent",
                                      hover_color=colors["background"])
        self.theme_btn.pack(side="right", padx=20)
        
        self.content_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.content_frame.grid(row=1, column=1, sticky="nsew", padx=20, pady=20)
        
        self.mostrar_dashboard()
    
    def fazer_logout(self):
        self.main_frame.grid_forget()
        self.criar_tela_login()
    
    def mostrar_dashboard(self):
        self.limpar_conteudo()
        colors = self.get_colors()
        self.page_title.configure(text="Dashboard")
        
        ctk.CTkLabel(self.content_frame, text="Visão Geral", 
                    font=("Roboto", 16, "bold"),
                    text_color=colors["text"]).pack(anchor="w", pady=(0, 20))
        
        cards_frame = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        cards_frame.pack(fill="x", pady=(0, 20))
        
        total_mensagens = sum([stat['enviados'] for stat in self.stats_data])
        msg_card = self.criar_card(cards_frame, "📩 Mensagens Enviadas", str(total_mensagens))
        msg_card.pack(side="left", fill="both", expand=True, padx=5)
        
        total_sucesso = sum([stat['enviados'] - stat['erros'] for stat in self.stats_data])
        success_card = self.criar_card(cards_frame, "✅ Sucesso", str(total_sucesso), colors["success"])
        success_card.pack(side="left", fill="both", expand=True, padx=5)
        
        total_erros = sum([stat['erros'] for stat in self.stats_data])
        error_card = self.criar_card(cards_frame, "❌ Erros", str(total_erros), colors["error"])
        error_card.pack(side="left", fill="both", expand=True, padx=5)
        
        camp_card = self.criar_card(cards_frame, "📋 Campanhas", str(len(self.campanhas)), colors["accent"])
        camp_card.pack(side="left", fill="both", expand=True, padx=5)
        
        ctk.CTkLabel(self.content_frame, text="Atividades Recentes", 
                    font=("Roboto", 16, "bold"),
                    text_color=colors["text"]).pack(anchor="w", pady=(20, 10))
        
        fig, ax = plt.subplots(figsize=(8, 4), facecolor=colors["surface"])
        ax.set_facecolor(colors["surface"])
        
        hoje = datetime.now().date()
        datas = [hoje - timedelta(days=i) for i in range(6, -1, -1)]
        enviados = []
        erros = []
        
        for data in datas:
            total_env = 0
            total_err = 0
            for stat in self.stats_data:
                stat_date = datetime.strptime(stat['data'], "%Y-%m-%d").date()
                if stat_date == data:
                    total_env += stat['enviados']
                    total_err += stat['erros']
            enviados.append(total_env)
            erros.append(total_err)
        
        ax.plot(datas, enviados, label='Enviados', color=colors["chart_line1"], marker='o')
        ax.plot(datas, erros, label='Erros', color=colors["chart_line2"], marker='o')
        ax.legend(facecolor=colors["surface"], labelcolor=colors["text"])
        ax.set_title('Mensagens nos últimos 7 dias', color=colors["text"])
        ax.tick_params(colors=colors["text_secondary"])
        ax.spines['bottom'].set_color(colors["text_secondary"])
        ax.spines['left'].set_color(colors["text_secondary"])
        
        chart_frame = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        chart_frame.pack(fill="both", expand=True)
        
        canvas = FigureCanvasTkAgg(fig, master=chart_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)
    
    def criar_card(self, parent, titulo, valor, cor=None):
        colors = self.get_colors()
        card = ctk.CTkFrame(parent, height=120, fg_color=colors["surface"], corner_radius=12)
        ctk.CTkLabel(card, text=titulo, font=("Roboto", 14), text_color=colors["text_secondary"]).pack(pady=(15, 5))
        lbl = ctk.CTkLabel(card, text=valor, font=("Roboto", 24, "bold"))
        if cor: lbl.configure(text_color=cor)
        lbl.pack()
        return card
    
    def mostrar_envio(self):
        self.limpar_conteudo()
        colors = self.get_colors()
        self.page_title.configure(text="Envio de Mensagens")
        
        self.content_frame.grid_columnconfigure(0, weight=3)
        self.content_frame.grid_columnconfigure(1, weight=1)
        self.content_frame.grid_rowconfigure(0, weight=1)
        
        self.control_panel = ctk.CTkFrame(self.content_frame, corner_radius=12, fg_color=colors["surface"])
        self.control_panel.grid(row=0, column=0, sticky="nsew", padx=(0,10), pady=0)
        
        self.status_panel = ctk.CTkFrame(self.content_frame, corner_radius=12, fg_color=colors["surface"])
        self.status_panel.grid(row=0, column=1, sticky="nsew", pady=0)
        
        self.setup_control_panel()
        self.setup_status_panel()
    
    def setup_control_panel(self):
        colors = self.get_colors()
        self.control_panel.grid_columnconfigure(0, weight=1)
        
        # Seção de seleção de campanha
        campanha_section = ctk.CTkFrame(self.control_panel, fg_color="transparent")
        campanha_section.pack(fill="x", padx=20, pady=(20, 10))
        
        ctk.CTkLabel(campanha_section, text="SELECIONAR CAMPANHA", 
                    font=("Roboto", 12, "bold"), 
                    text_color=colors["text_secondary"]).pack(anchor="w")
        
        self.campanha_dropdown = ctk.CTkOptionMenu(campanha_section, 
                                                 values=[c['nome'] for c in self.campanhas],
                                                 command=self.selecionar_campanha,
                                                 dynamic_resizing=False,
                                                 fg_color=colors["surface"], 
                                                 button_color=colors["primary"],
                                                 button_hover_color=colors["secondary"])
        self.campanha_dropdown.pack(fill="x", pady=(5, 0))
        self.campanha_dropdown.set("Selecione uma campanha")
        
        # Seção do arquivo
        file_section = ctk.CTkFrame(self.control_panel, fg_color="transparent")
        file_section.pack(fill="x", padx=20, pady=(10, 10))
        
        ctk.CTkLabel(file_section, text="PLANILHA DE CONTATOS", 
                    font=("Roboto", 12, "bold"), 
                    text_color=colors["text_secondary"]).pack(anchor="w")
        
        self.file_entry = ctk.CTkEntry(file_section, placeholder_text="Selecione o arquivo Excel...",
                                     border_color=colors["primary"])
        self.file_entry.pack(fill="x", pady=(5, 0))
        
        self.file_btn = ctk.CTkButton(file_section, text="Buscar Arquivo", 
                                     command=self.buscar_arquivo, 
                                     fg_color=colors["primary"], 
                                     hover_color=colors["secondary"],
                                     corner_radius=8)
        self.file_btn.pack(fill="x", pady=(10, 0))
        
        # Seção de colunas
        columns_section = ctk.CTkFrame(self.control_panel, fg_color="transparent")
        columns_section.pack(fill="x", padx=20, pady=10)
        
        ctk.CTkLabel(columns_section, text="MAPEAMENTO DE COLUNAS", 
                    font=("Roboto", 12, "bold"), 
                    text_color=colors["text_secondary"]).pack(anchor="w")
        
        self.numero_dropdown = ctk.CTkOptionMenu(columns_section, variable=self.coluna_numero,
                                               values=[], dynamic_resizing=False,
                                               fg_color=colors["surface"], button_color=colors["primary"],
                                               button_hover_color=colors["secondary"])
        self.numero_dropdown.pack(fill="x", pady=(5, 0))
        self.numero_dropdown.set("Coluna do número")
        
        self.nome_dropdown = ctk.CTkOptionMenu(columns_section, variable=self.coluna_nome,
                                             values=[], dynamic_resizing=False,
                                             fg_color=colors["surface"], button_color=colors["primary"],
                                             button_hover_color=colors["secondary"])
        self.nome_dropdown.pack(fill="x", pady=(10, 0))
        self.nome_dropdown.set("Coluna do nome")
        
        # Seção da mensagem
        message_section = ctk.CTkFrame(self.control_panel, fg_color="transparent")
        message_section.pack(fill="both", expand=True, padx=20, pady=10)
        
        ctk.CTkLabel(message_section, text="MENSAGEM PERSONALIZADA", 
                    font=("Roboto", 12, "bold"), 
                    text_color=colors["text_secondary"]).pack(anchor="w")
        
        self.message_text = ctk.CTkTextbox(message_section, wrap="word", font=("Roboto", 12),
                                         border_width=1, border_color=colors["primary"],
                                         fg_color=colors["surface"])
        self.message_text.pack(fill="both", expand=True, pady=(5, 0))
        self.message_text.insert("1.0", """Olá! Boa tarde! Como vai?
Sou a ATENDENTE, atendente do Hotel Fazenda Caluje.

Segue ficha para antecipação de check-in.
Qualquer duvida fico a disposição!

LEIA ATENTAMENTE ANTES DO PREENCHIMENTO

https://docs.google.com/forms/d/15bPlXAc1Ml-s13bzAWvwEQCMOlLzm8pAkwhKKQ9Y06A/edit""")
        
        # Seção de atendente
        attendant_section = ctk.CTkFrame(self.control_panel, fg_color="transparent")
        attendant_section.pack(fill="x", padx=20, pady=(10, 20))
        
        ctk.CTkLabel(attendant_section, text="ATENDENTE", 
                    font=("Roboto", 12, "bold"), 
                    text_color=colors["text_secondary"]).pack(anchor="w")
        
        attendant_controls = ctk.CTkFrame(attendant_section, fg_color="transparent")
        attendant_controls.pack(fill="x", pady=(5, 0))
        
        self.attendant_entry = ctk.CTkEntry(attendant_controls, placeholder_text="Nome do atendente",
                                          border_color=colors["primary"])
        self.attendant_entry.pack(side="left", fill="x", expand=True, padx=(0, 5))
        
        self.add_attendant_btn = ctk.CTkButton(attendant_controls, text="+", width=40,
                                              command=self.adicionar_atendente_controle,
                                              fg_color=colors["primary"], 
                                              hover_color=colors["secondary"],
                                              corner_radius=8)
        self.add_attendant_btn.pack(side="left", padx=(0, 5))
        
        self.remove_attendant_btn = ctk.CTkButton(attendant_controls, text="-", width=40,
                                                 command=self.remover_atendente_controle,
                                                 fg_color=colors["error"],
                                                 hover_color="#B71C1C",
                                                 corner_radius=8)
        self.remove_attendant_btn.pack(side="left")
        
        self.attendant_dropdown = ctk.CTkOptionMenu(attendant_section, variable=self.nome_atendente,
                                                  values=self.atendentes, dynamic_resizing=False,
                                                  fg_color=colors["surface"], button_color=colors["primary"],
                                                  button_hover_color=colors["secondary"])
        self.attendant_dropdown.pack(fill="x", pady=(10, 0))
        self.attendant_dropdown.set("Selecione o atendente")
        
        # Botão de envio
        self.send_btn = ctk.CTkButton(self.control_panel, text="INICIAR ENVIO", 
                                     command=self.iniciar_envio, 
                                     fg_color=colors["primary"], hover_color=colors["secondary"],
                                     font=("Roboto", 14, "bold"), height=50,
                                     corner_radius=12)
        self.send_btn.pack(fill="x", padx=20, pady=(0, 20))
    
    def selecionar_campanha(self, choice):
        self.campanha_selecionada = next((c for c in self.campanhas if c['nome'] == choice), None)
        if self.campanha_selecionada:
            self.file_entry.delete(0, ctk.END)
            self.file_entry.insert(0, self.campanha_selecionada['arquivo'])
            self.carregar_planilha(self.campanha_selecionada['arquivo'])
            self.coluna_numero.set(self.campanha_selecionada['coluna_numero'])
            self.coluna_nome.set(self.campanha_selecionada['coluna_nome'])
            self.message_text.delete("1.0", ctk.END)
            self.message_text.insert("1.0", self.campanha_selecionada['mensagem'])
            self.nome_atendente.set(self.campanha_selecionada['atendente'])
    
    def carregar_planilha(self, caminho):
        try:
            df = pd.read_excel(caminho)
            self.quantidade_contatos.set(f"{len(df)} contatos")
            self.colunas_planilha = list(df.columns)
            self.numero_dropdown.configure(values=self.colunas_planilha)
            self.nome_dropdown.configure(values=self.colunas_planilha)
            self.log_action(f"Arquivo carregado: {caminho} com {len(df)} contatos")
        except Exception as e:
            self.quantidade_contatos.set("0 contatos")
            messagebox.showerror("Erro", f"Não foi possível ler o arquivo:\n{str(e)}")
            self.log_action(f"Erro ao ler arquivo: {str(e)}", error=True)
    
    def setup_status_panel(self):
        colors = self.get_colors()
        self.status_panel.grid_columnconfigure(0, weight=1)
        self.status_panel.grid_rowconfigure(1, weight=1)
        
        ctk.CTkLabel(self.status_panel, text="STATUS DO ENVIO", 
                    font=("Roboto", 14, "bold"), 
                    text_color=colors["accent"]).pack(pady=(20, 10))
        
        info_card = ctk.CTkFrame(self.status_panel, corner_radius=12, 
                               fg_color=colors["surface"])
        info_card.pack(fill="x", padx=15, pady=5)
        
        ctk.CTkLabel(info_card, text="Contatos na planilha:", 
                    font=("Roboto", 11), 
                    text_color=colors["text_secondary"]).pack(anchor="w", padx=10, pady=(10, 0))
        self.contacts_count = ctk.CTkLabel(info_card, textvariable=self.quantidade_contatos, 
                                         font=("Roboto", 16, "bold"))
        self.contacts_count.pack(anchor="w", padx=10, pady=(0, 10))
        
        progress_card = ctk.CTkFrame(self.status_panel, corner_radius=12, 
                                   fg_color=colors["surface"])
        progress_card.pack(fill="x", padx=15, pady=5)
        
        ctk.CTkLabel(progress_card, text="Progresso:", 
                    font=("Roboto", 11), 
                    text_color=colors["text_secondary"]).pack(anchor="w", padx=10, pady=(10, 0))
        
        self.progress_bar = ctk.CTkProgressBar(progress_card, orientation="horizontal",
                                             height=15, fg_color=colors["background"],
                                             progress_color=colors["primary"])
        self.progress_bar.pack(fill="x", padx=10, pady=5)
        self.progress_bar.set(0)
        
        self.progress_label = ctk.CTkLabel(progress_card, text="0%", 
                                         font=("Roboto", 12))
        self.progress_label.pack(anchor="e", padx=10, pady=(0, 10))
        
        stats_card = ctk.CTkFrame(self.status_panel, corner_radius=12, 
                                fg_color=colors["surface"])
        stats_card.pack(fill="x", padx=15, pady=5)
        
        ctk.CTkLabel(stats_card, text="Mensagens enviadas:", 
                    font=("Roboto", 11), 
                    text_color=colors["text_secondary"]).pack(anchor="w", padx=10, pady=(10, 0))
        self.sent_count = ctk.CTkLabel(stats_card, text="0", 
                                      font=("Roboto", 16, "bold"), 
                                      text_color=colors["primary"])
        self.sent_count.pack(anchor="w", padx=10, pady=(0, 5))
        
        ctk.CTkLabel(stats_card, text="Erros encontrados:", 
                    font=("Roboto", 11), 
                    text_color=colors["text_secondary"]).pack(anchor="w", padx=10, pady=(5, 0))
        self.error_count = ctk.CTkLabel(stats_card, text="0", 
                                       font=("Roboto", 16, "bold"), 
                                       text_color=colors["error"])
        self.error_count.pack(anchor="w", padx=10, pady=(0, 10))
        
        status_card = ctk.CTkFrame(self.status_panel, corner_radius=12, 
                                 fg_color=colors["surface"])
        status_card.pack(fill="x", padx=15, pady=(5, 20))
        
        ctk.CTkLabel(status_card, text="Status atual:", 
                    font=("Roboto", 11), 
                    text_color=colors["text_secondary"]).pack(anchor="w", padx=10, pady=(10, 0))
        self.status_label = ctk.CTkLabel(status_card, text="Pronto", 
                                       font=("Roboto", 14), 
                                       text_color=colors["primary"])
        self.status_label.pack(anchor="w", padx=10, pady=(0, 10))
        
        log_frame = ctk.CTkFrame(self.status_panel, corner_radius=12, 
                               fg_color=colors["surface"])
        log_frame.pack(fill="both", expand=True, padx=15, pady=(0, 20))
        
        ctk.CTkLabel(log_frame, text="ÚLTIMAS AÇÕES:", 
                    font=("Roboto", 11), 
                    text_color=colors["text_secondary"]).pack(anchor="w", padx=10, pady=10)
        
        self.log_text = ctk.CTkTextbox(log_frame, wrap="word", font=("Consolas", 10),
                                      fg_color=colors["background"], border_width=0, state="disabled")
        self.log_text.pack(fill="both", expand=True, padx=10, pady=(0, 10))
    
    def buscar_arquivo(self):
        caminho = filedialog.askopenfilename(filetypes=[["Arquivos Excel", "*.xlsx *.xls"]])
        if caminho:
            self.file_entry.delete(0, ctk.END)
            self.file_entry.insert(0, caminho)
            self.carregar_planilha(caminho)

    def formatar_numero(self, numero):
        return re.sub(r'[^0-9]', '', str(numero))

    def validar_numero(self, numero):
        numero = self.formatar_numero(numero)
        return len(numero) >= 10 and len(numero) <= 13

    def iniciar_envio(self):
        if self.enviando:
            return

        if self.campanha_selecionada:
            caminho = self.campanha_selecionada['arquivo']
            mensagem = self.campanha_selecionada['mensagem']
            usar_api = self.campanha_selecionada['metodo'] == 'API'
        else:
            caminho = self.file_entry.get()
            mensagem = self.message_text.get("1.0", ctk.END).strip()
            usar_api = self.api_settings['enabled']
        
        if not caminho or not mensagem or not self.coluna_numero.get() or not self.coluna_nome.get():
            messagebox.showwarning("Aviso", "Por favor, preencha todos os campos e selecione as colunas corretamente.")
            return
            
        if not self.nome_atendente.get():
            messagebox.showwarning("Aviso", "Por favor, selecione ou cadastre um atendente.")
            return

        self.enviando = True
        self.send_btn.configure(state="disabled", fg_color="#777777")
        self.status_label.configure(text="Preparando...")
        self.enviados = 0
        self.erros = 0
        self.sent_count.configure(text="0")
        self.error_count.configure(text="0")
        self.progress_bar.set(0)        
        self.progress_label.configure(text="0%")
        
        threading.Thread(target=self.enviar_mensagens, args=(caminho, mensagem, usar_api), daemon=True).start()

    def enviar_mensagens(self, caminho, mensagem, usar_api):
        try:
            df = pd.read_excel(caminho)
            total_contatos = len(df)

            if usar_api:
                access_token = self.api_settings['access_token']
                phone_id = self.api_settings['phone_number_id']
                
                if not access_token or not phone_id:
                    messagebox.showerror("Erro", "Configurações da API incompletas!")
                    return

                for index, row in df.iterrows():
                    if not self.enviando:
                        break

                    numero = self.formatar_numero(row[self.coluna_numero.get()])
                    nome = str(row[self.coluna_nome.get()]).strip()
                    
                    if not self.validar_numero(numero):
                        self.erros += 1
                        self.log_action(f"Número inválido: {numero} - Pular contato", warning=True)
                        self.update_counts()
                        continue

                    msg_personalizada = mensagem.replace("{nome}", nome).replace("ATENDENTE", self.nome_atendente.get())
                    
                    headers = {
                        "Authorization": f"Bearer {access_token}",
                        "Content-Type": "application/json"
                    }
                    
                    payload = {
                        "messaging_product": "whatsapp",
                        "to": f"55{numero}",
                        "type": "text",
                        "text": {"body": msg_personalizada}
                    }

                    try:
                        response = requests.post(
                            f"https://graph.facebook.com/v18.0/{phone_id}/messages",
                            headers=headers,
                            json=payload
                        )
                        response.raise_for_status()
                        self.enviados += 1
                        self.log_action(f"Mensagem enviada via API para {numero} ({nome})")
                    except Exception as e:
                        self.erros += 1
                        self.log_action(f"Erro na API ao enviar para {numero}: {str(e)}", error=True)
                    
                    self.update_counts()
                    progresso = (self.enviados + self.erros) / total_contatos
                    self.progress_bar.set(progresso)
                    self.progress_label.configure(text=f"{int(progresso * 100)}%")
                    time.sleep(0.5)
            else:
                options = EdgeOptions()
                options.add_argument("--disable-dev-shm-usage")
                options.add_argument("--no-sandbox")
                options.add_argument("--disable-notifications")
                options.add_experimental_option("excludeSwitches", ["enable-logging"])

                edge_driver_path = os.path.join(os.getcwd(), "msedgedriver.exe")
                service = EdgeService(executable_path=edge_driver_path)

                try:
                    driver = webdriver.Edge(service=service, options=options)
                    driver.maximize_window()
                except Exception as e:
                    messagebox.showerror("Erro ao iniciar o Edge", f"Erro: {e}\nVerifique se o msedgedriver.exe está na pasta do aplicativo.")
                    self.log_action(f"Erro ao iniciar Edge: {str(e)}", error=True)
                    self.resetar_interface()
                    return

                driver.get("https://web.whatsapp.com")
                self.update_status("Aguardando login no WhatsApp...")
                self.log_action("Aguardando login no WhatsApp Web...")
                
                if not messagebox.askokcancel("Atenção", "Após escanear o QR Code (se necessário), clique em OK para iniciar o envio."):
                    driver.quit()
                    self.resetar_interface()
                    return

                self.update_status("Enviando mensagens...")
                self.log_action("Iniciando envio de mensagens...")
                
                for index, row in df.iterrows():
                    if not self.enviando:
                        break

                    numero = self.formatar_numero(row[self.coluna_numero.get()])
                    nome = str(row[self.coluna_nome.get()]).strip()
                    
                    if not self.validar_numero(numero):
                        self.log_action(f"Número inválido: {numero} - Pular contato", warning=True)
                        self.erros += 1
                        self.update_counts()
                        continue

                    msg_personalizada = mensagem.replace("{nome}", nome).replace("ATENDENTE", self.nome_atendente.get())
                    msg_codificada = urllib.parse.quote(msg_personalizada)

                    try:
                        driver.get(f"https://web.whatsapp.com/send?phone=55{numero}&text={msg_codificada}")
                        
                        WebDriverWait(driver, 30).until(
                            EC.presence_of_element_located((By.XPATH, '//div[@contenteditable="true"]'))
                        )
                        
                        time.sleep(2)
                        
                        WebDriverWait(driver, 15).until(
                            EC.element_to_be_clickable((By.XPATH, "//span[@data-icon='send']"))
                        ).click()
                        
                        self.enviados += 1
                        self.log_action(f"Mensagem enviada para {numero} ({nome})")
                        
                        self.update_counts()
                        progresso = (self.enviados + self.erros) / total_contatos
                        self.progress_bar.set(progresso)
                        self.progress_label.configure(text=f"{int(progresso * 100)}%")
                        
                        time.sleep(3)

                    except Exception as e:
                        self.erros += 1
                        self.log_action(f"Erro ao enviar para {numero}: {str(e)}", error=True)
                        self.update_counts()
                        time.sleep(5)
                
                driver.quit()

            stat_entry = {
                'data': datetime.now().strftime("%Y-%m-%d"),
                'enviados': self.enviados,
                'erros': self.erros,
                'campanha': self.nome_atendente.get(),
                'metodo': 'API' if usar_api else 'Web'
            }
            self.stats_data.append(stat_entry)
            self.salvar_dados()
            
            self.log_action(f"Processo concluído - Enviadas: {self.enviados}, Erros: {self.erros}")
            messagebox.showinfo("Concluído", f"Mensagens enviadas com sucesso!\nEnviadas: {self.enviados}\nErros: {self.erros}")

        except Exception as e:
            self.log_action(f"Erro inesperado: {str(e)}", error=True)
            messagebox.showerror("Erro", f"Ocorreu um erro inesperado:\n{str(e)}")
        finally:
            self.resetar_interface()

    def update_counts(self):
        self.sent_count.configure(text=str(self.enviados))
        self.error_count.configure(text=str(self.erros))
        self.app.update()

    def update_status(self, status):
        self.status_label.configure(text=status)
        self.app.update()

    def log_action(self, message, error=False, warning=False):
        timestamp = datetime.now().strftime("%H:%M:%S")
        if error:
            tag = "[ERRO]"
            color = COLORS[self.current_theme]["error"]
        elif warning:
            tag = "[AVISO]"
            color = COLORS[self.current_theme]["warning"]
        else:
            tag = "[INFO]"
            color = COLORS[self.current_theme]["success"]
        
        log_message = f"{timestamp} {tag} {message}\n"
        
        self.log_text.configure(state="normal")
        self.log_text.insert("end", log_message)
        self.log_text.tag_add(tag, "end-1c linestart", "end-1c lineend")
        self.log_text.tag_config(tag, foreground=color)
        self.log_text.see("end")
        self.log_text.configure(state="disabled")
        self.app.update()
        
        if error:
            logging.error(message)
        elif warning:
            logging.warning(message)
        else:
            logging.info(message)

    def resetar_interface(self):
        self.enviando = False
        self.send_btn.configure(state="normal", fg_color=self.get_colors()["primary"])
        self.status_label.configure(text="Pronto")
        self.progress_bar.set(0)
        self.progress_label.configure(text="0%")

    def alternar_tema(self):
        current_theme = ctk.get_appearance_mode()
        new_theme = "Light" if current_theme == "Dark" else "Dark"
        ctk.set_appearance_mode(new_theme)
        self.current_theme = new_theme.lower()
        self.atualizar_cores_interface()

    def atualizar_cores_interface(self):
        colors = self.get_colors()
        
        self.main_frame.configure(fg_color=colors["background"])
        self.sidebar.configure(fg_color=colors["surface"])
        self.header_frame.configure(fg_color=colors["surface"])
        
        if hasattr(self, 'control_panel'):
            self.control_panel.configure(fg_color=colors["surface"])
        if hasattr(self, 'status_panel'):
            self.status_panel.configure(fg_color=colors["surface"])
        
        self.page_title.configure(text_color=colors["text"])
        
        widgets = [self.file_btn, self.send_btn, self.add_attendant_btn]
        for widget in widgets:
            widget.configure(fg_color=colors["primary"], hover_color=colors["secondary"])
        
        self.remove_attendant_btn.configure(fg_color=colors["error"], hover_color="#B71C1C")

    def limpar_conteudo(self):
        for widget in self.content_frame.winfo_children():
            widget.destroy()

    def mostrar_campanhas(self):
        self.limpar_conteudo()
        colors = self.get_colors()
        self.page_title.configure(text="Campanhas")
        
        main_frame = ctk.CTkFrame(self.content_frame, fg_color=colors["surface"], corner_radius=12)
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Cabeçalho
        header_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        header_frame.pack(fill="x", padx=10, pady=10)
        
        ctk.CTkLabel(header_frame, text="Campanhas", 
                    font=("Roboto", 16, "bold")).pack(side="left")
        
        new_btn = ctk.CTkButton(header_frame, text="+ Nova Campanha", width=120,
                               command=self.criar_nova_campanha,
                               fg_color=colors["primary"], hover_color=colors["secondary"])
        new_btn.pack(side="right")
        
        # Lista de campanhas
        scroll_frame = ctk.CTkScrollableFrame(main_frame, height=400)
        scroll_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        for campanha in self.campanhas:
            camp_frame = ctk.CTkFrame(scroll_frame, corner_radius=8,
                                     fg_color=colors["background"])
            camp_frame.pack(fill="x", pady=5)
            
            # Informações principais
            info_frame = ctk.CTkFrame(camp_frame, fg_color="transparent")
            info_frame.pack(fill="x", padx=10, pady=10)
            
            ctk.CTkLabel(info_frame, text=campanha['nome'], 
                        font=("Roboto", 14, "bold")).pack(side="left")
            
            # Estatísticas
            stats_frame = ctk.CTkFrame(camp_frame, fg_color="transparent")
            stats_frame.pack(fill="x", padx=10, pady=(0, 10))
            
            ctk.CTkLabel(stats_frame, text=f"Arquivo: {os.path.basename(campanha['arquivo'])}",
                        font=("Roboto", 11)).pack(side="left", padx=10)
            
            ctk.CTkLabel(stats_frame, text=f"Método: {campanha['metodo']}",
                        font=("Roboto", 11)).pack(side="left", padx=10)
            
            ctk.CTkLabel(stats_frame, text=f"Atendente: {campanha['atendente']}",
                        font=("Roboto", 11)).pack(side="left", padx=10)
            
            # Botões de ação
            btn_frame = ctk.CTkFrame(camp_frame, fg_color="transparent")
            btn_frame.pack(side="right", padx=10)
            
            edit_btn = ctk.CTkButton(btn_frame, text="Editar", width=80,
                                    command=lambda c=campanha: self.editar_campanha(c),
                                    fg_color=colors["primary"], hover_color=colors["secondary"])
            edit_btn.pack(side="left", padx=5)
            
            delete_btn = ctk.CTkButton(btn_frame, text="Excluir", width=80,
                                      command=lambda c=campanha: self.excluir_campanha(c),
                                      fg_color=colors["error"], hover_color="#B71C1C")
            delete_btn.pack(side="left", padx=5)
    
    def criar_nova_campanha(self):
        dialog = ctk.CTkToplevel(self.app)
        dialog.title("Nova Campanha")
        dialog.geometry("700x900")
        dialog.grab_set()
        
        colors = self.get_colors()
        
        # Nome da Campanha
        ctk.CTkLabel(dialog, text="Nome da Campanha:", anchor="w").pack(padx=20, pady=(20, 5), fill="x")
        nome_entry = ctk.CTkEntry(dialog)
        nome_entry.pack(padx=20, pady=5, fill="x")
        
        # Arquivo Excel
        ctk.CTkLabel(dialog, text="Arquivo de Contatos:", anchor="w").pack(padx=20, pady=5, fill="x")
        file_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        file_frame.pack(padx=20, pady=5, fill="x")
        
        file_entry = ctk.CTkEntry(file_frame)
        file_entry.pack(side="left", fill="x", expand=True, padx=(0, 5))
        
        file_btn = ctk.CTkButton(file_frame, text="Procurar", width=100,
                                command=lambda: self.selecionar_arquivo_campanha(file_entry))
        file_btn.pack(side="left")
        
        # Colunas
        ctk.CTkLabel(dialog, text="Coluna do Número:", anchor="w").pack(padx=20, pady=5, fill="x")
        numero_dropdown = ctk.CTkOptionMenu(dialog, values=[])
        numero_dropdown.pack(padx=20, pady=5, fill="x")
        
        ctk.CTkLabel(dialog, text="Coluna do Nome:", anchor="w").pack(padx=20, pady=5, fill="x")
        nome_dropdown = ctk.CTkOptionMenu(dialog, values=[])
        nome_dropdown.pack(padx=20, pady=5, fill="x")
        
        # Atendente
        ctk.CTkLabel(dialog, text="Atendente:", anchor="w").pack(padx=20, pady=5, fill="x")
        atendente_dropdown = ctk.CTkOptionMenu(dialog, values=self.atendentes)
        atendente_dropdown.pack(padx=20, pady=5, fill="x")
        
        # Método de Envio
        ctk.CTkLabel(dialog, text="Método de Envio:", anchor="w").pack(padx=20, pady=5, fill="x")
        metodo_var = ctk.StringVar(value="Web")
        metodo_switch = ctk.CTkSwitch(dialog, text="Usar API WhatsApp", 
                                     variable=metodo_var, onvalue="API", offvalue="Web")
        metodo_switch.pack(padx=20, pady=5, fill="x")
        
        # Mensagem
        ctk.CTkLabel(dialog, text="Modelo de Mensagem:", anchor="w").pack(padx=20, pady=5, fill="x")
        msg_text = ctk.CTkTextbox(dialog, height=200)
        msg_text.pack(padx=20, pady=5, fill="x")
        msg_text.insert("1.0", """Olá! Boa tarde! Como vai?
Sou a ATENDENTE, atendente do Hotel Fazenda Caluje.

Segue ficha para antecipação de check-in.
Qualquer duvida fico a disposição!""")
        
        # Botões
        btn_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        btn_frame.pack(padx=20, pady=20, fill="x")
        
        cancel_btn = ctk.CTkButton(btn_frame, text="Cancelar", 
                                  command=dialog.destroy,
                                  fg_color=colors["error"], hover_color="#B71C1C")
        cancel_btn.pack(side="right", padx=5)
        
        save_btn = ctk.CTkButton(btn_frame, text="Salvar Campanha",
                                command=lambda: self.salvar_campanha(
                                    nome_entry.get(),
                                    file_entry.get(),
                                    numero_dropdown.get(),
                                    nome_dropdown.get(),
                                    atendente_dropdown.get(),
                                    metodo_var.get(),
                                    msg_text.get("1.0", ctk.END),
                                    dialog
                                ))
        save_btn.pack(side="right", padx=5)
        
        # Atualizar colunas quando arquivo for selecionado
        def atualizar_colunas():
            if file_entry.get():
                try:
                    df = pd.read_excel(file_entry.get())
                    cols = list(df.columns)
                    numero_dropdown.configure(values=cols)
                    nome_dropdown.configure(values=cols)
                except:
                    pass
        
        file_entry.bind("<KeyRelease>", lambda e: atualizar_colunas())
        file_btn.configure(command=lambda: [self.selecionar_arquivo_campanha(file_entry), atualizar_colunas()])
    
    def selecionar_arquivo_campanha(self, entry):
        caminho = filedialog.askopenfilename(filetypes=[["Arquivos Excel", "*.xlsx *.xls"]])
        if caminho:
            entry.delete(0, ctk.END)
            entry.insert(0, caminho)
    
    def salvar_campanha(self, nome, arquivo, coluna_numero, coluna_nome, atendente, metodo, mensagem, dialog):
        if not nome or not arquivo or not coluna_numero or not coluna_nome or not atendente:
            messagebox.showerror("Erro", "Preencha todos os campos obrigatórios!")
            return
        if not os.path.exists(arquivo):
            messagebox.showerror("Erro", "O arquivo selecionado não existe!")
            return

        
        nova_campanha = {
            'nome': nome,
            'arquivo': arquivo,
            'coluna_numero': coluna_numero,
            'coluna_nome': coluna_nome,
            'atendente': atendente,
            'metodo': metodo,
            'mensagem': mensagem.strip(),
            'data_criacao': datetime.now().strftime("%Y-%m-%d")
        }
        
        self.campanhas.append(nova_campanha)
        self.salvar_dados()
        self.campanha_dropdown.configure(values=[c['nome'] for c in self.campanhas])
        dialog.destroy()
        self.mostrar_campanhas()

        messagebox.showinfo("Sucesso", "Campanha salva com sucesso!")

    
    def editar_campanha(self, campanha):
        # Implementar edição similar à criação
        pass
    
    def excluir_campanha(self, campanha):
        confirm = messagebox.askyesno("Confirmar", f"Excluir a campanha {campanha['nome']} permanentemente?")
        if confirm:
            self.campanhas.remove(campanha)
            self.salvar_dados()
            self.campanha_dropdown.configure(values=[c['nome'] for c in self.campanhas])
            self.mostrar_campanhas()

    def mostrar_estatisticas(self):
        self.limpar_conteudo()
        colors = self.get_colors()
        self.page_title.configure(text="Estatísticas")
        
        stats_frame = ctk.CTkFrame(self.content_frame, fg_color=colors["surface"], corner_radius=12)
        stats_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        tabs = ctk.CTkTabview(stats_frame)
        tabs.pack(fill="both", expand=True, padx=10, pady=10)
        
        tabs.add("Gráficos")
        tabs.add("Tabelas")
        
        # Aba de Gráficos
        fig, ax = plt.subplots(figsize=(8, 4))
        ax.pie([self.enviados, self.erros], labels=['Enviadas', 'Erros'], 
              colors=[colors["success"], colors["error"]], autopct='%1.1f%%')
        ax.set_title('Distribuição de Mensagens')
        
        chart_canvas = FigureCanvasTkAgg(fig, master=tabs.tab("Gráficos"))
        chart_canvas.draw()
        chart_canvas.get_tk_widget().pack(fill="both", expand=True)
        
        # Aba de Tabelas
        columns = ['Data', 'Enviadas', 'Erros', 'Taxa Sucesso']
        tree = ttk.Treeview(tabs.tab("Tabelas"), columns=columns, show='headings')
        
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=100)
        
        for stat in self.stats_data:
            taxa = (stat['enviados'] - stat['erros']) / stat['enviados'] * 100 if stat['enviados'] > 0 else 0
            tree.insert('', 'end', values=(
                stat['data'],
                stat['enviados'],
                stat['erros'],
                f"{taxa:.1f}%"
            ))
        
        tree.pack(fill="both", expand=True)

    def mostrar_configuracoes(self):
        self.limpar_conteudo()
        colors = self.get_colors()
        self.page_title.configure(text="Configurações")
        
        config_frame = ctk.CTkFrame(self.content_frame, fg_color=colors["surface"], corner_radius=12)
        config_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Configurações da API
        ctk.CTkLabel(config_frame, text="Configurações da API WhatsApp", 
                    font=("Roboto", 14, "bold")).grid(row=0, column=0, padx=10, pady=10, sticky="w")
        
        ctk.CTkLabel(config_frame, text="Access Token:").grid(row=1, column=0, padx=10, sticky="w")
        self.api_token_entry = ctk.CTkEntry(config_frame, width=300)
        self.api_token_entry.grid(row=1, column=1, padx=10, pady=5)
        self.api_token_entry.insert(0, self.api_settings['access_token'])
        
        ctk.CTkLabel(config_frame, text="Phone Number ID:").grid(row=2, column=0, padx=10, sticky="w")
        self.phone_id_entry = ctk.CTkEntry(config_frame, width=300)
        self.phone_id_entry.grid(row=2, column=1, padx=10, pady=5)
        self.phone_id_entry.insert(0, self.api_settings['phone_number_id'])
        
        self.api_toggle = ctk.CTkSwitch(config_frame, text="Usar API WhatsApp",
                                      variable=ctk.BooleanVar(value=self.api_settings['enabled']),
                                      command=self.toggle_api)
        self.api_toggle.grid(row=3, column=0, columnspan=2, padx=10, pady=10)
        
        save_btn = ctk.CTkButton(config_frame, text="Salvar Configurações", 
                                command=self.salvar_config_api)
        save_btn.grid(row=4, column=0, columnspan=2, pady=20)

    def toggle_api(self):
        self.api_settings['enabled'] = self.api_toggle.get()
    
    def salvar_config_api(self):
        self.api_settings['access_token'] = self.api_token_entry.get()
        self.api_settings['phone_number_id'] = self.phone_id_entry.get()
        self.salvar_dados()
        messagebox.showinfo("Sucesso", "Configurações salvas com sucesso!")

    def adicionar_atendente_controle(self):
        novo_atendente = self.attendant_entry.get()
        if novo_atendente and novo_atendente not in self.atendentes:
            self.atendentes.append(novo_atendente)
            self.attendant_dropdown.configure(values=self.atendentes)
            self.attendant_entry.delete(0, ctk.END)
            self.salvar_dados()
    
    def remover_atendente_controle(self):
        atendente = self.nome_atendente.get()
        if atendente in self.atendentes:
            self.atendentes.remove(atendente)
            self.attendant_dropdown.configure(values=self.atendentes)
            self.salvar_dados()

    def executar(self):
        self.app.mainloop()

if __name__ == "__main__":
    splash = SplashScreen()
    splash.mainloop()
    app = WhatsAppSender()
    app.executar()