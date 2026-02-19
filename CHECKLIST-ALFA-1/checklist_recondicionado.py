import customtkinter as ctk
from tkinter import messagebox, filedialog
import wmi
import psutil
import platform
import os
import sys
import subprocess
import datetime
import webbrowser
import threading
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from PIL import Image
import json


# --- CONFIGURAÇÃO DE TEMA ---
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")  # Pode ser "blue" (padrão), "green", "dark-blue"

# Cores Personalizadas (Hex)
COLOR_BG = "#1a1a1a"       # Fundo Principal
COLOR_CARD = "#2b2b2b"     # Fundo dos Cartões/Frames
COLOR_ACCENT = "#2cc985"   # Verde Neon (Ações Principais)
COLOR_TEXT = "#ffffff"     # Texto Principal
COLOR_TEXT_DIM = "#a0a0a0" # Texto Secundário
COLOR_DANGER = "#e74c3c"   # Vermelho (Erros/Sair)
COLOR_INFO = "#3498db"     # Azul Informação
PASSWORD_REGISTOS = "admin" # Senha para acessar registos

# --- CONFIGURAÇÃO DE DIRETÓRIOS ---
# Define pasta de dados local (ao lado do script/executável) para portabilidade
# Se estiver congelado (PyInstaller), usa sys.executable, senão usa __file__
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

DATA_DIR = os.path.join(BASE_DIR, "Dados")

if not os.path.exists(DATA_DIR):
    try:
        os.makedirs(DATA_DIR)
    except Exception as e:
        print(f"Erro ao criar diretório de dados: {e}")

EXCEL_FILE = os.path.join(DATA_DIR, "registos_checklist.xlsx")



class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        # Configuração da Janela
        self.title("Sistema de Checklist Pro")
        self.geometry("900x700")
        self.minsize(800, 600)
        
        # Cache de Info do Sistema
        self.sys_info = {}
        
        # Container Principal (para gestão de 'páginas')
        self.container = ctk.CTkFrame(self, fg_color="transparent")
        self.container.pack(fill="both", expand=True)
        
        # Inicializa Telas
        self.frames = {}
        self.current_frame = None
        
        # Carregar Hardware em Background
        threading.Thread(target=self.load_system_info_bg, daemon=True).start()
        
        # Mostrar Menu Inicial
        self.show_frame("MenuPrincipal")

    def load_system_info_bg(self):
        """Carrega info do sistema sem travar a UI"""
        try:
            # WMI precisa de inicialização COM em threads separadas
            import pythoncom
            pythoncom.CoInitialize()
        except:
            pass
            
        self.sys_info = get_system_info()
        
        # Agendar atualização na thread principal
        self.after(0, self.update_ui_after_load)

    def update_ui_after_load(self):
        """Atualiza a UI na thread principal"""
        if "ChecklistFrame" in self.frames:
            self.frames["ChecklistFrame"].update_hardware_info()

    def show_frame(self, page_name):
        """Alterna entre telas com uma transição simples"""
        # Destrói o frame atual (limpeza simples para evitar sobreposição de estados)
        if self.current_frame:
            self.current_frame.pack_forget()
            
        # Cria a nova tela se ainda não existir ou recria para resetar
        if page_name == "MenuPrincipal":
            self.current_frame = MenuPrincipal(self.container, self)
        elif page_name == "ChecklistFrame":
            self.current_frame = ChecklistFrame(self.container, self)
        
        if self.current_frame:
            self.current_frame.pack(fill="both", expand=True)

# --- TELAS ---

class MenuPrincipal(ctk.CTkFrame):
    def __init__(self, parent, controller):
        super().__init__(parent, fg_color="transparent")
        self.controller = controller
        
        # Centralizar conteúdo
        self.center_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.center_frame.place(relx=0.5, rely=0.5, anchor="center")
        
        # Título
        ctk.CTkLabel(self.center_frame, text="CHECKLIST RECONDICIONADOS", 
                    font=("Roboto Medium", 32), text_color=COLOR_TEXT).pack(pady=(0, 10))
        
        ctk.CTkLabel(self.center_frame, text="Selecione uma opção para continuar", 
                    font=("Roboto", 14), text_color=COLOR_TEXT_DIM).pack(pady=(0, 40))
        
        # Botões (Cartões Grandes)
        self.create_menu_button("📝  NOVA CHECKLIST", COLOR_ACCENT, 
                              lambda: controller.show_frame("ChecklistFrame"))
        
        self.create_menu_button("📂  ABRIR REGISTOS", COLOR_INFO, 
                              self.check_password_registos)
        
        self.create_menu_button("⚠️  EXPORTAR DANOS", "#e67e22", 
                              exportar_danos_ui)
        
        self.create_menu_button("📄  EXPORTAR COMPRA (PDF)", "#8e44ad", 
                              exportar_compra_pdf_ui)

        # Rodapé

        ctk.CTkLabel(self, text="v3.0 - Design Minimalista", 
                    font=("Roboto", 10), text_color=COLOR_TEXT_DIM).pack(side="bottom", pady=20)

    def create_menu_button(self, text, color, command):
        btn = ctk.CTkButton(self.center_frame, text=text, command=command,
                           fg_color="transparent", border_width=2, border_color=color,
                           text_color=color, hover_color=color,
                           font=("Roboto Medium", 14), height=50, width=300,
                           corner_radius=25) # Botão arredondado ("Pílula")
        btn.pack(pady=10)
        # Hack para mudar cor do texto no hover (simulado, CTkButton nativo já lida bem com isso)
        def on_enter(e): btn.configure(text_color="#ffffff")
        def on_leave(e): btn.configure(text_color=color)
        btn.bind("<Enter>", on_enter)
        btn.bind("<Leave>", on_leave)

    def check_password_registos(self):
        """Solicita senha antes de abrir os registos"""
        dialog = ctk.CTkInputDialog(text="Digite a palavra-passe de administrador:", title="Acesso Restrito")
        password = dialog.get_input()
        
        if password == PASSWORD_REGISTOS:
            abrir_registos_excel()
        elif password is not None: # Se não cancelou
            messagebox.showerror("Erro", "Palavra-passe incorreta!")

class ChecklistFrame(ctk.CTkFrame):
    def __init__(self, parent, controller):
        super().__init__(parent, fg_color="transparent")
        self.controller = controller
        
        # Top Bar
        top_bar = ctk.CTkFrame(self, fg_color=COLOR_CARD, height=60, corner_radius=0)
        top_bar.pack(fill="x", side="top")
        
        ctk.CTkButton(top_bar, text="⬅ Voltar", command=lambda: controller.show_frame("MenuPrincipal"),
                     fg_color="transparent", text_color=COLOR_TEXT, width=80).pack(side="left", padx=10)
        
        ctk.CTkLabel(top_bar, text="NOVO RELATÓRIO", font=("Roboto Medium", 18)).pack(side="left", padx=20)
        
        # Scroll Area Principal
        self.scroll = ctk.CTkScrollableFrame(self, fg_color="transparent")
        self.scroll.pack(fill="both", expand=True, padx=20, pady=20)
        
        # --- SEÇÃO 1: TÉCNICO ---
        self.add_section_header("1. Técnico Responsável")
        self.user_frame = ctk.CTkFrame(self.scroll, fg_color=COLOR_CARD)
        self.user_frame.pack(fill="x", pady=(0, 20))
        
        users = ["Guilherme", "Alex", "Araujo", "Convidado"]
        self.user_var = ctk.StringVar(value=users[0])
        self.combo_user = ctk.CTkComboBox(self.user_frame, values=users, variable=self.user_var,
                                        command=self.check_guest, width=300)
        self.combo_user.pack(padx=20, pady=20, anchor="w")
        
        self.entry_guest = ctk.CTkEntry(self.user_frame, placeholder_text="Nome do Convidado", width=300)
        
        # --- SEÇÃO 2: HARDWARE ---
        self.add_section_header("2. Hardware Detectado")
        self.hw_frame = ctk.CTkFrame(self.scroll, fg_color=COLOR_CARD)
        self.hw_frame.pack(fill="x", pady=(0, 20))
        
        self.lbl_model = ctk.CTkLabel(self.hw_frame, text="Detectando...", font=("Roboto Medium", 16))
        self.lbl_model.pack(padx=20, pady=(15, 5), anchor="w")
        
        self.lbl_specs = ctk.CTkLabel(self.hw_frame, text="Aguarde...", text_color=COLOR_TEXT_DIM, justify="left")
        self.lbl_specs.pack(padx=20, pady=(0, 15), anchor="w")
        
        self.update_hardware_info()
        
        # Detalhes de Memória adicionais
        self.ram_frame = ctk.CTkFrame(self.hw_frame, fg_color="transparent")
        self.ram_frame.pack(fill="x", padx=20, pady=(0, 15))
        
        ctk.CTkLabel(self.ram_frame, text="Tipo Memória:", font=("Roboto", 12)).pack(side="left", padx=(0, 10))
        self.ram_type = ctk.CTkOptionMenu(self.ram_frame, values=["DIMM", "Onboard", "Mista"], width=100)
        self.ram_type.pack(side="left", padx=(0, 20))
        
        ctk.CTkLabel(self.ram_frame, text="Config. (Ex: 2x 8GB):", font=("Roboto", 12)).pack(side="left", padx=(0, 10))
        self.entry_ram_config = ctk.CTkEntry(self.ram_frame, width=150)
        self.entry_ram_config.pack(side="left")
        
        # --- SEÇÃO 3: COMPRA ---
        self.add_section_header("3. Referência de Compra")
        self.compra_frame = ctk.CTkFrame(self.scroll, fg_color=COLOR_CARD)
        self.compra_frame.pack(fill="x", pady=(0, 20))
        
        self.entry_compra = ctk.CTkEntry(self.compra_frame, placeholder_text="Ex: 123456", width=300)
        self.entry_compra.pack(padx=20, pady=20, anchor="w")
        
        # --- SEÇÃO 4: TESTES ---
        self.add_section_header("4. Checklist de Testes")
        self.test_vars = {}
        tests = ["Teclado", "Ecrã", "Touch Screen", "Wifi", "LAN", "Portas USB", "Webcam", "Microfone", "Colunas", "Saídas Vídeo"]
        
        self.tests_frame = ctk.CTkFrame(self.scroll, fg_color=COLOR_CARD)
        self.tests_frame.pack(fill="x", pady=(0, 20))
        
        # Grid para os testes
        # Grid para os testes
        for i, test in enumerate(tests):
            row = i // 2
            col = i % 2
            self.create_test_item(self.tests_frame, test, row, col)


        # --- SEÇÃO 5: NOTAS ---
        self.add_section_header("5. Observações e Danos")
        self.notes_frame = ctk.CTkFrame(self.scroll, fg_color=COLOR_CARD)
        self.notes_frame.pack(fill="x", pady=(0, 20))
        
        self.text_notes = ctk.CTkTextbox(self.notes_frame, height=100)
        self.text_notes.pack(fill="x", padx=20, pady=20)
        
        # --- BOTÃO AÇÃO ---
        self.btn_save = ctk.CTkButton(self.scroll, text="GERAR RELATÓRIO E GUARDAR", 
                                     fg_color=COLOR_ACCENT, hover_color="#27ae60",
                                     height=50, font=("Roboto Medium", 14),
                                     command=self.gerar_relatorio)
        self.btn_save.pack(fill="x", pady=20)

    def open_incognito(self, url):
        """Abre o link em modo anônimo (Chrome ou Edge)"""
        try:
            # Tenta Chrome Incognito
            if sys.platform == 'win32':
                subprocess.run(f'start chrome --incognito "{url}"', shell=True, check=True)
            else:
                webbrowser.open(url) # Fallback para Linux/Mac por enquanto
        except:
            try:
                # Tenta Edge InPrivate
                if sys.platform == 'win32':
                    subprocess.run(f'start msedge --inprivate "{url}"', shell=True, check=True)
            except:
                # Fallback padrão
                webbrowser.open(url)

    def add_section_header(self, text):
        ctk.CTkLabel(self.scroll, text=text, font=("Roboto Medium", 14), 
                    text_color=COLOR_ACCENT).pack(anchor="w", pady=(10, 5))

    def check_guest(self, choice):
        if choice == "Convidado":
            self.entry_guest.pack(padx=20, pady=(0, 20), anchor="w")
        else:
            self.entry_guest.pack_forget()

    def update_hardware_info(self):
        info = self.controller.sys_info
        if info:
            self.lbl_model.configure(text=info.get('modelo', 'Desconhecido'))
            specs = f"S/N: {info.get('serial', 'N/A')}\nCPU: {info.get('cpu', 'N/A')}\nRAM: {info.get('ram', 'N/A')}\nDISCO: {info.get('disk', 'N/A')}\nGPU: {info.get('gpu', 'N/A')}"
            self.lbl_specs.configure(text=specs)

    def create_test_item(self, parent, test_name, row, col):
        # Frame individual para cada teste
        frame = ctk.CTkFrame(parent, fg_color="transparent")
        frame.grid(row=row, column=col, sticky="ew", padx=20, pady=10)
        
        # Label
        ctk.CTkLabel(frame, text=test_name, font=("Roboto", 12)).pack(anchor="w")
        
        # Switch Moderno
        var = ctk.BooleanVar(value=False)
        self.test_vars[test_name] = var
        
        switch = ctk.CTkSwitch(frame, text="Aprovado", variable=var, 
                              progress_color=COLOR_ACCENT, button_color="#ffffff")
        switch.pack(anchor="w")
        
        # Link opcional
        test_urls = {
            "Teclado": "https://en.key-test.ru/",
            "Webcam": "https://pt.webcamtests.com/",
            "Microfone": "https://pt.mictests.com/",
            "Colunas": "https://pt.mictests.com/sound-test/",
            "Ecrã": "https://deadpixeltest.org/",
            "Touch Screen": "https://testmyscreen.com/"
        }
        if test_name in test_urls:
             link = ctk.CTkLabel(frame, text="Abrir Teste Online 🔗", text_color=COLOR_INFO, 
                                cursor="hand2", font=("Roboto", 10))
             link.pack(anchor="w")
             link.bind("<Button-1>", lambda e, u=test_urls[test_name]: self.open_incognito(u))

    def gerar_relatorio(self):
        # Coletar Dados
        usuario = self.user_var.get()
        if usuario == "Convidado":
            usuario = self.entry_guest.get() or "Convidado"
            
        compra_num = self.entry_compra.get()
        danos = self.text_notes.get("1.0", "end-1c")
        
        # Detalhes RAM
        ram_type = self.ram_type.get()
        ram_config = self.entry_ram_config.get()
        ram_details = {
            "type": ram_type,
            "config": ram_config
        }
        
        testes = {name: var.get() for name, var in self.test_vars.items()}
        
        # Mapeamento de nomes para o Excel/HTML (para compatibilidade com código antigo)
        testes_mapped = {
            "Teclado": testes.get("Teclado"),
            "Ecrã": testes.get("Ecrã"),
            "Touch Screen": testes.get("Touch Screen"),
            "Wifi": testes.get("Wifi"),
            "LAN": testes.get("LAN"),
            "Webcam": testes.get("Webcam"),
            "Microfone": testes.get("Microfone"),
            "Colunas": testes.get("Colunas"),
            "USB": testes.get("Portas USB"), 
            "Portas de Vídeo": testes.get("Saídas Vídeo")
        }
        
        # Chaman lógica de geração (reutilizando função externa refatorada ou movida)
        gerar_relatorio_logic(self.controller.sys_info, usuario, compra_num, testes_mapped, danos, ram_details)


# --- LÓGICA DO SISTEMA (ADAPTADA) ---

def get_system_info():
    info = {}
    try:
        c = wmi.WMI()
        # 1. Modelo
        try:
            sys_data = c.Win32_ComputerSystem()[0]
            info['modelo'] = f"{sys_data.Manufacturer} {sys_data.Model}".strip()
        except: info['modelo'] = "Modelo Desconhecido"
            
        # 2. Serial (Melhorado - Prioriza SystemProduct > BIOS > BaseBoard)
        try:
            # Tenta pegar o "Service Tag" ou Identificador do Sistema (Mais confiável)
            product = c.Win32_ComputerSystemProduct()[0]
            serial = product.IdentifyingNumber.strip()
            
            # Filtra valores genéricos comuns
            if serial.lower() in ["default string", "system serial number", "to be filled by o.e.m.", "0"]:
                raise ValueError("Serial genérico")
            
            info['serial'] = serial
        except:
            try:
                # Tenta BIOS (Muito comum em laptops)
                bios = c.Win32_Bios()[0]
                serial = bios.SerialNumber.strip()
                
                if serial.lower() in ["default string", "system serial number", "to be filled by o.e.m.", "0"]:
                    raise ValueError("Serial genérico")
                
                info['serial'] = serial
            except:
                try:
                     # Último recurso: BaseBoard (Pode ser o serial da placa mãe interna)
                    board = c.Win32_BaseBoard()[0]
                    info['serial'] = board.SerialNumber.strip()
                except: 
                    info['serial'] = "N/A"
        
        # 3. CPU
        try:
            proc = c.Win32_Processor()[0]
            info['cpu'] = proc.Name.strip()
        except: info['cpu'] = platform.processor()
        
        # 4. RAM
        try:
            ram_bytes = psutil.virtual_memory().total
            info['ram'] = f"{round(ram_bytes / (1024**3))} GB"
        except: info['ram'] = "N/A"
        
        # 5. Disco
        disks = []
        try:
            for disk in c.Win32_DiskDrive():
                size_gb = round(int(disk.Size) / (1024**3))
                disks.append(f"{disk.Model} ({size_gb} GB)")
            info['disk'] = " + ".join(disks)
        except: info['disk'] = "N/A"
        
        # 6. GPU
        gpus = []
        try:
            for gpu in c.Win32_VideoController():
                gpus.append(gpu.Name)
            info['gpu'] = " | ".join(gpus)
        except: info['gpu'] = "N/A"
        
    except Exception as e:
        info['error'] = str(e)
        return None
    return info

def gerar_relatorio_logic(sys_info, usuario, compra_num, testes, danos, ram_details=None):
    """Gera o HTML e salva no Excel"""
    
    # HTML Content (Design Melhorado)
    foto_user = f"{usuario.lower()}.jpg"
    html_content = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <title>Relatório {sys_info.get('modelo', 'PC')}</title>
        <style>
            body {{ font-family: 'Segoe UI', sans-serif; margin: 0; background: #f4f7f6; color: #333; }}
            .container {{ max_width: 800px; margin: 40px auto; background: white; padding: 40px; border-radius: 12px; box-shadow: 0 10px 30px rgba(0,0,0,0.1); }}
            .header {{ display: flex; align-items: center; border-bottom: 2px solid #eee; padding-bottom: 20px; margin-bottom: 30px; }}
            .header img {{ width: 80px; height: 80px; border-radius: 50%; object-fit: cover; margin-right: 20px; border: 3px solid #eee; }}
            .header h1 {{ margin: 0; font-size: 24px; color: #2c3e50; }}
            .header p {{ margin: 5px 0 0; color: #7f8c8d; }}
            h2 {{ color: #2980b9; font-size: 18px; border-left: 4px solid #2980b9; padding-left: 10px; margin-top: 30px; }}
            table {{ width: 100%; border-collapse: collapse; margin-top: 15px; }}
            th, td {{ padding: 12px 15px; border-bottom: 1px solid #eee; text-align: left; }}
            th {{ background-color: #f8f9fa; color: #2c3e50; font-weight: 600; }}
            .pass {{ color: #27ae60; font-weight: bold; }}
            .fail {{ color: #e74c3c; font-weight: bold; }}
            .notes {{ background: #fff8e1; padding: 20px; border-radius: 8px; border-left: 4px solid #ffc107; margin-top: 15px; }}
        </style>
    </head>
    <body>
        <div class="container">
            <div class="header">
                <img src="{foto_user}" onerror="this.src='https://ui-avatars.com/api/?name={usuario}&background=random'">
                <div>
                    <h1>Relatório Técnico de Certificação</h1>
                    <p>Técnico: <strong>{usuario}</strong> &bull; {datetime.datetime.now().strftime("%d/%m/%Y")}</p>
                </div>
            </div>

            <h2>📦 Identificação do Equipamento</h2>
            <table>
                <tr><th width="30%">Ref. Compra</th><td>{compra_num}</td></tr>
                <tr><th>Modelo</th><td><strong>{sys_info.get('modelo', 'N/A')}</strong></td></tr>
                <tr><th>Serial Number</th><td>{sys_info.get('serial', 'N/A')}</td></tr>
            </table>

            <h2>⚙️ Especificações Técnicas</h2>
            <table>
                <tr><th width="30%">Processador</th><td>{sys_info.get('cpu', 'N/A')}</td></tr>
                <tr><th>Memória RAM</th><td>{sys_info.get('ram', 'N/A')} <span style="font-size: 0.9em; color: #666;">({ram_details.get('type', '')} - {ram_details.get('config', '')})</span></td></tr>
                <tr><th>Armazenamento</th><td>{sys_info.get('disk', 'N/A')}</td></tr>
                <tr><th>Gráfica</th><td>{sys_info.get('gpu', 'N/A')}</td></tr>
            </table>

            <h2>✅ Resultados dos Testes</h2>
            <table>
                {''.join([f"<tr><td>{k}</td><td class='{'pass' if v else 'fail'}'>{'APROVADO' if v else 'REPROVADO'}</td></tr>" for k,v in testes.items()])}
            </table>

            <h2>📝 Observações</h2>
            <div class="notes">
                {danos.replace('\\n', '<br>') if danos and danos.strip() else "Nenhuma anomalia visual detetada. Equipamento em condições normais."}
            </div>
        </div>
    </body>
    </html>
    """
    
    # Salvar HTML
    safe_serial = "".join([c for c in sys_info.get('serial', 'SN') if c.isalnum()]).strip()
    default_filename = f"{safe_serial}.html"
    
    file_path = filedialog.asksaveasfilename(
        defaultextension=".html",
        filetypes=[("Ficheiros HTML", "*.html")],
        initialfile=default_filename,
        title="Guardar Relatório Como..."
    )
    
    if not file_path: return

    try:
        with open(file_path, "w", encoding="utf-8") as f:
            f.write(html_content)
            
        webbrowser.open('file://' + os.path.realpath(file_path))
        
        # Salvar Excel
        if guardar_em_excel(usuario, compra_num, sys_info, testes, danos, ram_details):
            messagebox.showinfo("Sucesso", "Processo concluído com sucesso!")
        else:
            messagebox.showwarning("Atenção", "Relatório HTML gerado, mas erro ao salvar no Excel.")
            
    except Exception as e:
        messagebox.showerror("Erro", f"Falha ao gravar: {e}")

def guardar_em_excel(usuario, compra_num, sys_info, testes, danos, ram_details=None):
    try:
        registo = {
            "Data": datetime.datetime.now().strftime("%d/%m/%Y %H:%M"),
            "Técnico": usuario,
            "Nº Compra": compra_num,
            "Modelo": sys_info.get('modelo', 'N/A'),
            "Serial": sys_info.get('serial', 'N/A'),
            "CPU": sys_info.get('cpu', 'N/A'),
            "RAM": sys_info.get('ram', 'N/A'),
            "Tipo RAM": ram_details.get('type', '') if ram_details else "",
            "Config RAM": ram_details.get('config', '') if ram_details else "",
            "Disco": sys_info.get('disk', 'N/A'),
            "GPU": sys_info.get('gpu', 'N/A'),
            "Teclado": "✓" if testes.get("Teclado") else "✗",
            "Ecrã": "✓" if testes.get("Ecrã") else "✗",
            "Touch Screen": "✓" if testes.get("Touch Screen") else "✗",
            "Wifi": "✓" if testes.get("Wifi") else "✗",
            "LAN": "✓" if testes.get("LAN") else "✗",
            "Webcam": "✓" if testes.get("Webcam") else "✗",
            "Microfone": "✓" if testes.get("Microfone") else "✗",
            "Colunas": "✓" if testes.get("Colunas") else "✗",
            "USB": "✓" if testes.get("USB") else "✗",
            "Portas de Vídeo": "✓" if testes.get("Portas de Vídeo") else "✗",
            "Notas": danos.strip() if danos.strip() else "Sem observações"
        }
        
        if os.path.exists(EXCEL_FILE):
             df_exist = pd.read_excel(EXCEL_FILE)
             df = pd.concat([df_exist, pd.DataFrame([registo])], ignore_index=True)
        else:
             df = pd.DataFrame([registo])
             
        # Forçar ordem das colunas
        cols_order = [
            "Data", "Técnico", "Nº Compra", "Modelo", "Serial", "CPU", "RAM", "Tipo RAM", "Config RAM", 
            "Disco", "GPU",
            "Teclado", "Ecrã", "Touch Screen", "Wifi", "LAN", "Webcam", "Microfone", "Colunas", "USB", 
            "Portas de Vídeo", "Notas"
        ]
        
        # Garantir que todas as colunas existem (se o excel antigo não tiver alguma)
        for col in cols_order:
            if col not in df.columns:
                df[col] = "N/A"
                
        # Reordenar
        df = df[cols_order]
             
        df.to_excel(EXCEL_FILE, index=False, sheet_name="Registos")
        formatar_excel(EXCEL_FILE)
        return True
    except Exception as e:
        print(e)
        return False

def formatar_excel(filepath):
    """Aplica formatação moderna ao ficheiro Excel"""
    try:
        wb = load_workbook(filepath)
        ws = wb.active
        
        # Definir estilos
        header_fill = PatternFill(start_color="2E75B6", end_color="2E75B6", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF", size=11)
        header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        
        alt_fill = PatternFill(start_color="D9E9F7", end_color="D9E9F7", fill_type="solid")
        white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
        
        thin_border = Border(
            left=Side(style='thin', color='B4C7E7'),
            right=Side(style='thin', color='B4C7E7'),
            top=Side(style='thin', color='B4C7E7'),
            bottom=Side(style='thin', color='B4C7E7')
        )
        
        center_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        left_alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        
        # Formatar cabeçalho
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = header_alignment
            cell.border = thin_border
        
        # Altura do cabeçalho
        ws.row_dimensions[1].height = 25
        
        # Formatar linhas de dados
        for row_idx, row in enumerate(ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column), 2):
            # Cor alternada para as linhas
            fill = alt_fill if (row_idx % 2 == 0) else white_fill
            
            for col_idx, cell in enumerate(row, 1):
                cell.fill = fill
                cell.border = thin_border
                
                # Alinhar colunas de testes (marcas ✓ e ✗)
                if col_idx > 11:  # Colunas dos testes (ajustado para novas colunas)
                    cell.alignment = center_alignment
                    cell.font = Font(size=12, bold=True)
                    
                    # Colorir as marcas
                    if cell.value == "✓":
                        cell.font = Font(size=12, bold=True, color="00B050")
                    elif cell.value == "✗":
                        cell.font = Font(size=12, bold=True, color="C00000")
                else:
                    cell.alignment = left_alignment
                    cell.font = Font(size=10)
                
                ws.row_dimensions[row_idx].height = 20
        
        # Ajustar largura das colunas
        column_widths = {
            'A': 16,  # Data
            'B': 12,  # Técnico
            'C': 14,  # Nº Compra
            'D': 22,  # Modelo
            'E': 18,  # Serial
            'F': 25,  # CPU
            'G': 12,  # RAM
            'H': 10,  # Tipo RAM
            'I': 15,  # Config RAM
            'J': 20,  # Disco
            'K': 20,  # GPU
            'L': 10,  # Teclado
            'M': 8,   # Ecrã
            'N': 12,  # Touch
            'O': 8,   # Wifi
            'P': 8,   # LAN
            'Q': 10,  # Webcam
            'R': 12,  # Microfone
            'S': 10,  # Colunas
            'T': 8,   # USB
            'U': 14,  # Portas de Vídeo
            'V': 30   # Notas
        }
        
        for col_letter, width in column_widths.items():
            ws.column_dimensions[col_letter].width = width
        
        # Congelar a primeira linha (cabeçalho)
        ws.freeze_panes = "A2"
        
        wb.save(filepath)
        return True
    except Exception as e:
        print(f"Erro ao formatar Excel: {e}")
        return False

def formatar_excel_danos(filepath):
    """Formata o Excel de Danos para ficar estético e ajustado"""
    try:
        wb = load_workbook(filepath)
        ws = wb.active
        
        # Estilos
        header_fill = PatternFill(start_color="E67E22", end_color="E67E22", fill_type="solid") # Laranja
        header_font = Font(bold=True, color="FFFFFF", size=12)
        header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        
        thin_border = Border(
            left=Side(style='thin', color='B4C7E7'),
            right=Side(style='thin', color='B4C7E7'),
            top=Side(style='thin', color='B4C7E7'),
            bottom=Side(style='thin', color='B4C7E7')
        )
        
        left_alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        
        # Formatar Cabeçalho
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = header_alignment
            cell.border = thin_border
            
        ws.row_dimensions[1].height = 25
        
        # Formatar Dados e Ajustar Colunas
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            for cell in row:
                cell.border = thin_border
                cell.alignment = left_alignment
                
                # Se for a coluna Serial (B), forçar formato de texto
                if cell.column_letter == 'B':
                    cell.number_format = '@'
                
                # Se for a coluna de Notas (C), permitir quebra de linha
                if cell.column_letter == 'C': 
                    cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        
        # Ajustar larguras
        column_widths = {'A': 30, 'B': 25, 'C': 60}
        
        for col in ws.columns:
            max_length = 0
            column = col[0].column_letter
            
            if column == 'C':
                ws.column_dimensions[column].width = 60
                continue
                
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except: pass
            
            adjusted_width = (max_length + 2)
            if adjusted_width < 10: adjusted_width = 10
            if adjusted_width > 50: adjusted_width = 50
            
            ws.column_dimensions[column].width = adjusted_width

        wb.save(filepath)
        return True
    except Exception as e:
        print(f"Erro ao formatar Excel de danos: {e}")
        return False

def exportar_danos_ui():
    # Verificar se existe ficheiro antes de pedir input
    if not os.path.exists(EXCEL_FILE):
        messagebox.showinfo("Aviso", "Ainda não existem registos para exportar.")
        return

    # Interface simples para input (CustomTkinter tem CTkInputDialog)
    dialog = ctk.CTkInputDialog(text="Digite o Nº de Compra para exportar:", title="Exportar Danos")
    compra_num = dialog.get_input()
    
    if not compra_num: return
    
    # Reutilizando lógica de exportação (simplificada)
    try:
        df = pd.read_excel(EXCEL_FILE)
        df['Nº Compra'] = df['Nº Compra'].astype(str)
        filtro = (df['Nº Compra'] == str(compra_num)) & (df['Notas'].notna()) & (df['Notas'] != "Sem observações")
        df_export = df[filtro][['Modelo', 'Serial', 'Notas']]
        
        if df_export.empty:
            messagebox.showinfo("Vazio", "Nenhum registo com danos encontrado para esta compra.")
            return

        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=f"Danos_{compra_num}.xlsx")
        if save_path:
            df_export.to_excel(save_path, index=False)
            formatar_excel_danos(save_path)
            os.startfile(save_path)
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao exportar: {str(e)}")

def formatar_excel_compra_pdf(filepath):
    """Formata o Excel temporário para ficar apresentável no PDF"""
    try:
        wb = load_workbook(filepath)
        ws = wb.active
        
        # Estilos
        header_fill = PatternFill(start_color="8E44AD", end_color="8E44AD", fill_type="solid") # Roxo
        header_font = Font(bold=True, color="FFFFFF", size=11)
        header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        
        thin_border = Border(
            left=Side(style='thin', color='000000'),
            right=Side(style='thin', color='000000'),
            top=Side(style='thin', color='000000'),
            bottom=Side(style='thin', color='000000')
        )
        
        left_alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = header_alignment
            cell.border = thin_border
            
        ws.row_dimensions[1].height = 25
        
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            for cell in row:
                cell.border = thin_border
                cell.alignment = left_alignment
                if cell.column_letter == 'B': # Serial
                    cell.number_format = '@'
                if cell.column_letter == 'G': # Observações
                    cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        
        # Ajustar larguras
        ws.column_dimensions['A'].width = 25 # Modelo
        ws.column_dimensions['B'].width = 15 # Serial
        ws.column_dimensions['C'].width = 20 # CPU
        ws.column_dimensions['D'].width = 12 # RAM
        ws.column_dimensions['E'].width = 15 # Disco
        ws.column_dimensions['F'].width = 15 # GPU
        ws.column_dimensions['G'].width = 45 # Observações
        
        wb.save(filepath)
        return True
    except Exception as e:
        print(f"Erro formatando PDF Excel: {e}")
        return False

def exportar_compra_pdf_ui():
    if not os.path.exists(EXCEL_FILE):
        messagebox.showinfo("Aviso", "Ainda não existem registos para exportar.")
        return

    dialog = ctk.CTkInputDialog(text="Digite o Nº de Compra para exportar em PDF:", title="Exportar Compra")
    compra_num = dialog.get_input()
    
    if not compra_num: return
    
    try:
        import win32com.client
        df = pd.read_excel(EXCEL_FILE)
        df['Nº Compra'] = df['Nº Compra'].astype(str)
        
        filtro = df['Nº Compra'] == str(compra_num)
        
        colunas_necessarias = ['Modelo', 'Serial', 'CPU', 'RAM', 'Disco', 'GPU', 'Notas']
        
        df_export = df[filtro].reindex(columns=colunas_necessarias)
        df_export.rename(columns={'Notas': 'Observações'}, inplace=True)
        
        if df_export.empty:
            messagebox.showinfo("Vazio", "Nenhum registo encontrado para esta compra.")
            return

        default_pdf = f"Compra_{compra_num}.pdf"
        save_path_pdf = filedialog.asksaveasfilename(defaultextension=".pdf", initialfile=default_pdf, filetypes=[("Documentos PDF", "*.pdf")])
        if not save_path_pdf: return
        
        temp_excel = os.path.join(DATA_DIR, f"temp_compra_{compra_num}.xlsx")
        
        df_export.to_excel(temp_excel, index=False)
        formatar_excel_compra_pdf(temp_excel)
        
        try:
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            
            wb = excel.Workbooks.Open(os.path.abspath(temp_excel))
            
            ws = wb.ActiveSheet
            ws.PageSetup.Orientation = 2 # xlLandscape
            ws.PageSetup.Zoom = False
            ws.PageSetup.FitToPagesWide = 1
            ws.PageSetup.FitToPagesTall = False
            
            wb.ExportAsFixedFormat(0, os.path.abspath(save_path_pdf))
            wb.Close(False)
            excel.Quit()
            
            if os.path.exists(temp_excel):
                try: os.remove(temp_excel)
                except: pass
                
            os.startfile(save_path_pdf)
        except Exception as com_err:
            messagebox.showerror("Erro PDF", f"Erro a comunicar com o Excel para gerar PDF: {str(com_err)}\n\nO ficheiro XLSX com formato foi gerado em {DATA_DIR}.")
            os.startfile(DATA_DIR)
            
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao exportar: {str(e)}")

def abrir_registos_excel():
    if os.path.exists(EXCEL_FILE): 
        try:
            os.startfile(EXCEL_FILE)
        except Exception as e:
            # Tentar abrir com subprocess se os.startfile falhar
            try:
                if sys.platform == 'win32':
                    subprocess.Popen(['start', '', EXCEL_FILE], shell=True)
                else:
                    subprocess.call(['xdg-open', EXCEL_FILE])
            except Exception as e2:
                 messagebox.showerror("Erro", f"Não foi possível abrir o Excel.\nCaminho: {EXCEL_FILE}\nErro: {e}\n{e2}")
    else: 
        # Se não existir, perguntar se quer criar um vazio
        resposta = messagebox.askyesno("Não Encontrado", "O ficheiro de registos ainda não existe. Deseja criar um novo agora?")
        if resposta:
            try:
                # Criar DataFrame vazio com as colunas corretas
                cols = [
                    "Data", "Técnico", "Nº Compra", "Modelo", "Serial", "CPU", "RAM", "Tipo RAM", "Config RAM", 
                    "Disco", "GPU",
                    "Teclado", "Ecrã", "Touch Screen", "Wifi", "LAN", "Webcam", "Microfone", "Colunas", "USB", 
                    "Portas de Vídeo", "Notas"
                ]
                df = pd.DataFrame(columns=cols)
                df.to_excel(EXCEL_FILE, index=False)
                formatar_excel(EXCEL_FILE)
                os.startfile(EXCEL_FILE)
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao criar ficheiro: {e}")

if __name__ == "__main__":
    app = App()
    app.mainloop()