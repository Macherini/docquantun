import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import datetime
import json
import os
import sys
import uuid
import shutil
import subprocess
import copy
import threading
import time
import csv
import logging # Módulo de logging importado
import argparse
import getpass
import traceback
from plyer import notification
from PIL import Image, ImageTk
from cryptography.fernet import Fernet
import hashlib
import base64
import math
import os.path
import pytz
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter, A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.units import inch
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
from pathlib import Path
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import numpy as np

# --- NOVA DEFINIÇÃO DE BASE_DIR ---
# Define o diretório base dentro de Central DocQuantum/DocQuantum Calendario
BASE_DIR = Path.home() / "Documents" / "Central DocQuantum" / "DocQuantum Calendario"
# Garante que o diretório base e seus pais existam
BASE_DIR.mkdir(parents=True, exist_ok=True)
# --- FIM DA NOVA DEFINIÇÃO ---

# --- Configuração Google Agenda ---
SCOPES = ['https://www.googleapis.com/auth/calendar.events']
CREDENTIALS_FILE = os.path.join(BASE_DIR, 'credentials.json') # O arquivo que você baixou
TOKEN_FILE = os.path.join(BASE_DIR, 'token.json') # Arquivo para salvar a autorização

# --- Configuração de Logging Estruturado ---
LOG_FILE = os.path.join(BASE_DIR, "calendario.log")
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(LOG_FILE, encoding='utf-8'),
        logging.StreamHandler() # Opcional: para ver logs no console também
    ]
)
logger = logging.getLogger(__name__)
logger.info(f"Diretório base do calendário definido como: {BASE_DIR}")
# -----------------------------------------

logger.info("Iniciando calendário moderno otimizado...")

ARQUIVO_EVENTOS = os.path.join(BASE_DIR, "eventos.json")
ARQUIVO_CONFIG = os.path.join(BASE_DIR, "config.json")
ARQUIVO_CHAVE = os.path.join(BASE_DIR, ".key")
ARQUIVO_METAS = os.path.join(BASE_DIR, "metas.json")
ARQUIVO_TEMPLATES = os.path.join(BASE_DIR, "templates.json")
PASTA_BACKUPS = os.path.join(BASE_DIR, "backups")
ALERTAS_OFFSET = [0, 1, 3, 7, 14, 21, 30]
PASTA_ANEXOS = os.path.join(BASE_DIR, "anexos")
EVENTOS_POR_PAGINA = 20

# Cores para os temas
CORES = {
    "light": {
        "primary": "#FFFFFF", "secondary": "#F5F5F5", "accent": "#3E7EBF", "text": "#000000",
        "card": "#E8E8E8", "card_hover": "#D8D8D8", "entry_bg": "#FFFFFF", "entry_fg": "#000000",
        "button_bg": "#E0E0E0", "button_fg": "#000000", "button_active": "#D0D0D0",
        "priority_high": "#F44336", "priority_medium": "#FFC107", "priority_low": "#E8E8E8",
        "status_success_bg": "#4CAF50", "status_danger_bg": "#F44336", "status_danger_hover": "#B71C1C",
        "status_today_bg": "#FFC107", "status_today_fg": "#000000", "status_tomorrow_bg": "#FFEB3B",
        "status_tomorrow_fg": "#000000", "status_3_days_bg": "#42A5F5", "status_3_days_fg": "#000000",
        "status_7_days_bg": "#B71C1C", "status_7_days_fg": "#FFFFFF",
        "status_14_days_bg": "#ed5c53", "status_21_days_bg": "#c501e2",
        "status_30_days_bg": "#4c007d", "status_30_days_fg": "#FFFFFF",
        "progress_bar_bg": "#CCCCCC", "progress_30_plus": "#2E7D32", "progress_7_30": "#F57C00",
        "progress_3_7": "#D84315", "progress_0_3": "#C62828", "progress_past": "#616161",
        "progress_completed": "#1565C0",
        "highlight_bg": "#FFF59D",
        "chart_1": "#4CAF50", "chart_2": "#2196F3", "chart_3": "#FFC107",
        "chart_4": "#F44336", "chart_5": "#9C27B0"
    },
    "dark": {
        "primary": "#2A2D37", "secondary": "#1E2026", "accent": "#3E7EBF", "text": "#FFFFFF",
        "card": "#3B3B3B", "card_hover": "#4A4A4A", "entry_bg": "#404040", "entry_fg": "#FFFFFF",
        "button_bg": "#404040", "button_fg": "#FFFFFF", "button_active": "#505050",
        "priority_high": "#F44336", "priority_medium": "#FF9800", "priority_low": "#3B3B3B",
        "status_success_bg": "#4CAF50", "status_danger_bg": "#F44336", "status_danger_hover": "#B71C1C",
        "status_today_bg": "#FF8C00", "status_today_fg": "#000000", "status_tomorrow_bg": "#FFD700",
        "status_tomorrow_fg": "#000000", "status_3_days_bg": "#1E90FF", "status_3_days_fg": "#FFFFFF",
        "status_7_days_bg": "#C62828", "status_7_days_fg": "#FFFFFF",
        "status_14_days_bg": "#ed5c53", "status_21_days_bg": "#c501e2",
        "status_30_days_bg": "#4c007d", "status_30_days_fg": "#FFFFFF",
        "progress_bar_bg": "#555555", "progress_30_plus": "#4CAF50", "progress_7_30": "#FFB300",
        "progress_3_7": "#FF7043", "progress_0_3": "#F44336", "progress_past": "#757575",
        "progress_completed": "#1976D2",
        "highlight_bg": "#6F6C47",
        "chart_1": "#4CAF50", "chart_2": "#2196F3", "chart_3": "#FFC107",
        "chart_4": "#F44336", "chart_5": "#9C27B0"
    }
}

# NOVA PALETA DE CORES PARA OS CARDS
NEW_STATUS_COLORS = {
    "base": {
        "completed": "#166534", # Verde
        "past":      "#D32F2F", # Vermelho Suave
        "today":     "#ff6961", # Laranja/Salmão
        "tomorrow":  "#FFFACD", # Amarelo (NOVO)
        "3_days":    "#fa34df", # Roxo - ALTERADO
        "7_days":    "#6B21A8", # Roxo
        "14_days":   "#581C1C", # Vinho
        "21_days":   "#0F0F0F", # Preto
        "30_days":   "#1E3A8A", # Azul Escuro
        "default":   "#3B3B3B", # Cinza padrão
    },
    "light": { # Usado como base para ambos os temas
        "bottom": "#FDFDFB", # Branco Pérola
        "text_top": "#FFFFFF"
    }
}


# Garantir que os diretórios necessários existam
for pasta in [PASTA_ANEXOS, PASTA_BACKUPS]:
    if not os.path.exists(pasta):
        try:
            os.makedirs(pasta)
            logger.info(f"Diretório '{pasta}' criado com sucesso.")
        except OSError as e:
            logger.error(f"Não foi possível criar o diretório '{pasta}': {e}")

class ConfigManager:
    """Gerenciador de configurações da aplicação"""

    def __init__(self):
        self.modo_aparencia = "light"
        self.modo_tela_cheia = False
        self.uppercase_enabled = True
        self.alertas_personalizados = [0, 1, 3, 7, 14, 21, 30] # Usando a lista atualizada diretamente
        self.carregar()
        logger.info("ConfigManager inicializado.")

    def carregar(self):
        default_alertas = [0, 1, 3, 7, 14, 21, 30] # Definindo o padrão aqui também
        if os.path.exists(ARQUIVO_CONFIG):
            try:
                with open(ARQUIVO_CONFIG, "r", encoding="utf-8") as f:
                    config = json.load(f)
                    self.modo_aparencia = config.get("appearance_mode", "light")
                    self.modo_tela_cheia = config.get("fullscreen_mode", False)
                    self.uppercase_enabled = config.get("uppercase_enabled", True)
                    self.alertas_personalizados = config.get("alertas_personalizados", ALERTAS_OFFSET.copy())
                logger.info("Configurações carregadas do arquivo.")
            except Exception as e:
                logger.error(f"Erro ao carregar configurações de {ARQUIVO_CONFIG}: {e}. Usando padrões.")
                self.modo_aparencia = "light"
                self.modo_tela_cheia = False
                self.uppercase_enabled = True
                self.alertas_personalizados = default_alertas
        else:
            logger.info(f"Arquivo de configuração '{ARQUIVO_CONFIG}' não encontrado. Usando configurações padrão.")

        if self.modo_aparencia not in CORES:
            self.modo_aparencia = "light"

    def salvar(self):
        try:
            with open(ARQUIVO_CONFIG, "w", encoding="utf-8") as f:
                json.dump({
                    "appearance_mode": self.modo_aparencia,
                    "fullscreen_mode": self.modo_tela_cheia,
                    "uppercase_enabled": self.uppercase_enabled,
                    "alertas_personalizados": self.alertas_personalizados
                }, f, ensure_ascii=False, indent=2)
            logger.info("Configurações salvas com sucesso.")
        except Exception as e:
            logger.error(f"Erro ao salvar o arquivo de configuração: {e}")
            messagebox.showerror("Erro", f"Não foi possível salvar as configurações: {e}")


class CryptManager:
    """Gerenciador de criptografia - DESABILITADO"""

    def __init__(self):
        self.criptografia_habilitada = False
        self.chave_criptografia = None
        logger.info("CryptManager inicializado com criptografia DESABILITADA.")

    def gerar_chave_de_senha(self, senha):
        """Função mantida por compatibilidade, mas não utilizada."""
        chave = hashlib.pbkdf2_hmac('sha256', senha.encode(), b'salt_estatico', 100000)
        return base64.urlsafe_b64encode(chave)

    def carregar_ou_criar_chave(self):
        """Sempre retorna True e garante que a criptografia está desabilitada."""
        self.criptografia_habilitada = False
        logger.info("Criptografia DESABILITADA permanentemente.")
        
        # Deleta o arquivo .key antigo se ele existir, para evitar confusão
        if os.path.exists(ARQUIVO_CHAVE):
            try:
                os.remove(ARQUIVO_CHAVE)
                logger.info("Arquivo de chave (.key) antigo removido.")
            except Exception as e:
                logger.warning(f"Não foi possível remover o arquivo .key antigo: {e}")
                
        return True

    def verificar_senha(self):
        """Sempre retorna True, nunca pedindo a senha."""
        return True

    def criptografar_dados(self, dados):
        """Retorna os dados originais, sem criptografar."""
        logger.info("Salvando dados sem criptografia.")
        return dados

    def descriptografar_dados(self, dados):
        """Verifica se os dados estão criptografados e avisa o usuário se estiverem."""
        
        # Se os dados parecem estar no formato antigo criptografado
        if isinstance(dados, dict) and dados.get("encrypted", False):
            logger.error("ERRO: Os dados em 'eventos.json' estão criptografados, mas a função de senha está desabilitada.")
            messagebox.showerror("Erro de Carga", 
                                 "Seus dados estão criptografados, mas a senha foi desabilitada no código.\n\n"
                                 "Para corrigir, delete o arquivo 'eventos.json' e comece de novo (seus dados atuais serão perdidos). "
                                 "Se você fez um backup, pode restaurá-lo.")
            return [] # Retorna uma lista vazia para evitar que o programa quebre

        # Se não estiverem criptografados, apenas retorna os dados como estão
        return dados


class BackupManager:
    """Gerenciador de backups"""

    def __init__(self):
        if not os.path.exists(PASTA_BACKUPS):
            os.makedirs(PASTA_BACKUPS)
        logger.info("BackupManager inicializado.")

    def criar_backup(self, eventos):
        """Cria um backup dos eventos"""
        try:
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            arquivo_backup = os.path.join(PASTA_BACKUPS, f"backup_{timestamp}.json")

            with open(arquivo_backup, 'w', encoding='utf-8') as f:
                json.dump(eventos, f, ensure_ascii=False, indent=4, default=str)

            logger.info(f"Backup criado com sucesso em: {arquivo_backup}")
            self.limitar_backups()

            return arquivo_backup
        except Exception as e:
            logger.error(f"Erro ao criar backup: {e}")
            messagebox.showerror("Erro", f"Erro ao criar backup: {str(e)}")
            return None

    def listar_backups(self):
        """Lista todos os backups disponíveis"""
        backups = []
        if os.path.exists(PASTA_BACKUPS):
            for arquivo in os.listdir(PASTA_BACKUPS):
                if arquivo.startswith("backup_") and arquivo.endswith(".json"):
                    caminho = os.path.join(PASTA_BACKUPS, arquivo)
                    stats = os.stat(caminho)
                    backups.append({
                        'arquivo': caminho,
                        'nome': arquivo,
                        'data': datetime.datetime.fromtimestamp(stats.st_mtime),
                        'tamanho': stats.st_size
                    })
        backups.sort(key=lambda x: x['data'], reverse=True)
        return backups

    def restaurar_backup(self, caminho_backup):
        """Restaura um backup"""
        try:
            with open(caminho_backup, 'r', encoding='utf-8') as f:
                eventos = json.load(f)
            logger.info(f"Backup '{caminho_backup}' carregado para restauração.")
            return eventos
        except Exception as e:
            logger.error(f"Erro ao restaurar backup '{caminho_backup}': {e}")
            messagebox.showerror("Erro", f"Erro ao restaurar backup: {str(e)}")
            return None

    def limitar_backups(self, max_backups=5):
        """Limita o número de backups mantidos"""
        backups = self.listar_backups()
        if len(backups) > max_backups:
            for backup in backups[max_backups:]:
                try:
                    os.remove(backup['arquivo'])
                    logger.info(f"Backup antigo removido: {backup['arquivo']}")
                except Exception as e:
                    logger.warning(f"Erro ao remover backup antigo '{backup['arquivo']}': {e}")

class TemplateManager:
    """Gerenciador de templates de eventos"""

    def __init__(self):
        self.templates = self.carregar_templates()
        logger.info("TemplateManager inicializado.")

    def carregar_templates(self):
        """Carrega templates do arquivo"""
        if os.path.exists(ARQUIVO_TEMPLATES):
            try:
                with open(ARQUIVO_TEMPLATES, "r", encoding="utf-8") as f:
                    templates = json.load(f)
                    logger.info(f"Templates carregados: {len(templates)}")
                    return templates
            except Exception as e:
                logger.error(f"Erro ao carregar templates: {e}. Iniciando com lista vazia.")
                return []
        return []

    def salvar_templates(self):
        """Salva templates no arquivo"""
        try:
            with open(ARQUIVO_TEMPLATES, "w", encoding="utf-8") as f:
                json.dump(self.templates, f, ensure_ascii=False, indent=4)
            logger.info(f"Templates salvos com sucesso. Total: {len(self.templates)}")
        except Exception as e:
            logger.error(f"Erro ao salvar templates: {e}")

    def adicionar_template(self, evento):
        """Adiciona um novo template baseado em um evento"""
        template = {
            'nome': evento.get('nome', ''),
            'local': evento.get('local', ''),
            'corpo': evento.get('corpo', ''),
            'categoria': evento.get('categoria', ''),
            'prioridade': evento.get('prioridade', ''),
            'grupo': evento.get('grupo', ''),
            'etiquetas': evento.get('etiquetas', ''),
            'hora': evento.get('hora', 0),
            'minuto': evento.get('minuto', 0)
        }

        self.templates.append(template)
        self.salvar_templates()
        logger.info(f"Novo template '{template['nome']}' adicionado.")
        return template

    def aplicar_template(self, template_id, data):
        """Aplica um template a uma data específica"""
        if 0 <= template_id < len(self.templates):
            template = self.templates[template_id].copy()
            template['data'] = data
            template['id'] = str(uuid.uuid4())
            template['alertas_enviados'] = {}
            template['concluido'] = False
            template['data_criacao'] = datetime.date.today()
            template['data_prorrogacao'] = None
            logger.info(f"Template '{template['nome']}' aplicado.")
            return template
        logger.warning(f"Tentativa de aplicar template com ID inválido: {template_id}")
        return None

    def remover_template(self, template_id):
        """Remove um template"""
        if 0 <= template_id < len(self.templates):
            template_removido = self.templates.pop(template_id)
            self.salvar_templates()
            logger.info(f"Template '{template_removido['nome']}' removido.")
            return True
        logger.warning(f"Tentativa de remover template com ID inválido: {template_id}")
        return False

class MetaManager:
    """Gerenciador de metas"""

    def __init__(self):
        self.metas = self.carregar_metas()
        logger.info("MetaManager inicializado.")

    def carregar_metas(self):
        """Carrega metas do arquivo"""
        if os.path.exists(ARQUIVO_METAS):
            try:
                with open(ARQUIVO_METAS, "r", encoding="utf-8") as f:
                    metas = json.load(f)
                    logger.info(f"Metas carregadas: {len(metas)}")
                    return metas
            except Exception as e:
                logger.error(f"Erro ao carregar metas: {e}. Iniciando com lista vazia.")
                return []
        return []

    def salvar_metas(self):
        """Salva metas no arquivo"""
        try:
            with open(ARQUIVO_METAS, "w", encoding="utf-8") as f:
                json.dump(self.metas, f, ensure_ascii=False, indent=4)
            logger.info(f"Metas salvas com sucesso. Total: {len(self.metas)}")
        except Exception as e:
            logger.error(f"Erro ao salvar metas: {e}")

    def adicionar_meta(self, tipo, valor, periodo, descricao=""):
        """Adiciona uma nova meta"""
        meta = {
            'id': str(uuid.uuid4()),
            'tipo': tipo,
            'valor': valor,
            'periodo': periodo,
            'descricao': descricao,
            'data_criacao': datetime.date.today().isoformat(),
            'progresso': 0,
            'concluida': False
        }

        self.metas.append(meta)
        self.salvar_metas()
        logger.info(f"Nova meta '{descricao}' adicionada.")
        return meta

    def atualizar_progresso(self, evento_manager):
        """Atualiza o progresso de todas as metas"""
        hoje = datetime.date.today()
        logger.info("Atualizando progresso das metas.")
        for meta in self.metas:
            if meta['concluida']:
                continue

            # Determinar período de análise
            if meta['periodo'] == 'diario':
                data_inicio = hoje
                data_fim = hoje
            elif meta['periodo'] == 'semanal':
                data_inicio = hoje - datetime.timedelta(days=hoje.weekday())
                data_fim = data_inicio + datetime.timedelta(days=6)
            else:  # mensal
                data_inicio = hoje.replace(day=1)
                ultimo_dia = (data_inicio.replace(month=data_inicio.month % 12 + 1, day=1) -
                             datetime.timedelta(days=1))
                data_fim = ultimo_dia

            # Calcular progresso baseado no tipo de meta
            if meta['tipo'] == 'eventos_concluidos':
                eventos_periodo = [
                    e for e in evento_manager.eventos
                    if data_inicio <= evento_manager.get_effective_date(e) <= data_fim and e['concluido']
                ]
                meta['progresso'] = len(eventos_periodo)

            elif meta['tipo'] == 'eventos_criados':
                eventos_periodo = [
                    e for e in evento_manager.eventos
                    if 'data_criacao' in e and data_inicio <= e['data_criacao'] <= data_fim
                ]
                meta['progresso'] = len(eventos_periodo)

            elif meta['tipo'] == 'produtividade':
                eventos_concluidos = [
                    e for e in evento_manager.eventos
                    if data_inicio <= evento_manager.get_effective_date(e) <= data_fim and e['concluido']
                ]
                eventos_totais = [
                    e for e in evento_manager.eventos
                    if data_inicio <= evento_manager.get_effective_date(e) <= data_fim
                ]
                meta['progresso'] = (len(eventos_concluidos) / len(eventos_totais) * 100) if eventos_totais else 0

            # Verificar se meta foi atingida
            if meta['progresso'] >= meta['valor']:
                if not meta['concluida']:
                    logger.info(f"Meta '{meta['descricao']}' foi concluída!")
                meta['concluida'] = True

        self.salvar_metas()

    def remover_meta(self, meta_id):
        """Remove uma meta"""
        meta_removida = next((m for m in self.metas if m['id'] == meta_id), None)
        self.metas = [m for m in self.metas if m['id'] != meta_id]
        self.salvar_metas()
        if meta_removida:
            logger.info(f"Meta '{meta_removida['descricao']}' removida.")
            return True
        return False

class EventoManager:
    """Gerenciador de eventos"""

    def __init__(self, crypt_manager):
        self.eventos = []
        self.crypt_manager = crypt_manager
        self.history = []
        self.history_index = -1
        self.max_history = 20
        self.backup_manager = BackupManager()
        logger.info("EventoManager inicializado.")

    def get_effective_date(self, evento):
        """Retorna a data de prorrogação se existir, senão a data original."""
        if evento.get("data_prorrogacao"):
            return evento["data_prorrogacao"]
        return evento["data"]

    def adicionar_ao_historico(self):
        """Adiciona estado atual ao histórico para undo/redo"""
        if self.history_index < len(self.history) - 1:
            self.history = self.history[:self.history_index + 1]
        estado = copy.deepcopy(self.eventos)
        self.history.append(estado)
        if len(self.history) > self.max_history:
            self.history.pop(0)
        self.history_index = len(self.history) - 1

    def undo(self):
        """Desfaz a última operação"""
        if self.history_index > 0:
            self.history_index -= 1
            self.eventos = copy.deepcopy(self.history[self.history_index])
            logger.info("Operação 'Desfazer' (Undo) executada.")
            return True
        logger.warning("Nenhuma operação para desfazer no histórico.")
        return False

    def redo(self):
        """Refaz a operação desfeita"""
        if self.history_index < len(self.history) - 1:
            self.history_index += 1
            self.eventos = copy.deepcopy(self.history[self.history_index])
            logger.info("Operação 'Refazer' (Redo) executada.")
            return True
        logger.warning("Nenhuma operação para refazer no histórico.")
        return False

    def garantir_campos_evento(self, e):
        defaults = {
            "id": str(uuid.uuid4()), "alertas_enviados": {}, "local": "", "corpo": "",
            "hora": 0, "minuto": 0, "concluido": False, "categoria": "",
            "anexos": [], "prioridade": "", "grupo": "", "etiquetas": "",
            "data_criacao": datetime.date.today().isoformat(),
            "data_prorrogacao": None,
            "google_event_id": None # Adiciona o campo padrão aqui
        }
        # ... (resto da função continua igual, convertendo datas str para date) ...
        for key, value in defaults.items():
            e.setdefault(key, value) # Garante que a chave exista

        # Conversões de data (MANTENHA ESTA PARTE)
        if isinstance(e.get("data"), str):
            try: e["data"] = datetime.datetime.strptime(e["data"], "%Y-%m-%d").date()
            except ValueError: e["data"] = datetime.date.today() # Fallback seguro
        if isinstance(e.get("data_criacao"), str):
            try: e["data_criacao"] = datetime.datetime.strptime(e["data_criacao"], "%Y-%m-%d").date()
            except ValueError: e["data_criacao"] = datetime.date.today() # Fallback
        if isinstance(e.get("data_prorrogacao"), str):
             try: e["data_prorrogacao"] = datetime.datetime.strptime(e["data_prorrogacao"], "%Y-%m-%d").date()
             except ValueError: e["data_prorrogacao"] = None # Se inválido, remove

        return e

    def limpar_eventos_antigos(self):
        """Remove eventos passados com mais de 60 dias e eventos concluídos com mais de 60 dias"""
        hoje = datetime.date.today()
        len_antes = len(self.eventos)
        self.eventos = [e for e in self.eventos if not (
            (e["concluido"] and (hoje - self.get_effective_date(e)).days > 60) or
            (not e["concluido"] and (hoje - self.get_effective_date(e)).days > 60)
        )]
        removidos = len_antes - len(self.eventos)
        if removidos > 0:
            logger.info(f"{removidos} evento(s) antigo(s) foram limpos.")


    def limpar_alertas_antigos(self):
        """Remove alertas enviados com mais de 30 dias"""
        data_limite = (datetime.date.today() - datetime.timedelta(days=30)).isoformat()
        for evento in self.eventos:
            alertas_limpos = {}
            for data_str, alertas in evento["alertas_enviados"].items():
                if data_str > data_limite:
                    alertas_limpos[data_str] = alertas
            evento["alertas_enviados"] = alertas_limpos

    def carregar_eventos(self):
        if os.path.exists(ARQUIVO_EVENTOS):
            try:
                with open(ARQUIVO_EVENTOS, "r", encoding="utf-8") as f:
                    data = json.load(f)

                data = self.crypt_manager.descriptografar_dados(data)
                self.eventos = [self.garantir_campos_evento(e) for e in data]
                logger.info(f"Eventos carregados: {len(self.eventos)}")
                self.limpar_eventos_antigos()
                self.limpar_alertas_antigos()
                self.adicionar_ao_historico()
            except (json.JSONDecodeError, TypeError) as e:
                logger.warning(f"Erro ao decodificar {ARQUIVO_EVENTOS}: {e}. Iniciando com lista vazia.")
                self.eventos = []
        else:
            logger.info(f"Arquivo de eventos '{ARQUIVO_EVENTOS}' não encontrado. Iniciando com lista vazia.")
            self.eventos = []

    def salvar_eventos(self):
        try:
            event_list_to_save = []
            for e in self.eventos:
                event_copy = e.copy()
                event_copy["data"] = e["data"].strftime("%Y-%m-%d")
                event_copy["data_criacao"] = e["data_criacao"].strftime("%Y-%m-%d")
                if e.get("data_prorrogacao"):
                    event_copy["data_prorrogacao"] = e["data_prorrogacao"].strftime("%Y-%m-%d")
                event_list_to_save.append(event_copy)

            data_to_save = self.crypt_manager.criptografar_dados(event_list_to_save)

            with open(ARQUIVO_EVENTOS, "w", encoding="utf-8") as f:
                json.dump(data_to_save, f, ensure_ascii=False, indent=4)
            logger.info(f"Eventos salvos com sucesso. Total: {len(self.eventos)}")

            self.backup_manager.criar_backup(event_list_to_save)
        except Exception as e:
            logger.error(f"Erro ao salvar eventos: {e}")

    def adicionar_evento(self, evento_data):
        novo_evento = {
            **evento_data,
            "id": str(uuid.uuid4()),
            "alertas_enviados": {},
            "concluido": False,
            "data_criacao": datetime.date.today(),
            "google_event_id": None # Inicializa o campo para o ID do Google
        }
        self.eventos.append(novo_evento)

        # Chama a sincronização e tenta obter o ID do Google
        sync_success, google_event_id = adicionar_evento_google_calendar(novo_evento)
        if sync_success and google_event_id:
            novo_evento['google_event_id'] = google_event_id # Salva o ID no evento local
            logger.info(f"ID do Google Calendar ({google_event_id}) associado ao evento local.")
        else:
            logger.warning(f"Não foi possível obter o ID do Google Calendar para o evento '{novo_evento['nome']}'. Edições/Exclusões não serão sincronizadas.")

        self.adicionar_ao_historico()
        logger.info(f"Evento '{novo_evento['nome']}' (ID local: {novo_evento['id']}) adicionado.")
        return novo_evento # Retorna o evento completo com o ID do Google (ou None)

    def atualizar_evento(self, evento_id, dados):
        evento_atualizado = None # Variável para guardar o evento após atualização local
        for e in self.eventos:
            if e["id"] == evento_id:
                google_event_id = e.get('google_event_id') # Pega o ID do Google antes de atualizar

                # Tenta atualizar no Google ANTES de atualizar localmente
                if google_event_id:
                    # Cria um dicionário temporário com os dados atualizados para enviar ao Google
                    dados_completos_para_google = e.copy() # Copia o evento existente
                    dados_completos_para_google.update(dados) # Aplica as novas alterações
                    
                    sync_success = atualizar_evento_google_calendar(google_event_id, dados_completos_para_google)
                    if not sync_success:
                        logger.warning(f"Falha ao sincronizar atualização do evento '{e['nome']}' com Google Calendar. Alteração será salva apenas localmente.")
                        # Opcional: Mostrar um aviso ao usuário na GUI
                        # messagebox.showwarning("Falha na Sincronização", "Não foi possível atualizar o evento no Google Agenda. A alteração foi salva apenas localmente.")
                else:
                     logger.warning(f"Evento '{e['nome']}' não possui ID do Google Calendar. Atualização não será sincronizada.")

                # Atualiza localmente DEPOIS de tentar sincronizar
                e.update(dados)
                evento_atualizado = e # Guarda a referência
                break # Sai do loop após encontrar e atualizar

        if evento_atualizado:
            self.adicionar_ao_historico()
            logger.info(f"Evento '{evento_atualizado['nome']}' (ID local: {evento_id}) atualizado localmente.")
            return evento_atualizado # Retorna o evento atualizado
        else:
            logger.warning(f"Tentativa de atualizar evento local com ID não encontrado: {evento_id}")
            return None

    def remover_evento(self, evento_id):
        evento_removido = self.buscar_evento_por_id(evento_id)
        google_event_id = None
        sync_success = True # Assume sucesso se não houver ID

        if evento_removido:
            google_event_id = evento_removido.get('google_event_id')
            if google_event_id:
                sync_success = excluir_evento_google_calendar(google_event_id)
                if not sync_success:
                     logger.warning(f"Falha ao sincronizar exclusão do evento '{evento_removido['nome']}' com Google Calendar. Evento será removido apenas localmente.")
                     # Opcional: Perguntar ao usuário se deseja remover localmente mesmo assim
                     # if not messagebox.askyesno("Falha na Sincronização", "Não foi possível excluir o evento no Google Agenda.\nDeseja remover o evento apenas localmente?", icon='warning'):
                     #     logger.info("Remoção local cancelada pelo usuário devido à falha na sincronização.")
                     #     return # Não remove localmente
            else:
                 logger.warning(f"Evento '{evento_removido['nome']}' não possui ID do Google Calendar. Exclusão não será sincronizada.")

        # Remove localmente APENAS se a sincronização deu certo OU se não havia ID OU se o usuário confirmou após falha (se implementado)
        if sync_success:
            self.eventos = [ev for ev in self.eventos if ev["id"] != evento_id]
            self.adicionar_ao_historico()
            if evento_removido:
                logger.info(f"Evento '{evento_removido['nome']}' (ID local: {evento_id}) removido localmente.")
        # else: # Se sync_success for False e não removemos localmente
        #     logger.info(f"Remoção local do evento '{evento_removido['nome']}' abortada devido à falha na sincronização com Google Calendar.")

    def remover_todos_eventos(self):
        """Remove todos os eventos"""
        total_removidos = len(self.eventos)
        self.eventos = []
        self.adicionar_ao_historico()
        logger.warning(f"Todos os {total_removidos} eventos foram removidos.")
        return True

    def buscar_evento_por_id(self, evento_id):
        return next((e for e in self.eventos if e["id"] == evento_id), None)

    def verificar_conflitos_agenda(self, novo_evento, evento_id_ignorar=None):
        """Verifica se há conflitos de horário entre o novo evento e os eventos existentes"""
        conflitos = []
        data_novo = self.get_effective_date(novo_evento)
        hora_novo = novo_evento["hora"]
        minuto_novo = novo_evento["minuto"]

        for evento in self.eventos:
            if evento_id_ignorar and evento["id"] == evento_id_ignorar:
                continue

            if self.get_effective_date(evento) == data_novo:
                hora_existente = evento["hora"]
                minuto_existente = evento["minuto"]

                tempo_novo = hora_novo * 60 + minuto_novo
                tempo_existente = hora_existente * 60 + minuto_existente

                if abs(tempo_novo - tempo_existente) < 60:
                    conflitos.append(evento)
        
        if conflitos:
            logger.warning(f"Detectado conflito de agenda para o evento '{novo_evento['nome']}'. Conflitos: {len(conflitos)}")
        return conflitos

    def filtrar_eventos(self, filtro="Todos", termo_busca="", ordem="Data", criterios_avancados=None):
        hoje = datetime.date.today()

        filtros_data = {
            "Todos": lambda e: self.get_effective_date(e) >= hoje,
            "Somente hoje": lambda e: self.get_effective_date(e) == hoje,
            "Vence em 1 dia": lambda e: self.get_effective_date(e) == hoje + datetime.timedelta(days=1),
            "Próximos 3 dias": lambda e: hoje <= self.get_effective_date(e) <= hoje + datetime.timedelta(days=3),
            "Próximos 7 dias": lambda e: hoje <= self.get_effective_date(e) <= hoje + datetime.timedelta(days=7),
            "Próximos 14 dias": lambda e: hoje <= self.get_effective_date(e) <= hoje + datetime.timedelta(days=14),
            "Próximos 21 dias": lambda e: hoje <= self.get_effective_date(e) <= hoje + datetime.timedelta(days=21),
            "Próximos 30 dias": lambda e: hoje <= self.get_effective_date(e) <= hoje + datetime.timedelta(days=30),
            "Passados": lambda e: self.get_effective_date(e) < hoje,
            "Concluídos": lambda e: e["concluido"]
        }

        if filtro in filtros_data:
            lista_filtrada = [e for e in self.eventos if filtros_data[filtro](e)]
        else:
            lista_filtrada = self.eventos[:]

        # Aplicar filtros avançados se fornecidos
        if criterios_avancados:
            # Tratar filtro de data primeiro
            data_inicio_str = criterios_avancados.get("data_inicio")
            data_fim_str = criterios_avancados.get("data_fim")
            if data_inicio_str and data_fim_str:
                try:
                    data_inicio = datetime.datetime.strptime(data_inicio_str, "%d/%m/%Y").date()
                    data_fim = datetime.datetime.strptime(data_fim_str, "%d/%m/%Y").date()
                    lista_filtrada = [e for e in lista_filtrada if data_inicio <= self.get_effective_date(e) <= data_fim]
                except ValueError:
                    messagebox.showerror("Erro de Formato", "Data inválida na pesquisa avançada. Use o formato DD/MM/AAAA.")
                    pass

            # Tratar outros filtros
            for campo, valor in criterios_avancados.items():
                if valor and campo not in ["data_inicio", "data_fim"]:
                    if campo == "prioridade":
                        lista_filtrada = [e for e in lista_filtrada if e.get("prioridade") == valor]
                    elif campo == "categoria":
                        lista_filtrada = [e for e in lista_filtrada if valor.lower() in e.get("categoria", "").lower()]
                    elif campo == "concluido":
                        concluido = valor.lower() == "sim"
                        lista_filtrada = [e for e in lista_filtrada if e["concluido"] == concluido]

        if termo_busca:
            busca_lower = termo_busca.lower()
            campos_busca = ["nome", "local", "categoria", "corpo", "grupo", "etiquetas", "prioridade"]
            lista_filtrada = [e for e in lista_filtrada if
                any(busca_lower in str(e.get(campo, '')).lower() for campo in campos_busca) or
                busca_lower in self.gerar_card_id(e).lower()]

        # Lógica de Ordenação
        if ordem == "Prioridade":
            prioridade_map = {"ALTA": 0, "MÉDIA": 1, "BAIXA": 2}
            return sorted(lista_filtrada, key=lambda e: (prioridade_map.get(e.get("prioridade", ""), 3), self.get_effective_date(e)))
        elif ordem == "Nome (A-Z)":
            return sorted(lista_filtrada, key=lambda e: e["nome"])
        elif ordem == "Data de Criação":
            return sorted(lista_filtrada, key=lambda e: e["data_criacao"], reverse=True)
        else:  # Padrão é "Data"
            return sorted(lista_filtrada, key=lambda e: (self.get_effective_date(e), e["hora"], e["minuto"]))

    def gerar_card_id(self, evento):
        """Gera o ID do card para busca"""
        nome_clean = ''.join(c for c in evento['nome'] if c.isalpha())[:3].upper()
        dia = self.get_effective_date(evento).day
        return f"{nome_clean}{dia:02d}"

    def gerar_relatorio(self, tipo_relatorio="mensal", data_inicio=None, data_fim=None):
        """Gera relatório de eventos"""
        logger.info(f"Gerando relatório do tipo '{tipo_relatorio}'.")
        hoje = datetime.date.today()

        # Definir período do relatório
        if tipo_relatorio == "mensal":
            data_inicio = hoje.replace(day=1)
            data_fim = (data_inicio.replace(month=data_inicio.month % 12 + 1, day=1) -
                       datetime.timedelta(days=1))
        elif tipo_relatorio == "semanal":
            data_inicio = hoje - datetime.timedelta(days=hoje.weekday())
            data_fim = data_inicio + datetime.timedelta(days=6)
        elif tipo_relatorio == "diario":
            data_inicio = hoje
            data_fim = hoje
        else:  # personalizado
            if not data_inicio or not data_fim:
                data_inicio = hoje.replace(day=1)
                data_fim = hoje

        # Filtrar eventos pelo período
        eventos_periodo = [
            e for e in self.eventos
            if data_inicio <= self.get_effective_date(e) <= data_fim
        ]

        # Calcular estatísticas
        total_eventos = len(eventos_periodo)
        eventos_concluidos = sum(1 for e in eventos_periodo if e["concluido"])
        taxa_conclusao = (eventos_concluidos / total_eventos * 100) if total_eventos > 0 else 0

        # Estatísticas por prioridade
        prioridades = {"ALTA": 0, "MÉDIA": 0, "BAIXA": 0, "NÃO DEFINIDA": 0}
        for e in eventos_periodo:
            prioridade = e.get("prioridade", "")
            if prioridade in prioridades:
                prioridades[prioridade] += 1
            else:
                prioridades["NÃO DEFINIDA"] += 1

        # Estatísticas por categoria
        categorias = {}
        for e in eventos_periodo:
            categoria = e.get("categoria", "Sem categoria")
            categorias[categoria] = categorias.get(categoria, 0) + 1

        # Eventos mais próximos (próximos 7 dias)
        eventos_proximos = [
            e for e in self.eventos
            if hoje <= self.get_effective_date(e) <= hoje + datetime.timedelta(days=7) and not e["concluido"]
        ]

        relatorio = {
            "periodo": {
                "inicio": data_inicio,
                "fim": data_fim,
                "tipo": tipo_relatorio
            },
            "estatisticas": {
                "total_eventos": total_eventos,
                "eventos_concluidos": eventos_concluidos,
                "taxa_conclusao": round(taxa_conclusao, 2),
                "prioridades": prioridades,
                "categorias": categorias
            },
            "eventos_proximos": eventos_proximos[:5],  # Top 5 eventos próximos
            "eventos_importantes": [
                e for e in eventos_periodo
                if e.get("prioridade") == "ALTA" and not e["concluido"]
            ]
        }

        return relatorio

    def analisar_produtividade(self, periodo="semanal"):
        """Analisa padrões de produtividade"""
        logger.info(f"Analisando produtividade para o período '{periodo}'.")
        hoje = datetime.date.today()
        dias_analise = 30 if periodo == "mensal" else 7

        dados = []
        for i in range(dias_analise):
            data = hoje - datetime.timedelta(days=i)
            eventos_dia = [e for e in self.eventos if self.get_effective_date(e) == data]

            concluidos = sum(1 for e in eventos_dia if e["concluido"])
            total = len(eventos_dia)

            dados.append({
                "data": data,
                "total_eventos": total,
                "eventos_concluidos": concluidos,
                "taxa_conclusao": (concluidos / total * 100) if total > 0 else 0
            })

        # Ordenar por data (mais antigo primeiro)
        dados.sort(key=lambda x: x["data"])

        return dados

class AlertaManager:
    """Gerenciador de alertas em thread separada"""

    def __init__(self, evento_manager, config_manager):
        self.evento_manager = evento_manager
        self.config_manager = config_manager
        self.thread_alertas = None
        self.stop_alertas = threading.Event()
        logger.info("AlertaManager inicializado.")

    def mostrar_notificacao(self, titulo, mensagem):
        try:
            notification.notify(title=titulo, message=mensagem, timeout=10)
            logger.info(f"Notificação enviada: '{titulo}' - '{mensagem}'")
        except (NotImplementedError, Exception) as e:
            logger.warning(f"Falha ao enviar notificação do sistema: {e}. Usando fallback.")
            try:
                # Fallback para messagebox se notificação falhar
                messagebox.showinfo(titulo, mensagem)
            except Exception as tk_e:
                logger.error(f"Falha no fallback de notificação (messagebox): {tk_e}")
                # Fallback final para console
                print(f"NOTIFICAÇÃO: {titulo} - {mensagem}")

    def thread_verificar_alertas(self, salvar_callback):
        """Thread separada para verificar alertas"""
        logger.info("Thread de verificação de alertas iniciada.")
        while not self.stop_alertas.is_set():
            try:
                hoje = datetime.date.today()
                hojeISO = hoje.isoformat()
                alertas_enviados = False

                for e in self.evento_manager.eventos:
                    if not e["concluido"]:
                        data_efetiva = self.evento_manager.get_effective_date(e)
                        dias_faltando = (data_efetiva - hoje).days
                        # Usar offsets personalizados da configuração
                        if dias_faltando in self.config_manager.alertas_personalizados and dias_faltando not in e["alertas_enviados"].get(hojeISO, []):
                            hora_str = f"{e['hora']:02d}:{e['minuto']:02d}"
                            msg = f"{e['nome']} é {'HOJE' if dias_faltando == 0 else f'em {dias_faltando} dia(s)'} ({self.formata_data_br(data_efetiva)} às {hora_str})"
                            self.mostrar_notificacao("📅 Alerta de Evento", msg)
                            e["alertas_enviados"].setdefault(hojeISO, []).append(dias_faltando)
                            alertas_enviados = True

                if alertas_enviados:
                    # Chamar callback para salvar
                    salvar_callback()

            except Exception as e:
                logger.error(f"Erro na thread de alertas: {e}", exc_info=True)

            # Aguardar 1 hora ou até sinal de parada
            self.stop_alertas.wait(3600)
        logger.info("Thread de verificação de alertas encerrada.")

    def formata_data_br(self, data_date):
        return data_date.strftime("%d/%m/%Y")

    def iniciar_thread_alertas(self, salvar_callback):
        """Inicia a thread de alertas"""
        if self.thread_alertas is None or not self.thread_alertas.is_alive():
            self.stop_alertas.clear()
            self.thread_alertas = threading.Thread(target=self.thread_verificar_alertas, args=(salvar_callback,), daemon=True)
            self.thread_alertas.start()
            logger.info("Thread de alertas iniciada ou reiniciada.")

    def parar_thread_alertas(self):
        """Para a thread de alertas"""
        logger.info("Solicitando parada da thread de alertas.")
        self.stop_alertas.set()
        if self.thread_alertas and self.thread_alertas.is_alive():
            self.thread_alertas.join(timeout=1)
            logger.info("Thread de alertas finalizada.")

class EventCard(tk.Canvas):
    """Card de evento reutilizável com novo layout"""

    def __init__(self, parent, evento, config_manager, evento_manager, on_click_callback, on_menu_callback, filtro_busca=""):
        super().__init__(parent, bg=CORES[config_manager.modo_aparencia]["primary"],
                        bd=0, highlightthickness=0, width=220, height=165)
        self.pack_propagate(False)
        self.evento = evento
        self.config_manager = config_manager
        self.evento_manager = evento_manager
        self.on_click_callback = on_click_callback
        self.on_menu_callback = on_menu_callback
        self.filtro_busca = filtro_busca
        self.menu_btn = None # Inicializa o atributo
        self.criar_card()

    def obter_cores_status_evento(self):
        hoje = datetime.date.today()
        data_efetiva = self.evento_manager.get_effective_date(self.evento)
        dias_faltando = (data_efetiva - hoje).days
        
        status_key = "default"
        if self.evento["concluido"]: status_key = "completed"
        elif dias_faltando < 0: status_key = "past"
        elif dias_faltando == 0: status_key = "today"
        elif dias_faltando == 1: status_key = "tomorrow"
        elif dias_faltando <= 3: status_key = "3_days"
        elif dias_faltando <= 7: status_key = "7_days"
        elif dias_faltando <= 14: status_key = "14_days"
        elif dias_faltando <= 21: status_key = "21_days"
        elif dias_faltando <= 30: status_key = "30_days"
        
        top_color = NEW_STATUS_COLORS["base"][status_key]
        
        # Define a cor do texto do topo (preto para 'tomorrow', branco para os demais)
        if status_key == "tomorrow":
            text_top_color = "#e49800"
        else:
            text_top_color = NEW_STATUS_COLORS["light"]["text_top"]

        # Define as cores da parte inferior (sempre fundo branco pérola, texto na cor do topo)
        bottom_color = NEW_STATUS_COLORS["light"]["bottom"]
        # CORREÇÃO: Para o card 'tomorrow' (amarelo), usar texto preto na parte inferior para legibilidade
        if status_key == "tomorrow":
            text_bottom_color = "#e49800"
        else:
            text_bottom_color = top_color
            
        return {
            "top": top_color, "bottom": bottom_color, 
            "text_top": text_top_color, "text_bottom": text_bottom_color
        }

    def criar_card(self):
        cores_gerais = CORES[self.config_manager.modo_aparencia]
        status_colors = self.obter_cores_status_evento()

        width, height = 220, 165
        top_part_height = height * 0.20
        border_width = 3

        prioridade_map = {"ALTA": "priority_high", "MÉDIA": "priority_medium"}
        prioridade_key = prioridade_map.get(self.evento.get("prioridade", ""), "priority_low")
        border_color = cores_gerais["status_success_bg"] if self.evento.get("concluido") else cores_gerais.get(prioridade_key, cores_gerais["priority_low"])

        self.delete("all")
        
        self.draw_rounded_rect(0, 0, width, height, radius=15, fill=border_color, outline="")
        self.draw_rounded_rect(border_width, border_width, width - border_width, height - border_width, radius=12, fill=status_colors["bottom"], outline="")
        self.create_rectangle(border_width, border_width, width - border_width, top_part_height, fill=status_colors["top"], outline="")

        top_frame = tk.Frame(self, bg=status_colors["top"], width=int(width - (2 * border_width)), height=int(top_part_height - border_width))
        top_frame.pack_propagate(False)
        self.create_window(border_width, border_width, window=top_frame, anchor="nw")
        top_frame.columnconfigure(0, weight=1)

        bottom_frame = tk.Frame(self, bg=status_colors["bottom"], width=int(width - (2 * border_width)), height=int(height - top_part_height - border_width))
        bottom_frame.pack_propagate(False)
        self.create_window(border_width, top_part_height, window=bottom_frame, anchor="nw")

        self.criar_conteudo(top_frame, bottom_frame, status_colors, cores_gerais)

        def bind_recursive(widget):
            if widget != self.menu_btn:
                widget.bind("<Button-1>", lambda event: self.on_click_callback(self.evento["id"]))
                for child in widget.winfo_children():
                    bind_recursive(child)
        
        bind_recursive(top_frame)
        bind_recursive(bottom_frame)
        self.bind("<Button-1>", lambda event: self.on_click_callback(self.evento["id"]))

    def criar_conteudo(self, top_frame, bottom_frame, status_colors, cores_gerais):
        text_top_color = status_colors["text_top"]
        info_text_color = status_colors["text_bottom"]
        bottom_color = status_colors["bottom"]

        nome_completo = self.evento['nome']
        nome_display = nome_completo[:18] + "..." if len(nome_completo) > 18 else nome_completo
        lbl_nome = tk.Label(top_frame, text=nome_display, font=("Arial", 10, "bold"), bg=status_colors["top"], fg=text_top_color, anchor="w")
        lbl_nome.grid(row=0, column=0, sticky="w", padx=6, pady=(4, 2))
        
        self.menu_btn = tk.Label(top_frame, text="⋯", font=("Arial", 12), bg=status_colors["top"], fg=text_top_color, cursor="hand2")
        self.menu_btn.grid(row=0, column=1, sticky="e", padx=(0, 10))
        self.menu_btn.bind("<Button-1>", lambda event: (self.on_menu_callback(self.evento["id"]), "break")[1])

        padding_x = 6
        def create_info_label(parent, text, font_size=7, bold=True):
            if not text or not str(text).strip(): return None
            font_style = ("Arial", font_size, "bold" if bold else "")
            label = tk.Label(parent, text=text, font=font_style, bg=bottom_color, fg=info_text_color, anchor="w")
            return label

        # Layout da primeira linha usando 'grid' para evitar que a data seja cortada
        row1 = tk.Frame(bottom_frame, bg=bottom_color)
        row1.pack(fill="x", padx=padding_x, pady=(5, 1))

        # Configura as colunas para que a do meio (prorrogação) se expanda
        row1.columnconfigure(0, weight=0)
        row1.columnconfigure(1, weight=1)
        row1.columnconfigure(2, weight=0)
        
        date_str = f"🗓 {self.formata_data_br(self.evento['data'])}"
        # CORREÇÃO: Removida a seta "➡️" para dar mais espaço à data
        pror_str = f"{self.formata_data_br(self.evento.get('data_prorrogacao'))}" if self.evento.get("data_prorrogacao") else ""
        time_str = f"🕔 {self.evento['hora']:02d}:{self.evento['minuto']:02d}"

        lbl_date = create_info_label(row1, date_str, 8)
        lbl_time = create_info_label(row1, time_str, 8)
        lbl_pror = create_info_label(row1, pror_str, 8, True)
        
        if lbl_date: lbl_date.grid(row=0, column=0, sticky="w")
        if lbl_pror:
            lbl_pror.config(anchor="center")
            lbl_pror.grid(row=0, column=1, sticky="ew")
        if lbl_time: lbl_time.grid(row=0, column=2, sticky="e")
        
        # Row 2: Categoria e Prioridade
        row2 = tk.Frame(bottom_frame, bg=bottom_color)
        row2.pack(fill="x", padx=padding_x)
        row2.columnconfigure(0, weight=0) # Coluna da esquerda não expande
        row2.columnconfigure(1, weight=1) # Coluna do meio expande
        row2.columnconfigure(2, weight=0) # Coluna da direita não expande
        cat_text = f"🏷 {self.evento.get('categoria', '')[:12]}"
        prio_text = f"⚠️ {self.evento.get('prioridade', '')}"
        lbl_cat = create_info_label(row2, cat_text)
        if lbl_cat:
            lbl_cat.grid(row=0, column=0, sticky="w") # Posição 0
        lbl_prio = create_info_label(row2, prio_text)
        if lbl_prio:
            lbl_prio.grid(row=0, column=2, sticky="e") # Posição 2

        # Row 3: Local e Grupo
        row3 = tk.Frame(bottom_frame, bg=bottom_color)
        row3.pack(fill="x", padx=padding_x)
        row3.columnconfigure(0, weight=0)
        row3.columnconfigure(1, weight=1)
        row3.columnconfigure(2, weight=0)
        local_text = f"📍 {self.evento.get('local', '')[:12]}"
        grupo_text = f"📂 {self.evento.get('grupo', '')}"
        lbl_local = create_info_label(row3, local_text)
        if lbl_local:
            lbl_local.grid(row=0, column=0, sticky="w") # Posição 0
        lbl_grupo = create_info_label(row3, grupo_text)
        if lbl_grupo:
            lbl_grupo.grid(row=0, column=2, sticky="e") # Posição 2

        # Row 4: Etiquetas e Anexo
        row4 = tk.Frame(bottom_frame, bg=bottom_color)
        row4.pack(fill="x", padx=padding_x)
        row4.columnconfigure(0, weight=0)
        row4.columnconfigure(1, weight=1)
        row4.columnconfigure(2, weight=0)
        etiquetas_text = f"📖 {self.evento.get('etiquetas', '')[:12]}"
        anexo_icon = "📎 Anexo" if self.evento.get("anexos") else ""
        lbl_etiq = create_info_label(row4, etiquetas_text)
        if lbl_etiq:
            lbl_etiq.grid(row=0, column=0, sticky="w") # Posição 0
        lbl_anexo = create_info_label(row4, anexo_icon)
        if lbl_anexo:
            lbl_anexo.grid(row=0, column=2, sticky="e") # Posição 2

        # Row 5: Data de Criação e ID
        row5 = tk.Frame(bottom_frame, bg=bottom_color)
        row5.pack(fill="x", padx=padding_x, pady=(1, 0))
        row5.columnconfigure(0, weight=0)
        row5.columnconfigure(1, weight=1)
        row5.columnconfigure(2, weight=0)
        data_criacao_str = self.formata_data_br(self.evento.get("data_criacao", datetime.date.today()))
        criado_text = f"Criado: {data_criacao_str}"
        card_id = f"ID: {self.gerar_card_id()}"
        lbl_criado = create_info_label(row5, criado_text)
        if lbl_criado:
            lbl_criado.grid(row=0, column=0, sticky="w") # Posição 0
        lbl_id = create_info_label(row5, card_id)
        if lbl_id:
            lbl_id.grid(row=0, column=2, sticky="e") # Posição 2
        
        progress_canvas = tk.Canvas(bottom_frame, height=4, bg=bottom_color, highlightthickness=0)
        progress_canvas.pack(side="bottom", fill="x", padx=padding_x, pady=(5, 3))
        
        progress_percent = self.calcular_progresso_barra()
        progress_color = self.obter_cor_barra_progresso()
        progress_bar_bg = cores_gerais["progress_bar_bg"]
        
        bar_width = 220 - (2 * 3) - (2 * padding_x) # Largura calculada e fixa
        
        progress_canvas.create_rectangle(0, 0, bar_width, 4, fill=progress_bar_bg, outline="")
        fill_width = (bar_width * progress_percent) / 100
        if fill_width > 0:
            progress_canvas.create_rectangle(0, 0, fill_width, 4, fill=progress_color, outline="")

    def obter_cor_barra_progresso(self):
        hoje = datetime.date.today()
        data_efetiva = self.evento_manager.get_effective_date(self.evento)
        dias_faltando = (data_efetiva - hoje).days
        cores = CORES[self.config_manager.modo_aparencia]
        if self.evento["concluido"]: return cores["progress_completed"]
        elif dias_faltando < 0: return cores["progress_past"]
        elif dias_faltando <= 3: return cores["progress_0_3"]
        elif dias_faltando <= 7: return cores["progress_3_7"]
        elif dias_faltando <= 30: return cores["progress_7_30"]
        else: return cores["progress_30_plus"]

    def calcular_progresso_barra(self):
        hoje = datetime.date.today()
        data_efetiva = self.evento_manager.get_effective_date(self.evento)
        dias_faltando = (data_efetiva - hoje).days
        if self.evento["concluido"]: return 100
        if dias_faltando < 0: return 15
        elif dias_faltando == 0: return 100
        elif dias_faltando == 1: return 85
        elif dias_faltando <= 3: return 85 - ((dias_faltando - 1) * 7)
        elif dias_faltando <= 7: return 70 - ((dias_faltando - 3) * 5)
        elif dias_faltando <= 14: return 50 - ((dias_faltando - 7) * 2)
        elif dias_faltando <= 30: return 35 - ((dias_faltando - 14) * 1)
        else: return 15

    def gerar_card_id(self):
        nome_clean = ''.join(c for c in self.evento['nome'] if c.isalpha())[:3].upper()
        dia = self.evento_manager.get_effective_date(self.evento).day
        return f"{nome_clean}{dia:02d}"

    def formata_data_br(self, data_date):
        if not data_date: return ""
        return data_date.strftime("%d/%m/%Y")

    def draw_rounded_rect(self, x1, y1, x2, y2, radius, **kwargs):
        points = [x1 + radius, y1, x1 + radius, y1, x2 - radius, y1, x2 - radius, y1,
                  x2, y1, x2, y1 + radius, x2, y1 + radius, x2, y2 - radius,
                  x2, y2 - radius, x2, y2, x2 - radius, y2, x2 - radius, y2,
                  x1 + radius, y2, x1 + radius, y2, x1, y2, x1, y2 - radius,
                  x1, y2 - radius, x1, y1 + radius, x1, y1 + radius, x1, y1]
        return self.create_polygon(points, **kwargs, smooth=True)

class ExportManager:
    """Gerenciador de exportações"""

    def __init__(self, evento_manager):
        self.evento_manager = evento_manager
        logger.info("ExportManager inicializado.")

    def exportar_para_ics(self, evento_id=None):
        """Exporta evento(s) para formato .ics"""
        try:
            if evento_id:
                eventos_exportar = [e for e in self.evento_manager.eventos if e["id"] == evento_id]
                if not eventos_exportar:
                    messagebox.showerror("Erro", "Evento não encontrado!")
                    return
                nome_limpo = ''.join(c for c in eventos_exportar[0]['nome'] if c.isalnum() or c in (' ', '-', '_'))[:20]
                nome_arquivo = f"evento_{nome_limpo}.ics"
            else:
                eventos_exportar = self.evento_manager.eventos
                nome_arquivo = "todos_eventos.ics"

            arquivo = filedialog.asksaveasfilename(
                title="Salvar arquivo ICS",
                filetypes=[("Arquivos ICS", "*.ics"), ("Todos os arquivos", "*.*")]
            )

            if not arquivo:
                logger.info("Exportação para ICS cancelada pelo usuário.")
                return

            if not arquivo.lower().endswith('.ics'):
                arquivo += '.ics'

            # Criar conteúdo ICS
            ics_content = self._gerar_conteudo_ics(eventos_exportar)

            try:
                with open(arquivo, 'w', encoding='utf-8') as f:
                    f.write('\r\n'.join(ics_content))

                logger.info(f"Eventos exportados para ICS com sucesso: {arquivo}")
                messagebox.showinfo("Sucesso", f"Arquivo ICS exportado com sucesso!\n\nLocal: {arquivo}\n\nVocê pode importar este arquivo no Google Calendar, Outlook ou outros aplicativos de calendário.")

            except PermissionError:
                messagebox.showerror("Erro", "Não foi possível salvar o arquivo. Verifique as permissões do diretório.")
            except Exception as save_error:
                messagebox.showerror("Erro", f"Erro ao salvar arquivo: {str(save_error)}")

        except Exception as e:
            logger.error(f"Erro ao exportar para ICS: {e}", exc_info=True)
            messagebox.showerror("Erro", f"Erro ao exportar para ICS: {str(e)}")

    def _gerar_conteudo_ics(self, eventos):
        """Gera o conteúdo do arquivo ICS"""
        ics_content = []
        ics_content.append("BEGIN:VCALENDAR")
        ics_content.append("VERSION:2.0")
        ics_content.append("PRODID:-//Calendario Moderno//NONSGML v1.0//PT")
        ics_content.append("CALSCALE:GREGORIAN")
        ics_content.append("METHOD:PUBLISH")

        for evento in eventos:
            ics_content.extend(self._gerar_evento_ics(evento))

        ics_content.append("END:VCALENDAR")
        return ics_content

    def _gerar_evento_ics(self, evento):
        """Gera as linhas ICS para um evento específico"""
        linhas = []

        def escapar_texto(texto):
            if not texto:
                return ""
            texto_str = str(texto)
            texto_str = texto_str.replace("\\", "\\\\")
            texto_str = texto_str.replace(",", "\\,")
            texto_str = texto_str.replace(";", "\\;")
            texto_str = texto_str.replace("\n", "\\n")
            texto_str = texto_str.replace("\r", "\\r")
            return texto_str
        
        data_evento = self.evento_manager.get_effective_date(evento)
        hora_evento = datetime.time(evento["hora"], evento["minuto"])
        dt_inicio = datetime.datetime.combine(data_evento, hora_evento)

        dtstart = dt_inicio.strftime("%Y%m%dT%H%M%S")
        dtend = (dt_inicio + datetime.timedelta(hours=1)).strftime("%Y%m%dT%H%M%S")

        linhas.append("BEGIN:VEVENT")
        linhas.append(f"UID:{evento['id']}@calendario-moderno")
        linhas.append(f"DTSTART:{dtstart}")
        linhas.append(f"DTEND:{dtend}")
        linhas.append(f"DTSTAMP:{datetime.datetime.now().strftime('%Y%m%dT%H%M%SZ')}")
        linhas.append(f"SUMMARY:{escapar_texto(evento['nome'])}")

        descricao = evento.get('corpo', '')
        descricao += f"\\n\\nData Original: {evento['data'].strftime('%d/%m/%Y')}"
        if evento.get('data_prorrogacao'):
            descricao += f"\\nData de Prorrogação: {evento['data_prorrogacao'].strftime('%d/%m/%Y')}"
        
        linhas.append(f"DESCRIPTION:{escapar_texto(descricao)}")

        if evento.get('local'):
            linhas.append(f"LOCATION:{escapar_texto(evento['local'])}")

        if evento.get('categoria'):
            linhas.append(f"CATEGORIES:{escapar_texto(evento['categoria'])}")

        if evento.get('prioridade'):
            prioridade_map = {"BAIXA": "9", "MÉDIA": "5", "ALTA": "1"}
            prioridade_ics = prioridade_map.get(evento['prioridade'], "0")
            linhas.append(f"PRIORITY:{prioridade_ics}")

        status = "COMPLETED" if evento.get('concluido') else "CONFIRMED"
        linhas.append(f"STATUS:{status}")
        linhas.append("END:VEVENT")

        return linhas

    def exportar_para_pdf(self, eventos, titulo="Relatório de Eventos"):
        """Exporta eventos para formato PDF"""
        try:
            arquivo = filedialog.asksaveasfilename(
                title="Salvar arquivo PDF",
                filetypes=[("Arquivos PDF", "*.pdf"), ("Todos os arquivos", "*.*")]
            )

            if not arquivo:
                logger.info("Exportação para PDF cancelada pelo usuário.")
                return

            if not arquivo.lower().endswith('.pdf'):
                arquivo += '.pdf'

            # Criar documento PDF
            doc = SimpleDocTemplate(arquivo, pagesize=A4)
            elements = []

            # Estilos
            styles = getSampleStyleSheet()
            title_style = styles['Title']
            heading_style = styles['Heading2']
            normal_style = styles['Normal']

            # Título
            elements.append(Paragraph(titulo, title_style))
            elements.append(Spacer(1, 12))

            # Data de geração
            data_geracao = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
            elements.append(Paragraph(f"Gerado em: {data_geracao}", normal_style))
            elements.append(Spacer(1, 24))

            # Tabela de eventos
            if eventos:
                # Cabeçalho da tabela
                data = [['Nome', 'Data Original', 'Prorrogação', 'Hora', 'Prioridade', 'Status']]

                for evento in eventos:
                    status = "Concluído" if evento['concluido'] else "Pendente"
                    data_prorrogacao_str = evento['data_prorrogacao'].strftime("%d/%m/%Y") if evento.get('data_prorrogacao') else ""
                    data.append([
                        evento['nome'],
                        evento['data'].strftime("%d/%m/%Y"),
                        data_prorrogacao_str,
                        f"{evento['hora']:02d}:{evento['minuto']:02d}",
                        evento.get('prioridade', ''),
                        status
                    ])

                # Criar tabela
                table = Table(data)
                table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, 0), 10),
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                    ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                    ('FONTSIZE', (0, 1), (-1, -1), 8),
                    ('GRID', (0, 0), (-1, -1), 1, colors.black)
                ]))

                elements.append(table)
            else:
                elements.append(Paragraph("Nenhum evento para exibir.", normal_style))

            # Gerar PDF
            doc.build(elements)
            logger.info(f"Relatório exportado para PDF com sucesso: {arquivo}")
            messagebox.showinfo("Sucesso", f"Arquivo PDF exportado com sucesso!\n\nLocal: {arquivo}")

        except Exception as e:
            logger.error(f"Erro ao exportar para PDF: {e}", exc_info=True)
            messagebox.showerror("Erro", f"Erro ao exportar para PDF: {str(e)}")

    def exportar_para_csv(self, eventos, arquivo=None):
        """Exporta eventos para formato CSV"""
        try:
            if not arquivo:
                arquivo = filedialog.asksaveasfilename(
                    title="Salvar arquivo CSV",
                    filetypes=[("Arquivos CSV", "*.csv"), ("Todos os arquivos", "*.*")]
                )

                if not arquivo:
                    logger.info("Exportação para CSV cancelada pelo usuário.")
                    return

                if not arquivo.lower().endswith('.csv'):
                    arquivo += '.csv'

            with open(arquivo, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f, delimiter=';')

                # Escrever cabeçalho
                writer.writerow(['Nome', 'Data Original', 'Data Prorrogação', 'Hora', 'Local', 'Categoria', 'Prioridade', 'Status', 'Data Criação'])

                # Escrever dados
                for evento in eventos:
                    status = "Concluído" if evento['concluido'] else "Pendente"
                    data_prorrogacao_str = evento['data_prorrogacao'].strftime("%d/%m/%Y") if evento.get('data_prorrogacao') else ""
                    writer.writerow([
                        evento['nome'],
                        evento['data'].strftime("%d/%m/%Y"),
                        data_prorrogacao_str,
                        f"{evento['hora']:02d}:{evento['minuto']:02d}",
                        evento.get('local', ''),
                        evento.get('categoria', ''),
                        evento.get('prioridade', ''),
                        status,
                        evento['data_criacao'].strftime("%d/%m/%Y")
                    ])
            logger.info(f"Relatório exportado para CSV com sucesso: {arquivo}")
            return arquivo

        except Exception as e:
            logger.error(f"Erro ao exportar para CSV: {e}", exc_info=True)
            messagebox.showerror("Erro", f"Erro ao exportar para CSV: {str(e)}")
            return None

class CalendarioApp:
    """Aplicação principal do calendário"""

    def __init__(self):
        logger.info("Inicializando CalendarioApp...")
        self.config_manager = ConfigManager()
        self.crypt_manager = CryptManager()
        self.evento_manager = EventoManager(self.crypt_manager)
        self.alerta_manager = AlertaManager(self.evento_manager, self.config_manager)
        self.export_manager = ExportManager(self.evento_manager)
        self.template_manager = TemplateManager()
        self.meta_manager = MetaManager()
        logger.info("Todos os gerenciadores foram inicializados.")

        # Estado da aplicação
        self.editando_id = None
        self.temp_anexos = []
        self.PAGINA_ATUAL = 1
        self.cards_por_id = {}
        self.filtro_atual = "Todos"
        self.filtro_busca = ""
        self.ordem_atual = "Data"
        self.modo_responsivo = False

        # Widgets (serão inicializados na criação da interface)
        self.root = None
        self.widgets = {}

    def inicializar(self):
        """Inicialização da aplicação"""
        logger.info("Iniciando processo de inicialização da aplicação.")
        if not self.crypt_manager.carregar_ou_criar_chave():
            logger.error("Falha ao carregar ou criar a chave de criptografia.")
            return False
        if self.crypt_manager.criptografia_habilitada and not self.crypt_manager.verificar_senha():
            logger.warning("Verificação de senha falhou ou foi cancelada.")
            return False

        self.config_manager.carregar()
        self.evento_manager.carregar_eventos()

        self.root = tk.Tk()
        self.root.title("📅 Calendário DocQuantum - Versão SmarthCards")
        self.root.geometry("1200x700")
        self.root.minsize(1000, 600)
        self.root.bind('<F11>', self.alternar_tela_cheia)
        self.root.bind('<Escape>', self.sair_tela_cheia)
        self.root.bind('<Control-z>', lambda e: self.undo())
        self.root.bind('<Control-y>', lambda e: self.redo())
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        # Verificar modo responsivo
        self.verificar_modo_responsivo()
        logger.info("Aplicação inicializada com sucesso.")
        return True

    def verificar_modo_responsivo(self):
        """Verifica se a aplicação deve entrar em modo responsivo"""
        largura = self.root.winfo_screenwidth()
        self.modo_responsivo = largura < 1280
        logger.info(f"Modo responsivo: {self.modo_responsivo}")
        return self.modo_responsivo

    def undo(self):
        """Desfaz a última operação"""
        if self.evento_manager.undo():
            self.atualizar_lista()
            self.evento_manager.salvar_eventos()

    def redo(self):
        """Refaz a operação desfeita"""
        if self.evento_manager.redo():
            self.atualizar_lista()
            self.evento_manager.salvar_eventos()

    def criar_interface(self):
        """Cria toda a interface da aplicação"""
        logger.info("Criando a interface gráfica...")
        self._criar_variaveis()
        self._criar_layout_principal()
        self._configurar_eventos()
        self._criar_menu_principal()
        logger.info("Interface criada com sucesso.")

    def _criar_variaveis(self):
        """Cria as variáveis tkinter"""
        self.nome_var = tk.StringVar()
        self.local_var = tk.StringVar()
        self.categoria_var = tk.StringVar()
        self.grupo_var = tk.StringVar()
        self.etiquetas_var = tk.StringVar()
        self.hora_var = tk.StringVar(value="00")
        self.minuto_var = tk.StringVar(value="00")
        self.prioridade_var = tk.StringVar()
        self.filtro_var = tk.StringVar(value="Todos")
        self.contador_var = tk.StringVar(value="Eventos: 0/0")
        self.template_var = tk.StringVar(value="Selecionar Template")
        self.prorrogacao_var = tk.StringVar()

    def _criar_layout_principal(self):
        """Cria o layout principal da interface"""
        # Header
        header = tk.Frame(self.root, height=50)
        header.grid(row=0, column=0, columnspan=2, sticky="ew")
        header.grid_propagate(False)

        tk.Label(header, text="📅 Calendário Financeiro", font=("Arial", 18, "bold"), fg="white").pack(side="left", padx=15)
        tk.Button(header, text="🎨", command=self.alternar_tema, relief="flat", fg="white", cursor="hand2").pack(side="right", padx=15)
        tk.Button(header, text="⛶", command=self.alternar_tela_cheia, relief="flat", fg="white", cursor="hand2").pack(side="right", padx=5)

        self.widgets['header'] = header

        # Painel esquerdo
        left_panel = tk.Frame(self.root)
        left_panel.grid(row=1, column=0, sticky="ns", padx=(15, 5), pady=10)
        self.widgets['left_panel'] = left_panel

        self._criar_painel_filtros(left_panel)
        self._criar_painel_formulario(left_panel)

        # Painel direito
        right_panel = tk.Frame(self.root)
        right_panel.grid(row=1, column=1, sticky="nsew", padx=(5, 5), pady=10)
        self.widgets['right_panel'] = right_panel

        self._criar_painel_eventos(right_panel)

        # Configurar grid
        self.root.columnconfigure(1, weight=1)
        self.root.rowconfigure(1, weight=1)

        # Ajustar layout para modo responsivo
        if self.modo_responsivo:
            self._ajustar_layout_responsivo()

    def _ajustar_layout_responsivo(self):
        """Ajusta o layout para modo responsivo"""
        # Esconder painel esquerdo inicialmente
        self.widgets['left_panel'].grid_remove()

        # Adicionar botão para toggle do painel esquerdo
        btn_toggle = tk.Button(self.widgets['header'], text="☰", command=self.toggle_painel_esquerdo,
                              relief="flat", fg="white", cursor="hand2")
        btn_toggle.pack(side="left", padx=5)
        self.widgets['btn_toggle'] = btn_toggle

    def toggle_painel_esquerdo(self):
        """Alterna a visibilidade do painel esquerdo no modo responsivo"""
        if self.widgets['left_panel'].winfo_ismapped():
            self.widgets['left_panel'].grid_remove()
        else:
            self.widgets['left_panel'].grid()

    def _criar_painel_filtros(self, parent):
        """Cria o painel de filtros"""
        frame_filtros = tk.Frame(parent, relief="solid", bd=1)
        frame_filtros.pack(fill="x", pady=(0, 10))
        self.widgets['frame_filtros'] = frame_filtros

        tk.Label(frame_filtros, text="🔍").pack(side="left", padx=(12, 5))

        filtro_menu = ttk.Combobox(frame_filtros, textvariable=self.filtro_var,
                                  values=["Todos", "Somente hoje", "Vence em 1 dia", "Próximos 3 dias", "Próximos 7 dias",
                                         "Próximos 14 dias", "Próximos 21 dias", "Próximos 30 dias", "Passados", "Concluídos"],
                                  state="readonly", width=15, style="Themed.TCombobox")
        filtro_menu.pack(side="left", padx=0, pady=5)
        filtro_menu.bind("<<ComboboxSelected>>", self.alterar_filtro)
        self.widgets['filtro_menu'] = filtro_menu

        tk.Label(frame_filtros, text="Pesquisa:", font=("Arial", 10)).pack(side="left", padx=(10, 2), pady=5)

        entry_busca = ttk.Entry(frame_filtros, width=15, style="Themed.TEntry")
        entry_busca.pack(side="left", fill="x", expand=True, padx=(0, 12), pady=5)
        entry_busca.bind("<KeyRelease>", self.filtrar_por_busca)
        entry_busca.bind("<FocusOut>", self.filtrar_por_busca)
        self.widgets['entry_busca'] = entry_busca

        # Botão para pesquisa avançada
        btn_avancado = ttk.Button(frame_filtros, text="Avançado", command=self.mostrar_pesquisa_avancada,
                                 style="Themed.TButton", width=10)
        btn_avancado.pack(side="right", padx=(0, 5), pady=5)
        self.widgets['btn_avancado'] = btn_avancado

    def _criar_painel_formulario(self, parent):
        """Cria o painel do formulário"""
        frame_form = tk.Frame(parent, relief="solid", bd=1)
        frame_form.pack(fill="both", expand=True)
        self.widgets['frame_form'] = frame_form

        # Título
        titulo_frame = tk.Frame(frame_form)
        titulo_frame.grid(row=0, column=0, columnspan=2, pady=(5, 10), sticky="ew")
        titulo_frame.columnconfigure(0, weight=1)
        tk.Label(titulo_frame, text="📝 Novo Evento", font=("Arial", 14, "bold")).grid(row=0, column=0)

        row = 1

        # Template Label
        tk.Label(frame_form, text="Template:", font=("Arial", 10), width=8, anchor="w").grid(row=row, column=0, padx=(12, 5), pady=2, sticky="w")

        # Frame for the template widgets (Combobox + Button)
        template_widget_frame = tk.Frame(frame_form)
        template_widget_frame.grid(row=row, column=1, padx=(0, 12), pady=2, sticky="ew")

        template_menu = ttk.Combobox(template_widget_frame, textvariable=self.template_var, state="readonly", style="Themed.TCombobox")
        template_menu.pack(side="left", fill="x", expand=True, padx=(0, 5))
        template_menu.bind("<<ComboboxSelected>>", self.aplicar_template)
        self.widgets['template_menu'] = template_menu

        btn_gerenciar_templates = ttk.Button(template_widget_frame, text="Gerir", command=self.gerenciar_templates,
                                            style="Themed.TButton", width=5)
        btn_gerenciar_templates.pack(side="left")
        self.widgets['btn_gerenciar_templates'] = btn_gerenciar_templates
        row += 1

        # Nome
        entry_nome = ttk.Entry(frame_form, textvariable=self.nome_var, width=15, style="Themed.TEntry")
        self._criar_linha_formulario("Nome:", entry_nome, row, frame_form)
        self.widgets['entry_nome'] = entry_nome
        row += 1

        # Campo de data original e prorrogação
        data_frame = tk.Frame(frame_form)
        entry_data = ttk.Entry(data_frame, width=10, style="Themed.TEntry")
        entry_data.pack(side="left")
        btn_cal_orig = ttk.Button(data_frame, text="📅", width=3, command=lambda: self.selecionar_data(entry_data), style="Themed.TButton")
        btn_cal_orig.pack(side="left", padx=(2, 5))

        tk.Label(data_frame, text="PR:").pack(side="left")
        
        entry_prorrogacao = ttk.Entry(data_frame, textvariable=self.prorrogacao_var, width=10, style="Themed.TEntry")
        entry_prorrogacao.pack(side="left", padx=(5,0))
        btn_cal_pror = ttk.Button(data_frame, text="📅", width=3, command=lambda: self.selecionar_data(entry_prorrogacao), style="Themed.TButton")
        btn_cal_pror.pack(side="left", padx=(2,0))

        self._criar_linha_formulario("Datas:", data_frame, row, frame_form)
        self.widgets['entry_data'] = entry_data
        self.widgets['entry_prorrogacao'] = entry_prorrogacao
        row += 1

        # Outros campos
        campos = [
            ('Categoria:', 'categoria_var', 'entry_categoria'),
            ('Grupo:', 'grupo_var', 'entry_grupo'),
            ('Etiquetas:', 'etiquetas_var', 'entry_etiquetas'),
            ('Local:', 'local_var', 'entry_local')
        ]

        for label, var_name, widget_name in campos:
            var = getattr(self, var_name)
            entry = ttk.Entry(frame_form, textvariable=var, width=15, style="Themed.TEntry")
            self._criar_linha_formulario(label, entry, row, frame_form)
            self.widgets[widget_name] = entry
            if self.config_manager.uppercase_enabled:
                self._setup_uppercase_entry(entry)
            row += 1

        # Prioridade
        prioridade_menu = ttk.Combobox(frame_form, textvariable=self.prioridade_var,
                                      values=["", "BAIXA", "MÉDIA", "ALTA"],
                                      state="readonly", width=15, style="Themed.TCombobox")
        self._criar_linha_formulario("Prioridade:", prioridade_menu, row, frame_form)
        self.widgets['prioridade_menu'] = prioridade_menu
        row += 1

        # Hora
        hora_frame = tk.Frame(frame_form)
        hora_combo = ttk.Combobox(hora_frame, textvariable=self.hora_var,
                                 values=[f"{h:02d}" for h in range(24)],
                                 width=5, style="Themed.TCombobox", state="readonly")
        hora_combo.pack(side="left")
        tk.Label(hora_frame, text=":").pack(side="left", padx=5)
        minuto_combo = ttk.Combobox(hora_frame, textvariable=self.minuto_var,
                                   values=[f"{m:02d}" for m in range(60)],
                                   width=5, style="Themed.TCombobox", state="readonly")
        minuto_combo.pack(side="left")
        self._criar_linha_formulario("Hora:", hora_frame, row, frame_form)
        self.widgets['hora_combo'] = hora_combo
        self.widgets['minuto_combo'] = minuto_combo
        row += 1

        # Detalhes
        tk.Label(frame_form, text="Detalhes:", font=("Arial", 10), anchor="nw").grid(row=row, column=0, padx=(12, 5), pady=2, sticky="nw")
        text_corpo = tk.Text(frame_form, height=5, width=15, wrap="word")
        text_corpo.grid(row=row, column=1, padx=(0, 12), pady=2, sticky="ew")
        self.widgets['text_corpo'] = text_corpo
        if self.config_manager.uppercase_enabled:
            self._uppercase_text_handler(text_corpo)
        row += 1

        # Anexar arquivo
        ttk.Button(frame_form, text="📎 Anexar Arquivo", command=self.anexar_arquivo,
                  style="Themed.TButton").grid(row=row, column=0, columnspan=2, padx=12, pady=5)
        row += 1

        # Botões
        btn_frame = tk.Frame(frame_form)
        btn_frame.grid(row=row, column=0, columnspan=2, pady=10, sticky="ew")

        btn_adicionar = ttk.Button(btn_frame, text="Adicionar evento", command=self.adicionar_evento,
                                  style="Themed.TButton", width=15)
        btn_salvar = ttk.Button(btn_frame, text="Salvar Alterações", command=self.salvar_alteracoes,
                               style="Themed.TButton", state='disabled', width=15)
        btn_cancelar = ttk.Button(btn_frame, text="Cancelar", command=self.cancelar_edicao,
                                 style="Themed.TButton", width=15)

        tk.Label(btn_frame, text="").pack(side="left", expand=True)
        btn_adicionar.pack(side="left", padx=5)
        btn_salvar.pack(side="left", padx=5)
        btn_cancelar.pack(side="left", padx=5)
        tk.Label(btn_frame, text="").pack(side="right", expand=True)

        self.widgets['btn_adicionar'] = btn_adicionar
        self.widgets['btn_salvar'] = btn_salvar
        self.widgets['btn_cancelar'] = btn_cancelar

        frame_form.columnconfigure(1, weight=1)
        frame_form.rowconfigure(row, weight=1)

        if self.config_manager.uppercase_enabled:
            self._setup_uppercase_entry(entry_nome)
        
        self.widgets['template_widget_frame'] = template_widget_frame

    def _criar_linha_formulario(self, label, widget, row, parent):
        """Helper para criar linha do formulário"""
        tk.Label(parent, text=label, font=("Arial", 10), width=8, anchor="w").grid(row=row, column=0, padx=(12, 5), pady=2, sticky="w")
        widget.grid(row=row, column=1, padx=(0, 12), pady=2, sticky="ew")

    def _criar_painel_eventos(self, parent):
        """Cria o painel de eventos"""
        label_titulo_eventos = tk.Label(parent, text="Eventos Financeiros Agendados", font=("Arial", 14, "bold"))
        label_titulo_eventos.pack(pady=(0, 8))
        self.widgets['label_titulo_eventos'] = label_titulo_eventos

        scroll_frame = tk.Frame(parent)
        scroll_frame.pack(fill="both", expand=True)
        self.widgets['scroll_frame'] = scroll_frame

        canvas = tk.Canvas(scroll_frame, highlightthickness=0)
        scrollbar = ttk.Scrollbar(scroll_frame, orient="vertical", command=canvas.yview)
        frame_lista = tk.Frame(canvas)

        frame_lista.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=frame_lista, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.widgets['canvas'] = canvas
        self.widgets['frame_lista'] = frame_lista

        # Configurar scroll do mouse
        self._configurar_scroll_mouse(canvas)

        # Contador de eventos
        tk.Label(parent, textvariable=self.contador_var).pack(side="bottom", pady=5)

        # Configurar grid do frame_lista
        for i in range(4):
            frame_lista.columnconfigure(i, weight=1)

    def _configurar_scroll_mouse(self, canvas):
        """Configura o scroll do mouse"""
        def on_mousewheel(event):
            if hasattr(event, 'delta') and event.delta:
                canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
            elif event.num == 4:
                canvas.yview_scroll(-1, "units")
            elif event.num == 5:
                canvas.yview_scroll(1, "units")

        def _bind_mousewheel(event):
            canvas.bind_all("<MouseWheel>", on_mousewheel)
            canvas.bind_all("<Button-4>", on_mousewheel)
            canvas.bind_all("<Button-5>", on_mousewheel)

        def _unbind_mousewheel(event):
            canvas.unbind_all("<MouseWheel>")
            canvas.unbind_all("<Button-4>")
            canvas.unbind_all("<Button-5>")

        canvas.bind("<Enter>", _bind_mousewheel)
        canvas.bind("<Leave>", _unbind_mousewheel)

    def _configurar_eventos(self):
        """Configura eventos da aplicação"""
        if self.config_manager.uppercase_enabled:
            self._setup_search_field()

        # Atualizar lista de templates
        self._atualizar_templates()

    def _setup_uppercase_entry(self, entry_widget):
        """Configura entrada para maiúsculas automáticas"""
        def to_uppercase(event=None):
            cursor_pos = entry_widget.index(tk.INSERT)
            content = entry_widget.get()
            if content != content.upper():
                entry_widget.delete(0, tk.END)
                entry_widget.insert(0, content.upper())
                entry_widget.icursor(cursor_pos)

        entry_widget.bind("<KeyRelease>", to_uppercase)
        entry_widget.bind("<FocusOut>", to_uppercase)

    def _uppercase_text_handler(self, text_widget):
        """Configura text widget para maiúsculas automáticas"""
        def to_uppercase(event=None):
            cursor_pos = text_widget.index(tk.INSERT)
            content = text_widget.get("1.0", "end-1c")
            if content != content.upper():
                text_widget.delete("1.0", "end")
                text_widget.insert("1.0", content.upper())
                text_widget.mark_set(tk.INSERT, cursor_pos)

        text_widget.bind("<KeyRelease>", to_uppercase)
        text_widget.bind("<FocusOut>", to_uppercase)

    def _setup_search_field(self):
        """Configura campo de busca"""
        entry_busca = self.widgets['entry_busca']

        def handle_search_input(event=None):
            if self.config_manager.uppercase_enabled:
                cursor_pos = entry_busca.index(tk.INSERT)
                content = entry_busca.get()
                if content != content.upper():
                    entry_busca.delete(0, tk.END)
                    entry_busca.insert(0, content.upper())
                    entry_busca.icursor(cursor_pos)
            self.filtrar_por_busca()

        entry_busca.bind("<KeyRelease>", handle_search_input)
        entry_busca.bind("<FocusOut>", handle_search_input)

    def _atualizar_templates(self):
        """Atualiza a lista de templates no menu"""
        templates = self.template_manager.templates
        nomes_templates = [f"Template {i+1}: {t['nome'][:20]}" for i, t in enumerate(templates)]

        self.widgets['template_menu']['values'] = ["Selecionar Template"] + nomes_templates
        self.template_var.set("Selecionar Template")

    def _criar_menu_principal(self):
        """Cria o menu principal da aplicação"""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)

        # Menu Arquivo
        arquivo_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Arquivo", menu=arquivo_menu)
        arquivo_menu.add_command(label="Importar de ICS", command=self.importar_de_ics)
        arquivo_menu.add_separator()
        arquivo_menu.add_command(label="Exportar Evento para ICS",
                                command=lambda: messagebox.showinfo("Info", "Selecione um evento e use o menu do card"))
        arquivo_menu.add_command(label="Exportar Todos para ICS",
                                command=lambda: self.export_manager.exportar_para_ics())
        arquivo_menu.add_separator()

        # Submenu Backup
        backup_menu = tk.Menu(arquivo_menu, tearoff=0)
        arquivo_menu.add_cascade(label="Backup", menu=backup_menu)
        backup_menu.add_command(label="Criar Backup", command=self.criar_backup)
        backup_menu.add_command(label="Restaurar Backup", command=self.restaurar_backup)

        arquivo_menu.add_separator()
        arquivo_menu.add_command(label="Exportar para PDF", command=self.exportar_para_pdf)
        arquivo_menu.add_command(label="Exportar para CSV", command=self.exportar_para_csv)
        arquivo_menu.add_separator()
        #arquivo_menu.add_command(label="Configurar Criptografia", command=self.configurar_criptografia)
        arquivo_menu.add_separator()
        arquivo_menu.add_command(label="Sair", command=self.on_closing)

        # Menu Editar
        editar_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Editar", menu=editar_menu)
        editar_menu.add_command(label="Desfazer (Ctrl+Z)", command=self.undo)
        editar_menu.add_command(label="Refazer (Ctrl+Y)", command=self.redo)
        editar_menu.add_separator()
        editar_menu.add_command(label="Excluir Todos os Eventos", command=self.excluir_todos_eventos_confirmacao)
        editar_menu.add_separator()
        editar_menu.add_command(label="Configurações", command=self.mostrar_configuracoes)

        # Menu Visualizar
        view_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Visualizar", menu=view_menu)
        view_menu.add_command(label="Alternar Tema", command=self.alternar_tema)
        view_menu.add_command(label="Tela Cheia (F11)", command=self.alternar_tela_cheia)

        ordem_menu = tk.Menu(view_menu, tearoff=0)
        view_menu.add_cascade(label="Ordenar por", menu=ordem_menu)

        def set_ordem(ordem):
            self.ordem_atual = ordem
            self.atualizar_lista()

        ordem_menu.add_command(label="Data do Evento", command=lambda: set_ordem("Data"))
        ordem_menu.add_command(label="Prioridade", command=lambda: set_ordem("Prioridade"))
        ordem_menu.add_command(label="Nome (A-Z)", command=lambda: set_ordem("Nome (A-Z)"))
        ordem_menu.add_command(label="Data de Criação", command=lambda: set_ordem("Data de Criação"))

        # Menu Relatórios
        relatorios_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Relatórios", menu=relatorios_menu)
        relatorios_menu.add_command(label="Dashboard", command=self.mostrar_dashboard)
        relatorios_menu.add_command(label="Relatório Mensal", command=lambda: self.gerar_relatorio("mensal"))
        relatorios_menu.add_command(label="Relatório Semanal", command=lambda: self.gerar_relatorio("semanal"))
        relatorios_menu.add_command(label="Análise de Produtividade", command=self.mostrar_analise_produtividade)

        # Menu Metas
        metas_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Metas", menu=metas_menu)
        metas_menu.add_command(label="Gerenciar Metas", command=self.gerenciar_metas)
        metas_menu.add_command(label="Ver Progresso", command=self.verificar_progresso_metas)

        # Menu Ajuda
        ajuda_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Ajuda", menu=ajuda_menu)
        ajuda_menu.add_command(label="Sobre", command=self.mostrar_sobre)

    def mostrar_configuracoes(self):
        """Mostra janela de configurações"""
        cores = CORES[self.config_manager.modo_aparencia]
        config_window = tk.Toplevel(self.root)
        config_window.title("Configurações")
        config_window.geometry("400x400")
        config_window.configure(bg=cores["primary"])
        config_window.transient(self.root)
        config_window.grab_set()

        config_window.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() - config_window.winfo_width()) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - config_window.winfo_height()) // 2
        config_window.geometry(f"+{x}+{y}")

        tk.Label(config_window, text="Configurações", font=("Arial", 16, "bold"),
                bg=cores["primary"], fg=cores["text"]).pack(pady=10)

        # Configuração de maiúsculas
        uppercase_var = tk.BooleanVar(value=self.config_manager.uppercase_enabled)
        uppercase_frame = tk.Frame(config_window, bg=cores["primary"])
        uppercase_frame.pack(fill="x", padx=20, pady=10)

        tk.Checkbutton(uppercase_frame, text="Converter texto para maiúsculas automaticamente",
                      variable=uppercase_var, bg=cores["primary"], fg=cores["text"],
                      selectcolor=cores["secondary"]).pack(side="left")

        # Configuração de alertas personalizados
        alertas_frame = tk.Frame(config_window, bg=cores["primary"])
        alertas_frame.pack(fill="x", padx=20, pady=10)

        tk.Label(alertas_frame, text="Alertas (dias antes):", bg=cores["primary"], fg=cores["text"]).pack(anchor="w")

        alertas_entry = ttk.Entry(alertas_frame, width=30)
        alertas_entry.insert(0, ",".join(map(str, self.config_manager.alertas_personalizados)))
        alertas_entry.pack(fill="x", pady=5)

        def salvar_config():
            self.config_manager.uppercase_enabled = uppercase_var.get()

            # Processar alertas personalizados
            try:
                alertas_texto = alertas_entry.get().strip()
                if alertas_texto:
                    alertas = [int(a.strip()) for a in alertas_texto.split(",")]
                    self.config_manager.alertas_personalizados = sorted(alertas)
            except ValueError:
                messagebox.showerror("Erro", "Os alertas devem ser números separados por vírgula")
                return

            self.config_manager.salvar()
            messagebox.showinfo("Sucesso", "Configurações salvas!")
            config_window.destroy()

        def cancelar_config():
            config_window.destroy()

        btn_frame = tk.Frame(config_window, bg=cores["primary"])
        btn_frame.pack(side="bottom", pady=20)

        tk.Button(btn_frame, text="Salvar", command=salvar_config,
                 bg=cores["accent"], fg="white").pack(side="left", padx=5)
        tk.Button(btn_frame, text="Cancelar", command=cancelar_config,
                 bg=cores["accent"], fg="white").pack(side="left", padx=5)

    def mostrar_sobre(self):
        """Mostra informações sobre o aplicativo"""
        cores = CORES[self.config_manager.modo_aparencia]
        about_window = tk.Toplevel(self.root)
        about_window.title("Sobre o DocQuantum Calendário")
        about_window.geometry("500x600")
        about_window.configure(bg=cores["primary"])
        about_window.transient(self.root)
        about_window.grab_set()

        about_window.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() - about_window.winfo_width()) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - about_window.winfo_height()) // 2
        about_window.geometry(f"+{x}+{y}")

        info_text = """
DocQuantum Calendário - Versão Aprimorada

Funcionalidades:
• Gerenciamento completo de eventos
• Criptografia de dados com senha
• Exportação para formatos ICS, PDF e CSV
• Interface moderna com temas
• Notificações automáticas
• Anexação de arquivos
• Busca avançada e filtros
• Sistema de Undo/Redo (Ctrl+Z/Ctrl+Y)
• Backup e restauração de dados
• Templates de eventos
• Sistema de metas e produtividade
• Relatórios e dashboard
• Análise de produtividade

Melhorias desta versão:
✓ Sistema de backup automático
✓ Templates de eventos reutilizáveis
✓ Relatórios e estatísticas detalhadas
✓ Exportação para PDF
✓ Sistema de metas e acompanhamento
✓ Dashboard visual
✓ Análise de produtividade
✓ Pesquisa avançada com múltiplos critérios
✓ Interface responsiva para diferentes telas
✓ Alertas personalizáveis

Desenvolvido com Python + Tkinter
        """

        tk.Label(about_window, text=info_text, bg=cores["primary"], fg=cores["text"],
                justify="left", font=("Arial", 10)).pack(padx=20, pady=20)

        tk.Button(about_window, text="Fechar", command=about_window.destroy,
                 bg=cores["accent"], fg="white").pack(pady=10)

    def aplicar_tema(self):
        """Aplica o tema atual na interface"""
        cores = CORES[self.config_manager.modo_aparencia]

        # Configurar estilo TTK
        style = ttk.Style()
        style.theme_use('clam')

        style.configure("Themed.TEntry",
                       fieldbackground=cores["entry_bg"],
                       background=cores["entry_bg"],
                       foreground=cores["entry_fg"],
                       bordercolor=cores["accent"],
                       lightcolor=cores["accent"],
                       darkcolor=cores["accent"],
                       insertcolor=cores["text"],
                       padding=4, relief="flat", borderwidth=1)

        style.configure("Themed.TCombobox",
                       fieldbackground=cores["entry_bg"],
                       background=cores["secondary"],
                       foreground=cores["entry_fg"],
                       arrowcolor=cores["text"],
                       bordercolor=cores["accent"],
                       lightcolor=cores["accent"],
                       darkcolor=cores["accent"],
                       arrowsize=12, padding=2, relief="flat", borderwidth=1)

        style.map("Themed.TCombobox",
                 background=[('active', cores["card_hover"]), ('!disabled', cores["secondary"]), ('readonly', cores["secondary"])],
                 fieldbackground=[('active', cores["entry_bg"]), ('!disabled', cores["entry_bg"]), ('readonly', cores["entry_bg"])],
                 foreground=[('active', cores["text"]), ('!disabled', cores["entry_fg"]), ('readonly', cores["entry_fg"])],
                 arrowcolor=[('active', cores["text"]), ('!disabled', cores["text"]), ('readonly', cores["text"])])

        style.configure("Themed.TButton",
                       background=cores["button_bg"],
                       foreground=cores["button_fg"],
                       bordercolor=cores["accent"],
                       lightcolor=cores["accent"],
                       darkcolor=cores["accent"],
                       padding=6)

        style.map("Themed.TButton",
                 background=[('active', cores["button_active"]), ('pressed', cores["button_active"])])

        # Aplicar cores nos widgets principais
        self.root.configure(bg=cores["primary"])

        for widget_name, widget in self.widgets.items():
            try:
                if 'header' in widget_name:
                    widget.configure(bg=cores["accent"])
                    self._aplicar_tema_recursivo(widget, cores["accent"], "white")
                elif 'frame_filtros' in widget_name or 'frame_form' in widget_name or 'template_widget_frame' in widget_name:
                    widget.configure(bg=cores["secondary"])
                    self._aplicar_tema_recursivo(widget, cores["secondary"], cores["text"])
                elif 'label_titulo_eventos' in widget_name:
                    widget.configure(bg=cores["primary"], fg=cores["text"])
                elif isinstance(widget, (tk.Frame, tk.Canvas)):
                    widget.configure(bg=cores["primary"])
                elif isinstance(widget, tk.Text):
                    widget.configure(bg=cores["entry_bg"], fg=cores["entry_fg"], insertbackground=cores["text"])
            except Exception:
                pass

        self.atualizar_lista()

    def _aplicar_tema_recursivo(self, widget, bg_color, fg_color):
        """Aplica tema recursivamente nos widgets filhos"""
        try:
            widget.configure(bg=bg_color) # Aplica no próprio frame
            for child in widget.winfo_children():
                if isinstance(child, (tk.Label, tk.Button)):
                    child.configure(bg=bg_color, fg=fg_color)
                elif isinstance(child, tk.Frame):
                    self._aplicar_tema_recursivo(child, bg_color, fg_color)
        except Exception:
            pass

    def alternar_tema(self):
        """Alterna entre tema claro e escuro"""
        self.config_manager.modo_aparencia = "dark" if self.config_manager.modo_aparencia == "light" else "light"
        self.config_manager.salvar()
        self.aplicar_tema()
        logger.info(f"Tema alterado para: {self.config_manager.modo_aparencia}")


    def alternar_tela_cheia(self, event=None):
        """Alterna modo tela cheia"""
        self.config_manager.modo_tela_cheia = not self.config_manager.modo_tela_cheia
        self.root.attributes('-fullscreen', self.config_manager.modo_tela_cheia)
        self.config_manager.salvar()
        logger.info(f"Modo tela cheia alterado para: {self.config_manager.modo_tela_cheia}")

    def sair_tela_cheia(self, event=None):
        """Sai do modo tela cheia"""
        if self.config_manager.modo_tela_cheia:
            self.config_manager.modo_tela_cheia = False
            self.root.attributes('-fullscreen', False)
            self.config_manager.salvar()
            logger.info("Saindo do modo tela cheia.")

    def configurar_criptografia(self):
        """Permite reconfigurar a criptografia"""
        logger.info("Abrindo diálogo para configurar criptografia.")
        if self.crypt_manager.criptografia_habilitada:
            resposta = messagebox.askyesno("Criptografia Ativa",
                                         "A criptografia já está ativa. Deseja alterá-la?\n"
                                         "(Isso requer a senha atual)")
            if not resposta:
                return

            if not self.crypt_manager.verificar_senha():
                return

        nova_senha = simpledialog.askstring("Nova Criptografia",
                                           "Digite a nova senha para criptografia\n(deixe vazio para desativar):",
                                           show='*')

        if nova_senha:
            self.crypt_manager.chave_criptografia = self.crypt_manager.gerar_chave_de_senha(nova_senha)
            try:
                with open(ARQUIVO_CHAVE, 'wb') as f:
                    f.write(self.crypt_manager.chave_criptografia)
                self.crypt_manager.criptografia_habilitada = True
                messagebox.showinfo("Sucesso", "Criptografia configurada! Salvando dados...")
                self.evento_manager.salvar_eventos()
                logger.info("Criptografia reconfigurada com sucesso.")
                messagebox.showinfo("Concluído", "Dados protegidos com nova criptografia!")
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao configurar criptografia: {str(e)}")
        else:
            if os.path.exists(ARQUIVO_CHAVE):
                os.remove(ARQUIVO_CHAVE)
            self.crypt_manager.criptografia_habilitada = False
            self.crypt_manager.chave_criptografia = None
            messagebox.showinfo("Sucesso", "Criptografia desativada! Salvando dados...")
            self.evento_manager.salvar_eventos()
            logger.info("Criptografia desativada pelo usuário.")
            messagebox.showinfo("Concluído", "Dados agora são salvos sem criptografia!")

    def criar_backup(self):
        """Cria um backup dos eventos"""
        logger.info("Iniciando criação de backup manual.")
        # O backup já é logado dentro do backup_manager, então aqui só confirmamos a ação do usuário
        caminho_backup = self.evento_manager.backup_manager.criar_backup(self.evento_manager.eventos)
        if caminho_backup:
            messagebox.showinfo("Sucesso", f"Backup criado com sucesso!\n\nLocal: {caminho_backup}")

    def restaurar_backup(self):
        """Restaura um backup"""
        logger.info("Iniciando processo de restauração de backup.")
        backups = self.evento_manager.backup_manager.listar_backups()
        if not backups:
            messagebox.showinfo("Info", "Nenhum backup disponível para restaurar.")
            return

        # Criar janela de seleção de backup
        cores = CORES[self.config_manager.modo_aparencia]
        backup_window = tk.Toplevel(self.root)
        backup_window.title("Restaurar Backup")
        backup_window.geometry("600x400")
        backup_window.configure(bg=cores["primary"])
        backup_window.transient(self.root)
        backup_window.grab_set()

        backup_window.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() - backup_window.winfo_width()) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - backup_window.winfo_height()) // 2
        backup_window.geometry(f"+{x}+{y}")

        tk.Label(backup_window, text="Selecione o backup para restaurar", font=("Arial", 14, "bold"),
                bg=cores["primary"], fg=cores["text"]).pack(pady=10)

        # Frame para lista de backups
        frame_backups = tk.Frame(backup_window, bg=cores["primary"])
        frame_backups.pack(fill="both", expand=True, padx=20, pady=10)

        # Lista de backups
        lista_backups = tk.Listbox(frame_backups, height=10, selectmode="single")
        scrollbar = ttk.Scrollbar(frame_backups, orient="vertical", command=lista_backups.yview)
        lista_backups.configure(yscrollcommand=scrollbar.set)

        for backup in backups:
            lista_backups.insert(tk.END, f"{backup['nome']} - {backup['data'].strftime('%d/%m/%Y %H:%M')} - {backup['tamanho']} bytes")

        lista_backups.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Botões
        btn_frame = tk.Frame(backup_window, bg=cores["primary"])
        btn_frame.pack(side="bottom", pady=10)

        def restaurar_selecionado():
            selecao = lista_backups.curselection()
            if not selecao:
                messagebox.showwarning("Aviso", "Selecione um backup para restaurar.")
                return

            backup = backups[selecao[0]]
            resposta = messagebox.askyesno("Confirmar",
                                         f"Tem certeza que deseja restaurar o backup?\n\n"
                                         f"Arquivo: {backup['nome']}\n"
                                         f"Data: {backup['data'].strftime('%d/%m/%Y %H:%M')}\n\n"
                                         f"Todos os eventos atuais serão substituídos.",
                                         icon="warning")

            if resposta:
                logger.warning(f"Usuário confirmou a restauração do backup: {backup['nome']}")
                eventos_restaurados_raw = self.evento_manager.backup_manager.restaurar_backup(backup['arquivo'])
                if eventos_restaurados_raw is not None:
                    # Garantir que todos os campos, incluindo os novos, existam
                    self.evento_manager.eventos = [self.evento_manager.garantir_campos_evento(e) for e in eventos_restaurados_raw]
                    self.evento_manager.salvar_eventos()
                    self.atualizar_lista()
                    logger.info("Backup restaurado com sucesso.")
                    messagebox.showinfo("Sucesso", "Backup restaurado com sucesso!")
                    backup_window.destroy()

        tk.Button(btn_frame, text="Restaurar Selecionado", command=restaurar_selecionado,
                 bg=cores["accent"], fg="white").pack(side="left", padx=5)

        tk.Button(btn_frame, text="Cancelar", command=backup_window.destroy,
                 bg=cores["accent"], fg="white").pack(side="left", padx=5)

    def exportar_para_pdf(self):
        """Exporta eventos para PDF"""
        eventos = self.evento_manager.filtrar_eventos(self.filtro_atual, self.filtro_busca, self.ordem_atual)
        self.export_manager.exportar_para_pdf(eventos, "Relatório de Eventos")

    def exportar_para_csv(self):
        """Exporta eventos para CSV"""
        eventos = self.evento_manager.filtrar_eventos(self.filtro_atual, self.filtro_busca, self.ordem_atual)
        self.export_manager.exportar_para_csv(eventos)

    def aplicar_template(self, event):
        """Aplica um template selecionado"""
        selecao = self.widgets['template_menu'].current()
        if selecao > 0:  # 0 é "Selecionar Template"
            template_id = selecao - 1
            # Obter data atual ou usar a data do formulário se existir
            data_str = self.widgets['entry_data'].get()
            if data_str:
                data = self.parse_data_br(data_str)
            else:
                data = datetime.date.today()

            if data:
                template = self.template_manager.aplicar_template(template_id, data)
                if template:
                    self.nome_var.set(template['nome'])
                    self.local_var.set(template['local'])
                    self.widgets['text_corpo'].delete("1.0", "end")
                    self.widgets['text_corpo'].insert("1.0", template['corpo'])
                    self.categoria_var.set(template['categoria'])
                    self.prioridade_var.set(template['prioridade'])
                    self.grupo_var.set(template['grupo'])
                    self.etiquetas_var.set(template['etiquetas'])
                    self.hora_var.set(f"{template['hora']:02d}")
                    self.minuto_var.set(f"{template['minuto']:02d}")
                    self.prorrogacao_var.set("")

                    messagebox.showinfo("Sucesso", "Template aplicado com sucesso!")

    def gerenciar_templates(self):
        """Mostra janela de gerenciamento de templates"""
        cores = CORES[self.config_manager.modo_aparencia]
        template_window = tk.Toplevel(self.root)
        template_window.title("Gerenciar Templates")
        template_window.geometry("600x400")
        template_window.configure(bg=cores["primary"])
        template_window.transient(self.root)
        template_window.grab_set()

        template_window.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() - template_window.winfo_width()) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - template_window.winfo_height()) // 2
        template_window.geometry(f"+{x}+{y}")

        tk.Label(template_window, text="Gerenciar Templates de Eventos", font=("Arial", 14, "bold"),
                bg=cores["primary"], fg=cores["text"]).pack(pady=10)

        # Frame para lista de templates
        frame_templates = tk.Frame(template_window, bg=cores["primary"])
        frame_templates.pack(fill="both", expand=True, padx=20, pady=10)

        # Lista de templates
        lista_templates = tk.Listbox(frame_templates, height=10, selectmode="single")
        scrollbar = ttk.Scrollbar(frame_templates, orient="vertical", command=lista_templates.yview)
        lista_templates.configure(yscrollcommand=scrollbar.set)

        for i, template in enumerate(self.template_manager.templates):
            lista_templates.insert(tk.END, f"Template {i+1}: {template['nome']}")

        lista_templates.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Botões
        btn_frame = tk.Frame(template_window, bg=cores["primary"])
        btn_frame.pack(side="bottom", pady=10)

        def criar_template():
            # Verificar se há um evento sendo editado
            if self.editando_id:
                evento = self.evento_manager.buscar_evento_por_id(self.editando_id)
                if evento:
                    self.template_manager.adicionar_template(evento)
                    self._atualizar_templates()
                    # Atualizar lista
                    lista_templates.delete(0, tk.END)
                    for i, template in enumerate(self.template_manager.templates):
                        lista_templates.insert(tk.END, f"Template {i+1}: {template['nome']}")
                    messagebox.showinfo("Sucesso", "Template criado com sucesso!")
                else:
                    messagebox.showwarning("Aviso", "Nenhum evento selecionado para criar template.")
            else:
                messagebox.showwarning("Aviso", "Edite um evento primeiro para criar um template.")

        def excluir_template():
            selecao = lista_templates.curselection()
            if not selecao:
                messagebox.showwarning("Aviso", "Selecione um template para excluir.")
                return

            resposta = messagebox.askyesno("Confirmar", "Tem certeza que deseja excluir este template?")
            if resposta:
                self.template_manager.remover_template(selecao[0])
                self._atualizar_templates()
                # Atualizar lista
                lista_templates.delete(selecao[0])
                messagebox.showinfo("Sucesso", "Template excluído com sucesso!")

        tk.Button(btn_frame, text="Criar Template do Evento Atual", command=criar_template,
                 bg=cores["accent"], fg="white").pack(side="left", padx=5)

        tk.Button(btn_frame, text="Excluir Template Selecionado", command=excluir_template,
                 bg=cores["accent"], fg="white").pack(side="left", padx=5)

        tk.Button(btn_frame, text="Fechar", command=template_window.destroy,
                 bg=cores["accent"], fg="white").pack(side="left", padx=5)

    def mostrar_pesquisa_avancada(self):
        """Mostra janela de pesquisa avançada"""
        cores = CORES[self.config_manager.modo_aparencia]
        pesquisa_window = tk.Toplevel(self.root)
        pesquisa_window.title("Pesquisa Avançada")
        pesquisa_window.geometry("500x400")
        pesquisa_window.configure(bg=cores["primary"])
        pesquisa_window.transient(self.root)
        pesquisa_window.grab_set()

        pesquisa_window.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() - pesquisa_window.winfo_width()) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - pesquisa_window.winfo_height()) // 2
        pesquisa_window.geometry(f"+{x}+{y}")

        tk.Label(pesquisa_window, text="Pesquisa Avançada", font=("Arial", 14, "bold"),
                bg=cores["primary"], fg=cores["text"]).pack(pady=10)

        # Frame principal
        main_frame = tk.Frame(pesquisa_window, bg=cores["primary"])
        main_frame.pack(fill="both", expand=True, padx=20, pady=10)

        # Variáveis para os campos
        data_inicio_var = tk.StringVar()
        data_fim_var = tk.StringVar()
        prioridade_var = tk.StringVar()
        categoria_var = tk.StringVar()
        concluido_var = tk.StringVar()

        # Campos de pesquisa
        tk.Label(main_frame, text="Data Início (DD/MM/AAAA):", bg=cores["primary"], fg=cores["text"]).pack(anchor="w")
        entry_data_inicio = ttk.Entry(main_frame, textvariable=data_inicio_var, width=15)
        entry_data_inicio.pack(fill="x", pady=(0, 10))

        tk.Label(main_frame, text="Data Fim (DD/MM/AAAA):", bg=cores["primary"], fg=cores["text"]).pack(anchor="w")
        entry_data_fim = ttk.Entry(main_frame, textvariable=data_fim_var, width=15)
        entry_data_fim.pack(fill="x", pady=(0, 10))

        tk.Label(main_frame, text="Prioridade:", bg=cores["primary"], fg=cores["text"]).pack(anchor="w")
        combo_prioridade = ttk.Combobox(main_frame, textvariable=prioridade_var,
                                       values=["", "ALTA", "MÉDIA", "BAIXA"], state="readonly")
        combo_prioridade.pack(fill="x", pady=(0, 10))

        tk.Label(main_frame, text="Categoria:", bg=cores["primary"], fg=cores["text"]).pack(anchor="w")
        entry_categoria = ttk.Entry(main_frame, textvariable=categoria_var)
        entry_categoria.pack(fill="x", pady=(0, 10))

        tk.Label(main_frame, text="Concluído:", bg=cores["primary"], fg=cores["text"]).pack(anchor="w")
        combo_concluido = ttk.Combobox(main_frame, textvariable=concluido_var,
                                      values=["", "Sim", "Não"], state="readonly")
        combo_concluido.pack(fill="x", pady=(0, 20))

        def pesquisar():
            criterios = {}

            if data_inicio_var.get():
                criterios["data_inicio"] = data_inicio_var.get()
            if data_fim_var.get():
                criterios["data_fim"] = data_fim_var.get()
            if prioridade_var.get():
                criterios["prioridade"] = prioridade_var.get()
            if categoria_var.get():
                criterios["categoria"] = categoria_var.get()
            if concluido_var.get():
                criterios["concluido"] = concluido_var.get()

            # Aplicar filtros
            self.filtro_busca = ""  # Limpar busca simples
            eventos_filtrados = self.evento_manager.filtrar_eventos(
                self.filtro_atual, "", self.ordem_atual, criterios)

            # Atualizar lista
            self.PAGINA_ATUAL = 1
            self.atualizar_lista_com_eventos(eventos_filtrados)

            pesquisa_window.destroy()

        def limpar():
            data_inicio_var.set("")
            data_fim_var.set("")
            prioridade_var.set("")
            categoria_var.set("")
            concluido_var.set("")

        # Botões
        btn_frame = tk.Frame(main_frame, bg=cores["primary"])
        btn_frame.pack(side="bottom", pady=10)

        tk.Button(btn_frame, text="Pesquisar", command=pesquisar,
                 bg=cores["accent"], fg="white").pack(side="left", padx=5)

        tk.Button(btn_frame, text="Limpar", command=limpar,
                 bg=cores["accent"], fg="white").pack(side="left", padx=5)

        tk.Button(btn_frame, text="Cancelar", command=pesquisa_window.destroy,
                 bg=cores["accent"], fg="white").pack(side="left", padx=5)

    def atualizar_lista_com_eventos(self, eventos):
        """Atualiza a lista com eventos específicos (para pesquisa avançada)"""
        total_eventos = len(eventos)

        inicio = (self.PAGINA_ATUAL - 1) * EVENTOS_POR_PAGINA
        fim = inicio + EVENTOS_POR_PAGINA
        lista_a_exibir = eventos[inicio:fim]

        # Limpar cards existentes
        for widget in self.widgets['frame_lista'].winfo_children():
            widget.destroy()
        self.cards_por_id = {}


        # Criar novos cards
        frame_lista = self.widgets['frame_lista']

        for i, evento in enumerate(lista_a_exibir):
            row, col = divmod(i, 4)
            evento_id = evento["id"]

            card = EventCard(frame_lista, evento, self.config_manager, self.evento_manager,
                           self.mostrar_detalhes_evento, self.mostrar_menu_opcoes_popup,
                           self.filtro_busca)
            card.grid(row=row, column=col, padx=2, pady=5, sticky="nsew")
            self.cards_por_id[evento_id] = card

        # Mostrar mensagem se não houver eventos
        cores = CORES[self.config_manager.modo_aparencia]
        if not lista_a_exibir:
            empty_label = tk.Label(frame_lista, text="📅\n\nNenhum evento para exibir",
                                  font=("Arial", 16), bg=cores["primary"], fg=cores["text"],
                                  justify="center")
            empty_label.place(relx=0.5, rely=0.4, anchor="center")

        total_paginas = (total_eventos + EVENTOS_POR_PAGINA - 1) // EVENTOS_POR_PAGINA if EVENTOS_POR_PAGINA > 0 else 0
        self.contador_var.set(f"Eventos: {len(lista_a_exibir)}/{total_eventos} (Página {self.PAGINA_ATUAL}/{total_paginas})")

        self.adicionar_controles_paginacao(total_eventos)

    def gerar_relatorio(self, tipo_relatorio):
        """Gera e exibe um relatório"""
        relatorio = self.evento_manager.gerar_relatorio(tipo_relatorio)

        cores = CORES[self.config_manager.modo_aparencia]
        relatorio_window = tk.Toplevel(self.root)
        relatorio_window.title(f"Relatório {tipo_relatorio.capitalize()}")
        relatorio_window.geometry("800x600")
        relatorio_window.configure(bg=cores["primary"])
        relatorio_window.transient(self.root)
        relatorio_window.grab_set()

        relatorio_window.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() - relatorio_window.winfo_width()) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - relatorio_window.winfo_height()) // 2
        relatorio_window.geometry(f"+{x}+{y}")

        # Frame principal com scroll
        main_frame = tk.Frame(relatorio_window, bg=cores["primary"])
        main_frame.pack(fill="both", expand=True, padx=20, pady=10)

        canvas = tk.Canvas(main_frame, bg=cores["primary"], highlightthickness=0)
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg=cores["primary"])

        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # CORREÇÃO: Adicionando o pack do canvas e scrollbar que estava faltando
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        self._configurar_scroll_mouse(canvas)

        # Título
        titulo_periodo = f"{relatorio['periodo']['inicio'].strftime('%d/%m/%Y')} a {relatorio['periodo']['fim'].strftime('%d/%m/%Y')}"
        tk.Label(scrollable_frame, text=f"Relatório {tipo_relatorio.capitalize()} - {titulo_periodo}",
                font=("Arial", 16, "bold"), bg=cores["primary"], fg=cores["text"]).pack(pady=10)

        # Estatísticas
        stats_frame = tk.Frame(scrollable_frame, bg=cores["secondary"], relief="solid", bd=1)
        stats_frame.pack(fill="x", pady=10, padx=10)

        stats_text = f"""
        Estatísticas:
        • Total de Eventos: {relatorio['estatisticas']['total_eventos']}
        • Eventos Concluídos: {relatorio['estatisticas']['eventos_concluidos']}
        • Taxa de Conclusão: {relatorio['estatisticas']['taxa_conclusao']}%

        Prioridades:
        • ALTA: {relatorio['estatisticas']['prioridades']['ALTA']}
        • MÉDIA: {relatorio['estatisticas']['prioridades']['MÉDIA']}
        • BAIXA: {relatorio['estatisticas']['prioridades']['BAIXA']}
        • NÃO DEFINIDA: {relatorio['estatisticas']['prioridades']['NÃO DEFINIDA']}
        """

        tk.Label(stats_frame, text=stats_text, font=("Arial", 10),
                bg=cores["secondary"], fg=cores["text"], justify="left").pack(padx=10, pady=10)

        # Categorias
        if relatorio['estatisticas']['categorias']:
            cats_frame = tk.Frame(scrollable_frame, bg=cores["secondary"], relief="solid", bd=1)
            cats_frame.pack(fill="x", pady=10, padx=10)

            cats_text = "Eventos por Categoria:\n"
            for categoria, quantidade in relatorio['estatisticas']['categorias'].items():
                cats_text += f"• {categoria}: {quantidade}\n"

            tk.Label(cats_frame, text=cats_text, font=("Arial", 10),
                    bg=cores["secondary"], fg=cores["text"], justify="left").pack(padx=10, pady=10)

        # Eventos Importantes
        if relatorio['eventos_importantes']:
            imp_frame = tk.Frame(scrollable_frame, bg=cores["secondary"], relief="solid", bd=1)
            imp_frame.pack(fill="x", pady=10, padx=10)

            imp_text = "Eventos Importantes Pendentes:\n"
            for evento in relatorio['eventos_importantes']:
                imp_text += f"• {evento['nome']} ({self.evento_manager.get_effective_date(evento).strftime('%d/%m/%Y')})\n"

            tk.Label(imp_frame, text=imp_text, font=("Arial", 10),
                    bg=cores["secondary"], fg=cores["text"], justify="left").pack(padx=10, pady=10)

        # Eventos Próximos
        if relatorio['eventos_proximos']:
            prox_frame = tk.Frame(scrollable_frame, bg=cores["secondary"], relief="solid", bd=1)
            prox_frame.pack(fill="x", pady=10, padx=10)

            prox_text = "Próximos Eventos:\n"
            for evento in relatorio['eventos_proximos']:
                dias = (self.evento_manager.get_effective_date(evento) - datetime.date.today()).days
                status = "HOJE" if dias == 0 else f"em {dias} dias"
                prox_text += f"• {evento['nome']} ({status})\n"

            tk.Label(prox_frame, text=prox_text, font=("Arial", 10),
                    bg=cores["secondary"], fg=cores["text"], justify="left").pack(padx=10, pady=10)

        # Botões
        btn_frame = tk.Frame(scrollable_frame, bg=cores["primary"])
        btn_frame.pack(side="bottom", pady=20)

        def exportar_pdf():
            eventos_do_periodo = [e for e in self.evento_manager.eventos if relatorio['periodo']['inicio'] <= self.evento_manager.get_effective_date(e) <= relatorio['periodo']['fim']]
            self.export_manager.exportar_para_pdf(eventos_do_periodo, f"Relatório {tipo_relatorio.capitalize()}")

        def exportar_csv():
            eventos_do_periodo = [e for e in self.evento_manager.eventos if relatorio['periodo']['inicio'] <= self.evento_manager.get_effective_date(e) <= relatorio['periodo']['fim']]
            self.export_manager.exportar_para_csv(eventos_do_periodo)

        tk.Button(btn_frame, text="Exportar PDF", command=exportar_pdf,
                 bg=cores["accent"], fg="white").pack(side="left", padx=5)

        tk.Button(btn_frame, text="Exportar CSV", command=exportar_csv,
                 bg=cores["accent"], fg="white").pack(side="left", padx=5)

        tk.Button(btn_frame, text="Fechar", command=relatorio_window.destroy,
                 bg=cores["accent"], fg="white").pack(side="left", padx=5)

    def mostrar_dashboard(self):
        """Mostra o dashboard com visão geral"""
        cores = CORES[self.config_manager.modo_aparencia]
        dashboard_window = tk.Toplevel(self.root)
        dashboard_window.title("Dashboard")
        dashboard_window.geometry("1000x700")
        dashboard_window.configure(bg=cores["primary"])
        dashboard_window.transient(self.root)
        dashboard_window.grab_set()

        dashboard_window.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() - dashboard_window.winfo_width()) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - dashboard_window.winfo_height()) // 2
        dashboard_window.geometry(f"+{x}+{y}")

        # Frame principal com scroll
        main_frame = tk.Frame(dashboard_window, bg=cores["primary"])
        main_frame.pack(fill="both", expand=True, padx=20, pady=10)

        canvas = tk.Canvas(main_frame, bg=cores["primary"], highlightthickness=0)
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg=cores["primary"])

        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        self._configurar_scroll_mouse(canvas)

        # Título
        tk.Label(scrollable_frame, text="Dashboard - Visão Geral",
                font=("Arial", 18, "bold"), bg=cores["primary"], fg=cores["text"]).pack(pady=10)

        # Gerar relatório mensal
        relatorio = self.evento_manager.gerar_relatorio("mensal")

        # Estatísticas rápidas
        stats_frame = tk.Frame(scrollable_frame, bg=cores["secondary"], relief="solid", bd=1)
        stats_frame.pack(fill="x", pady=10, padx=10)

        stats_text = f"""
        Estatísticas do Mês:
        • Total de Eventos: {relatorio['estatisticas']['total_eventos']}
        • Eventos Concluídos: {relatorio['estatisticas']['eventos_concluidos']}
        • Taxa de Conclusão: {relatorio['estatisticas']['taxa_conclusao']}%
        • Eventos por Prioridade:
          - ALTA: {relatorio['estatisticas']['prioridades']['ALTA']}
          - MÉDIA: {relatorio['estatisticas']['prioridades']['MÉDIA']}
          - BAIXA: {relatorio['estatisticas']['prioridades']['BAIXA']}
        """

        tk.Label(stats_frame, text=stats_text, font=("Arial", 10),
                bg=cores["secondary"], fg=cores["text"], justify="left").pack(padx=10, pady=10)

        # Gráfico de pizza de prioridades - COM CORREÇÃO DO ERRO
        try:
            # Verificar se há dados válidos para o gráfico
            prioridades = ['ALTA', 'MÉDIA', 'BAIXA', 'NÃO DEFINIDA']
            valores = [
                relatorio['estatisticas']['prioridades']['ALTA'],
                relatorio['estatisticas']['prioridades']['MÉDIA'],
                relatorio['estatisticas']['prioridades']['BAIXA'],
                relatorio['estatisticas']['prioridades']['NÃO DEFINIDA']
            ]

            # Verificar se todos os valores são zero ou a soma é zero
            total_valores = sum(valores)
            
            if total_valores == 0:
                # Exibir mensagem em vez do gráfico
                no_data_frame = tk.Frame(scrollable_frame, bg=cores["secondary"], relief="solid", bd=1)
                no_data_frame.pack(fill="x", pady=10, padx=10)
                tk.Label(no_data_frame, text="📊 Não há dados suficientes para exibir o gráfico de prioridades\n\nAdicione eventos com diferentes prioridades para visualizar estatísticas.",
                        font=("Arial", 10), bg=cores["secondary"], fg=cores["text"], 
                        justify="center").pack(padx=10, pady=20)
            else:
                # Criar gráfico apenas se houver dados válidos
                fig = Figure(figsize=(5, 4), dpi=100)
                ax = fig.add_subplot(111)

                cores_grafico = [
                    CORES[self.config_manager.modo_aparencia]['priority_high'],
                    CORES[self.config_manager.modo_aparencia]['priority_medium'],
                    CORES[self.config_manager.modo_aparencia]['priority_low'],
                    CORES[self.config_manager.modo_aparencia]['progress_past']
                ]

                # Filtrar valores zero para evitar divisão por zero
                labels_filtrados = []
                valores_filtrados = []
                cores_filtradas = []
                
                for i, valor in enumerate(valores):
                    if valor > 0:
                        labels_filtrados.append(prioridades[i])
                        valores_filtrados.append(valor)
                        cores_filtradas.append(cores_grafico[i])
                
                if valores_filtrados:  # Se ainda houver valores após filtrar zeros
                    ax.pie(valores_filtrados, labels=labels_filtrados, autopct='%1.1f%%', colors=cores_filtradas)
                    ax.set_title('Distribuição por Prioridade')

                    chart_frame = tk.Frame(scrollable_frame, bg=cores["primary"])
                    chart_frame.pack(fill="x", pady=10, padx=10)

                    canvas_chart = FigureCanvasTkAgg(fig, chart_frame)
                    canvas_chart.draw()
                    canvas_chart.get_tk_widget().pack(fill="both", expand=True)
                else:
                    # Caso todos os valores sejam zero após filtrar
                    no_data_frame = tk.Frame(scrollable_frame, bg=cores["secondary"], relief="solid", bd=1)
                    no_data_frame.pack(fill="x", pady=10, padx=10)
                    tk.Label(no_data_frame, text="📊 Não há eventos com prioridades definidas para exibir no gráfico",
                            font=("Arial", 10), bg=cores["secondary"], fg=cores["text"], 
                            justify="center").pack(padx=10, pady=20)

        except Exception as e:
            logger.error(f"Erro ao criar gráfico do dashboard: {e}", exc_info=True)
            # Fallback: exibir mensagem de erro
            error_frame = tk.Frame(scrollable_frame, bg=cores["secondary"], relief="solid", bd=1)
            error_frame.pack(fill="x", pady=10, padx=10)
            tk.Label(error_frame, text=f"❌ Erro ao gerar gráfico: {str(e)}",
                    font=("Arial", 10), bg=cores["secondary"], fg=cores["text"], 
                    justify="center").pack(padx=10, pady=20)

        # Eventos importantes
        if relatorio['eventos_importantes']:
            imp_frame = tk.Frame(scrollable_frame, bg=cores["secondary"], relief="solid", bd=1)
            imp_frame.pack(fill="x", pady=10, padx=10)

            imp_text = "📌 Eventos Importantes Pendentes:\n\n"
            for evento in relatorio['eventos_importantes']:
                dias = (self.evento_manager.get_effective_date(evento) - datetime.date.today()).days
                status = "HOJE" if dias == 0 else f"em {dias} dias" if dias > 0 else f"{abs(dias)} dias atrás"
                imp_text += f"• {evento['nome']} ({status})\n"

            tk.Label(imp_frame, text=imp_text, font=("Arial", 10),
                    bg=cores["secondary"], fg=cores["text"], justify="left").pack(padx=10, pady=10)

        # Próximos eventos
        if relatorio['eventos_proximos']:
            prox_frame = tk.Frame(scrollable_frame, bg=cores["secondary"], relief="solid", bd=1)
            prox_frame.pack(fill="x", pady=10, padx=10)

            prox_text = "🔔 Próximos Eventos:\n\n"
            for evento in relatorio['eventos_proximos']:
                dias = (self.evento_manager.get_effective_date(evento) - datetime.date.today()).days
                status = "HOJE" if dias == 0 else f"em {dias} dias"
                prox_text += f"• {evento['nome']} ({status})\n"

            tk.Label(prox_frame, text=prox_text, font=("Arial", 10),
                    bg=cores["secondary"], fg=cores["text"], justify="left").pack(padx=10, pady=10)

        # Botões de ação
        btn_frame = tk.Frame(scrollable_frame, bg=cores["primary"])
        btn_frame.pack(side="bottom", pady=20)

        tk.Button(btn_frame, text="Ver Relatório Completo",
                 command=lambda: self.gerar_relatorio("mensal"),
                 bg=cores["accent"], fg="white").pack(side="left", padx=5)

        tk.Button(btn_frame, text="Fechar", command=dashboard_window.destroy,
                 bg=cores["accent"], fg="white").pack(side="left", padx=5)

    def mostrar_analise_produtividade(self):
        """Mostra análise de produtividade"""
        dados = self.evento_manager.analisar_produtividade("mensal")

        cores = CORES[self.config_manager.modo_aparencia]
        analise_window = tk.Toplevel(self.root)
        analise_window.title("Análise de Produtividade")
        analise_window.geometry("900x600")
        analise_window.configure(bg=cores["primary"])
        analise_window.transient(self.root)
        analise_window.grab_set()

        analise_window.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() - analise_window.winfo_width()) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - analise_window.winfo_height()) // 2
        analise_window.geometry(f"+{x}+{y}")

        # Frame principal
        main_frame = tk.Frame(analise_window, bg=cores["primary"])
        main_frame.pack(fill="both", expand=True, padx=20, pady=10)

        # Título
        tk.Label(main_frame, text="Análise de Produtividade (Últimos 30 dias)",
                font=("Arial", 16, "bold"), bg=cores["primary"], fg=cores["text"]).pack(pady=10)

        # Gráfico de linha - Taxa de conclusão
        try:
            fig = Figure(figsize=(8, 6), dpi=100)
            ax = fig.add_subplot(111)

            datas = [d['data'].strftime('%d/%m') for d in dados]
            taxas = [d['taxa_conclusao'] for d in dados]

            ax.plot(datas, taxas, marker='o', color=cores['accent'])
            ax.set_xlabel('Data')
            ax.set_ylabel('Taxa de Conclusão (%)')
            ax.set_title('Evolução da Produtividade')
            ax.grid(True)

            # Rotacionar labels do eixo X para melhor visualização
            plt.setp(ax.xaxis.get_majorticklabels(), rotation=45)
            
            fig.tight_layout()

            chart_frame = tk.Frame(main_frame, bg=cores["primary"])
            chart_frame.pack(fill="both", expand=True, pady=10)

            canvas_chart = FigureCanvasTkAgg(fig, chart_frame)
            canvas_chart.draw()
            canvas_chart.get_tk_widget().pack(fill="both", expand=True)
        except Exception as e:
            logger.error(f"Erro ao criar gráfico de produtividade: {e}", exc_info=True)
            tk.Label(main_frame, text="Erro ao gerar gráfico de produtividade",
                    bg=cores["primary"], fg=cores["text"]).pack(pady=10)

        # Estatísticas resumidas
        stats_frame = tk.Frame(main_frame, bg=cores["secondary"], relief="solid", bd=1)
        stats_frame.pack(fill="x", pady=10)

        total_eventos = sum(d['total_eventos'] for d in dados)
        total_concluidos = sum(d['eventos_concluidos'] for d in dados)
        taxa_media = (total_concluidos / total_eventos * 100) if total_eventos > 0 else 0

        stats_text = f"""
        Estatísticas dos Últimos 30 Dias:
        • Total de Eventos: {total_eventos}
        • Eventos Concluídos: {total_concluidos}
        • Taxa Média de Conclusão: {taxa_media:.2f}%
        • Melhor Dia: {max(dados, key=lambda x: x['taxa_conclusao'])['data'].strftime('%d/%m/%Y')} ({max(dados, key=lambda x: x['taxa_conclusao'])['taxa_conclusao']:.2f}%)
        """

        tk.Label(stats_frame, text=stats_text, font=("Arial", 10),
                bg=cores["secondary"], fg=cores["text"], justify="left").pack(padx=10, pady=10)

        # Botão fechar
        tk.Button(main_frame, text="Fechar", command=analise_window.destroy,
                 bg=cores["accent"], fg="white").pack(pady=10)

    def gerenciar_metas(self):
        """Mostra janela de gerenciamento de metas"""
        cores = CORES[self.config_manager.modo_aparencia]
        metas_window = tk.Toplevel(self.root)
        metas_window.title("Gerenciar Metas")
        metas_window.geometry("600x500")
        metas_window.configure(bg=cores["primary"])
        metas_window.transient(self.root)
        metas_window.grab_set()

        metas_window.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() - metas_window.winfo_width()) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - metas_window.winfo_height()) // 2
        metas_window.geometry(f"+{x}+{y}")

        tk.Label(metas_window, text="Gerenciar Metas", font=("Arial", 16, "bold"),
                bg=cores["primary"], fg=cores["text"]).pack(pady=10)

        # Frame para lista de metas
        frame_metas = tk.Frame(metas_window, bg=cores["primary"])
        frame_metas.pack(fill="both", expand=True, padx=20, pady=10)

        # Lista de metas
        lista_metas = tk.Listbox(frame_metas, height=10, selectmode="single")
        scrollbar = ttk.Scrollbar(frame_metas, orient="vertical", command=lista_metas.yview)
        lista_metas.configure(yscrollcommand=scrollbar.set)

        for meta in self.meta_manager.metas:
            status = "✓" if meta['concluida'] else f"{meta['progresso']}/{meta['valor']}"
            lista_metas.insert(tk.END, f"{meta['descricao']} - {status}")

        lista_metas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Frame para adicionar nova meta
        frame_nova_meta = tk.Frame(metas_window, bg=cores["primary"])
        frame_nova_meta.pack(fill="x", padx=20, pady=10)

        tk.Label(frame_nova_meta, text="Nova Meta:", bg=cores["primary"], fg=cores["text"]).pack(anchor="w")

        # Campos para nova meta
        frame_campos = tk.Frame(frame_nova_meta, bg=cores["primary"])
        frame_campos.pack(fill="x", pady=5)

        tk.Label(frame_campos, text="Descrição:", bg=cores["primary"], fg=cores["text"], width=10).grid(row=0, column=0, sticky="w")
        descricao_var = tk.StringVar()
        entry_descricao = ttk.Entry(frame_campos, textvariable=descricao_var, width=20)
        entry_descricao.grid(row=0, column=1, padx=5, sticky="ew")

        tk.Label(frame_campos, text="Tipo:", bg=cores["primary"], fg=cores["text"], width=10).grid(row=0, column=2, sticky="w")
        tipo_var = tk.StringVar()
        combo_tipo = ttk.Combobox(frame_campos, textvariable=tipo_var,
                                 values=["eventos_concluidos", "eventos_criados", "produtividade"],
                                 state="readonly", width=15)
        combo_tipo.grid(row=0, column=3, padx=5, sticky="ew")

        tk.Label(frame_campos, text="Valor:", bg=cores["primary"], fg=cores["text"], width=10).grid(row=1, column=0, sticky="w")
        valor_var = tk.StringVar()
        entry_valor = ttk.Entry(frame_campos, textvariable=valor_var, width=10)
        entry_valor.grid(row=1, column=1, padx=5, sticky="w")

        tk.Label(frame_campos, text="Período:", bg=cores["primary"], fg=cores["text"], width=10).grid(row=1, column=2, sticky="w")
        periodo_var = tk.StringVar()
        combo_periodo = ttk.Combobox(frame_campos, textvariable=periodo_var,
                                    values=["diario", "semanal", "mensal"],
                                    state="readonly", width=15)
        combo_periodo.grid(row=1, column=3, padx=5, sticky="ew")

        frame_campos.columnconfigure(1, weight=1)
        frame_campos.columnconfigure(3, weight=1)

        # Botões
        btn_frame = tk.Frame(metas_window, bg=cores["primary"])
        btn_frame.pack(side="bottom", pady=10)

        def adicionar_meta():
            descricao = descricao_var.get().strip()
            tipo = tipo_var.get()
            valor = valor_var.get()
            periodo = periodo_var.get()

            if not all([descricao, tipo, valor, periodo]):
                messagebox.showwarning("Aviso", "Preencha todos os campos.")
                return

            try:
                valor_num = float(valor) if tipo == "produtividade" else int(valor)
                meta = self.meta_manager.adicionar_meta(tipo, valor_num, periodo, descricao)

                # Atualizar lista
                lista_metas.insert(tk.END, f"{meta['descricao']} - 0/{meta['valor']}")

                # Limpar campos
                descricao_var.set("")
                tipo_var.set("")
                valor_var.set("")
                periodo_var.set("")

                messagebox.showinfo("Sucesso", "Meta adicionada com sucesso!")
            except ValueError:
                messagebox.showerror("Erro", "Valor deve ser um número válido.")

        def excluir_meta():
            selecao = lista_metas.curselection()
            if not selecao:
                messagebox.showwarning("Aviso", "Selecione uma meta para excluir.")
                return

            meta = self.meta_manager.metas[selecao[0]]
            resposta = messagebox.askyesno("Confirmar", f"Tem certeza que deseja excluir a meta?\n\n{meta['descricao']}")

            if resposta:
                self.meta_manager.remover_meta(meta['id'])
                lista_metas.delete(selecao[0])
                messagebox.showinfo("Sucesso", "Meta excluída com sucesso!")

        tk.Button(btn_frame, text="Adicionar Meta", command=adicionar_meta,
                 bg=cores["accent"], fg="white").pack(side="left", padx=5)

        tk.Button(btn_frame, text="Excluir Meta", command=excluir_meta,
                 bg=cores["accent"], fg="white").pack(side="left", padx=5)

        tk.Button(btn_frame, text="Fechar", command=metas_window.destroy,
                 bg=cores["accent"], fg="white").pack(side="left", padx=5)

    def verificar_progresso_metas(self):
        """Verifica e mostra o progresso das metas"""
        self.meta_manager.atualizar_progresso(self.evento_manager)

        cores = CORES[self.config_manager.modo_aparencia]
        progresso_window = tk.Toplevel(self.root)
        progresso_window.title("Progresso das Metas")
        progresso_window.geometry("500x400")
        progresso_window.configure(bg=cores["primary"])
        progresso_window.transient(self.root)
        progresso_window.grab_set()

        progresso_window.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() - progresso_window.winfo_width()) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - progresso_window.winfo_height()) // 2
        progresso_window.geometry(f"+{x}+{y}")

        tk.Label(progresso_window, text="Progresso das Metas", font=("Arial", 16, "bold"),
                bg=cores["primary"], fg=cores["text"]).pack(pady=10)

        # Frame para lista de metas
        frame_metas = tk.Frame(progresso_window, bg=cores["primary"])
        frame_metas.pack(fill="both", expand=True, padx=20, pady=10)

        # Lista de metas com progresso
        for meta in self.meta_manager.metas:
            meta_frame = tk.Frame(frame_metas, bg=cores["secondary"], relief="solid", bd=1)
            meta_frame.pack(fill="x", pady=5)

            # Barra de progresso
            progresso = min(meta['progresso'] / meta['valor'] * 100, 100) if meta['valor'] > 0 else 0
            progresso_bar = tk.Canvas(meta_frame, height=20, bg=cores["secondary"], highlightthickness=0)
            progresso_bar.create_rectangle(0, 0, progresso, 20, fill=cores["accent"], outline="")
            progresso_bar.create_text(50, 10, text=f"{progresso:.1f}%", fill=cores["text"])
            progresso_bar.pack(fill="x", padx=5, pady=5)

            # Informações da meta
            status = "Concluída" if meta['concluida'] else "Em andamento"
            meta_text = f"{meta['descricao']} - {meta['progresso']}/{meta['valor']} - {status}"
            tk.Label(meta_frame, text=meta_text, bg=cores["secondary"], fg=cores["text"]).pack(padx=5, pady=5)

        # Botão fechar
        tk.Button(progresso_window, text="Fechar", command=progresso_window.destroy,
                 bg=cores["accent"], fg="white").pack(pady=10)

    # Métodos de manipulação de eventos
    def alterar_filtro(self, event):
        """Altera o filtro de eventos"""
        self.filtro_atual = self.filtro_var.get()
        self.PAGINA_ATUAL = 1
        self.atualizar_lista()

    def filtrar_por_busca(self, event=None):
        """Filtra eventos por termo de busca"""
        self.filtro_busca = self.widgets['entry_busca'].get().strip()
        self.PAGINA_ATUAL = 1
        self.atualizar_lista()

    def parse_data_br(self, data_str):
        """Converte string de data BR para objeto date"""
        try:
            return datetime.datetime.strptime(data_str, "%d/%m/%Y").date()
        except ValueError:
            return None

    def formata_data_br(self, data_date):
        """Formata data para padrão brasileiro"""
        if data_date:
            return data_date.strftime("%d/%m/%Y")
        return ""

    def obter_dados_formulario(self):
        """Obtém e valida dados do formulário"""
        nome = self.nome_var.get().strip()
        data_str = self.widgets['entry_data'].get().strip()
        prorrogacao_str = self.prorrogacao_var.get().strip()
        local = self.local_var.get().strip()
        categoria = self.categoria_var.get().strip()
        prioridade = self.prioridade_var.get()
        grupo = self.grupo_var.get().strip()
        etiquetas = self.etiquetas_var.get().strip()

        if not nome:
            messagebox.showwarning("Atenção", "Informe o nome do evento.")
            return None

        if len(nome) > 100:
            messagebox.showwarning("Atenção", "O nome do evento deve ter no máximo 100 caracteres.")
            return None

        data = self.parse_data_br(data_str)
        if data is None:
            messagebox.showerror("Erro", "Data original inválida! Use o formato DD/MM/AAAA.")
            return None

        data_prorrogacao = None
        if prorrogacao_str:
            data_prorrogacao = self.parse_data_br(prorrogacao_str)
            if data_prorrogacao is None:
                messagebox.showerror("Erro", "Data de prorrogação inválida! Use o formato DD/MM/AAAA.")
                return None

        hoje = datetime.date.today()
        data_efetiva = data_prorrogacao if data_prorrogacao else data

        if data_efetiva < hoje:
            if not messagebox.askyesno("Confirmação", "A data efetiva do evento é anterior a hoje. Deseja continuar?"):
                return None

        try:
            hora = int(self.hora_var.get())
            if not 0 <= hora <= 23:
                messagebox.showerror("Erro", "Hora inválida! Deve estar entre 00 e 23.")
                return None
        except ValueError:
            messagebox.showerror("Erro", "Hora inválida! Deve ser um número entre 00 e 23.")
            return None

        try:
            minuto = int(self.minuto_var.get())
            if not 0 <= minuto <= 59:
                messagebox.showerror("Erro", "Minutos inválidos! Deve estar entre 00 e 59.")
                return None
        except ValueError:
            messagebox.showerror("Erro", "Minutos inválidos! Deve ser um número entre 00 e 59.")
            return None

        campos_validar = [
            (local, "Local", 200),
            (categoria, "Categoria", 50),
            (grupo, "Grupo", 50),
            (etiquetas, "Etiquetas", 100)
        ]

        for valor, nome_campo, tamanho_max in campos_validar:
            if len(valor) > tamanho_max:
                messagebox.showerror("Erro", f"{nome_campo} deve ter no máximo {tamanho_max} caracteres.")
                return None

        return {
            "nome": nome, "data": data, "local": local,
            "data_prorrogacao": data_prorrogacao,
            "corpo": self.widgets['text_corpo'].get("1.0", "end-1c").strip(),
            "hora": hora, "minuto": minuto, "categoria": categoria,
            "prioridade": prioridade, "grupo": grupo, "etiquetas": etiquetas
        }

    def adicionar_evento(self):
        """Adiciona novo evento"""
        dados = self.obter_dados_formulario()
        if dados is None:
            return

        conflitos = self.evento_manager.verificar_conflitos_agenda(dados)
        if conflitos:
            conflito_msg = "Conflito de agenda com os seguintes eventos:\n\n" + \
                           "\n".join([f"- {e['nome']} ({self.formata_data_br(self.evento_manager.get_effective_date(e))} às {e['hora']:02d}:{e['minuto']:02d})" for e in conflitos]) + \
                           "\n\nDeseja continuar mesmo assim?"
            if not messagebox.askyesno("Conflito de Agenda", conflito_msg):
                return

        # Adicionar anexos temporários
        dados["anexos"] = self.temp_anexos[:]

        self.evento_manager.adicionar_evento(dados)
        self.temp_anexos = []
        messagebox.showinfo("Sucesso", "Evento adicionado com sucesso!")
        self.evento_manager.salvar_eventos()
        self.limpar_form()
        self.atualizar_lista()

    def salvar_alteracoes(self):
        """Salva alterações em evento existente"""
        if self.editando_id is None:
            return

        dados = self.obter_dados_formulario()
        if dados is None:
            return

        conflitos = self.evento_manager.verificar_conflitos_agenda(dados, self.editando_id)
        if conflitos:
            conflito_msg = "Conflito de agenda com os seguintes eventos:\n\n" + \
                           "\n".join([f"- {e['nome']} ({self.formata_data_br(self.evento_manager.get_effective_date(e))} às {e['hora']:02d}:{e['minuto']:02d})" for e in conflitos]) + \
                           "\n\nDeseja continuar mesmo assim?"
            if not messagebox.askyesno("Conflito de Agenda", conflito_msg):
                return

        evento = self.evento_manager.atualizar_evento(self.editando_id, dados)
        if evento:
            evento["anexos"].extend(self.temp_anexos)
            messagebox.showinfo("Sucesso", "Evento atualizado com sucesso!")
            self.evento_manager.salvar_eventos()
            self.cancelar_edicao()
            self.atualizar_lista()

    def preparar_edicao(self, evento_id):
        """Prepara formulário para edição"""
        evento = self.evento_manager.buscar_evento_por_id(evento_id)
        if not evento:
            return

        self.editando_id = evento_id
        self.temp_anexos = []

        self.nome_var.set(evento["nome"])
        self.widgets['entry_data'].delete(0, "end")
        self.widgets['entry_data'].insert(0, self.formata_data_br(evento["data"]))
        self.prorrogacao_var.set(self.formata_data_br(evento.get("data_prorrogacao")))

        self.local_var.set(evento["local"])
        self.widgets['text_corpo'].delete("1.0", "end")
        self.widgets['text_corpo'].insert("1.0", evento["corpo"])
        self.hora_var.set(f"{evento['hora']:02d}")
        self.minuto_var.set(f"{evento['minuto']:02d}")
        self.categoria_var.set(evento["categoria"])
        self.prioridade_var.set(evento.get("prioridade", "BAIXA"))
        self.grupo_var.set(evento.get("grupo", ""))
        self.etiquetas_var.set(evento.get("etiquetas", ""))

        self.widgets['btn_salvar'].config(state='normal')
        self.widgets['btn_adicionar'].config(state='disabled')
        self.widgets['entry_nome'].focus()

    def limpar_form(self):
        """Limpa o formulário"""
        self.nome_var.set("")
        self.widgets['entry_data'].delete(0, "end")
        self.prorrogacao_var.set("")
        self.local_var.set("")
        self.widgets['text_corpo'].delete("1.0", "end")
        self.hora_var.set("00")
        self.minuto_var.set("00")
        self.categoria_var.set("")
        self.prioridade_var.set("")
        self.grupo_var.set("")
        self.etiquetas_var.set("")
        self.editando_id = None
        self.temp_anexos = []

    def cancelar_edicao(self):
        """Cancela a edição atual"""
        self.editando_id = None
        self.limpar_form()
        self.widgets['btn_salvar'].config(state='disabled')
        self.widgets['btn_adicionar'].config(state='normal')

    def anexar_arquivo(self):
        """Anexa arquivo ao evento"""
        if not os.path.exists(PASTA_ANEXOS):
            os.makedirs(PASTA_ANEXOS)

        arquivo = filedialog.askopenfilename(title="Selecionar arquivo", filetypes=[("Todos os arquivos", "*.*")])
        if arquivo:
            nome_arquivo = os.path.basename(arquivo)
            destino = os.path.join(PASTA_ANEXOS, f"{uuid.uuid4()}_{nome_arquivo}")
            try:
                shutil.copy2(arquivo, destino)
                messagebox.showinfo("Sucesso", f"Arquivo '{nome_arquivo}' anexado com sucesso!")
                if self.editando_id is not None:
                    evento = self.evento_manager.buscar_evento_por_id(self.editando_id)
                    if evento:
                        evento["anexos"].append(destino)
                else:
                    self.temp_anexos.append(destino)
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao anexar arquivo: {str(e)}")

    def selecionar_data(self, target_entry):
        """Abre seletor de data"""
        try:
            import tkcalendar
        except ImportError:
            messagebox.showerror("Erro", "Biblioteca tkcalendar não instalada. Instale com: pip install tkcalendar")
            return

        cores = CORES[self.config_manager.modo_aparencia]

        def grab_date():
            target_entry.delete(0, "end")
            target_entry.insert(0, cal.get_date())
            top.destroy()

        top = tk.Toplevel(self.root)
        top.title("Selecionar Data")
        top.transient(self.root)
        top.grab_set()
        top.configure(bg=cores["primary"])

        cal = tkcalendar.Calendar(top, selectmode="day", date_pattern="dd/mm/yyyy")
        cal.pack(pady=20, padx=20)

        btn_confirmar = tk.Button(top, text="Confirmar", command=grab_date,
                                 bg=cores["accent"], fg="white")
        btn_confirmar.pack(pady=10)

        top.update_idletasks()
        width, height = top.winfo_width(), top.winfo_height()
        x = (self.root.winfo_width() // 2) - (width // 2) + self.root.winfo_x()
        y = (self.root.winfo_height() // 2) - (height // 2) + self.root.winfo_y()
        top.geometry(f"{width}x{height}+{x}+{y}")

    # Métodos de interface de lista
    def atualizar_lista(self):
        """Atualiza a lista de eventos na interface"""
        lista_filtrada = self.evento_manager.filtrar_eventos(
            self.filtro_atual, self.filtro_busca, self.ordem_atual)
        total_eventos = len(lista_filtrada)

        inicio = (self.PAGINA_ATUAL - 1) * EVENTOS_POR_PAGINA
        fim = inicio + EVENTOS_POR_PAGINA
        lista_a_exibir = lista_filtrada[inicio:fim]

        # Remover cards que não devem mais ser exibidos
        ids_a_exibir = {e["id"] for e in lista_a_exibir}
        ids_a_remover = set(self.cards_por_id.keys()) - ids_a_exibir

        for evento_id in ids_a_remover:
            if evento_id in self.cards_por_id:
                card = self.cards_por_id.pop(evento_id)
                card.destroy()

        # Criar/atualizar cards
        frame_lista = self.widgets['frame_lista']

        for i, evento in enumerate(lista_a_exibir):
            row, col = divmod(i, 4)
            evento_id = evento["id"]

            if evento_id in self.cards_por_id:
                # Atualizar card existente
                card = self.cards_por_id[evento_id]
                card.evento = evento
                card.filtro_busca = self.filtro_busca
                card.criar_card()
                card.grid(row=row, column=col, padx=2, pady=5, sticky="nsew")
            else:
                # Criar novo card
                card = EventCard(frame_lista, evento, self.config_manager, self.evento_manager,
                               self.mostrar_detalhes_evento, self.mostrar_menu_opcoes_popup,
                               self.filtro_busca)
                card.grid(row=row, column=col, padx=2, pady=5, sticky="nsew")
                self.cards_por_id[evento_id] = card

        # Remover labels vazios antigos
        for widget in frame_lista.winfo_children():
            if isinstance(widget, tk.Label) and widget.cget("text").startswith(("📅\n\nNenhum evento", "Mostrando os primeiros")):
                widget.destroy()

        cores = CORES[self.config_manager.modo_aparencia]
        if not lista_a_exibir:
            empty_label = tk.Label(frame_lista, text="📅\n\nNenhum evento para exibir",
                                  font=("Arial", 16), bg=cores["primary"], fg=cores["text"],
                                  justify="center")
            empty_label.place(relx=0.6, rely=0.4, anchor="center")

        total_paginas = (total_eventos + EVENTOS_POR_PAGINA - 1) // EVENTOS_POR_PAGINA if EVENTOS_POR_PAGINA > 0 else 0
        self.contador_var.set(f"Eventos: {len(lista_a_exibir)}/{total_eventos} (Página {self.PAGINA_ATUAL}/{total_paginas})")

        self.adicionar_controles_paginacao(total_eventos)

    def adicionar_controles_paginacao(self, total_eventos):
        """Adiciona controles de paginação"""
        right_panel = self.widgets['right_panel']

        # Remover controles existentes
        for widget in right_panel.winfo_children():
            if hasattr(widget, '_is_pagination_control'):
                widget.destroy()

        total_paginas = (total_eventos + EVENTOS_POR_PAGINA - 1) // EVENTOS_POR_PAGINA

        if total_paginas > 1:
            cores = CORES[self.config_manager.modo_aparencia]
            pagination_frame = tk.Frame(right_panel, bg=cores["primary"])
            pagination_frame.pack(side="bottom", pady=5)
            pagination_frame._is_pagination_control = True

            btn_anterior = tk.Button(
                pagination_frame,
                text="◀ Anterior",
                command=lambda: self.mudar_pagina(-1, total_paginas),
                bg=cores["button_bg"],
                fg=cores["button_fg"]
            )
            btn_anterior.pack(side="left", padx=5)

            lbl_pagina = tk.Label(
                pagination_frame,
                text=f"Página {self.PAGINA_ATUAL} de {total_paginas}",
                bg=cores["primary"],
                fg=cores["text"]
            )
            lbl_pagina.pack(side="left", padx=5)

            btn_proximo = tk.Button(
                pagination_frame,
                text="Próximo ▶",
                command=lambda: self.mudar_pagina(1, total_paginas),
                bg=cores["button_bg"],
                fg=cores["button_fg"]
            )
            btn_proximo.pack(side="left", padx=5)

    def mudar_pagina(self, direcao, total_paginas):
        """Muda de página na paginação"""
        nova_pagina = self.PAGINA_ATUAL + direcao

        if 1 <= nova_pagina <= total_paginas:
            self.PAGINA_ATUAL = nova_pagina
            self.root.after_idle(self.atualizar_lista)

    # Métodos de ações nos eventos
    def mostrar_menu_opcoes_popup(self, evento_id):
        """Mostra menu de opções do evento"""
        menu = tk.Menu(self.root, tearoff=0)
        menu.add_command(label="📝 Editar", command=lambda: self.preparar_edicao(evento_id))
        menu.add_command(label="📋 Copiar", command=lambda: self.copiar_evento(evento_id))
        menu.add_command(label="📅 Exportar ICS", command=lambda: self.export_manager.exportar_para_ics(evento_id))
        menu.add_command(label="🗑️ Excluir", command=lambda: self.excluir_evento(evento_id))
        menu.add_command(label="✅ Concluído", command=lambda: self.marcar_concluido(evento_id))
        menu.add_command(label="👀 Ver detalhes", command=lambda: self.mostrar_detalhes_evento(evento_id))
        menu.tk_popup(self.root.winfo_pointerx(), self.root.winfo_pointery())

    def copiar_evento(self, evento_id):
        """Copia um evento existente"""
        evento_original = self.evento_manager.buscar_evento_por_id(evento_id)
        if not evento_original:
            messagebox.showerror("Erro", "Evento não encontrado!")
            return

        novo_evento = copy.deepcopy(evento_original)
        novo_evento["id"] = str(uuid.uuid4())
        novo_evento["alertas_enviados"] = {}
        novo_evento["concluido"] = False
        novo_evento["nome"] = f"{novo_evento['nome']} (CÓPIA)"

        self.evento_manager.eventos.append(novo_evento)
        self.evento_manager.salvar_eventos()
        self.atualizar_lista()
        messagebox.showinfo("Sucesso", "Evento copiado com sucesso!")

    def excluir_evento(self, evento_id):
        """Exclui um evento"""
        if not messagebox.askyesno("Confirmação", "Deseja excluir este evento?"):
            return

        self.evento_manager.remover_evento(evento_id)
        if self.editando_id == evento_id:
            self.cancelar_edicao()

        self.evento_manager.salvar_eventos()
        self.atualizar_lista()
        messagebox.showinfo("Sucesso", "Evento excluído com sucesso!")

    def excluir_todos_eventos_confirmacao(self):
        """Pede confirmação e exclui todos os eventos."""
        resposta = messagebox.askyesno("Confirmar Exclusão Total",
                                     "ATENÇíO: Isso excluirá TODOS os eventos permanentemente.\n\n"
                                     "Deseja continuar?",
                                     icon='warning')
        if resposta:
            self.evento_manager.remover_todos_eventos()
            self.evento_manager.salvar_eventos()
            self.atualizar_lista()
            messagebox.showinfo("Sucesso", "Todos os eventos foram excluídos.")

    def marcar_concluido(self, evento_id):
        """Marca/desmarca evento como concluído"""
        evento = self.evento_manager.buscar_evento_por_id(evento_id)
        if evento:
            evento["concluido"] = not evento["concluido"]
            self.evento_manager.salvar_eventos()
            self.atualizar_lista()

    def mostrar_detalhes_evento(self, evento_id):
        """Mostra janela com detalhes do evento"""
        evento = self.evento_manager.buscar_evento_por_id(evento_id)
        if not evento:
            return

        cores = CORES[self.config_manager.modo_aparencia]
        top = tk.Toplevel(self.root)
        top.title("Detalhes do Evento")
        top.geometry("450x620")
        top.minsize(450, 620)
        top.configure(bg=cores["primary"])
        top.transient(self.root)
        top.grab_set()

        top.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() - top.winfo_width()) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - top.winfo_height()) // 2
        top.geometry(f"+{x}+{y}")

        main_frame = tk.Frame(top, bg=cores["primary"])
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)

        # Botões na parte inferior
        button_frame = tk.Frame(main_frame, bg=cores["primary"])
        button_frame.pack(side="bottom", fill="x", pady=(10, 0))

        content_frame = tk.Frame(main_frame, bg=cores["primary"])
        content_frame.pack(side="top", fill="both", expand=True)

        btn_exportar = tk.Button(button_frame, text="📅 Exportar ICS",
                               command=lambda: self.export_manager.exportar_para_ics(evento_id),
                               bg=cores["accent"], fg="white")

        btn_fechar = tk.Button(button_frame, text="Fechar", command=top.destroy,
                              bg=cores["accent"], fg="white")

        tk.Label(button_frame, text="", bg=cores["primary"]).pack(side="left", expand=True)
        btn_exportar.pack(side="left", padx=5)
        btn_fechar.pack(side="left", padx=5)
        tk.Label(button_frame, text="", bg=cores["primary"]).pack(side="right", expand=True)

        # Conteúdo do evento
        tk.Label(content_frame, text=evento['nome'], font=("Arial", 18, "bold"),
                bg=cores["primary"], fg=cores["text"], wraplength=400).pack(pady=(0, 10))

        info_frame = tk.Frame(content_frame, bg=cores["secondary"], relief="solid", bd=1)
        info_frame.pack(fill="x", pady=5)
        
        informacoes = [
            ("Data Original:", self.formata_data_br(evento["data"])),
            ("Data Prorrogação:", self.formata_data_br(evento.get("data_prorrogacao"))),
            ("Hora:", f"{evento['hora']:02d}:{evento['minuto']:02d}"),
            ("Local:", evento["local"]),
            ("Categoria:", evento["categoria"]),
            ("Prioridade:", evento["prioridade"]),
            ("Grupo:", evento["grupo"]),
            ("Etiquetas:", evento["etiquetas"]),
            ("Criado em:", self.formata_data_br(evento["data_criacao"]))
        ]

        for key, value in informacoes:
            if value:
                tk.Label(info_frame, text=f"{key} {value}", font=("Arial", 10),
                        bg=cores["secondary"], fg=cores["text"], anchor="w",
                        wraplength=380).pack(anchor="w", padx=10, pady=2)

        tk.Label(content_frame, text="Detalhes:", font=("Arial", 12, "bold"),
                bg=cores["primary"], fg=cores["text"]).pack(anchor="w")

        text_widget = tk.Text(content_frame, height=6, width=50, wrap="word",
                             bg=cores["entry_bg"], fg=cores["entry_fg"])
        text_widget.insert("1.0", evento.get("corpo", ""))
        text_widget.pack(fill="x", padx=10, pady=5)
        text_widget.config(state="disabled")

        # Anexos
        if evento.get("anexos"):
            tk.Label(content_frame, text="Anexos:", font=("Arial", 12, "bold"),
                    bg=cores["primary"], fg=cores["text"]).pack(anchor="w", pady=(5, 0))

            anexos_container = tk.Frame(content_frame, height=120, bg=cores["primary"])
            anexos_container.pack(fill="x", expand=False, padx=10, pady=(5, 0))
            anexos_container.pack_propagate(False)

            canvas_anexos = tk.Canvas(anexos_container, bg=cores["primary"], highlightthickness=0)
            scrollbar_anexos = ttk.Scrollbar(anexos_container, orient="vertical", command=canvas_anexos.yview)
            scrollable_frame = tk.Frame(canvas_anexos, bg=cores["primary"])

            scrollable_frame.bind("<Configure>",
                                lambda ev: canvas_anexos.configure(scrollregion=canvas_anexos.bbox("all")))
            canvas_anexos.create_window((0, 0), window=scrollable_frame, anchor="nw")
            canvas_anexos.configure(yscrollcommand=scrollbar_anexos.set)

            def _on_mousewheel(event):
                canvas_anexos.yview_scroll(int(-1 * (event.delta / 120)), "units")

            canvas_anexos.bind("<Enter>", lambda ev: canvas_anexos.bind_all("<MouseWheel>", _on_mousewheel))
            canvas_anexos.bind("<Leave>", lambda ev: canvas_anexos.unbind_all("<MouseWheel>"))

            canvas_anexos.pack(side="left", fill="both", expand=True)
            scrollbar_anexos.pack(side="right", fill="y")

            for anexo in evento["anexos"]:
                anexo_frame = tk.Frame(scrollable_frame, bg=cores["primary"])
                anexo_frame.pack(anchor="w", fill='x', expand=True, pady=2)

                nome_base = os.path.basename(anexo)
                nome_exibido = nome_base[:40] + '...' if len(nome_base) > 40 else nome_base

                btn_abrir = tk.Button(anexo_frame, text=nome_exibido,
                                     command=lambda path=anexo: self.abrir_arquivo(path),
                                     bg=cores["button_bg"], fg=cores["button_fg"],
                                     relief="flat", anchor="w")
                btn_abrir.pack(side="left", fill='x', expand=True)

                btn_excluir = tk.Button(anexo_frame, text="🗑",
                                       command=lambda path=anexo, ev=evento, top_window=top: self.excluir_anexo(path, ev, top_window),
                                       bg=cores["status_danger_bg"], fg="white",
                                       relief="flat", font=("Arial", 10))
                btn_excluir.pack(side="left", padx=(5, 0))

    def excluir_anexo(self, caminho_arquivo, evento, janela_detalhes):
        """Exclui um anexo do evento"""
        if messagebox.askyesno("Confirmar", "Deseja realmente excluir este anexo?"):
            try:
                if caminho_arquivo in evento["anexos"]:
                    evento["anexos"].remove(caminho_arquivo)
                    if os.path.exists(caminho_arquivo):
                        os.remove(caminho_arquivo)
                    self.evento_manager.salvar_eventos()
                    messagebox.showinfo("Sucesso", "Anexo excluído com sucesso!")
                    janela_detalhes.destroy()
                    self.mostrar_detalhes_evento(evento["id"])
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao excluir anexo: {str(e)}")

    def abrir_arquivo(self, caminho_arquivo):
        """Abre um arquivo anexado"""
        try:
            if os.path.exists(caminho_arquivo):
                if sys.platform == "win32":
                    os.startfile(caminho_arquivo)
                elif sys.platform == "darwin":
                    subprocess.run(["open", caminho_arquivo], check=True)
                else:
                    subprocess.run(["xdg-open", caminho_arquivo], check=True)
            else:
                messagebox.showerror("Erro", "Arquivo não encontrado!")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao abrir arquivo: {str(e)}")

    def importar_de_ics(self):
        """Importa eventos de um arquivo .ics"""
        arquivo = filedialog.askopenfilename(
            title="Selecionar arquivo ICS para importar",
            filetypes=[("Arquivos ICS", "*.ics"), ("Todos os arquivos", "*.*")]
        )
        if not arquivo:
            return

        try:
            with open(arquivo, 'r', encoding='utf-8') as f:
                linhas = f.readlines()

            eventos_importados = 0
            evento_atual = None
            for linha in linhas:
                linha = linha.strip()
                if linha == "BEGIN:VEVENT":
                    evento_atual = {
                        "nome": "Evento Sem Nome",
                        "data": datetime.date.today(),
                        "hora": 0, "minuto": 0,
                        "local": "", "corpo": "", "categoria": "",
                        "prioridade": "", "grupo": "", "etiquetas": "",
                        "anexos": [], "data_prorrogacao": None
                    }
                elif linha == "END:VEVENT" and evento_atual is not None:
                    # Adicionar o evento parseado usando o manager
                    self.evento_manager.adicionar_evento(evento_atual)
                    eventos_importados += 1
                    evento_atual = None
                elif evento_atual is not None and ':' in linha:
                    chave, valor = linha.split(':', 1)
                    if chave.startswith("SUMMARY"):
                        evento_atual["nome"] = valor
                    elif chave.startswith("DTSTART"):
                        dt_str = valor
                        if 'T' in dt_str:
                            date_part, time_part = dt_str.split('T')
                            evento_atual["data"] = datetime.datetime.strptime(date_part, "%Y%m%d").date()
                            evento_atual["hora"] = int(time_part[0:2])
                            evento_atual["minuto"] = int(time_part[2:4])
                        else:
                            evento_atual["data"] = datetime.datetime.strptime(dt_str, "%Y%m%d").date()
                    elif chave.startswith("LOCATION"):
                        evento_atual["local"] = valor
                    elif chave.startswith("DESCRIPTION"):
                        evento_atual["corpo"] = valor.replace('\\n', '\n')

            if eventos_importados > 0:
                self.evento_manager.salvar_eventos()
                self.atualizar_lista()
                messagebox.showinfo("Sucesso", f"{eventos_importados} evento(s) importado(s) com sucesso!")
            else:
                messagebox.showwarning("Aviso", "Nenhum evento válido foi encontrado no arquivo ICS.")

        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro ao importar o arquivo ICS: {str(e)}")


    def on_closing(self):
        """Trata o fechamento da aplicação"""
        self.alerta_manager.parar_thread_alertas()

        # Verificar se há dados não salvos no formulário
        is_dirty = any([
            self.nome_var.get(),
            self.widgets['entry_data'].get(),
            self.prorrogacao_var.get(),
            self.local_var.get(),
            self.widgets['text_corpo'].get("1.0", "end-1c").strip(),
            self.categoria_var.get(),
            self.grupo_var.get(),
            self.etiquetas_var.get()
        ])

        if is_dirty and self.editando_id is None:
            if messagebox.askyesno("Sair", "Você tem dados não salvos no formulário. Deseja realmente sair?"):
                self.root.destroy()
        else:
            self.root.destroy()

    def executar(self):
        """Executa o loop principal da aplicação gráfica"""
        try:
            self.criar_interface()
            self.aplicar_tema()
            self.atualizar_lista()
            self.alerta_manager.iniciar_thread_alertas(self.evento_manager.salvar_eventos)
            
            self.root.mainloop()
        except Exception as e:
            logger.error(f"Erro fatal durante a execução da GUI: {e}", exc_info=True)
            messagebox.showerror("Erro Fatal", f"Ocorreu um erro inesperado e a aplicação será encerrada:\n\n{e}")
            self.root.destroy()

# --- INÍCIO DA INTEGRAÇÃO (MODO CLI) ---

def adicionar_evento_via_cli(dados_evento_json_str, anexo_path_str):
    """
    Função dedicada para adicionar um evento via linha de comando (CLI), sem iniciar a GUI.
    É chamada pelo extraçao.py.
    """
    # Configura um logger básico para o modo CLI, caso o logger principal não tenha sido configurado
    cli_logger = logging.getLogger('calendario_cli')
    if not cli_logger.hasHandlers():
        handler = logging.StreamHandler(sys.stderr) # Envia logs para o erro padrão
        formatter = logging.Formatter('[CLI] %(levelname)s: %(message)s')
        handler.setFormatter(formatter)
        cli_logger.addHandler(handler)
        cli_logger.setLevel(logging.INFO)

    cli_logger.info("Iniciando processo de adição de evento via linha de comando...")

    try:
        # 1. Inicializa os gerenciadores de dados
        config_manager = ConfigManager()
        crypt_manager = CryptManager()
        
        # 2. Lida com a criptografia
        if not crypt_manager.carregar_ou_criar_chave():
            if os.path.exists(ARQUIVO_CHAVE):
                cli_logger.error("Falha ao carregar a chave de criptografia. O arquivo .key pode estar corrompido.")
                return False

        if crypt_manager.criptografia_habilitada:
            senha = None
            try:
                # Usa getpass para pedir a senha de forma segura no terminal (se executado interativamente)
                senha = getpass.getpass("Digite a senha do calendário para registrar o evento do boleto: ")
            except Exception as e:
                cli_logger.error(f"Não foi possível solicitar a senha no terminal. Configure uma variável de ambiente ou execute interativamente. Erro: {e}")
                return False

            if not senha or not crypt_manager.gerar_chave_de_senha(senha) == crypt_manager.chave_criptografia:
                cli_logger.error("Senha do calendário incorreta ou não fornecida. Evento não foi adicionado.")
                return False
            cli_logger.info("Senha do calendário verificada com sucesso.")

        # 3. Prepara o gerenciador de eventos
        evento_manager = EventoManager(crypt_manager)
        evento_manager.carregar_eventos()

        # 4. Processa os dados do evento recebidos do extraçao.py
        dados_evento = json.loads(dados_evento_json_str)

        # Converte a data de string (DD/MM/AAAA) para um objeto 'date'
        data_str = dados_evento.get("data")
        data_obj = None
        if data_str:
            try:
                data_obj = datetime.datetime.strptime(data_str, "%d/%m/%Y").date()
            except (ValueError, TypeError):
                cli_logger.error(f"Formato de data inválido ('{data_str}'). Use DD/MM/AAAA. Evento não adicionado.")
                return False
        else:
            cli_logger.error("Campo 'data' (vencimento) não encontrado nos dados do evento.")
            return False
        dados_evento["data"] = data_obj

        # Mapeia 'detalhes' do extraçao.py para o campo 'corpo' do calendário
        dados_evento['corpo'] = dados_evento.pop('detalhes', dados_evento.get('corpo', ''))
        
        # Garante valores padrão para campos opcionais
        dados_evento.setdefault('prioridade', 'MÉDIA')
        dados_evento.setdefault('data_prorrogacao', None)
        # (outros campos como local, categoria, etc., já são tratados no EventoManager)

        # 5. Processa o anexo
        anexos_finais = []
        if anexo_path_str and os.path.exists(anexo_path_str):
            nome_base_anexo = os.path.basename(anexo_path_str)
            destino_anexo = os.path.join(PASTA_ANEXOS, f"{uuid.uuid4()}_{nome_base_anexo}")
            try:
                shutil.copy2(anexo_path_str, destino_anexo)
                anexos_finais.append(destino_anexo)
                cli_logger.info(f"Anexo '{nome_base_anexo}' copiado com sucesso para a pasta de anexos.")
            except Exception as e_copy:
                cli_logger.error(f"Falha ao copiar anexo '{anexo_path_str}': {e_copy}")
        elif anexo_path_str:
            cli_logger.warning(f"Arquivo de anexo não encontrado: {anexo_path_str}")
        
        dados_evento["anexos"] = anexos_finais

        # 6. Adiciona e salva o evento
        evento_adicionado = evento_manager.adicionar_evento(dados_evento)
        if evento_adicionado:
            #adicionar_evento_google_calendar(evento_adicionado) # Chama a sincronização
            evento_manager.salvar_eventos()
            cli_logger.info(f"Evento '{evento_adicionado['nome']}' foi adicionado com sucesso ao calendário.")
            return True
        else:
            cli_logger.error("Falha ao adicionar o evento (método evento_manager.adicionar_evento falhou).")
            return False

    except json.JSONDecodeError as e_json:
        cli_logger.error(f"Erro ao processar os dados do evento (JSON inválido): {e_json}")
        return False
    except Exception as e:
        cli_logger.error(f"Erro fatal ao adicionar evento via CLI: {e}")
        cli_logger.error(traceback.format_exc()) # Loga o traceback completo para depuração
        return False
    

def obter_credenciais_google():
    creds = None
    if os.path.exists(TOKEN_FILE):
        try:
            creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
        except Exception as e_token:
            logger.error(f"Erro ao carregar token.json: {e_token}. Iniciando fluxo de autenticação.")
            creds = None # Força novo login se token estiver inválido

    # Se não houver credenciais válidas...
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
             logger.info("Token expirado, tentando refresh.") # Log útil mantido
             try:
                creds.refresh(Request())
                logger.info("Token atualizado com sucesso via refresh.") # Log útil mantido
             except Exception as e_refresh:
                logger.error(f"Erro ao atualizar token via refresh: {e_refresh}. Iniciando fluxo de autenticação.")
                creds = None # Força novo login se refresh falhar
        else:
             logger.info("Nenhum token válido encontrado. Iniciando novo fluxo de autenticação.") # Log útil mantido
             if not os.path.exists(CREDENTIALS_FILE):
                 logger.error(f"Arquivo de credenciais '{CREDENTIALS_FILE}' não encontrado!")
                 try:
                     messagebox.showerror("Erro de Configuração", f"Arquivo '{CREDENTIALS_FILE}' não encontrado.")
                 except: pass
                 return None

             try:
                 flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
                 creds = flow.run_local_server(port=0) # Esta linha abre o navegador
             except Exception as e_flow:
                 logger.error(f"!!! ERRO DURANTE o fluxo de autenticação (run_local_server): {e_flow}", exc_info=True)
                 try:
                     messagebox.showerror("Erro de Autenticação", f"Falha no fluxo: {e_flow}")
                 except: pass
                 return None

        # Salva as credenciais para a próxima execução
        if creds:
             try:
                with open(TOKEN_FILE, 'w') as token:
                    token.write(creds.to_json())
                logger.info(f"Token salvo em {TOKEN_FILE}") # Log útil mantido
             except Exception as e_save:
                 logger.error(f"Erro ao salvar o token em {TOKEN_FILE}: {e_save}")
        # else: # Log de warning removido para limpar
            # logger.warning("Não foi possível obter credenciais após o fluxo.")

    # Logs de verificação final removidos para limpar
    return creds

def adicionar_evento_google_calendar(evento_local):
    logger.info(f"Tentando adicionar evento '{evento_local.get('nome')}' ao Google Calendar.")
    google_event_id = None # Inicializa o ID como None
    try:
        creds = obter_credenciais_google()
    except Exception as e_creds:
        logger.error(f"!!! ERRO AO CHAMAR obter_credenciais_google(): {e_creds}", exc_info=True)
        creds = None

    if not creds:
        logger.error("Não foi possível obter credenciais do Google. Evento não sincronizado.")
        return False, None # Retorna Falha e ID None

    try:
        service = build('calendar', 'v3', credentials=creds)
        # ... (Lógica para determinar data_evento, hora, minuto) ...
        data_evento_raw = evento_local.get("data_prorrogacao") or evento_local.get("data")
        # ...(Restante da lógica de data/hora/fuso horário)...
        if isinstance(data_evento_raw, str):
             try: data_evento = datetime.datetime.strptime(data_evento_raw, "%Y-%m-%d").date()
             except ValueError:
                  try: data_evento = datetime.datetime.strptime(data_evento_raw, "%d/%m/%Y").date()
                  except ValueError:
                      logger.error(f"Formato de data inválido para Google Calendar: {data_evento_raw}")
                      return False, None
        elif isinstance(data_evento_raw, datetime.date): data_evento = data_evento_raw
        else:
            logger.error(f"Tipo de data inesperado para Google Calendar: {type(data_evento_raw)}")
            return False, None
        hora = evento_local.get("hora", 0)
        minuto = evento_local.get("minuto", 0)
        try:
             local_tz_name_tuple = time.tzname
             local_tz_name = local_tz_name_tuple[1] if time.daylight and len(local_tz_name_tuple) > 1 and local_tz_name_tuple[1] != local_tz_name_tuple[0] else local_tz_name_tuple[0]
             if local_tz_name in ['BRT', 'BRST', 'Horário de Brasília']: local_tz_pytz = 'America/Sao_Paulo'
             else:
                 try:
                    pytz.timezone(local_tz_name)
                    local_tz_pytz = local_tz_name
                 except pytz.exceptions.UnknownTimeZoneError:
                    logger.warning(f"Nome de fuso horário local '{local_tz_name}' não reconhecido por pytz, usando 'America/Sao_Paulo'.")
                    local_tz_pytz = 'America/Sao_Paulo'
             local_tz = pytz.timezone(local_tz_pytz)
        except Exception as e_tz:
             logger.warning(f"Não foi possível detectar timezone local, usando 'America/Sao_Paulo'. Erro: {e_tz}", exc_info=True)
             local_tz = pytz.timezone('America/Sao_Paulo')
        dt_inicio_naive = datetime.datetime.combine(data_evento, datetime.time(hora, minuto))
        dt_inicio_aware = local_tz.localize(dt_inicio_naive)
        dt_fim_aware = dt_inicio_aware + datetime.timedelta(hours=1)
        start_rfc = dt_inicio_aware.isoformat()
        end_rfc = dt_fim_aware.isoformat()
        # ...(Fim da lógica de data/hora/fuso horário)...

        # >>> ADIÇÃO DO PASSO 2 (COR) <<<
        google_color_id = _determinar_google_color_id(evento_local)
        # >>> FIM DA ADIÇÃO DO PASSO 2 <<<

        evento_google = {
          'summary': evento_local.get('nome', 'Evento Sem Nome'),
          'location': evento_local.get('local', ''),
          'description': ( f"Detalhes: {evento_local.get('corpo', '')}\n"
                           f"Categoria: {evento_local.get('categoria', '')}\n"
                           f"Grupo: {evento_local.get('grupo', '')}\n"
                           f"Etiquetas: {evento_local.get('etiquetas', '')}\n"
                           f"Prioridade: {evento_local.get('prioridade', '')}\n"
                           f"Valor: R$ {evento_local.get('valor', 'N/A')}\n"
                           f"Valor Total c/ Atraso: R$ {evento_local.get('valor_total_com_atraso') or evento_local.get('VALOR_TOTAL_COM_ATRASO', 'N/A')}\n"
                           f"\n---\nEvento criado por DocQuantum" ),
          'start': {'dateTime': start_rfc}, 'end': {'dateTime': end_rfc},
          # 'colorId': '9', # << REMOVA ou COMENTE a linha antiga que fixava a cor 9
          # >>> MODIFICAÇÃO DO PASSO 2 (COR) <<<
          'colorId': google_color_id, # Usa a cor determinada dinamicamente
          # >>> FIM DA MODIFICAÇÃO DO PASSO 2 <<<
          'reminders': { 'useDefault': False, 'overrides': [{'method': 'popup', 'minutes': 10}], },
        }
        # Log opcional para ver a cor usada
        logger.info(f"--> Usando Google Color ID: {google_color_id}")

        event = service.events().insert(calendarId='primary', body=evento_google).execute()
        google_event_id = event.get('id') # Captura o ID retornado pelo Google (Passo 1)
        logger.info(f"Evento criado no Google Calendar com sucesso! ID: {google_event_id}, Link: {event.get('htmlLink')}")
        return True, google_event_id # Retorna Sucesso e o ID (Passo 1)

    except HttpError as error:
        logger.error(f'!!! Ocorreu um erro na API do Google Calendar (insert): Código {error.resp.status}, Motivo: {error.content.decode("utf-8")}')
        return False, None # Retorna Falha e ID None (Passo 1)
    except Exception as e:
         logger.error(f'!!! Erro inesperado ao criar evento no Google Calendar: {e}', exc_info=True)
         return False, None # Retorna Falha e ID None (Passo 1)

def atualizar_evento_google_calendar(google_event_id, evento_local_atualizado):
    logger.info(f"Tentando atualizar evento '{evento_local_atualizado.get('nome')}' (ID Google: {google_event_id}) no Google Calendar.")
    try:
        creds = obter_credenciais_google()
    except Exception as e_creds:
        logger.error(f"!!! ERRO AO CHAMAR obter_credenciais_google() para update: {e_creds}", exc_info=True)
        creds = None

    if not creds:
        logger.error("Não foi possível obter credenciais do Google. Evento não atualizado.")
        return False

    try:
        service = build('calendar', 'v3', credentials=creds)

        # Reutiliza a mesma lógica de formatação de data/hora/fuso da criação
        data_evento_raw = evento_local_atualizado.get("data_prorrogacao") or evento_local_atualizado.get("data")
        # ...(Lógica completa de data/hora/fuso horário)...
        if isinstance(data_evento_raw, str):
             try: data_evento = datetime.datetime.strptime(data_evento_raw, "%Y-%m-%d").date()
             except ValueError:
                  try: data_evento = datetime.datetime.strptime(data_evento_raw, "%d/%m/%Y").date()
                  except ValueError:
                      logger.error(f"Formato de data inválido para Google Calendar (update): {data_evento_raw}")
                      return False
        elif isinstance(data_evento_raw, datetime.date): data_evento = data_evento_raw
        else:
            logger.error(f"Tipo de data inesperado para Google Calendar (update): {type(data_evento_raw)}")
            return False
        hora = evento_local_atualizado.get("hora", 0)
        minuto = evento_local_atualizado.get("minuto", 0)
        try:
             local_tz_name_tuple = time.tzname
             local_tz_name = local_tz_name_tuple[1] if time.daylight and len(local_tz_name_tuple) > 1 and local_tz_name_tuple[1] != local_tz_name_tuple[0] else local_tz_name_tuple[0]
             if local_tz_name in ['BRT', 'BRST', 'Horário de Brasília']: local_tz_pytz = 'America/Sao_Paulo'
             else:
                 try:
                    pytz.timezone(local_tz_name)
                    local_tz_pytz = local_tz_name
                 except pytz.exceptions.UnknownTimeZoneError:
                    local_tz_pytz = 'America/Sao_Paulo' # Mantém o fallback
             local_tz = pytz.timezone(local_tz_pytz)
        except Exception as e_tz: # Captura exceção genérica aqui também
             logger.warning(f"Não foi possível detectar timezone local (update), usando 'America/Sao_Paulo'. Erro: {e_tz}", exc_info=True)
             local_tz = pytz.timezone('America/Sao_Paulo') # Mantém o fallback
        dt_inicio_naive = datetime.datetime.combine(data_evento, datetime.time(hora, minuto))
        dt_inicio_aware = local_tz.localize(dt_inicio_naive)
        dt_fim_aware = dt_inicio_aware + datetime.timedelta(hours=1)
        start_rfc = dt_inicio_aware.isoformat()
        end_rfc = dt_fim_aware.isoformat()
        # ...(Fim da lógica de data/hora/fuso horário)...

        # >>> ADIÇÃO DO PASSO 3 (COR) <<<
        google_color_id = _determinar_google_color_id(evento_local_atualizado)
        # >>> FIM DA ADIÇÃO DO PASSO 3 <<<

        # Monta o corpo do evento ATUALIZADO
        evento_google_atualizado = {
          'summary': evento_local_atualizado.get('nome', 'Evento Sem Nome'),
          'location': evento_local_atualizado.get('local', ''),
          'description': ( f"Detalhes: {evento_local_atualizado.get('corpo', '')}\n"
                           f"Categoria: {evento_local_atualizado.get('categoria', '')}\n"
                           f"Grupo: {evento_local_atualizado.get('grupo', '')}\n"
                           f"Etiquetas: {evento_local_atualizado.get('etiquetas', '')}\n"
                           f"Prioridade: {evento_local_atualizado.get('prioridade', '')}\n"
                           f"Valor: R$ {evento_local_atualizado.get('valor', 'N/A')}\n"
                           f"Valor Total c/ Atraso: R$ {evento_local_atualizado.get('valor_total_com_atraso') or evento_local_atualizado.get('VALOR_TOTAL_COM_ATRASO', 'N/A')}\n"
                           f"\n---\nEvento criado por DocQuantum" ),
          'start': {'dateTime': start_rfc}, 'end': {'dateTime': end_rfc},
          # 'colorId': '9', # << REMOVA ou COMENTE a linha antiga que fixava a cor 9
          # >>> MODIFICAÇÃO DO PASSO 3 (COR) <<<
          'colorId': google_color_id, # Usa a cor determinada dinamicamente
          # >>> FIM DA MODIFICAÇÃO DO PASSO 3 <<<
          'reminders': { 'useDefault': False, 'overrides': [{'method': 'popup', 'minutes': 10}], },
        }
        # Log opcional para ver a cor usada na atualização
        logger.info(f"--> Atualizando com Google Color ID: {google_color_id}")

        # Chama a API de UPDATE
        updated_event = service.events().update(calendarId='primary', eventId=google_event_id, body=evento_google_atualizado).execute()
        logger.info(f"Evento atualizado no Google Calendar com sucesso! Link: {updated_event.get('htmlLink')}")
        return True

    except HttpError as error:
        logger.error(f'!!! Ocorreu um erro na API do Google Calendar (update): Código {error.resp.status}, Motivo: {error.content.decode("utf-8")}')
        # messagebox.showerror("Erro Google Agenda", f"Não foi possível atualizar o evento: {error.content.decode('utf-8')}")
        return False
    except Exception as e:
         logger.error(f'!!! Erro inesperado ao atualizar evento no Google Calendar: {e}', exc_info=True)
         # messagebox.showerror("Erro Google Agenda", f"Erro inesperado ao atualizar: {e}")
         return False

def excluir_evento_google_calendar(google_event_id):
    logger.info(f"Tentando excluir evento (ID Google: {google_event_id}) do Google Calendar.")
    try:
        creds = obter_credenciais_google()
    except Exception as e_creds:
        logger.error(f"!!! ERRO AO CHAMAR obter_credenciais_google() para delete: {e_creds}", exc_info=True)
        creds = None

    if not creds:
        logger.error("Não foi possível obter credenciais do Google. Evento não excluído.")
        return False

    try:
        service = build('calendar', 'v3', credentials=creds)

        # Chama a API de DELETE
        service.events().delete(calendarId='primary', eventId=google_event_id).execute()
        logger.info(f"Evento (ID Google: {google_event_id}) excluído do Google Calendar com sucesso!")
        return True

    except HttpError as error:
        # Verifica se o erro é "404 Not Found" ou "410 Gone" (evento já excluído)
        if error.resp.status in [404, 410]:
            logger.warning(f"Evento (ID Google: {google_event_id}) não encontrado no Google Calendar (provavelmente já excluído). Considerado sucesso.")
            return True # Considera sucesso se já não existe lá
        else:
            logger.error(f'!!! Ocorreu um erro na API do Google Calendar (delete): Código {error.resp.status}, Motivo: {error.content.decode("utf-8")}')
            # messagebox.showerror("Erro Google Agenda", f"Não foi possível excluir o evento: {error.content.decode('utf-8')}")
            return False
    except Exception as e:
         logger.error(f'!!! Erro inesperado ao excluir evento no Google Calendar: {e}', exc_info=True)
         # messagebox.showerror("Erro Google Agenda", f"Erro inesperado ao excluir: {e}")
         return False

def _determinar_google_color_id(evento_local):
    """Determina o Google Calendar colorId com base no status do evento DocQuantum."""
    # Mapeamento (Ajuste as cores/IDs conforme sua preferência visual)
    # IDs comuns: 1:Azul, 2:Verde, 3:Roxo, 4:Vermelho, 5:Amarelo, 6:Laranja, 7:Turquesa, 8:Cinza, 9:Azul Escuro, 10:Verde Escuro, 11:Vermelho Escuro
    mapa_status_para_google_id = {
        "completed": "10", # Verde Escuro
        "past":      "11", # Vermelho Escuro
        "today":     "6",  # Laranja
        "tomorrow":  "5",  # Amarelo
        "3_days":    "3",  # Roxo
        "7_days":    "3",  # Roxo (mesmo de 3 dias, ou escolha outro ID como 8-Cinza)
        "14_days":   "8",  # Cinza
        "21_days":   "8",  # Cinza
        "30_days":   "9",  # Azul Escuro
        "default":   "1"   # Azul Padrão (para eventos futuros > 30 dias)
    }

    # Lógica similar a obter_cores_status_evento para determinar o status
    hoje = datetime.date.today()
    try:
        # Tenta obter a data efetiva (pode ser string ou date)
        data_efetiva_raw = evento_local.get("data_prorrogacao") or evento_local.get("data")
        if isinstance(data_efetiva_raw, str):
            try: data_efetiva = datetime.datetime.strptime(data_efetiva_raw, "%Y-%m-%d").date()
            except ValueError: data_efetiva = datetime.datetime.strptime(data_efetiva_raw, "%d/%m/%Y").date()
        elif isinstance(data_efetiva_raw, datetime.date):
            data_efetiva = data_efetiva_raw
        else: # Se não for data válida, retorna o default
             return mapa_status_para_google_id["default"]
    except Exception: # Em caso de erro ao processar data
        return mapa_status_para_google_id["default"]

    dias_faltando = (data_efetiva - hoje).days

    status_key = "default"
    if evento_local.get("concluido"): status_key = "completed"
    elif dias_faltando < 0: status_key = "past"
    elif dias_faltando == 0: status_key = "today"
    elif dias_faltando == 1: status_key = "tomorrow"
    elif dias_faltando <= 3: status_key = "3_days"
    elif dias_faltando <= 7: status_key = "7_days"
    elif dias_faltando <= 14: status_key = "14_days"
    elif dias_faltando <= 21: status_key = "21_days"
    elif dias_faltando <= 30: status_key = "30_days"

    return mapa_status_para_google_id.get(status_key, "1") # Retorna o ID mapeado ou o default (Azul)

# Bloco principal que decide entre o modo CLI (linha de comando) e o modo GUI (gráfico)
if __name__ == "__main__":
    # Configura o parser para ler argumentos da linha de comando
    parser = argparse.ArgumentParser(description="Calendário DocQuantum - App Gráfico ou Adicionador de Eventos via CLI")
    parser.add_argument("--add-event-data", help="String JSON com os dados do evento para adicionar via CLI.")
    parser.add_argument("--attach-file", help="Caminho completo do arquivo para anexar ao evento via CLI.")

    args = parser.parse_args()

    # Se o argumento --add-event-data foi passado, executa em MODO CLI
    if args.add_event_data:
        if adicionar_evento_via_cli(args.add_event_data, args.attach_file):
            sys.exit(0)  # Sucesso
        else:
            sys.exit(1)  # Falha
            
    # Se nenhum argumento foi passado, executa em MODO GRÁFICO (GUI)
    else:
        # Este é o código que inicia a interface gráfica
        try:
            app = CalendarioApp()
            if app.inicializar():
                app.executar()
            else:
                print("Falha na inicialização da aplicação.")
                sys.exit(1)
        except ImportError as e:
            missing_lib = str(e).split("'")[1] if "'" in str(e) else str(e)
            print(f"\nERRO: Biblioteca necessária não encontrada: {missing_lib}")
            print(f"Para instalar, execute: pip install {missing_lib} cryptography plyer pillow reportlab matplotlib tkcalendar")
            try:
                # Tenta mostrar um messagebox se o tkinter estiver minimamente funcional
                root_temp = tk.Tk()
                root_temp.withdraw()
                messagebox.showerror("Erro de Dependência",
                                     f"A biblioteca '{missing_lib}' não foi encontrada.\n\n"
                                     f"Por favor, instale as dependências com o comando:\n"
                                     f"pip install cryptography plyer pillow reportlab matplotlib tkcalendar")
            except:
                pass # Se o tkinter falhar, a mensagem no console já foi exibida
            sys.exit(1)
        except Exception as e:
            print(f"Erro inesperado durante a inicialização da GUI: {e}")
            traceback.print_exc() # Imprime o traceback completo para facilitar a depuração
            sys.exit(1)


