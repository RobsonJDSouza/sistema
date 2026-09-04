import sqlite3
import os
import csv
import json
import re
import urllib.request
from datetime import datetime, date, timedelta
import tkinter as tk
from tkinter import ttk, filedialog
import hashlib
import shutil
import sys

try:
    import openpyxl
    HAS_OPENPYXL = True
except:
    HAS_OPENPYXL = False

DB_PATH = os.path.join(os.path.dirname(__file__), "sistema.db")

BACKUP_DIR = os.path.join(os.path.dirname(__file__), "backups")

def garantir_pasta_backup():
    try:
        os.makedirs(BACKUP_DIR, exist_ok=True)
    except Exception:
        pass
    return BACKUP_DIR

def gerar_nome_backup():
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    login = ""
    try:
        if usuario_logado and usuario_logado.get("login"):
            login = "_" + str(usuario_logado["login"]).replace(" ", "_")
    except Exception:
        pass
    return f"backup{login}_{ts}.db"

TABELAS_SISTEMA = (
    "usuarios", "clientes", "fornecedores", "produtos", "vendas",
    "caixa", "contas_a_pagar", "contas_a_receber",
)

# Hook preenchido em criar_interface() para atualizar a lista de backups da tela
_hook_atualizar_lista_backups = None

def _copiar_banco_sqlite(origem, destino):
    """Copia um banco SQLite de `origem` para `destino`.

    Usa a API de backup do próprio SQLite (cópia consistente, página a página,
    funciona mesmo com conexões abertas e substitui TODO o conteúdo do destino).
    Se a API não estiver disponível/falhar, usa cópia simples de arquivo.
    """
    try:
        src = sqlite3.connect(origem, timeout=10)
        try:
            if not hasattr(src, "backup"):
                raise RuntimeError("API de backup indisponível")
            dst = sqlite3.connect(destino, timeout=10)
            try:
                src.backup(dst)
            finally:
                dst.close()
        finally:
            src.close()
        return True
    except Exception as e:
        print("Aviso: API de backup do SQLite falhou, usando cópia de arquivo:", e)
    shutil.copy2(origem, destino)
    return True

def validar_arquivo_backup(caminho):
    """Verifica se o arquivo é um banco SQLite íntegro e do sistema.
    Retorna (ok, mensagem_de_erro)."""
    try:
        if not caminho or not os.path.isfile(caminho):
            return False, "Arquivo não encontrado."
        if os.path.getsize(caminho) < 100:
            return False, "Arquivo vazio ou inválido."
        with open(caminho, "rb") as f:
            cabecalho = f.read(16)
        if not cabecalho.startswith(b"SQLite format 3"):
            return False, "O arquivo selecionado não é um banco de dados SQLite válido (.db)."
        conn = sqlite3.connect(caminho, timeout=5)
        try:
            cur = conn.cursor()
            cur.execute("PRAGMA quick_check")
            res = cur.fetchone()
            if not res or str(res[0]).lower() != "ok":
                return False, f"Banco de dados do backup está corrompido: {res[0] if res else 'sem resposta'}"
            cur.execute("SELECT name FROM sqlite_master WHERE type='table'")
            tabelas = {r[0] for r in cur.fetchall()}
        finally:
            conn.close()
        if not any(t in tabelas for t in TABELAS_SISTEMA):
            return False, "O arquivo não contém as tabelas do sistema (não é um backup deste programa)."
        return True, ""
    except Exception as e:
        return False, str(e)

def fazer_backup(destino=None, silencioso=False):
    """Copia sistema.db para pasta backups/ (ou caminho informado).
    Retorna caminho do arquivo gerado ou None em erro."""
    garantir_pasta_backup()
    if not os.path.exists(DB_PATH):
        if not silencioso:
            mostrar_erro("Arquivo de banco de dados não encontrado.")
        return None
    try:
        if destino is None:
            destino = os.path.join(BACKUP_DIR, gerar_nome_backup())
            # Evita sobrescrever um backup gerado no mesmo segundo
            base, ext = os.path.splitext(destino)
            n = 1
            while os.path.exists(destino):
                destino = f"{base}_{n}{ext}"
                n += 1
        # Garante que conexões pendentes gravem (sqlite)
        try:
            conn = conectar()
            conn.execute("PRAGMA wal_checkpoint(FULL)")
            conn.close()
        except Exception:
            pass
        _copiar_banco_sqlite(DB_PATH, destino)
        # Atualiza a lista da tela de backup (se a interface estiver aberta)
        try:
            if callable(_hook_atualizar_lista_backups):
                _hook_atualizar_lista_backups()
        except Exception:
            pass
        if not silencioso:
            mostrar_sucesso(
                f"Backup realizado com sucesso!\n\nArquivo:\n{destino}",
                "Backup",
            )
        return destino
    except Exception as e:
        if not silencioso:
            mostrar_erro(f"Falha ao gerar backup:\n{e}")
        return None

def listar_backups_locais():
    garantir_pasta_backup()
    itens = []
    try:
        for nome in os.listdir(BACKUP_DIR):
            if not nome.lower().endswith(".db"):
                continue
            caminho = os.path.join(BACKUP_DIR, nome)
            if not os.path.isfile(caminho):
                continue
            try:
                st = os.stat(caminho)
                tamanho = st.st_size
                mtime_num = st.st_mtime
                mtime = datetime.fromtimestamp(st.st_mtime).strftime("%d/%m/%Y %H:%M:%S")
            except Exception:
                tamanho, mtime, mtime_num = 0, "-", 0
            itens.append((nome, mtime, tamanho, caminho, mtime_num))
    except Exception:
        pass
    # Mais recentes primeiro (ordena pela data real, não pelo texto)
    itens.sort(key=lambda x: x[4], reverse=True)
    return [(n, m, t, c) for n, m, t, c, _ in itens]

def recarregar_dados_sistema(limpar_formularios=False):
    """Recarrega TODAS as listas, combos, cards e o dashboard a partir do banco.
    Usado após restaurar um backup e pelo botão 'Atualizar dados'.
    Retorna True se tudo foi recarregado sem erro."""
    if root is None:
        return False
    erros = []

    def _tenta(fn):
        try:
            fn()
        except Exception as e:
            erros.append(f"{getattr(fn, '__name__', fn)}: {e}")

    if limpar_formularios:
        for fn in (limpar_form_cliente, limpar_form_fornecedor, limpar_form_produto,
                   limpar_form_usuario, limpar_form_venda):
            _tenta(fn)
    for fn in (atualizar_combos, listar_clientes, listar_fornecedores, listar_produtos,
               listar_estoque, listar_vendas, listar_caixa, listar_contas_pagar,
               listar_contas_receber, listar_todas_lixeiras, listar_usuarios,
               atualizar_dashboard):
        _tenta(fn)
    if callable(_hook_atualizar_lista_backups):
        _tenta(_hook_atualizar_lista_backups)
    try:
        root.update_idletasks()
    except Exception:
        pass
    if erros:
        print("Erros ao recarregar dados:", erros)
    return not erros

def restaurar_backup(caminho_origem):
    """Restaura o banco a partir de um arquivo .db de backup.
    Cria um backup de segurança do banco atual antes e, ao final,
    recarrega todas as telas do sistema com os dados restaurados."""
    if not caminho_origem or not os.path.isfile(caminho_origem):
        mostrar_aviso("Selecione um arquivo de backup válido (.db).")
        return False
    try:
        if os.path.abspath(caminho_origem) == os.path.abspath(DB_PATH):
            mostrar_aviso("Este arquivo é o próprio banco de dados em uso.\nSelecione um arquivo de backup.")
            return False
    except Exception:
        pass
    ok, msg = validar_arquivo_backup(caminho_origem)
    if not ok:
        mostrar_erro(f"Não foi possível restaurar:\n\n{msg}", "Backup inválido")
        return False
    if not confirmar_moderno(
        "Restaurar backup",
        "ATENÇÃO: os dados atuais serão substituídos pelo backup.\n\n"
        "Antes da restauração será gerado um backup de segurança dos dados atuais.\n\n"
        "Deseja continuar?",
    ):
        return False
    seguranca = None
    try:
        # Backup de segurança do estado atual
        garantir_pasta_backup()
        seguranca = os.path.join(
            BACKUP_DIR,
            f"antes_restore_{datetime.now().strftime('%Y%m%d_%H%M%S')}.db",
        )
        if os.path.exists(DB_PATH):
            try:
                conn = conectar()
                conn.execute("PRAGMA wal_checkpoint(FULL)")
                conn.close()
            except Exception:
                pass
            _copiar_banco_sqlite(DB_PATH, seguranca)
        else:
            seguranca = None

        # Substitui o banco atual pelo conteúdo do backup
        _copiar_banco_sqlite(caminho_origem, DB_PATH)

        # limpa possíveis arquivos WAL/SHM antigos
        for ext in ("-wal", "-shm"):
            extra = DB_PATH + ext
            if os.path.exists(extra):
                try:
                    os.remove(extra)
                except Exception:
                    pass
    except Exception as e:
        mostrar_erro(f"Falha ao restaurar backup:\n{e}")
        return False

    # Aplica migrações (caso o backup seja de uma versão anterior) e
    # recarrega TODAS as telas com os dados restaurados.
    try:
        init_db()
    except Exception as e:
        print("Erro init_db após restauração:", e)
    tudo_ok = recarregar_dados_sistema(limpar_formularios=True)

    msg = "Backup restaurado com sucesso!\n\n"
    if tudo_ok:
        msg += "Todos os dados já foram atualizados nas telas do sistema."
    else:
        msg += ("Os dados foram restaurados, mas algumas telas não puderam ser atualizadas.\n"
                "Use o botão 'Atualizar dados' ou faça logout e login.")
    if seguranca:
        msg += f"\n\nCópia de segurança dos dados anteriores:\n{os.path.basename(seguranca)}"
    mostrar_modal_moderno("Restauração concluída", msg, "sucesso", 0)
    return True

def perguntar_backup_ao_sair(titulo="Sair"):
    """Pergunta se deseja backup e onde salvar.
    Retorna:
      'cancel' -> não sair
      'sem_backup' -> sair sem backup
      caminho(str) -> backup feito e caminho
      True -> backup ok (silencioso path)
    """
    global root
    result = {"acao": "cancel"}

    modal = tk.Toplevel(root)
    modal.title(titulo)
    modal.configure(bg="white")
    modal.transient(root)
    modal.grab_set()
    modal.resizable(False, False)
    try:
        modal.update_idletasks()
        x = root.winfo_x() + max(0, (root.winfo_width() // 2) - 240)
        y = root.winfo_y() + max(0, (root.winfo_height() // 2) - 160)
        modal.geometry(f"480x320+{x}+{y}")
    except Exception:
        modal.geometry("480x320")

    header = tk.Frame(modal, bg=CORES["primary"], height=52)
    header.pack(fill="x")
    header.pack_propagate(False)
    tk.Label(header, text=titulo, bg=CORES["primary"], fg="white",
             font=("Arial", 13, "bold")).pack(pady=12)

    body = tk.Frame(modal, bg="white")
    body.pack(fill="both", expand=True, padx=20, pady=12)
    tk.Label(
        body,
        text="Deseja gerar um backup antes de sair?",
        bg="white",
        fg=CORES["text_dark"],
        font=("Arial", 11, "bold"),
    ).pack(anchor="w", pady=(0, 8))
    tk.Label(
        body,
        text="Recomendado para não perder vendas, caixa e cadastros.",
        bg="white",
        fg=CORES["text_gray"],
        font=("Arial", 9),
    ).pack(anchor="w", pady=(0, 12))

    def fechar(acao, caminho=None):
        result["acao"] = acao
        result["caminho"] = caminho
        try:
            modal.grab_release()
        except Exception:
            pass
        modal.destroy()

    def backup_escolher_local():
        garantir_pasta_backup()
        destino = filedialog.asksaveasfilename(
            parent=modal,
            title="Onde salvar o backup?",
            defaultextension=".db",
            filetypes=[("Banco SQLite", "*.db"), ("Todos", "*.*")],
            initialfile=gerar_nome_backup(),
            initialdir=BACKUP_DIR,
        )
        if not destino:
            return
        caminho = fazer_backup(destino=destino, silencioso=True)
        if caminho:
            fechar("ok", caminho)
        else:
            mostrar_erro("Não foi possível salvar o backup no local escolhido.")

    tk.Button(
        body, text="Sim — Escolher onde salvar",
        command=backup_escolher_local,
        bg=CORES["primary"], fg="white", font=("Arial", 10, "bold"),
        bd=0, padx=12, pady=8, cursor="hand2",
    ).pack(fill="x", pady=4)

    tk.Button(
        body, text="Não — Sair sem backup",
        command=lambda: fechar("sem_backup"),
        bg="#64748b", fg="white", font=("Arial", 10, "bold"),
        bd=0, padx=12, pady=8, cursor="hand2",
    ).pack(fill="x", pady=4)

    tk.Button(
        body, text="Cancelar",
        command=lambda: fechar("cancel"),
        bg="white", fg=CORES["text_gray"], font=("Arial", 9),
        bd=1, relief="solid", padx=12, pady=6, cursor="hand2",
    ).pack(fill="x", pady=(10, 0))

    modal.bind("<Escape>", lambda e: fechar("cancel"))
    modal.protocol("WM_DELETE_WINDOW", lambda: fechar("cancel"))
    modal.wait_window()
    return result




CORES = {
    "bg_dark": "#1e293b",
    "bg_dark_hover": "#334155",
    "bg_top": "#0f172a",
    "bg_light": "#f1f5f9",
    "bg_white": "#ffffff",
    "bg_tab": "#e2e8f0",
    "bg_tab_active": "#ffffff",
    "primary": "#3b82f6",
    "primary_hover": "#2563eb",
    "success": "#10b981",
    "success_hover": "#059669",
    "warning": "#f59e0b",
    "warning_hover": "#d97706",
    "danger": "#ef4444",
    "danger_hover": "#dc2626",
    "info": "#06b6d4",
    "purple": "#8b5cf6",
    "pink": "#ec4899",
    "card_blue": "#dbeafe",
    "card_green": "#dcfce7",
    "card_yellow": "#fef3c7",
    "card_red": "#fee2e2",
    "card_purple": "#ede9fe",
    "card_orange": "#ffedd5",
    "text_dark": "#1e293b",
    "text_gray": "#64748b",
    "border": "#e2e8f0",
}

# GLOBAIS
usuario_logado = None
carrinho_venda = []
telas = {}
open_tabs = {}
active_tab_key = None
listas_combos = {"clientes": [], "fornecedores": [], "produtos": []}
root = None
frame_abas = None

# DATA BR
def iso_para_br_data(data_str):
    """Converte ISO/data para dd/mm/aaaa (sem hora)."""
    if not data_str:
        return ""
    try:
        s = str(data_str).strip().split(" ")[0]
        if "-" in s and len(s) >= 10:
            dt = datetime.strptime(s[:10], "%Y-%m-%d")
            return dt.strftime("%d/%m/%Y")
        if "/" in s:
            return s[:10]
        return s
    except Exception:
        return str(data_str).split(" ")[0]

def iso_para_br(data_str):

    if not data_str:
        return ""
    try:
        data_str = str(data_str).strip()
        if " " in data_str:
            partes = data_str.split(" ")
            data_part = partes[0]
            hora_part = partes[1] if len(partes) > 1 else ""
            if "-" in data_part:
                dt = datetime.strptime(data_part[:10], "%Y-%m-%d")
                data_br = dt.strftime("%d/%m/%Y")
                if hora_part:
                    try:
                        if ":" in hora_part:
                            h = hora_part[:5]
                            return f"{data_br} {h}"
                    except:
                        pass
                    return f"{data_br} {hora_part}"
                return data_br
        else:
            if "-" in data_str:
                dt = datetime.strptime(data_str[:10], "%Y-%m-%d")
                return dt.strftime("%d/%m/%Y")
            elif "/" in data_str:
                return data_str
        return data_str
    except:
        return str(data_str)

def br_para_iso(data_br):
    if not data_br:
        return None
    data_br = str(data_br).strip()
    if not data_br:
        return None
    if len(data_br) >= 10 and data_br[4] == "-" and "-" in data_br:
        try:
            parte = data_br.split(" ")[0]
            datetime.strptime(parte, "%Y-%m-%d")
            return parte
        except:
            pass
    try:
        if "/" in data_br:
            parte = data_br.split(" ")[0]
            dt = datetime.strptime(parte, "%d/%m/%Y")
            return dt.strftime("%Y-%m-%d")
    except:
        pass
    try:
        if "-" in data_br and "/" not in data_br:
            parte = data_br.split(" ")[0]
            if len(parte.split("-")[0]) == 2:
                dt = datetime.strptime(parte, "%d-%m-%Y")
                return dt.strftime("%Y-%m-%d")
    except:
        pass
    return None

def hoje_str():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def hoje_iso():
    return date.today().isoformat()

def hoje_br():
    return datetime.now().strftime("%d/%m/%Y")

def hoje_br_completo():
    return datetime.now().strftime("%d/%m/%Y %H:%M")


import calendar as _cal_mod

def abrir_calendario(entry, parent=None):
    """Modal de calendário: escolha dia, mês e ano (dd/mm/aaaa)."""
    global root
    parent = parent or root
    top = tk.Toplevel(parent if parent is not None else root)
    top.title("Escolher data")
    top.configure(bg="white")
    try:
        top.transient(parent if parent is not None else root)
        top.grab_set()
    except Exception:
        pass
    top.resizable(False, False)

    ano = date.today().year
    mes = date.today().month
    try:
        txt = (entry.get() or "").strip().split(" ")[0]
        if "/" in txt:
            p = txt.split("/")
            if len(p) >= 3:
                d0, m0, a0 = int(p[0]), int(p[1]), int(p[2])
                if 1 <= m0 <= 12 and 1900 <= a0 <= 2100:
                    mes, ano = m0, a0
        elif "-" in txt:
            p = txt.split("-")
            if len(p) >= 3 and len(p[0]) == 4:
                ano, mes = int(p[0]), int(p[1])
    except Exception:
        pass

    estado = {"ano": ano, "mes": mes}
    meses_pt = ["", "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
                "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]

    header = tk.Frame(top, bg=CORES["primary"])
    header.pack(fill="x")
    lbl_mes = tk.Label(header, text="", bg=CORES["primary"], fg="white", font=("Arial", 11, "bold"))
    lbl_mes.pack(side="left", expand=True, pady=8)

    corpo = tk.Frame(top, bg="white", padx=8, pady=8)
    corpo.pack()
    for i, n in enumerate(["Seg", "Ter", "Qua", "Qui", "Sex", "Sáb", "Dom"]):
        tk.Label(corpo, text=n, width=4, bg="white", fg=CORES["text_gray"], font=("Arial", 8, "bold")).grid(row=0, column=i, padx=1, pady=2)

    cells = []

    def _desenhar():
        for w in cells:
            try:
                w.destroy()
            except Exception:
                pass
        cells.clear()
        lbl_mes.config(text=f"{meses_pt[estado['mes']]} {estado['ano']}")
        cal = _cal_mod.Calendar(firstweekday=0)
        row = 1
        for week in cal.monthdayscalendar(estado["ano"], estado["mes"]):
            for col, day in enumerate(week):
                if day == 0:
                    tk.Label(corpo, text="", width=4, bg="white").grid(row=row, column=col)
                    continue
                d = day

                def _escolhe(dd=d, mm=estado["mes"], yy=estado["ano"]):
                    try:
                        entry.delete(0, tk.END)
                        entry.insert(0, f"{dd:02d}/{mm:02d}/{yy}")
                        entry.event_generate("<KeyRelease>")
                    except Exception:
                        pass
                    try:
                        top.grab_release()
                    except Exception:
                        pass
                    top.destroy()

                hoje = date.today()
                is_hoje = (d == hoje.day and estado["mes"] == hoje.month and estado["ano"] == hoje.year)
                btn = tk.Button(
                    corpo, text=str(d), width=4, bd=0, font=("Arial", 9),
                    bg="#bfdbfe" if is_hoje else "white",
                    command=_escolhe, cursor="hand2",
                )
                btn.grid(row=row, column=col, padx=1, pady=1)
                cells.append(btn)
            row += 1

    def _ant():
        estado["mes"] -= 1
        if estado["mes"] < 1:
            estado["mes"] = 12
            estado["ano"] -= 1
        _desenhar()

    def _prox():
        estado["mes"] += 1
        if estado["mes"] > 12:
            estado["mes"] = 1
            estado["ano"] += 1
        _desenhar()

    tk.Button(header, text="◀", command=_ant, bg=CORES["primary"], fg="white", bd=0, padx=10, cursor="hand2").pack(side="left", padx=4)
    tk.Button(header, text="▶", command=_prox, bg=CORES["primary"], fg="white", bd=0, padx=10, cursor="hand2").pack(side="right", padx=4)
    _desenhar()

    try:
        top.update_idletasks()
        px = (parent.winfo_rootx() if parent else 100) + 40
        py = (parent.winfo_rooty() if parent else 100) + 40
        top.geometry(f"+{px}+{py}")
    except Exception:
        pass
    top.bind("<Escape>", lambda e: (top.grab_release(), top.destroy()))


def ativar_seletor_data(entry):
    """Ao clicar ou focar no campo de data, abre o modal de calendário."""
    if getattr(entry, "_cal_bound", False):
        return entry

    def _abrir(event=None):
        try:
            abrir_calendario(entry, entry.winfo_toplevel())
        except Exception:
            try:
                abrir_calendario(entry, root)
            except Exception:
                pass
        return "break"

    entry.bind("<Button-1>", _abrir)
    entry.bind("<FocusIn>", _abrir)
    entry._cal_bound = True
    return entry


def formatar_moeda(valor):
    try:
        return f"R$ {float(valor):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        return "R$ 0,00"

def formatar_status(status):
    """Exibe status com inicial maiúscula (Em aberto, Recebido, Cancelado, etc.)."""
    if not status:
        return ""
    mapa = {
        "em_aberto": "Em aberto",
        "em aberto": "Em aberto",
        "recebido": "Recebido",
        "pago": "Pago",
        "cancelado": "Cancelado",
        "em_atraso": "Em atraso",
        "em atraso": "Em atraso",
        "vencido": "Em atraso",
        "pendente": "Em aberto",
        "concluida": "Em aberto",
    }
    s = str(status).strip().lower()
    if s in mapa:
        return mapa[s]
    # fallback: capitaliza cada palavra
    return " ".join(p.capitalize() if p else p for p in str(status).replace("_", " ").split())

def formatar_telefone(event=None, entry=None):
    """Formata telefone no padrão (11) 9-0000-0000"""
    if entry is None and event is not None:
        entry = event.widget
    if entry is None:
        return
    texto = entry.get()
    digitos = re.sub(r"\D", "", texto)
    if len(digitos) > 11:
        digitos = digitos[:11]
    formatado = ""
    if len(digitos) == 0:
        formatado = ""
    elif len(digitos) <= 2:
        formatado = f"({digitos}"
    elif len(digitos) <= 3:
        formatado = f"({digitos[:2]}) {digitos[2:]}"
    elif len(digitos) <= 7:
        formatado = f"({digitos[:2]}) {digitos[2:3]}-{digitos[3:]}"
    else:
        formatado = f"({digitos[:2]}) {digitos[2:3]}-{digitos[3:7]}-{digitos[7:]}"
    # Evita loop infinito de KeyRelease
    if entry.get() != formatado:
        cursor = entry.index(tk.INSERT)
        entry.delete(0, tk.END)
        entry.insert(0, formatado)
        # tenta manter cursor no final
        try:
            entry.icursor(len(formatado))
        except:
            pass

def formatar_cep(event=None, entry=None):
    """Formata CEP 00000-000"""
    if entry is None and event is not None:
        entry = event.widget
    if entry is None:
        return
    digitos = re.sub(r"\D", "", entry.get())[:8]
    if len(digitos) > 5:
        formatado = f"{digitos[:5]}-{digitos[5:]}"
    else:
        formatado = digitos
    if entry.get() != formatado:
        entry.delete(0, tk.END)
        entry.insert(0, formatado)
        try:
            entry.icursor(len(formatado))
        except:
            pass

def formatar_cpf_cnpj(event=None, entry=None):
    """Formata automaticamente CPF (000.000.000-00) ou CNPJ (00.000.000/0000-00)"""
    if entry is None and event is not None:
        entry = event.widget
    if entry is None:
        return
    digitos = re.sub(r"\D", "", entry.get())[:14]
    if len(digitos) <= 11:
        # CPF: 000.000.000-00
        if len(digitos) <= 3:
            formatado = digitos
        elif len(digitos) <= 6:
            formatado = f"{digitos[:3]}.{digitos[3:]}"
        elif len(digitos) <= 9:
            formatado = f"{digitos[:3]}.{digitos[3:6]}.{digitos[6:]}"
        else:
            formatado = f"{digitos[:3]}.{digitos[3:6]}.{digitos[6:9]}-{digitos[9:]}"
    else:
        # CNPJ: 00.000.000/0000-00
        if len(digitos) <= 2:
            formatado = digitos
        elif len(digitos) <= 5:
            formatado = f"{digitos[:2]}.{digitos[2:]}"
        elif len(digitos) <= 8:
            formatado = f"{digitos[:2]}.{digitos[2:5]}.{digitos[5:]}"
        elif len(digitos) <= 12:
            formatado = f"{digitos[:2]}.{digitos[2:5]}.{digitos[5:8]}/{digitos[8:]}"
        else:
            formatado = f"{digitos[:2]}.{digitos[2:5]}.{digitos[5:8]}/{digitos[8:12]}-{digitos[12:]}"
    if entry.get() != formatado:
        entry.delete(0, tk.END)
        entry.insert(0, formatado)
        try:
            entry.icursor(len(formatado))
        except:
            pass

def buscar_cep_viacep(cep):
    """Consulta ViaCEP e retorna dict com logradouro, bairro, localidade, uf ou None"""
    cep_limpo = re.sub(r"\D", "", str(cep or ""))
    if len(cep_limpo) != 8:
        return None
    try:
        url = f"https://viacep.com.br/ws/{cep_limpo}/json/"
        req = urllib.request.Request(url, headers={"User-Agent": "SistemaGestao/3.3"})
        with urllib.request.urlopen(req, timeout=8) as resp:
            data = json.loads(resp.read().decode("utf-8"))
        if data.get("erro"):
            return None
        return data
    except Exception:
        return None

# BANCO
def conectar():
    conn = sqlite3.connect(DB_PATH)
    conn.execute("PRAGMA foreign_keys = ON")
    return conn

def hash_senha(senha):
    return hashlib.sha256(senha.encode()).hexdigest()

def migrar_tabela_check(nome_tabela, create_sql_novo):
    """Migra tabela antiga com CHECK restritivo para nova sem CHECK"""
    conn = conectar()
    cur = conn.cursor()
    try:
        cur.execute("SELECT sql FROM sqlite_master WHERE type='table' AND name=?", (nome_tabela,))
        row = cur.fetchone()
        if row and row[0]:
            sql_old = row[0].lower()
            # Detecta CHECK antigo com pendente
            if "check" in sql_old and ("pendente" in sql_old or "concluida" in sql_old):
                print(f"🔧 Migrando tabela {nome_tabela} - removendo CHECK antigo...")
                # Renomeia antiga
                cur.execute(f"ALTER TABLE {nome_tabela} RENAME TO {nome_tabela}_old")
                # Cria nova
                cur.execute(create_sql_novo)
                # Copia colunas comuns
                cur.execute(f"PRAGMA table_info({nome_tabela}_old)")
                cols_old = [r[1] for r in cur.fetchall()]
                cur.execute(f"PRAGMA table_info({nome_tabela})")
                cols_new = [r[1] for r in cur.fetchall()]
                cols_comum = [c for c in cols_old if c in cols_new]
                if cols_comum:
                    cols_str = ", ".join(cols_comum)
                    cur.execute(f"INSERT INTO {nome_tabela} ({cols_str}) SELECT {cols_str} FROM {nome_tabela}_old")
                cur.execute(f"DROP TABLE {nome_tabela}_old")
                conn.commit()
                print(f"✅ Tabela {nome_tabela} migrada com sucesso!")
                conn.close()
                return True
    except Exception as e:
        print(f"Erro migrar {nome_tabela}: {e}")
        try:
            conn.rollback()
        except:
            pass
    conn.close()
    return False

def init_db():
    conn = conectar()
    cur = conn.cursor()
    
    # Cria tabelas novas já sem CHECK restritivo (ou com CHECK novo)
    cur.execute("""
    CREATE TABLE IF NOT EXISTS usuarios (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nome TEXT NOT NULL,
        login TEXT UNIQUE NOT NULL,
        senha TEXT NOT NULL,
        perfil TEXT NOT NULL,
        email TEXT,
        data_cadastro TEXT DEFAULT CURRENT_TIMESTAMP,
        excluido INTEGER DEFAULT 0
    )""")
    
    cur.execute("""
    CREATE TABLE IF NOT EXISTS clientes (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nome TEXT NOT NULL,
        cpf_cnpj TEXT,
        telefone TEXT,
        email TEXT,
        endereco TEXT,
        numero TEXT,
        bairro TEXT,
        cidade TEXT,
        cep TEXT,
        data_cadastro TEXT DEFAULT CURRENT_TIMESTAMP,
        excluido INTEGER DEFAULT 0,
        data_exclusao TEXT,
        excluido_por TEXT
    )""")
    
    cur.execute("""
    CREATE TABLE IF NOT EXISTS fornecedores (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nome TEXT NOT NULL,
        cnpj TEXT,
        telefone TEXT,
        email TEXT,
        endereco TEXT,
        numero TEXT,
        bairro TEXT,
        cidade TEXT,
        cep TEXT,
        data_cadastro TEXT DEFAULT CURRENT_TIMESTAMP,
        excluido INTEGER DEFAULT 0,
        data_exclusao TEXT,
        excluido_por TEXT
    )""")
    
    cur.execute("""
    CREATE TABLE IF NOT EXISTS produtos (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nome TEXT NOT NULL,
        codigo TEXT UNIQUE,
        descricao TEXT,
        preco_custo REAL DEFAULT 0,
        preco_venda REAL NOT NULL,
        estoque_atual REAL DEFAULT 0,
        estoque_minimo REAL DEFAULT 5,
        fornecedor_id INTEGER,
        categoria TEXT,
        data_cadastro TEXT DEFAULT CURRENT_TIMESTAMP,
        excluido INTEGER DEFAULT 0,
        data_exclusao TEXT,
        excluido_por TEXT,
        FOREIGN KEY (fornecedor_id) REFERENCES fornecedores(id) ON DELETE SET NULL
    )""")
    
    cur.execute("""
    CREATE TABLE IF NOT EXISTS vendas (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        cliente_id INTEGER,
        data TEXT NOT NULL,
        total REAL NOT NULL,
        desconto REAL DEFAULT 0,
        forma_pagamento TEXT NOT NULL,
        status TEXT DEFAULT 'em_aberto',
        observacao TEXT,
        parcelas INTEGER DEFAULT 1,
        taxa REAL DEFAULT 0,
        excluido INTEGER DEFAULT 0,
        data_exclusao TEXT,
        excluido_por TEXT,
        FOREIGN KEY (cliente_id) REFERENCES clientes(id) ON DELETE SET NULL
    )""")
    
    cur.execute("""
    CREATE TABLE IF NOT EXISTS venda_itens (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        venda_id INTEGER NOT NULL,
        produto_id INTEGER NOT NULL,
        quantidade REAL NOT NULL,
        preco_unitario REAL NOT NULL,
        subtotal REAL NOT NULL,
        FOREIGN KEY (venda_id) REFERENCES vendas(id) ON DELETE CASCADE,
        FOREIGN KEY (produto_id) REFERENCES produtos(id)
    )""")
    
    cur.execute("""
    CREATE TABLE IF NOT EXISTS caixa (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        data TEXT NOT NULL,
        tipo TEXT NOT NULL,
        valor REAL NOT NULL,
        descricao TEXT NOT NULL,
        origem TEXT,
        referencia_id INTEGER,
        forma_pagamento TEXT,
        excluido INTEGER DEFAULT 0,
        data_exclusao TEXT,
        excluido_por TEXT
    )""")
    
    cur.execute("""
    CREATE TABLE IF NOT EXISTS contas_a_pagar (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        fornecedor_id INTEGER,
        descricao TEXT NOT NULL,
        valor REAL NOT NULL,
        vencimento TEXT NOT NULL,
        data_pagamento TEXT,
        status TEXT DEFAULT 'em_aberto',
        forma_pagamento TEXT,
        excluido INTEGER DEFAULT 0,
        data_exclusao TEXT,
        excluido_por TEXT,
        parcela_atual INTEGER DEFAULT 1,
        total_parcelas INTEGER DEFAULT 1,
        data_emissao TEXT,
        numero_documento TEXT,
        FOREIGN KEY (fornecedor_id) REFERENCES fornecedores(id) ON DELETE SET NULL
    )""")
    
    cur.execute("""
    CREATE TABLE IF NOT EXISTS contas_a_receber (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        cliente_id INTEGER,
        venda_id INTEGER,
        descricao TEXT NOT NULL,
        valor REAL NOT NULL,
        vencimento TEXT NOT NULL,
        data_recebimento TEXT,
        status TEXT DEFAULT 'em_aberto',
        forma_pagamento TEXT,
        parcela_atual INTEGER DEFAULT 1,
        total_parcelas INTEGER DEFAULT 1,
        excluido INTEGER DEFAULT 0,
        data_exclusao TEXT,
        excluido_por TEXT,
        data_emissao TEXT,
        numero_documento TEXT,
        FOREIGN KEY (cliente_id) REFERENCES clientes(id) ON DELETE SET NULL,
        FOREIGN KEY (venda_id) REFERENCES vendas(id) ON DELETE SET NULL
    )""")
    
    cur.execute("""
    CREATE TABLE IF NOT EXISTS movimentacao_estoque (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        produto_id INTEGER NOT NULL,
        tipo TEXT NOT NULL,
        quantidade REAL NOT NULL,
        motivo TEXT,
        data TEXT NOT NULL,
        referencia_id INTEGER,
        FOREIGN KEY (produto_id) REFERENCES produtos(id)
    )""")
    
    conn.commit()
    conn.close()
    
    # Migração de tabelas antigas com CHECK restritivo
    # SQL novo sem CHECK
    sql_cp_novo = """
    CREATE TABLE contas_a_pagar (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        fornecedor_id INTEGER,
        descricao TEXT NOT NULL,
        valor REAL NOT NULL,
        vencimento TEXT NOT NULL,
        data_pagamento TEXT,
        status TEXT DEFAULT 'em_aberto',
        forma_pagamento TEXT,
        excluido INTEGER DEFAULT 0,
        data_exclusao TEXT,
        excluido_por TEXT,
        parcela_atual INTEGER DEFAULT 1,
        total_parcelas INTEGER DEFAULT 1,
        FOREIGN KEY (fornecedor_id) REFERENCES fornecedores(id) ON DELETE SET NULL
    )"""
    
    sql_cr_novo = """
    CREATE TABLE contas_a_receber (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        cliente_id INTEGER,
        venda_id INTEGER,
        descricao TEXT NOT NULL,
        valor REAL NOT NULL,
        vencimento TEXT NOT NULL,
        data_recebimento TEXT,
        status TEXT DEFAULT 'em_aberto',
        forma_pagamento TEXT,
        parcela_atual INTEGER DEFAULT 1,
        total_parcelas INTEGER DEFAULT 1,
        excluido INTEGER DEFAULT 0,
        data_exclusao TEXT,
        excluido_por TEXT,
        FOREIGN KEY (cliente_id) REFERENCES clientes(id) ON DELETE SET NULL,
        FOREIGN KEY (venda_id) REFERENCES vendas(id) ON DELETE SET NULL
    )"""
    
    sql_vendas_novo = """
    CREATE TABLE vendas (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        cliente_id INTEGER,
        data TEXT NOT NULL,
        total REAL NOT NULL,
        desconto REAL DEFAULT 0,
        forma_pagamento TEXT NOT NULL,
        status TEXT DEFAULT 'em_aberto',
        observacao TEXT,
        parcelas INTEGER DEFAULT 1,
        taxa REAL DEFAULT 0,
        excluido INTEGER DEFAULT 0,
        data_exclusao TEXT,
        excluido_por TEXT,
        FOREIGN KEY (cliente_id) REFERENCES clientes(id) ON DELETE SET NULL
    )"""
    
    sql_caixa_novo = """
    CREATE TABLE caixa (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        data TEXT NOT NULL,
        tipo TEXT NOT NULL,
        valor REAL NOT NULL,
        descricao TEXT NOT NULL,
        origem TEXT,
        referencia_id INTEGER,
        forma_pagamento TEXT,
        excluido INTEGER DEFAULT 0,
        data_exclusao TEXT,
        excluido_por TEXT
    )"""
    
    migrar_tabela_check("contas_a_pagar", sql_cp_novo)
    migrar_tabela_check("contas_a_receber", sql_cr_novo)
    migrar_tabela_check("vendas", sql_vendas_novo)
    migrar_tabela_check("caixa", sql_caixa_novo)
    
    # Migração colunas adicionais
    conn = conectar()
    cur = conn.cursor()
    tabelas_colunas = {
        "clientes": [("excluido", "INTEGER DEFAULT 0"), ("data_exclusao", "TEXT"), ("excluido_por", "TEXT"), ("numero", "TEXT"), ("bairro", "TEXT"), ("cidade", "TEXT"), ("cep", "TEXT")],
        "fornecedores": [("excluido", "INTEGER DEFAULT 0"), ("data_exclusao", "TEXT"), ("excluido_por", "TEXT"), ("numero", "TEXT"), ("bairro", "TEXT"), ("cidade", "TEXT"), ("cep", "TEXT")],
        "produtos": [("excluido", "INTEGER DEFAULT 0"), ("data_exclusao", "TEXT"), ("excluido_por", "TEXT")],
        "vendas": [("excluido", "INTEGER DEFAULT 0"), ("data_exclusao", "TEXT"), ("excluido_por", "TEXT"), ("parcelas", "INTEGER DEFAULT 1"), ("taxa", "REAL DEFAULT 0")],
        "caixa": [("excluido", "INTEGER DEFAULT 0"), ("data_exclusao", "TEXT"), ("excluido_por", "TEXT")],
        "contas_a_pagar": [("excluido", "INTEGER DEFAULT 0"), ("data_exclusao", "TEXT"), ("excluido_por", "TEXT"), ("parcela_atual", "INTEGER DEFAULT 1"), ("total_parcelas", "INTEGER DEFAULT 1"), ("data_emissao", "TEXT"), ("numero_documento", "TEXT")],
        "contas_a_receber": [("excluido", "INTEGER DEFAULT 0"), ("data_exclusao", "TEXT"), ("excluido_por", "TEXT"), ("parcela_atual", "INTEGER DEFAULT 1"), ("total_parcelas", "INTEGER DEFAULT 1"), ("data_emissao", "TEXT"), ("numero_documento", "TEXT")],
        "usuarios": [("excluido", "INTEGER DEFAULT 0"), ("email", "TEXT")],
    }
    for tabela, colunas in tabelas_colunas.items():
        for col_nome, col_tipo in colunas:
            try:
                cur.execute(f"SELECT {col_nome} FROM {tabela} LIMIT 1")
            except sqlite3.OperationalError:
                try:
                    cur.execute(f"ALTER TABLE {tabela} ADD COLUMN {col_nome} {col_tipo}")
                except:
                    pass
    conn.commit()
    
    # Migração de status antigos
    try:
        cur.execute("UPDATE vendas SET status='em_aberto' WHERE status='concluida' AND excluido=0")
        cur.execute("UPDATE vendas SET status='em_aberto' WHERE status='pendente' AND excluido=0")
        cur.execute("UPDATE contas_a_receber SET status='em_aberto' WHERE status='pendente' AND excluido=0")
        cur.execute("UPDATE contas_a_pagar SET status='em_aberto' WHERE status='pendente' AND excluido=0")
        cur.execute("UPDATE contas_a_receber SET status='em_atraso' WHERE status='vencido' AND excluido=0")
        cur.execute("UPDATE contas_a_pagar SET status='em_atraso' WHERE status='vencido' AND excluido=0")
        cur.execute("UPDATE contas_a_receber SET status='recebido' WHERE status='pago' AND excluido=0")
        conn.commit()
    except:
        pass
    
    try:
        hoje = hoje_iso()
        cur.execute("UPDATE contas_a_receber SET status='em_atraso' WHERE status='em_aberto' AND date(vencimento) < date(?) AND excluido=0", (hoje,))
        cur.execute("UPDATE contas_a_pagar SET status='em_atraso' WHERE status='em_aberto' AND date(vencimento) < date(?) AND excluido=0", (hoje,))
        conn.commit()
    except:
        pass
    
    cur.execute("SELECT COUNT(*) FROM usuarios WHERE excluido=0")
    if cur.fetchone()[0] == 0:
        cur.execute("INSERT INTO usuarios (nome, login, senha, perfil, email) VALUES (?,?,?,?,?)",
                    ("Administrador", "admin", hash_senha("admin123"), "admin", "admin@sistema.local"))
        cur.execute("INSERT INTO usuarios (nome, login, senha, perfil, email) VALUES (?,?,?,?,?)",
                    ("Operador", "operador", hash_senha("operador123"), "operador", "operador@sistema.local"))
    else:
        # Garante email nos usuários padrão existentes
        try:
            cur.execute("UPDATE usuarios SET email='admin@sistema.local' WHERE login='admin' AND (email IS NULL OR email='')")
            cur.execute("UPDATE usuarios SET email='operador@sistema.local' WHERE login='operador' AND (email IS NULL OR email='')")
        except:
            pass
    
    conn.commit()
    conn.close()

# PERMISSÕES
def eh_admin():
    return usuario_logado and usuario_logado.get('perfil') == 'admin'

def verificar_permissao_exclusao():
    if not eh_admin():
        mostrar_aviso("🔒 Somente usuário ADM pode excluir!\n\nSeu perfil: " + (usuario_logado.get('perfil') if usuario_logado else 'desconhecido'))
        return False
    return True

# MODAL MODERNO SEM SOM
def mostrar_modal_moderno(titulo, mensagem, tipo="sucesso", duracao_ms=None):
    global root
    if root is None:
        return
    cores = {"sucesso": CORES["success"], "erro": CORES["danger"], "aviso": CORES["warning"], "info": CORES["primary"]}
    icones = {"sucesso": "✅", "erro": "❌", "aviso": "⚠️", "info": "ℹ️"}
    
    modal = tk.Toplevel(root)
    modal.title(titulo)
    modal.geometry("460x240")
    modal.configure(bg="white")
    modal.transient(root)
    modal.grab_set()
    modal.resizable(False, False)
    
    modal.update_idletasks()
    try:
        x = root.winfo_x() + (root.winfo_width()//2) - 230
        y = root.winfo_y() + (root.winfo_height()//2) - 120
        modal.geometry(f"+{x}+{y}")
    except:
        pass
    
    header = tk.Frame(modal, bg=cores.get(tipo, CORES["success"]), height=70)
    header.pack(fill='x')
    header.pack_propagate(False)
    tk.Label(header, text=f"{icones.get(tipo, '✅')}  {titulo}", bg=cores.get(tipo, CORES["success"]), fg="white", font=('Arial', 14, 'bold')).pack(pady=20)
    
    body = tk.Frame(modal, bg="white")
    body.pack(fill='both', expand=True, padx=25, pady=15)
    tk.Label(body, text=mensagem, bg="white", fg="#334155", font=('Arial', 11), wraplength=400, justify='center').pack(pady=10)
    
    def fechar():
        try:
            modal.grab_release()
        except:
            pass
        modal.destroy()
    
    btn = tk.Button(body, text="OK", command=fechar, bg=cores.get(tipo, CORES["success"]), fg="white", font=('Arial', 11, 'bold'), bd=0, padx=30, pady=8, cursor='hand2')
    btn.pack(pady=12)
    btn.focus()
    modal.bind('<Return>', lambda e: fechar())
    modal.bind('<Escape>', lambda e: fechar())
    
    if duracao_ms is None and tipo == "sucesso":
        duracao_ms = 2500
    if duracao_ms:
        modal.after(duracao_ms, fechar)

    # Ajusta a altura ao conteúdo (mensagens longas não escondem o botão OK)
    try:
        modal.update_idletasks()
        h_req = int(modal.winfo_reqheight())
        if h_req > 240:
            modal.geometry(f"460x{h_req}")
    except Exception:
        pass
    
    modal.wait_window()

def mostrar_sucesso(mensagem, titulo="Ação Concluída"):
    mostrar_modal_moderno(titulo, mensagem, "sucesso", 2500)

def mostrar_erro(mensagem, titulo="Erro"):
    mostrar_modal_moderno(titulo, mensagem, "erro", 4000)

def mostrar_aviso(mensagem, titulo="Aviso"):
    mostrar_modal_moderno(titulo, mensagem, "aviso", 3500)

def mostrar_info(mensagem, titulo="Informação"):
    mostrar_modal_moderno(titulo, mensagem, "info", 3000)

def confirmar_moderno(titulo, mensagem):
    global root
    result = {"value": False}
    
    modal = tk.Toplevel(root)
    modal.title(titulo)
    modal.geometry("480x260")
    modal.configure(bg="white")
    modal.transient(root)
    modal.grab_set()
    modal.resizable(False, False)
    
    modal.update_idletasks()
    try:
        x = root.winfo_x() + (root.winfo_width()//2) - 240
        y = root.winfo_y() + (root.winfo_height()//2) - 130
        modal.geometry(f"+{x}+{y}")
    except:
        pass
    
    header = tk.Frame(modal, bg=CORES["warning"], height=60)
    header.pack(fill='x')
    header.pack_propagate(False)
    tk.Label(header, text=f"⚠️  {titulo}", bg=CORES["warning"], fg="white", font=('Arial', 13, 'bold')).pack(pady=15)
    
    body = tk.Frame(modal, bg="white")
    body.pack(fill='both', expand=True, padx=25, pady=15)
    tk.Label(body, text=mensagem, bg="white", fg="#334155", font=('Arial', 11), wraplength=420, justify='center').pack(pady=10)
    
    def sim():
        result["value"] = True
        try:
            modal.grab_release()
        except:
            pass
        modal.destroy()
    
    def nao():
        result["value"] = False
        try:
            modal.grab_release()
        except:
            pass
        modal.destroy()
    
    frame_btn = tk.Frame(body, bg="white")
    frame_btn.pack(pady=15)
    tk.Button(frame_btn, text="❌ Não", command=nao, bg="#64748b", fg="white", font=('Arial', 10, 'bold'), bd=0, padx=20, pady=8, cursor='hand2').pack(side='left', padx=10)
    tk.Button(frame_btn, text="✅ Sim", command=sim, bg=CORES["success"], fg="white", font=('Arial', 10, 'bold'), bd=0, padx=20, pady=8, cursor='hand2').pack(side='left', padx=10)
    
    modal.bind('<Escape>', lambda e: nao())
    modal.bind('<Return>', lambda e: sim())

    # Ajusta a altura ao conteúdo (mensagens longas não escondem os botões)
    try:
        modal.update_idletasks()
        h_req = int(modal.winfo_reqheight())
        if h_req > 260:
            modal.geometry(f"480x{h_req}")
    except Exception:
        pass
    
    modal.wait_window()
    return result["value"]


def confirmar_contas_lote(titulo, contas, total, acao_label="Confirmar"):
    """Modal com lista rolável, total e botões sempre visíveis (Receber/Pagar)."""
    global root
    result = {"value": False}

    modal = tk.Toplevel(root)
    modal.title(titulo)
    modal.configure(bg="white")
    modal.transient(root)
    modal.grab_set()
    modal.resizable(True, True)
    modal.minsize(560, 420)

    try:
        x = root.winfo_x() + max(0, (root.winfo_width() // 2) - 310)
        y = root.winfo_y() + max(0, (root.winfo_height() // 2) - 250)
        modal.geometry(f"620x500+{x}+{y}")
    except Exception:
        modal.geometry("620x500")

    is_receber = ("Receber" in titulo) or ("receb" in titulo.lower())
    cor_header = CORES["primary"] if is_receber else CORES["danger"]

    # Header
    header = tk.Frame(modal, bg=cor_header, height=56)
    header.pack(fill='x')
    header.pack_propagate(False)
    tk.Label(header, text=titulo, bg=cor_header, fg="white", font=('Arial', 13, 'bold')).pack(pady=14)

    # Botões PRIMEIRO no rodapé (sempre visíveis)
    frame_btn = tk.Frame(modal, bg="#f1f5f9", height=64)
    frame_btn.pack(fill='x', side='bottom')
    frame_btn.pack_propagate(False)

    def fechar(ok=False):
        result["value"] = bool(ok)
        try:
            modal.grab_release()
        except Exception:
            pass
        try:
            modal.destroy()
        except Exception:
            pass

    tk.Button(
        frame_btn, text="❌ Cancelar", command=lambda: fechar(False),
        bg="#64748b", fg="white", font=('Arial', 10, 'bold'),
        bd=0, padx=18, pady=10, cursor='hand2',
    ).pack(side='left', padx=16, pady=12)

    tk.Button(
        frame_btn, text=f"✅ {acao_label}", command=lambda: fechar(True),
        bg=CORES["success"], fg="white", font=('Arial', 12, 'bold'),
        bd=0, padx=24, pady=10, cursor='hand2',
    ).pack(side='right', padx=16, pady=12)

    # Total acima dos botões
    frame_total = tk.Frame(modal, bg="#fef3c7")
    frame_total.pack(fill='x', side='bottom')
    tk.Label(frame_total, text="TOTAL", bg="#fef3c7", font=('Arial', 11, 'bold'), fg="#92400e").pack(side='left', padx=16, pady=10)
    tk.Label(frame_total, text=formatar_moeda(total), bg="#fef3c7", font=('Arial', 15, 'bold'), fg=CORES["danger"]).pack(side='right', padx=16, pady=10)

    # Corpo com lista
    body = tk.Frame(modal, bg="white")
    body.pack(fill='both', expand=True, padx=14, pady=10)

    tk.Label(
        body, text=f"{len(contas)} conta(s) selecionada(s)",
        bg="white", fg=CORES["text_dark"], font=('Arial', 10, 'bold'),
    ).pack(anchor='w')

    head = tk.Frame(body, bg="#e2e8f0")
    head.pack(fill='x', pady=(8, 0))
    tk.Label(head, text="ID", bg="#e2e8f0", font=('Arial', 9, 'bold'), width=8, anchor='w').pack(side='left', padx=6, pady=4)
    tk.Label(head, text="Descrição", bg="#e2e8f0", font=('Arial', 9, 'bold'), anchor='w').pack(side='left', fill='x', expand=True, padx=4, pady=4)
    tk.Label(head, text="Valor", bg="#e2e8f0", font=('Arial', 9, 'bold'), width=14, anchor='e').pack(side='right', padx=10, pady=4)

    frame_lista = tk.Frame(body, bg="white", bd=1, relief='solid')
    frame_lista.pack(fill='both', expand=True)

    canvas = tk.Canvas(frame_lista, bg="white", highlightthickness=0)
    sb = ttk.Scrollbar(frame_lista, orient='vertical', command=canvas.yview)
    inner = tk.Frame(canvas, bg="white")
    win_id = canvas.create_window((0, 0), window=inner, anchor='nw')

    def _cfg_inner(event=None):
        canvas.configure(scrollregion=canvas.bbox("all"))
        try:
            canvas.itemconfigure(win_id, width=canvas.winfo_width())
        except Exception:
            pass

    inner.bind("<Configure>", _cfg_inner)
    canvas.bind("<Configure>", _cfg_inner)
    canvas.configure(yscrollcommand=sb.set)
    canvas.pack(side='left', fill='both', expand=True)
    sb.pack(side='right', fill='y')

    def _wheel(event):
        delta = getattr(event, 'delta', 0)
        if delta:
            canvas.yview_scroll(int(-1 * (delta / 120)), "units")
        elif getattr(event, 'num', None) == 4:
            canvas.yview_scroll(-1, "units")
        elif getattr(event, 'num', None) == 5:
            canvas.yview_scroll(1, "units")

    canvas.bind("<Enter>", lambda e: canvas.bind_all("<MouseWheel>", _wheel))
    canvas.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))
    canvas.bind("<Button-4>", _wheel)
    canvas.bind("<Button-5>", _wheel)

    for i, c in enumerate(contas):
        bg = "#f8fafc" if i % 2 == 0 else "white"
        row = tk.Frame(inner, bg=bg)
        row.pack(fill='x')
        tk.Label(row, text=str(c.get("id", "")), bg=bg, font=('Arial', 9), width=8, anchor='w').pack(side='left', padx=6, pady=4)
        desc = str(c.get("desc") or "-")
        if len(desc) > 55:
            desc = desc[:52] + "..."
        tk.Label(row, text=desc, bg=bg, font=('Arial', 9), anchor='w').pack(side='left', fill='x', expand=True, padx=4, pady=4)
        tk.Label(row, text=formatar_moeda(c.get("valor", 0)), bg=bg, font=('Arial', 9, 'bold'), width=14, anchor='e').pack(side='right', padx=10, pady=4)

    modal.bind('<Escape>', lambda e: fechar(False))
    modal.protocol("WM_DELETE_WINDOW", lambda: fechar(False))
    modal.wait_window()
    return result["value"]


# VALIDAÇÃO
def validar_obrigatorio(valor, nome_campo):
    if not valor or not str(valor).strip():
        mostrar_aviso(f"Campo obrigatório não preenchido:\n\n📌 {nome_campo}\n\nPreencha antes de salvar!", "Campo Obrigatório")
        return False
    return True

def validar_numero(valor, nome_campo, permitir_zero=False):
    try:
        num = float(str(valor).replace(",", "."))
        if not permitir_zero and num <= 0:
            mostrar_aviso(f"{nome_campo} deve ser maior que zero!", "Valor Inválido")
            return False
        return True
    except:
        mostrar_aviso(f"{nome_campo} inválido! Digite um número.", "Valor Inválido")
        return False

# AUTOCOMPLETE
def configurar_autocomplete_combo(combo, lista_completa_ref, min_chars=3):
    """Autocomplete em Combobox: a partir de min_chars caracteres mostra opções filtradas."""
    # Evita empilhar binds em rechamadas de atualizar_combos
    for seq in ('<KeyRelease>', '<FocusIn>', '<Button-1>'):
        try:
            combo.unbind(seq)
        except Exception:
            pass

    def _lista():
        # lista_completa_ref pode ser lista mutável compartilhada
        return list(lista_completa_ref) if not callable(lista_completa_ref) else list(lista_completa_ref())

    def on_keyrelease(event):
        if event.keysym in ('Up', 'Down', 'Left', 'Right', 'Return', 'Escape', 'Tab'):
            return
        valor_digitado = combo.get().strip()
        base = _lista()
        if len(valor_digitado) < min_chars:
            combo['values'] = base
            return
        filtrados = [item for item in base if valor_digitado.lower() in str(item).lower()]
        combo['values'] = filtrados if filtrados else base
        combo.focus()
        if filtrados:
            try:
                combo.event_generate('<Down>')
            except Exception:
                pass

    def on_focus_in(event):
        combo['values'] = _lista()

    combo.bind('<KeyRelease>', on_keyrelease)
    combo.bind('<FocusIn>', on_focus_in)
    combo.bind('<Button-1>', on_focus_in)

# Popup de sugestões para campos Entry de busca (após 3 caracteres)
_sugestao_popup = {"win": None}

def fechar_sugestao_popup():
    try:
        if _sugestao_popup["win"] is not None:
            _sugestao_popup["win"].destroy()
    except Exception:
        pass
    _sugestao_popup["win"] = None

def configurar_busca_com_opcoes(entry, get_opcoes_fn, on_escolher=None, min_chars=3, max_itens=12):
    """
    No Entry de busca: ao digitar min_chars caracteres, abre lista de opções.
    get_opcoes_fn(texto) -> lista de strings
    on_escolher(texto_escolhido) -> callback opcional (ex.: filtrar grid)
    """
    for seq in ('<KeyRelease>', '<FocusOut>', '<Escape>'):
        try:
            entry.unbind(seq)
        except Exception:
            pass

    def mostrar_opcoes(event=None):
        if event and getattr(event, 'keysym', None) in ('Up', 'Down', 'Left', 'Right', 'Return', 'Tab'):
            return
        texto = entry.get().strip()
        fechar_sugestao_popup()
        if len(texto) < min_chars:
            if on_escolher:
                on_escolher(texto)
            return
        try:
            opcoes = get_opcoes_fn(texto) or []
        except Exception:
            opcoes = []
        opcoes = opcoes[:max_itens]
        if not opcoes:
            if on_escolher:
                on_escolher(texto)
            return

        # Janela flutuante logo abaixo do entry
        root_win = entry.winfo_toplevel()
        popup = tk.Toplevel(root_win)
        popup.overrideredirect(True)
        popup.attributes('-topmost', True)
        _sugestao_popup["win"] = popup

        x = entry.winfo_rootx()
        y = entry.winfo_rooty() + entry.winfo_height()
        largura = max(entry.winfo_width(), 280)
        altura = min(220, 24 * len(opcoes) + 8)
        popup.geometry(f"{largura}x{altura}+{x}+{y}")

        lb = tk.Listbox(popup, font=('Arial', 10), activestyle='dotbox',
                        bg='white', fg='#1e293b', selectbackground='#3b82f6',
                        selectforeground='white', relief='solid', bd=1)
        lb.pack(fill='both', expand=True)
        for op in opcoes:
            lb.insert(tk.END, op)

        def escolher(evt=None):
            sel = lb.curselection()
            if not sel:
                return
            valor = lb.get(sel[0])
            entry.delete(0, tk.END)
            entry.insert(0, valor)
            fechar_sugestao_popup()
            if on_escolher:
                on_escolher(valor)
            entry.focus_set()

        def navegar(evt):
            if not lb.size():
                return
            if evt.keysym == 'Down':
                idx = lb.curselection()[0] + 1 if lb.curselection() else 0
                if idx >= lb.size():
                    idx = 0
                lb.selection_clear(0, tk.END)
                lb.selection_set(idx)
                lb.activate(idx)
                lb.see(idx)
            elif evt.keysym == 'Up':
                idx = lb.curselection()[0] - 1 if lb.curselection() else lb.size() - 1
                if idx < 0:
                    idx = lb.size() - 1
                lb.selection_clear(0, tk.END)
                lb.selection_set(idx)
                lb.activate(idx)
                lb.see(idx)
            elif evt.keysym == 'Return':
                escolher()
            elif evt.keysym == 'Escape':
                fechar_sugestao_popup()

        lb.bind('<Double-Button-1>', escolher)
        lb.bind('<ButtonRelease-1>', escolher)
        entry.bind('<Down>', navegar, add='+')
        entry.bind('<Up>', navegar, add='+')
        entry.bind('<Return>', lambda e: (escolher() if _sugestao_popup["win"] else None), add='+')
        entry.bind('<Escape>', lambda e: fechar_sugestao_popup(), add='+')
        popup.bind('<Escape>', lambda e: fechar_sugestao_popup())

        # Filtra grid em paralelo enquanto digita
        if on_escolher:
            on_escolher(texto)

    def ao_sair(event=None):
        # Pequeno atraso para permitir clique na lista
        entry.after(180, fechar_sugestao_popup)

    entry.bind('<KeyRelease>', mostrar_opcoes)
    entry.bind('<FocusOut>', ao_sair)

# HELPERS
def limpar_tree(tree):
    for item in tree.get_children():
        tree.delete(item)

# ============================================================
# SELEÇÃO MÚLTIPLA (checkbox) — reutilizável em qualquer lista
# ============================================================
_selecoes_multiplas = {}  # id(tree) -> set de iids marcados

def _sel_key(tree):
    return id(tree)

def habilitar_selecao_multipla(tree):
    """Torna a ÚLTIMA coluna da tree um checkbox clicável (☐ / ☑) para seleção múltipla.
    Usar a última coluna (e não a primeira) evita quebrar qualquer código existente
    que leia valores[0] como o ID do registro."""
    _selecoes_multiplas[_sel_key(tree)] = set()
    try:
        tree.tag_configure("marcado", background="#dbeafe")
    except Exception:
        pass

    def _on_click(event):
        regiao = tree.identify_region(event.x, event.y)
        if regiao != "cell":
            return
        col = tree.identify_column(event.x)
        linha = tree.identify_row(event.y)
        if not linha:
            return
        n_colunas = len(tree["columns"])
        if col != f"#{n_colunas}":  # apenas a última coluna é o checkbox
            return
        marcados = _selecoes_multiplas.setdefault(_sel_key(tree), set())
        valores = list(tree.item(linha, "values"))
        if not valores:
            return
        tags_atuais = [t for t in tree.item(linha, "tags") if t != "marcado"]
        if linha in marcados:
            marcados.discard(linha)
            valores[-1] = "☐"
        else:
            marcados.add(linha)
            valores[-1] = "☑"
            tags_atuais.append("marcado")
        tree.item(linha, values=valores, tags=tuple(tags_atuais))

    tree.bind("<Button-1>", _on_click, add="+")

def limpar_selecao_multipla(tree):
    _selecoes_multiplas[_sel_key(tree)] = set()

def obter_ids_selecionados(tree):
    """IDs (1ª coluna, inalterada) das linhas marcadas na tree."""
    marcados = _selecoes_multiplas.get(_sel_key(tree), set())
    ids = []
    for iid in marcados:
        try:
            if tree.exists(iid):
                valores = tree.item(iid, "values")
                ids.append(valores[0])
        except Exception:
            pass
    return ids

def criar_barra_selecao_multipla(parent, tree, on_excluir_ids=None, texto_botao="Excluir selecionados", mostrar_excluir=True):
    """Barra com 'Marcar todos'. Botão excluir opcional (em CP/CR já existe botão abaixo)."""
    barra = tk.Frame(parent, bg=CORES["bg_light"])
    var_todos = tk.BooleanVar(value=False)

    def _marcar_todos():
        marcados = _selecoes_multiplas.setdefault(_sel_key(tree), set())
        marcar = var_todos.get()
        for iid in tree.get_children():
            valores = list(tree.item(iid, "values"))
            if not valores:
                continue
            if marcar:
                marcados.add(iid)
                valores[-1] = "☑"
                tree.item(iid, values=valores, tags=tuple([t for t in tree.item(iid, "tags") if t != "marcado"] + ["marcado"]))
            else:
                marcados.discard(iid)
                valores[-1] = "☐"
                tree.item(iid, values=valores, tags=tuple([t for t in tree.item(iid, "tags") if t != "marcado"]))

    tk.Checkbutton(barra, text="Marcar todos", variable=var_todos, command=_marcar_todos,
                   bg=CORES["bg_light"], font=('Arial', 9)).pack(side='left', padx=(0, 14))

    if mostrar_excluir and on_excluir_ids is not None:
        def _excluir():
            ids = obter_ids_selecionados(tree)
            if not ids:
                mostrar_aviso("Marque ao menos um item (clique no ☐) para excluir.")
                return
            plural = "ns" if len(ids) > 1 else ""
            if not confirmar_moderno("Excluir selecionados", f"Confirma a exclusão de {len(ids)} ite{plural} selecionado{plural}?"):
                return
            on_excluir_ids(ids)
            var_todos.set(False)
            limpar_selecao_multipla(tree)

        tk.Button(barra, text=texto_botao, command=_excluir, bg=CORES["danger"], fg="white",
                  bd=0, padx=10, pady=4, font=('Arial', 9, 'bold'), cursor='hand2').pack(side='left')
    return barra



def _configurar_tree_com_checkbox(tree, colunas, larguras=None):
    """Configura Treeview com coluna Sel (checkbox) sempre visível à direita."""
    for c in colunas:
        tree.heading(c, text=c)
    if larguras is None:
        larguras = {}
    for c in colunas:
        if c == "Sel":
            tree.column(c, width=50, minwidth=45, anchor='center', stretch=False)
            tree.heading(c, text="☐")
        else:
            w = larguras.get(c, 120)
            tree.column(c, width=w, minwidth=40, anchor='w' if c not in ("ID", "Valor", "Venc BR", "Data Venda", "Status", "Parcela") else 'center', stretch=True)
    habilitar_selecao_multipla(tree)


def criar_barra_lixeira(parent, tree, tabela):
    """Barra com marcar todos + restaurar em massa + excluir definitivo em massa (Lixeira)."""
    barra = tk.Frame(parent, bg=CORES["bg_light"])
    var_todos = tk.BooleanVar(value=False)

    def _marcar_todos():
        marcados = _selecoes_multiplas.setdefault(_sel_key(tree), set())
        marcar = var_todos.get()
        for iid in tree.get_children():
            valores = list(tree.item(iid, "values"))
            if not valores:
                continue
            if marcar:
                marcados.add(iid)
                valores[-1] = "☑"
                tree.item(iid, values=valores, tags=tuple([t for t in tree.item(iid, "tags") if t != "marcado"] + ["marcado"]))
            else:
                marcados.discard(iid)
                valores[-1] = "☐"
                tree.item(iid, values=valores, tags=tuple([t for t in tree.item(iid, "tags") if t != "marcado"]))

    def _restaurar():
        restaurar_itens_em_massa(tabela, tree)
        var_todos.set(False)

    def _excluir_def():
        excluir_definitivo_em_massa(tabela, tree)
        var_todos.set(False)

    tk.Checkbutton(barra, text="Marcar todos", variable=var_todos, command=_marcar_todos,
                   bg=CORES["bg_light"], font=('Arial', 9)).pack(side='left', padx=(0, 14))
    tk.Button(barra, text="Restaurar selecionados", command=_restaurar, bg=CORES["success"], fg="white",
              bd=0, padx=10, pady=4, font=('Arial', 9, 'bold'), cursor='hand2').pack(side='left', padx=4)
    tk.Button(barra, text="Excluir definitivo", command=_excluir_def, bg=CORES["danger"], fg="white",
              bd=0, padx=10, pady=4, font=('Arial', 9, 'bold'), cursor='hand2').pack(side='left', padx=4)
    return barra


# ============================================================
# LISTAGEM (sem paginação — mostra todos os registros)
# ============================================================
_estado_pagina = {}  # mantido só por compatibilidade

def criar_controle_paginacao(parent, chave, bg=None):
    """Paginação removida: retorna frame vazio (compatibilidade com chamadas existentes)."""
    bg = bg or CORES["bg_light"]
    frame = tk.Frame(parent, bg=bg)
    # não exibe controles de página
    return frame

def definir_dados_paginados(chave, tree, lista_valores, lista_tags=None, reset_pagina=True):
    """Carrega TODOS os registros na tree (sem paginação)."""
    if tree is None:
        return
    try:
        limpar_tree(tree)
    except Exception:
        pass
    try:
        limpar_selecao_multipla(tree)
    except Exception:
        pass
    dados = list(lista_valores or [])
    tags_list = list(lista_tags) if lista_tags is not None else [() for _ in dados]
    for i, vals in enumerate(dados):
        tag = tags_list[i] if i < len(tags_list) else ()
        try:
            tree.insert("", "end", values=vals, tags=tag if tag else ())
        except Exception:
            try:
                tree.insert("", "end", values=vals)
            except Exception:
                pass

def exportar_dados(colunas, dados, nome_padrao):
    if not dados:
        mostrar_aviso("Nenhum dado para exportar!")
        return
    arquivo = filedialog.asksaveasfilename(
        defaultextension=".xlsx" if HAS_OPENPYXL else ".csv",
        filetypes=[("Excel", "*.xlsx"), ("CSV", "*.csv"), ("Todos", "*.*")],
        initialfile=nome_padrao
    )
    if not arquivo:
        return
    try:
        if arquivo.endswith(".xlsx") and HAS_OPENPYXL:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.append(colunas)
            for linha in dados:
                ws.append(linha)
            for col in ws.columns:
                max_len = 0
                col_letter = col[0].column_letter
                for cell in col:
                    if cell.value:
                        max_len = max(max_len, len(str(cell.value)))
                ws.column_dimensions[col_letter].width = min(max_len+2, 40)
            wb.save(arquivo)
        else:
            if not arquivo.endswith(".csv"):
                arquivo += ".csv"
            with open(arquivo, 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f, delimiter=';')
                writer.writerow(colunas)
                writer.writerows(dados)
        mostrar_sucesso(f"Relatório exportado com sucesso!\n\n{arquivo}", "Exportado")
    except Exception as e:
        mostrar_erro(f"Erro ao exportar: {e}")

def atualizar_status_atraso():
    try:
        conn = conectar()
        cur = conn.cursor()
        hoje = hoje_iso()
        cur.execute("UPDATE contas_a_receber SET status='em_atraso' WHERE status='em_aberto' AND date(vencimento) < date(?) AND excluido=0", (hoje,))
        cur.execute("UPDATE contas_a_pagar SET status='em_atraso' WHERE status='em_aberto' AND date(vencimento) < date(?) AND excluido=0", (hoje,))
        cur.execute("""
            UPDATE vendas SET status='em_atraso' 
            WHERE id IN (SELECT venda_id FROM contas_a_receber WHERE status='em_atraso' AND venda_id IS NOT NULL AND excluido=0)
            AND status='em_aberto' AND excluido=0
        """)
        conn.commit()
        conn.close()
    except:
        pass

def atualizar_status_venda(venda_id):
    try:
        conn = conectar()
        cur = conn.cursor()
        cur.execute("SELECT status FROM vendas WHERE id=? AND excluido=0", (venda_id,))
        venda = cur.fetchone()
        if not venda or venda[0] == 'cancelado':
            conn.close()
            return
        cur.execute("SELECT status FROM contas_a_receber WHERE venda_id=? AND excluido=0", (venda_id,))
        contas = cur.fetchall()
        if not contas:
            conn.close()
            return
        statuses = [c[0] for c in contas]
        if all(s == 'recebido' for s in statuses):
            cur.execute("UPDATE vendas SET status='recebido' WHERE id=?", (venda_id,))
        elif any(s == 'em_atraso' for s in statuses):
            cur.execute("UPDATE vendas SET status='em_atraso' WHERE id=?", (venda_id,))
        elif any(s == 'em_aberto' for s in statuses):
            cur.execute("UPDATE vendas SET status='em_aberto' WHERE id=?", (venda_id,))
        conn.commit()
        conn.close()
    except:
        pass

# DASHBOARD
def atualizar_dashboard():
    try:
        atualizar_status_atraso()
        conn = conectar()
        cur = conn.cursor()
        cur.execute("SELECT COALESCE(SUM(CASE WHEN tipo='entrada' THEN valor ELSE -valor END),0) FROM caixa WHERE excluido=0")
        saldo = cur.fetchone()[0] or 0
        cur.execute("SELECT COUNT(*), COALESCE(SUM(total),0) FROM vendas WHERE date(data)=date('now') AND status!='cancelado' AND excluido=0")
        qtd_hoje, total_hoje = cur.fetchone()
        cur.execute("SELECT COALESCE(SUM(valor),0) FROM contas_a_receber WHERE status IN ('em_aberto','em_atraso') AND excluido=0")
        a_receber = cur.fetchone()[0] or 0
        cur.execute("SELECT COALESCE(SUM(valor),0) FROM contas_a_pagar WHERE status IN ('em_aberto','em_atraso') AND excluido=0")
        a_pagar = cur.fetchone()[0] or 0
        cur.execute("SELECT COUNT(*) FROM produtos WHERE estoque_atual <= estoque_minimo AND excluido=0")
        estoque_baixo = cur.fetchone()[0]
        cur.execute("SELECT COUNT(*) FROM clientes WHERE excluido=0")
        total_clientes = cur.fetchone()[0]
        cur.execute("SELECT COUNT(*) FROM produtos WHERE excluido=0")
        total_produtos = cur.fetchone()[0]
        conn.close()
        
        lbl_saldo_val.config(text=formatar_moeda(saldo))
        lbl_vendas_hoje_val.config(text=f"{qtd_hoje} vendas\n{formatar_moeda(total_hoje)}")
        lbl_receber_val.config(text=formatar_moeda(a_receber))
        lbl_pagar_val.config(text=formatar_moeda(a_pagar))
        lbl_estoque_baixo_val.config(text=f"{estoque_baixo} produtos")
        lbl_clientes_val.config(text=str(total_clientes))
        lbl_produtos_val.config(text=str(total_produtos))
        
        limpar_tree(tree_dashboard_vendas)
        conn = conectar()
        cur = conn.cursor()
        cur.execute("""
            SELECT v.id, v.data, COALESCE(c.nome,'Avulso'), v.total, v.forma_pagamento, v.status
            FROM vendas v LEFT JOIN clientes c ON v.cliente_id=c.id
            WHERE v.excluido=0
            ORDER BY v.id DESC LIMIT 10
        """)
        for row in cur.fetchall():
            id_, data_, cliente, total_, forma, status = row
            tree_dashboard_vendas.insert("", "end", values=(id_, iso_para_br(data_), cliente, formatar_moeda(total_), forma, formatar_status(status)))
        conn.close()
    except Exception as e:
        print("Erro dashboard:", e)

# CLIENTES
def listar_clientes():
    conn = conectar()
    cur = conn.cursor()
    filtro = entry_busca_cliente.get().strip().lower()
    aplicar_filtro = len(filtro) >= 3
    cur.execute("SELECT id,nome,cpf_cnpj,telefone,email,cidade FROM clientes WHERE excluido=0 ORDER BY id DESC")
    dados = []
    for row in cur.fetchall():
        if aplicar_filtro and filtro not in (str(row[1]).lower() + str(row[2] or "").lower() + str(row[3] or "").lower() + str(row[5] or "").lower()):
            continue
        dados.append(tuple(row) + ("☐",))
    conn.close()
    definir_dados_paginados("clientes", tree_clientes, dados)

def salvar_cliente():
    nome = entry_cli_nome.get().strip()
    telefone = entry_cli_tel.get().strip()
    cpf = entry_cli_cpf.get().strip()
    email = entry_cli_email.get().strip()
    end = entry_cli_end.get().strip()
    numero = entry_cli_numero.get().strip()
    bairro = entry_cli_bairro.get().strip()
    cidade = entry_cli_cidade.get().strip()
    cep = entry_cli_cep.get().strip()
    
    if not validar_obrigatorio(nome, "Nome do Cliente"):
        entry_cli_nome.focus()
        return
    if not validar_obrigatorio(telefone, "Telefone do Cliente"):
        entry_cli_tel.focus()
        return
    if not validar_obrigatorio(cpf, "CPF/CNPJ do Cliente"):
        entry_cli_cpf.focus()
        return
    
    cli_id = entry_cli_id.get().strip()
    conn = conectar()
    cur = conn.cursor()
    try:
        if cli_id:
            cur.execute("""UPDATE clientes SET nome=?, cpf_cnpj=?, telefone=?, email=?, endereco=?,
                           numero=?, bairro=?, cidade=?, cep=? WHERE id=?""",
                        (nome, cpf, telefone, email, end, numero, bairro, cidade, cep, cli_id))
        else:
            cur.execute("""INSERT INTO clientes (nome, cpf_cnpj, telefone, email, endereco, numero, bairro, cidade, cep)
                           VALUES (?,?,?,?,?,?,?,?,?)""",
                        (nome, cpf, telefone, email, end, numero, bairro, cidade, cep))
        conn.commit()
        limpar_form_cliente()
        listar_clientes()
        atualizar_combos()
        mostrar_sucesso(f"Cliente '{nome}' salvo com sucesso!", "Cliente Salvo")
    except Exception as e:
        mostrar_erro(str(e))
    finally:
        conn.close()

def editar_cliente(event=None):
    sel = tree_clientes.selection()
    if not sel:
        return
    vals = tree_clientes.item(sel[0])['values']
    cli_id = vals[0]
    conn = conectar()
    cur = conn.cursor()
    cur.execute("SELECT id, nome, cpf_cnpj, telefone, email, endereco, numero, bairro, cidade, cep FROM clientes WHERE id=?", (cli_id,))
    row = cur.fetchone()
    conn.close()
    if row:
        entry_cli_id.delete(0, tk.END); entry_cli_id.insert(0, row[0])
        entry_cli_nome.delete(0, tk.END); entry_cli_nome.insert(0, row[1] or "")
        entry_cli_cpf.delete(0, tk.END); entry_cli_cpf.insert(0, row[2] or "")
        entry_cli_tel.delete(0, tk.END); entry_cli_tel.insert(0, row[3] or "")
        entry_cli_email.delete(0, tk.END); entry_cli_email.insert(0, row[4] or "")
        entry_cli_end.delete(0, tk.END); entry_cli_end.insert(0, row[5] or "")
        entry_cli_numero.delete(0, tk.END); entry_cli_numero.insert(0, row[6] or "")
        entry_cli_bairro.delete(0, tk.END); entry_cli_bairro.insert(0, row[7] or "")
        entry_cli_cidade.delete(0, tk.END); entry_cli_cidade.insert(0, row[8] or "")
        entry_cli_cep.delete(0, tk.END); entry_cli_cep.insert(0, row[9] or "")

def excluir_cliente():
    if not verificar_permissao_exclusao():
        return
    sel = tree_clientes.selection()
    if not sel:
        mostrar_aviso("Selecione um cliente para excluir!")
        return
    cli_id = tree_clientes.item(sel[0])['values'][0]
    nome = tree_clientes.item(sel[0])['values'][1]
    if not confirmar_moderno("Mover para Lixeira", f"Mover cliente para Lixeira?\n\n{nome}\n\nEle poderá ser restaurado na aba Lixeira."):
        return
    conn = conectar()
    cur = conn.cursor()
    try:
        cur.execute("UPDATE clientes SET excluido=1, data_exclusao=?, excluido_por=? WHERE id=?", (hoje_str(), usuario_logado['login'], cli_id))
        conn.commit()
        listar_clientes()
        listar_lixeira_clientes()
        atualizar_combos()
        mostrar_sucesso(f"Cliente '{nome}' movido para lixeira!", "Movido para Lixeira")
    except Exception as e:
        mostrar_erro(str(e))
    finally:
        conn.close()

def excluir_clientes_em_massa(ids):
    if not verificar_permissao_exclusao():
        return
    conn = conectar()
    cur = conn.cursor()
    try:
        for cli_id in ids:
            cur.execute("UPDATE clientes SET excluido=1, data_exclusao=?, excluido_por=? WHERE id=?",
                        (hoje_str(), usuario_logado['login'], cli_id))
        conn.commit()
        listar_clientes()
        listar_lixeira_clientes()
        atualizar_combos()
        mostrar_sucesso(f"{len(ids)} cliente(s) movido(s) para lixeira!", "Movido para Lixeira")
    except Exception as e:
        conn.rollback()
        mostrar_erro(str(e))
    finally:
        conn.close()

def limpar_form_cliente():
    entry_cli_id.delete(0, tk.END)
    entry_cli_nome.delete(0, tk.END)
    entry_cli_cpf.delete(0, tk.END)
    entry_cli_tel.delete(0, tk.END)
    entry_cli_email.delete(0, tk.END)
    entry_cli_end.delete(0, tk.END)
    entry_cli_numero.delete(0, tk.END)
    entry_cli_bairro.delete(0, tk.END)
    entry_cli_cidade.delete(0, tk.END)
    entry_cli_cep.delete(0, tk.END)

def buscar_cep_cliente():
    cep = entry_cli_cep.get().strip()
    if not cep:
        mostrar_aviso("Informe o CEP para buscar o endereço.")
        entry_cli_cep.focus()
        return
    data = buscar_cep_viacep(cep)
    if not data:
        mostrar_aviso("CEP não encontrado ou sem conexão com a internet.\nVerifique o CEP digitado.")
        return
    # Preenche endereço, bairro e cidade
    logradouro = data.get("logradouro") or ""
    bairro = data.get("bairro") or ""
    cidade = data.get("localidade") or ""
    uf = data.get("uf") or ""
    if logradouro:
        entry_cli_end.delete(0, tk.END)
        entry_cli_end.insert(0, logradouro)
    if bairro:
        entry_cli_bairro.delete(0, tk.END)
        entry_cli_bairro.insert(0, bairro)
    if cidade:
        cidade_txt = f"{cidade}/{uf}" if uf else cidade
        entry_cli_cidade.delete(0, tk.END)
        entry_cli_cidade.insert(0, cidade_txt)
    # Formata CEP
    formatar_cep(entry=entry_cli_cep)
    entry_cli_numero.focus()
    mostrar_sucesso(f"Endereço preenchido!\n{logradouro}\n{bairro} - {cidade}/{uf}", "CEP encontrado")

# FORNECEDORES
def listar_fornecedores():
    conn = conectar()
    cur = conn.cursor()
    filtro = entry_busca_forn.get().strip().lower()
    aplicar_filtro = len(filtro) >= 3
    cur.execute("SELECT id,nome,cnpj,telefone,email,cidade FROM fornecedores WHERE excluido=0 ORDER BY id DESC")
    dados = []
    for row in cur.fetchall():
        if aplicar_filtro and filtro not in (str(row[1]).lower() + str(row[2] or "").lower() + str(row[5] or "").lower()):
            continue
        dados.append(tuple(row) + ("☐",))
    conn.close()
    definir_dados_paginados("fornecedores", tree_fornecedores, dados)

def salvar_fornecedor():
    nome = entry_forn_nome.get().strip()
    cnpj = entry_forn_cnpj.get().strip()
    tel = entry_forn_tel.get().strip()
    
    if not validar_obrigatorio(nome, "Nome do Fornecedor"):
        entry_forn_nome.focus()
        return
    if not validar_obrigatorio(cnpj, "CNPJ/CPF do Fornecedor"):
        entry_forn_cnpj.focus()
        return
    if not validar_obrigatorio(tel, "Telefone do Fornecedor"):
        entry_forn_tel.focus()
        return
    
    email = entry_forn_email.get().strip()
    end = entry_forn_end.get().strip()
    numero = entry_forn_numero.get().strip()
    bairro = entry_forn_bairro.get().strip()
    cidade = entry_forn_cidade.get().strip()
    cep = entry_forn_cep.get().strip()
    forn_id = entry_forn_id.get().strip()
    conn = conectar()
    cur = conn.cursor()
    try:
        if forn_id:
            cur.execute("""UPDATE fornecedores SET nome=?, cnpj=?, telefone=?, email=?, endereco=?,
                           numero=?, bairro=?, cidade=?, cep=? WHERE id=?""",
                        (nome, cnpj, tel, email, end, numero, bairro, cidade, cep, forn_id))
        else:
            cur.execute("""INSERT INTO fornecedores (nome,cnpj,telefone,email,endereco,numero,bairro,cidade,cep)
                           VALUES (?,?,?,?,?,?,?,?,?)""",
                        (nome, cnpj, tel, email, end, numero, bairro, cidade, cep))
        conn.commit()
        limpar_form_fornecedor()
        listar_fornecedores()
        atualizar_combos()
        mostrar_sucesso(f"Fornecedor '{nome}' salvo com sucesso!", "Fornecedor Salvo")
    except Exception as e:
        mostrar_erro(str(e))
    finally:
        conn.close()

def editar_fornecedor(event=None):
    sel = tree_fornecedores.selection()
    if not sel:
        return
    forn_id = tree_fornecedores.item(sel[0])['values'][0]
    conn = conectar()
    cur = conn.cursor()
    cur.execute("SELECT id, nome, cnpj, telefone, email, endereco, numero, bairro, cidade, cep FROM fornecedores WHERE id=?", (forn_id,))
    row = cur.fetchone()
    conn.close()
    if row:
        entry_forn_id.delete(0, tk.END); entry_forn_id.insert(0, row[0])
        entry_forn_nome.delete(0, tk.END); entry_forn_nome.insert(0, row[1] or "")
        entry_forn_cnpj.delete(0, tk.END); entry_forn_cnpj.insert(0, row[2] or "")
        entry_forn_tel.delete(0, tk.END); entry_forn_tel.insert(0, row[3] or "")
        entry_forn_email.delete(0, tk.END); entry_forn_email.insert(0, row[4] or "")
        entry_forn_end.delete(0, tk.END); entry_forn_end.insert(0, row[5] or "")
        entry_forn_numero.delete(0, tk.END); entry_forn_numero.insert(0, row[6] or "")
        entry_forn_bairro.delete(0, tk.END); entry_forn_bairro.insert(0, row[7] or "")
        entry_forn_cidade.delete(0, tk.END); entry_forn_cidade.insert(0, row[8] or "")
        entry_forn_cep.delete(0, tk.END); entry_forn_cep.insert(0, row[9] or "")

def excluir_fornecedor():
    if not verificar_permissao_exclusao():
        return
    sel = tree_fornecedores.selection()
    if not sel:
        mostrar_aviso("Selecione um fornecedor!")
        return
    forn_id = tree_fornecedores.item(sel[0])['values'][0]
    nome = tree_fornecedores.item(sel[0])['values'][1]
    if not confirmar_moderno("Mover para Lixeira", f"Mover fornecedor para LIXEIRA?\n\n{nome}"):
        return
    conn = conectar()
    cur = conn.cursor()
    try:
        cur.execute("UPDATE fornecedores SET excluido=1, data_exclusao=?, excluido_por=? WHERE id=?", (hoje_str(), usuario_logado['login'], forn_id))
        conn.commit()
        listar_fornecedores()
        listar_lixeira_fornecedores()
        atualizar_combos()
        mostrar_sucesso(f"Fornecedor '{nome}' movido para lixeira!", "Movido para Lixeira")
    except Exception as e:
        mostrar_erro(f"Erro: {e}")
    finally:
        conn.close()

def excluir_fornecedores_em_massa(ids):
    if not verificar_permissao_exclusao():
        return
    conn = conectar()
    cur = conn.cursor()
    try:
        for forn_id in ids:
            cur.execute("UPDATE fornecedores SET excluido=1, data_exclusao=?, excluido_por=? WHERE id=?",
                        (hoje_str(), usuario_logado['login'], forn_id))
        conn.commit()
        listar_fornecedores()
        listar_lixeira_fornecedores()
        atualizar_combos()
        mostrar_sucesso(f"{len(ids)} fornecedor(es) movido(s) para lixeira!", "Movido para Lixeira")
    except Exception as e:
        conn.rollback()
        mostrar_erro(str(e))
    finally:
        conn.close()

def limpar_form_fornecedor():
    entry_forn_id.delete(0, tk.END)
    entry_forn_nome.delete(0, tk.END)
    entry_forn_cnpj.delete(0, tk.END)
    entry_forn_tel.delete(0, tk.END)
    entry_forn_email.delete(0, tk.END)
    entry_forn_end.delete(0, tk.END)
    entry_forn_numero.delete(0, tk.END)
    entry_forn_bairro.delete(0, tk.END)
    entry_forn_cidade.delete(0, tk.END)
    entry_forn_cep.delete(0, tk.END)

def buscar_cep_fornecedor():
    cep = entry_forn_cep.get().strip()
    if not cep:
        mostrar_aviso("Informe o CEP para buscar o endereço.")
        entry_forn_cep.focus()
        return
    data = buscar_cep_viacep(cep)
    if not data:
        mostrar_aviso("CEP não encontrado ou sem conexão com a internet.\nVerifique o CEP digitado.")
        return
    logradouro = data.get("logradouro") or ""
    bairro = data.get("bairro") or ""
    cidade = data.get("localidade") or ""
    uf = data.get("uf") or ""
    if logradouro:
        entry_forn_end.delete(0, tk.END)
        entry_forn_end.insert(0, logradouro)
    if bairro:
        entry_forn_bairro.delete(0, tk.END)
        entry_forn_bairro.insert(0, bairro)
    if cidade:
        cidade_txt = f"{cidade}/{uf}" if uf else cidade
        entry_forn_cidade.delete(0, tk.END)
        entry_forn_cidade.insert(0, cidade_txt)
    formatar_cep(entry=entry_forn_cep)
    entry_forn_numero.focus()
    mostrar_sucesso(f"Endereço preenchido!\n{logradouro}\n{bairro} - {cidade}/{uf}", "CEP encontrado")

# PRODUTOS
def listar_produtos():
    conn = conectar()
    cur = conn.cursor()
    filtro = entry_busca_prod.get().strip().lower()
    aplicar_filtro = len(filtro) >= 3
    cur.execute("""
        SELECT p.id, p.codigo, p.nome, p.preco_venda, p.estoque_atual, COALESCE(f.nome,'-')
        FROM produtos p LEFT JOIN fornecedores f ON p.fornecedor_id=f.id
        WHERE p.excluido=0
        ORDER BY p.id DESC
    """)
    dados, tags_list = [], []
    for row in cur.fetchall():
        if aplicar_filtro and filtro not in (str(row[1]).lower() + str(row[2]).lower()):
            continue
        preco_fmt = formatar_moeda(row[3])
        estoque = row[4]
        tags = ()
        if estoque <= 0:
            tags = ('zerado',)
        elif estoque <= 5:
            tags = ('baixo',)
        dados.append((row[0], row[1], row[2], preco_fmt, estoque, row[5], "☐"))
        tags_list.append(tags)
    conn.close()
    definir_dados_paginados("produtos", tree_produtos, dados, tags_list)

def salvar_produto():
    nome = entry_prod_nome.get().strip()
    codigo = entry_prod_codigo.get().strip()
    preco_venda_str = entry_prod_venda.get().strip()
    fornecedor = combo_prod_forn.get().strip()
    estoque_str = entry_prod_estoque.get().strip()
    custo_str = entry_prod_custo.get().strip()
    
    if not validar_obrigatorio(nome, "Nome do Produto"):
        entry_prod_nome.focus()
        return
    if not validar_obrigatorio(codigo, "Código do Produto"):
        entry_prod_codigo.focus()
        return
    if not validar_obrigatorio(preco_venda_str, "Preço de Venda"):
        entry_prod_venda.focus()
        return
    if not validar_obrigatorio(fornecedor, "Fornecedor do Produto"):
        combo_prod_forn.focus()
        return
    if not validar_obrigatorio(estoque_str, "Estoque Atual"):
        entry_prod_estoque.focus()
        return
    if not validar_obrigatorio(custo_str, "Preço de Custo"):
        entry_prod_custo.focus()
        return
    
    try:
        preco_custo = float(custo_str.replace(",", ".") or 0)
        preco_venda = float(preco_venda_str.replace(",", ".") or 0)
        estoque = float(estoque_str.replace(",", ".") or 0)
        estoque_min = float(entry_prod_estmin.get().replace(",", ".") or 5)
    except:
        mostrar_aviso("Valores numéricos inválidos! Verifique preço e estoque.")
        return
    
    if not validar_numero(preco_venda, "Preço de Venda", permitir_zero=False):
        return
    
    descricao = entry_prod_desc.get().strip()
    forn_id = None
    if fornecedor and " - " in fornecedor:
        try:
            forn_id = int(fornecedor.split(" - ")[0])
        except:
            pass
    prod_id = entry_prod_id.get().strip()
    
    conn = conectar()
    cur = conn.cursor()
    try:
        if prod_id:
            cur.execute("""
                UPDATE produtos SET nome=?, codigo=?, descricao=?, preco_custo=?, preco_venda=?, estoque_atual=?, estoque_minimo=?, fornecedor_id=?
                WHERE id=?
            """, (nome, codigo, descricao, preco_custo, preco_venda, estoque, estoque_min, forn_id, prod_id))
        else:
            cur.execute("""
                INSERT INTO produtos (nome,codigo,descricao,preco_custo,preco_venda,estoque_atual,estoque_minimo,fornecedor_id)
                VALUES (?,?,?,?,?,?,?,?)
            """, (nome, codigo, descricao, preco_custo, preco_venda, estoque, estoque_min, forn_id))
            novo_id = cur.lastrowid
            if estoque > 0:
                cur.execute("INSERT INTO movimentacao_estoque (produto_id,tipo,quantidade,motivo,data) VALUES (?,?,?,?,?)",
                            (novo_id, 'entrada', estoque, 'Cadastro inicial', hoje_str()))
        conn.commit()
        limpar_form_produto()
        listar_produtos()
        listar_estoque()
        atualizar_combos()
        mostrar_sucesso(f"Produto '{nome}' salvo com sucesso!", "Produto Salvo")
    except sqlite3.IntegrityError:
        mostrar_erro("Código já existe! Use outro código.")
    except Exception as e:
        mostrar_erro(str(e))
    finally:
        conn.close()

def editar_produto(event=None):
    sel = tree_produtos.selection()
    if not sel:
        return
    prod_id = tree_produtos.item(sel[0])['values'][0]
    conn = conectar()
    cur = conn.cursor()
    cur.execute("SELECT id, nome, codigo, descricao, preco_custo, preco_venda, estoque_atual, estoque_minimo, fornecedor_id FROM produtos WHERE id=?", (prod_id,))
    row = cur.fetchone()
    conn.close()
    if row:
        entry_prod_id.delete(0, tk.END); entry_prod_id.insert(0, row[0])
        entry_prod_nome.delete(0, tk.END); entry_prod_nome.insert(0, row[1])
        entry_prod_codigo.delete(0, tk.END); entry_prod_codigo.insert(0, row[2] or "")
        entry_prod_desc.delete(0, tk.END); entry_prod_desc.insert(0, row[3] or "")
        entry_prod_custo.delete(0, tk.END); entry_prod_custo.insert(0, str(row[4] or 0))
        entry_prod_venda.delete(0, tk.END); entry_prod_venda.insert(0, str(row[5] or 0))
        entry_prod_estoque.delete(0, tk.END); entry_prod_estoque.insert(0, str(row[6] or 0))
        entry_prod_estmin.delete(0, tk.END); entry_prod_estmin.insert(0, str(row[7] or 5))
        combo_prod_forn.set("")
        if row[8]:
            conn = conectar()
            cur = conn.cursor()
            cur.execute("SELECT id,nome FROM fornecedores WHERE id=?", (row[8],))
            f = cur.fetchone()
            conn.close()
            if f:
                combo_prod_forn.set(f"{f[0]} - {f[1]}")

def excluir_produto():
    if not verificar_permissao_exclusao():
        return
    sel = tree_produtos.selection()
    if not sel:
        mostrar_aviso("Selecione um produto!")
        return
    prod_id = tree_produtos.item(sel[0])['values'][0]
    nome = tree_produtos.item(sel[0])['values'][2]
    if not confirmar_moderno("Mover para Lixeira", f"Mover produto para LIXEIRA?\n\n{nome}\n\nEle sairá do estoque e vendas."):
        return
    conn = conectar()
    cur = conn.cursor()
    try:
        cur.execute("UPDATE produtos SET excluido=1, data_exclusao=?, excluido_por=? WHERE id=?", (hoje_str(), usuario_logado['login'], prod_id))
        conn.commit()
        listar_produtos()
        listar_estoque()
        listar_lixeira_produtos()
        atualizar_combos()
        mostrar_sucesso(f"Produto '{nome}' movido para lixeira!", "Movido para Lixeira")
    except Exception as e:
        mostrar_erro(f"Não foi possível: {e}")
    finally:
        conn.close()

def excluir_produtos_em_massa(ids):
    if not verificar_permissao_exclusao():
        return
    conn = conectar()
    cur = conn.cursor()
    try:
        for prod_id in ids:
            cur.execute("UPDATE produtos SET excluido=1, data_exclusao=?, excluido_por=? WHERE id=?",
                        (hoje_str(), usuario_logado['login'], prod_id))
        conn.commit()
        listar_produtos()
        listar_estoque()
        listar_lixeira_produtos()
        atualizar_combos()
        mostrar_sucesso(f"{len(ids)} produto(s) movido(s) para lixeira!", "Movido para Lixeira")
    except Exception as e:
        conn.rollback()
        mostrar_erro(f"Não foi possível: {e}")
    finally:
        conn.close()

def limpar_form_produto():
    entry_prod_id.delete(0, tk.END)
    entry_prod_nome.delete(0, tk.END)
    entry_prod_codigo.delete(0, tk.END)
    entry_prod_desc.delete(0, tk.END)
    entry_prod_custo.delete(0, tk.END)
    entry_prod_venda.delete(0, tk.END)
    entry_prod_estoque.delete(0, tk.END)
    entry_prod_estmin.delete(0, tk.END)
    entry_prod_estmin.insert(0, "5")
    combo_prod_forn.set("")

# ESTOQUE
def listar_estoque():
    conn = conectar()
    cur = conn.cursor()
    cur.execute("""
        SELECT p.id, p.codigo, p.nome, p.estoque_atual, p.estoque_minimo
        FROM produtos p WHERE p.excluido=0 ORDER BY p.id DESC
    """)
    dados, tags_list = [], []
    for row in cur.fetchall():
        status = "OK"
        tags = ()
        if row[3] <= 0:
            status = "ZERADO"
            tags = ('zerado',)
        elif row[3] <= row[4]:
            status = "BAIXO"
            tags = ('baixo',)
        dados.append((row[0], row[1], row[2], row[3], row[4], status))
        tags_list.append(tags)
    conn.close()
    definir_dados_paginados("estoque", tree_estoque, dados, tags_list)
    listar_mov_estoque()

def listar_mov_estoque():
    conn = conectar()
    cur = conn.cursor()
    prod_filtro = entry_est_filtro_prod.get().strip().lower()
    aplicar_filtro = len(prod_filtro) >= 3
    cur.execute("""
        SELECT m.data, p.nome, m.tipo, m.quantidade, m.motivo
        FROM movimentacao_estoque m JOIN produtos p ON m.produto_id=p.id
        WHERE p.excluido=0
        ORDER BY m.id DESC LIMIT 2000
    """)
    dados = []
    for row in cur.fetchall():
        if aplicar_filtro and prod_filtro not in row[1].lower():
            continue
        dados.append((iso_para_br(row[0]), row[1], row[2], row[3], row[4]))
    conn.close()
    definir_dados_paginados("mov_estoque", tree_mov_estoque, dados)

def movimentar_estoque():
    prod_text = combo_est_prod.get().strip()
    qtd_str = entry_est_qtd.get().strip()
    
    if not validar_obrigatorio(prod_text, "Produto para movimentação"):
        combo_est_prod.focus()
        return
    if not validar_obrigatorio(qtd_str, "Quantidade"):
        entry_est_qtd.focus()
        return
    
    try:
        prod_id = int(prod_text.split(" - ")[0])
        qtd = float(qtd_str.replace(",", "."))
        tipo = combo_est_tipo.get()
        motivo = entry_est_motivo.get().strip() or "Ajuste manual"
    except:
        mostrar_aviso("Dados inválidos! Verifique produto e quantidade.")
        return
    if qtd <= 0:
        mostrar_aviso("Quantidade deve ser maior que zero!")
        return
    
    conn = conectar()
    cur = conn.cursor()
    try:
        cur.execute("SELECT estoque_atual FROM produtos WHERE id=? AND excluido=0", (prod_id,))
        atual = cur.fetchone()
        if not atual:
            mostrar_erro("Produto não encontrado!")
            return
        estoque_atual = atual[0] or 0
        novo_estoque = estoque_atual
        if tipo == "entrada":
            novo_estoque += qtd
        elif tipo == "saida":
            if estoque_atual < qtd:
                if not confirmar_moderno("Estoque Baixo", f"Estoque atual {estoque_atual} menor que saída {qtd}. Continuar e ficar negativo?"):
                    return
            novo_estoque -= qtd
        elif tipo == "ajuste":
            novo_estoque = qtd
            qtd = qtd - estoque_atual
        
        cur.execute("UPDATE produtos SET estoque_atual=? WHERE id=?", (novo_estoque, prod_id))
        cur.execute("INSERT INTO movimentacao_estoque (produto_id,tipo,quantidade,motivo,data) VALUES (?,?,?,?,?)",
                    (prod_id, tipo, qtd, motivo, hoje_str()))
        conn.commit()
        listar_estoque()
        listar_produtos()
        mostrar_sucesso(f"Estoque atualizado: {estoque_atual} → {novo_estoque}", "Estoque Atualizado")
        entry_est_qtd.delete(0, tk.END)
        entry_est_motivo.delete(0, tk.END)
    except Exception as e:
        mostrar_erro(str(e))
    finally:
        conn.close()

# VENDAS
def atualizar_combos():
    conn = conectar()
    cur = conn.cursor()
    cur.execute("SELECT id,nome FROM clientes WHERE excluido=0 ORDER BY id DESC")
    clientes = [f"{r[0]} - {r[1]}" for r in cur.fetchall()]
    cur.execute("SELECT id,nome FROM fornecedores WHERE excluido=0 ORDER BY id DESC")
    fornecedores = [f"{r[0]} - {r[1]}" for r in cur.fetchall()]
    cur.execute("SELECT id,nome,preco_venda,estoque_atual,codigo FROM produtos WHERE excluido=0 ORDER BY id DESC")
    produtos = cur.fetchall()
    prod_list = [f"{r[0]} - {r[1]} ({r[4]}) | Estoque:{r[3]} | {formatar_moeda(r[2])}" for r in produtos]
    conn.close()
    
    listas_combos["clientes"] = ["0 - Consumidor Final"] + clientes
    listas_combos["fornecedores"] = fornecedores
    listas_combos["produtos"] = prod_list
    
    try:
        combo_venda_cliente['values'] = listas_combos["clientes"]
        combo_venda_prod['values'] = listas_combos["produtos"]
        combo_prod_forn['values'] = listas_combos["fornecedores"]
        combo_est_prod['values'] = listas_combos["produtos"]
        combo_cp_forn['values'] = listas_combos["fornecedores"]
        combo_cr_cliente['values'] = clientes
        
        configurar_autocomplete_combo(combo_venda_cliente, listas_combos["clientes"])
        configurar_autocomplete_combo(combo_venda_prod, listas_combos["produtos"])
        configurar_autocomplete_combo(combo_prod_forn, listas_combos["fornecedores"])
        configurar_autocomplete_combo(combo_est_prod, listas_combos["produtos"])
        configurar_autocomplete_combo(combo_cp_forn, listas_combos["fornecedores"])
        configurar_autocomplete_combo(combo_cr_cliente, clientes)
    except:
        pass

def adicionar_item_venda():
    global carrinho_venda
    prod_text = combo_venda_prod.get().strip()
    qtd_str = entry_venda_qtd.get().strip()
    
    if not validar_obrigatorio(prod_text, "Produto para venda"):
        combo_venda_prod.focus()
        return
    if not validar_obrigatorio(qtd_str, "Quantidade"):
        entry_venda_qtd.focus()
        return
    
    try:
        prod_id = int(prod_text.split(" - ")[0])
        qtd = float(qtd_str.replace(",", ".") or 1)
    except:
        mostrar_aviso("Quantidade inválida!")
        return
    if qtd <=0:
        mostrar_aviso("Quantidade deve ser maior que zero!")
        return
    conn = conectar()
    cur = conn.cursor()
    cur.execute("SELECT nome, preco_venda, estoque_atual FROM produtos WHERE id=? AND excluido=0", (prod_id,))
    row = cur.fetchone()
    conn.close()
    if not row:
        mostrar_erro("Produto não encontrado!")
        return
    nome, preco, estoque = row
    if estoque < qtd:
        if not confirmar_moderno("Estoque Baixo", f"Estoque atual: {estoque}. Vender mesmo assim?"):
            return
    subtotal = preco * qtd
    for item in carrinho_venda:
        if item['produto_id'] == prod_id:
            item['quantidade'] += qtd
            item['subtotal'] = item['quantidade'] * item['preco']
            atualizar_carrinho_tree()
            return
    carrinho_venda.append({
        'produto_id': prod_id,
        'nome': nome,
        'quantidade': qtd,
        'preco': preco,
        'subtotal': subtotal
    })
    atualizar_carrinho_tree()
    combo_venda_prod.set("")
    entry_venda_qtd.delete(0, tk.END)
    entry_venda_qtd.insert(0, "1")
    atualizar_calculo_cartao()

def _fator_taxa_carrinho():
    """Retorna (fator_exibicao, total_base, total_final, valor_extra).
    fator aplica desconto e, se cartão, a taxa — para o subtotal do carrinho."""
    total_bruto = sum(i['subtotal'] for i in carrinho_venda) if carrinho_venda else 0.0
    try:
        desc = float(str(entry_venda_desc.get()).replace(",", ".") or 0)
    except Exception:
        desc = 0.0
    if desc < 0:
        desc = 0.0
    total_base = max(0.0, total_bruto - desc)
    forma = ""
    try:
        forma = combo_venda_forma.get().strip()
    except Exception:
        pass
    # fator do desconto sobre o bruto
    if total_bruto <= 0:
        return 1.0, 0.0, 0.0, 0.0
    fator_desc = total_base / total_bruto if total_bruto else 1.0
    if forma != "Cartão Crédito":
        return fator_desc, total_base, total_base, desc
    try:
        taxa = float(str(entry_taxa.get()).replace(",", ".") or 0)
    except Exception:
        taxa = 0.0
    tipo = "Porcentagem (%)"
    try:
        tipo = combo_tipo_taxa.get().strip() or tipo
    except Exception:
        pass
    total_com_taxa, valor_taxa = calcular_total_com_taxa(total_base, taxa, tipo)
    fator = (total_com_taxa / total_bruto) if total_bruto else 1.0
    return fator, total_base, total_com_taxa, valor_taxa


def atualizar_carrinho_tree():
    limpar_tree(tree_venda_carrinho)
    total = 0
    fator, total_base, total_com_taxa, valor_taxa = _fator_taxa_carrinho()
    total_bruto = sum(float(i.get('subtotal') or 0) for i in carrinho_venda) if carrinho_venda else 0.0
    try:
        desc_global = float(str(entry_venda_desc.get()).replace(",", ".") or 0)
    except Exception:
        desc_global = 0.0
    if desc_global < 0:
        desc_global = 0.0
    for item in carrinho_venda:
        sub = float(item['subtotal'] or 0)
        preco = float(item['preco'] or 0)
        total += sub
        # desconto rateado por item
        desc_item = (sub / total_bruto * desc_global) if total_bruto > 0 and desc_global > 0 else 0.0
        # acréscimo (taxa do cartão) rateado por item, só aparece se houver taxa
        acresc_item = (sub / total_bruto * valor_taxa) if total_bruto > 0 and valor_taxa > 0 else 0.0
        # subtotal com desconto/taxa (fator); preço unitário sempre BRUTO
        sub_exib = sub * fator if total_bruto > 0 else sub
        tree_venda_carrinho.insert(
            "", "end",
            values=(
                item['produto_id'],
                item['nome'],
                item['quantidade'],
                formatar_moeda(preco),
                formatar_moeda(desc_item),
                formatar_moeda(acresc_item) if acresc_item > 0 else "-",
                formatar_moeda(sub_exib),
            ),
        )
    try:
        desc = float(str(entry_venda_desc.get()).replace(",", ".") or 0)
    except Exception:
        desc = 0
    total_final = max(0.0, total - desc)
    try:
        lbl_venda_total.config(text=formatar_moeda(total_final))
        if desc > 0:
            lbl_venda_total_sub.config(text=f"Bruto {formatar_moeda(total)} - desc. {formatar_moeda(desc)}")
        else:
            lbl_venda_total_sub.config(text="")
    except Exception:
        pass
    atualizar_calculo_cartao()


def limpar_form_venda():
    """Zera campos do PDV após finalizar venda (evita carregar dados na próxima)."""
    global carrinho_venda
    carrinho_venda = []
    try:
        limpar_tree(tree_venda_carrinho)
    except Exception:
        pass
    try:
        combo_venda_cliente.set("0 - Consumidor Final")
    except Exception:
        pass
    try:
        combo_venda_prod.set("")
    except Exception:
        pass
    try:
        entry_venda_qtd.delete(0, tk.END)
        entry_venda_qtd.insert(0, "1")
    except Exception:
        pass
    try:
        entry_venda_desc.delete(0, tk.END)
        entry_venda_desc.insert(0, "0")
    except Exception:
        pass
    try:
        entry_venda_obs.delete(0, tk.END)
    except Exception:
        pass
    try:
        combo_venda_forma.set("Dinheiro")
    except Exception:
        pass
    try:
        entry_venda_venc.delete(0, tk.END)
        entry_venda_venc.insert(0, hoje_br())
    except Exception:
        pass
    try:
        combo_parcelas.set("1")
        entry_taxa.delete(0, tk.END)
        entry_taxa.insert(0, "0")
        combo_tipo_taxa.set("Porcentagem (%)")
    except Exception:
        pass
    try:
        lbl_taxa_campo.config(text="Taxa %:")
    except Exception:
        pass
    try:
        frame_cartao_credito.pack_forget()
    except Exception:
        pass
    try:
        lbl_venda_total.config(text="R$ 0,00", fg="#15803d")
        lbl_venda_total_titulo.config(text="TOTAL DA VENDA")
        lbl_venda_total_sub.config(text="")
    except Exception:
        pass
    try:
        lbl_cartao_total_base.config(text="R$ 0,00")
        lbl_cartao_total_com_taxa.config(text="R$ 0,00")
        lbl_cartao_parcela.config(text="R$ 0,00")
        lbl_cartao_taxa_info.config(text="")
    except Exception:
        pass

def remover_item_venda():
    global carrinho_venda
    sel = tree_venda_carrinho.selection()
    if not sel:
        mostrar_aviso("Selecione um item do carrinho para remover!")
        return
    vals = tree_venda_carrinho.item(sel[0])['values']
    prod_id = vals[0]
    carrinho_venda = [i for i in carrinho_venda if i['produto_id'] != prod_id]
    atualizar_carrinho_tree()

def _parse_float_br(texto, default=0.0):
    try:
        return float(str(texto).replace("R$", "").replace(" ", "").replace(".", "").replace(",", ".") if False else str(texto).replace(",", ".") or 0)
    except Exception:
        try:
            return float(str(texto).replace(",", ".") or 0)
        except Exception:
            return default


def calcular_total_com_taxa(valor_base, taxa, tipo_taxa="Porcentagem (%)"):
    """Aplica taxa em % ou em valor (R$). Retorna (total_com_taxa, valor_taxa_aplicada)."""
    try:
        base = float(valor_base or 0)
    except Exception:
        base = 0.0
    try:
        t = float(taxa or 0)
    except Exception:
        t = 0.0
    if t < 0:
        t = 0.0
    tipo = (tipo_taxa or "Porcentagem (%)").strip()
    if tipo.startswith("Valor"):
        # taxa em reais: soma no total
        total = base + t
        return total, t
    # porcentagem
    valor_taxa = base * (t / 100.0)
    total = base + valor_taxa
    return total, valor_taxa


def atualizar_calculo_cartao():
    try:
        if 'combo_venda_forma' not in globals():
            return
        forma = combo_venda_forma.get().strip()

        total_bruto = sum(i['subtotal'] for i in carrinho_venda)
        try:
            desconto = float(str(entry_venda_desc.get()).replace(",", ".") or 0)
        except Exception:
            desconto = 0
        total_final = max(0.0, total_bruto - desconto)

        if forma != "Cartão Crédito":
            try:
                lbl_cartao_total_base.config(text="R$ 0,00")
                lbl_cartao_total_com_taxa.config(text="R$ 0,00")
                lbl_cartao_parcela.config(text="R$ 0,00")
                lbl_cartao_taxa_info.config(text="")
            except Exception:
                pass
            try:
                lbl_venda_total.config(text=formatar_moeda(total_final), fg="#15803d")
                lbl_venda_total_titulo.config(text="TOTAL DA VENDA")
                lbl_venda_total_sub.config(text="")
            except Exception:
                pass
            return

        try:
            parcelas = int(combo_parcelas.get() or 1)
        except Exception:
            parcelas = 1
        if parcelas < 1:
            parcelas = 1
        if parcelas > 12:
            parcelas = 12

        try:
            taxa = float(str(entry_taxa.get()).replace(",", ".") or 0)
        except Exception:
            taxa = 0

        tipo = "Porcentagem (%)"
        try:
            tipo = combo_tipo_taxa.get().strip() or tipo
        except Exception:
            pass

        total_com_taxa, valor_taxa = calcular_total_com_taxa(total_final, taxa, tipo)
        valor_parcela = total_com_taxa / parcelas if parcelas > 0 else total_com_taxa

        try:
            lbl_cartao_total_base.config(text=formatar_moeda(total_final))
            lbl_cartao_total_com_taxa.config(text=formatar_moeda(total_com_taxa))
            lbl_cartao_parcela.config(text=f"{parcelas}x de {formatar_moeda(valor_parcela)}")
            if tipo.startswith("Valor"):
                lbl_cartao_taxa_info.config(
                    text=f"Taxa: + {formatar_moeda(valor_taxa)} (valor fixo)  →  TOTAL DA VENDA = {formatar_moeda(total_final)} + {formatar_moeda(valor_taxa)} = {formatar_moeda(total_com_taxa)}"
                )
            else:
                lbl_cartao_taxa_info.config(
                    text=f"Taxa: {taxa:g}% = + {formatar_moeda(valor_taxa)}  →  TOTAL DA VENDA = {formatar_moeda(total_final)} + {formatar_moeda(valor_taxa)} = {formatar_moeda(total_com_taxa)}"
                )
        except Exception as e:
            print("Erro labels cartao", e)

        # Total principal da tela = valor já com a taxa somada
        try:
            lbl_venda_total.config(text=formatar_moeda(total_com_taxa), fg="#b45309")
            lbl_venda_total_titulo.config(text="TOTAL DA VENDA (com taxa)")
            lbl_venda_total_sub.config(
                text=f"Base {formatar_moeda(total_final)} + taxa {formatar_moeda(valor_taxa)}"
            )
        except Exception:
            pass
    except Exception as e:
        print("Erro calculo cartao", e)

def _fmt_num(valor):
    """Formata número para exibição em campos (3 -> '3', 3.5 -> '3.5', 1500 -> '1500')."""
    try:
        txt = f"{float(valor):.4f}".rstrip("0").rstrip(".")
        return txt if txt and txt != "-0" else "0"
    except Exception:
        return str(valor)


def abrir_modal_parcelas_cartao():
    """Modal para configurar as parcelas do Cartão de Crédito (1-12x), a taxa
    (% ou R$) e o vencimento da 1ª parcela, com preview das parcelas.

    Os botões de ação ficam em um rodapé fixo (sempre visíveis, independente do
    tamanho da tela) e há a opção de aplicar a configuração e já FINALIZAR a
    venda direto do modal.
    """
    global root, combo_parcelas, entry_taxa, entry_venda_venc, carrinho_venda, entry_venda_desc
    global lbl_cartao_total_com_taxa, lbl_cartao_parcela

    if root is None:
        return

    # Total atual do carrinho (já com desconto)
    total_bruto = sum(i['subtotal'] for i in carrinho_venda) if carrinho_venda else 0
    try:
        desconto = float(str(entry_venda_desc.get()).replace(",", ".") or 0)
    except Exception:
        desconto = 0
    total_final = max(0.0, total_bruto - desconto)

    if total_final <= 0:
        mostrar_aviso("Adicione produtos ao carrinho antes de configurar parcelas!\n\nTotal atual: R$ 0,00", "Carrinho Vazio")
        return

    # Valores atuais do PDV
    try:
        parcelas_atual = int(combo_parcelas.get() or 1)
    except Exception:
        parcelas_atual = 1
    parcelas_atual = min(12, max(1, parcelas_atual))
    try:
        taxa_atual = float(str(entry_taxa.get()).replace(",", ".") or 0)
    except Exception:
        taxa_atual = 0
    try:
        tipo_atual = combo_tipo_taxa.get().strip() or "Porcentagem (%)"
    except Exception:
        tipo_atual = "Porcentagem (%)"
    if tipo_atual not in ("Porcentagem (%)", "Valor (R$)"):
        tipo_atual = "Porcentagem (%)"

    venc_base_br = entry_venda_venc.get().strip() or hoje_br()
    if not br_para_iso(venc_base_br):
        venc_base_br = hoje_br()

    # ---------------- Janela ----------------
    modal = tk.Toplevel(root)
    modal.title("Configurar Parcelas - Cartão de Crédito (até 12x)")
    modal.configure(bg="white")
    modal.transient(root)
    modal.grab_set()
    modal.resizable(True, True)
    modal.minsize(620, 500)

    largura, altura = 720, 640
    try:
        modal.update_idletasks()
        sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
        largura = min(largura, max(620, sw - 40))
        altura = min(altura, max(500, sh - 80))
        x = root.winfo_x() + (root.winfo_width() // 2) - (largura // 2)
        y = root.winfo_y() + (root.winfo_height() // 2) - (altura // 2)
        x = max(0, min(x, sw - largura))
        y = max(0, min(y, sh - altura))
        modal.geometry(f"{largura}x{altura}+{x}+{y}")
    except Exception:
        modal.geometry(f"{largura}x{altura}")

    def fechar_modal():
        try:
            modal.grab_release()
        except Exception:
            pass
        try:
            modal.destroy()
        except Exception:
            pass

    # ---------------- Cabeçalho ----------------
    header = tk.Frame(modal, bg="#f59e0b", height=70)
    header.pack(fill='x', side='top')
    header.pack_propagate(False)
    tk.Label(header, text="Cartão de Crédito - Parcelar em até 12x", bg="#f59e0b", fg="white", font=('Arial', 13, 'bold')).pack(pady=(10, 0))
    tk.Label(header, text=f"Total da venda: {formatar_moeda(total_final)}", bg="#f59e0b", fg="white", font=('Arial', 11, 'bold')).pack()

    # ---------------- Rodapé (botões) ----------------
    # Empacotado ANTES do corpo e ancorado embaixo: assim os botões nunca são
    # "empurrados" para fora da janela, mesmo em telas pequenas.
    footer = tk.Frame(modal, bg="#f8fafc", highlightthickness=1, highlightbackground=CORES["border"])
    footer.pack(fill='x', side='bottom')

    # ---------------- Corpo ----------------
    body = tk.Frame(modal, bg="white")
    body.pack(fill='both', expand=True, side='top', padx=20, pady=(12, 6))

    # Configuração
    frame_config = tk.LabelFrame(body, text="Configuração das Parcelas", font=('Arial', 10, 'bold'), bg="white", padx=15, pady=10)
    frame_config.pack(fill='x', pady=(0, 6))

    lbl_cfg = dict(bg="white", font=('Arial', 9, 'bold'), fg=CORES["text_dark"])

    tk.Label(frame_config, text="Qtd. parcelas (1-12)*:", **lbl_cfg).grid(row=0, column=0, sticky='w', padx=(0, 6), pady=5)
    combo_modal_parcelas = ttk.Combobox(frame_config, width=6, values=[str(i) for i in range(1, 13)], font=('Arial', 11, 'bold'))
    combo_modal_parcelas.grid(row=0, column=1, sticky='w', padx=(0, 20), pady=5)
    combo_modal_parcelas.set(str(parcelas_atual))

    tk.Label(frame_config, text="1ª parcela vence em*:", **lbl_cfg).grid(row=0, column=2, sticky='w', padx=(0, 6), pady=5)
    entry_modal_venc = tk.Entry(frame_config, width=14, font=('Arial', 10))
    entry_modal_venc.grid(row=0, column=3, sticky='w', pady=5)
    entry_modal_venc.insert(0, venc_base_br)
    ativar_seletor_data(entry_modal_venc)
    tk.Label(frame_config, text="(dd/mm/aaaa)", bg="white", fg=CORES["text_gray"], font=('Arial', 8)).grid(row=0, column=4, sticky='w', padx=(6, 0))

    tk.Label(frame_config, text="Tipo da taxa*:", **lbl_cfg).grid(row=1, column=0, sticky='w', padx=(0, 6), pady=5)
    combo_modal_tipo = ttk.Combobox(frame_config, width=16, values=["Porcentagem (%)", "Valor (R$)"], state="readonly", font=('Arial', 10))
    combo_modal_tipo.grid(row=1, column=1, sticky='w', padx=(0, 20), pady=5)
    combo_modal_tipo.set(tipo_atual)

    lbl_modal_taxa = tk.Label(frame_config, text="Taxa %:", **lbl_cfg)
    lbl_modal_taxa.grid(row=1, column=2, sticky='w', padx=(0, 6), pady=5)
    entry_modal_taxa = tk.Entry(frame_config, width=10, font=('Arial', 11, 'bold'))
    entry_modal_taxa.grid(row=1, column=3, sticky='w', pady=5)
    entry_modal_taxa.insert(0, _fmt_num(taxa_atual))

    def _sync_lbl_modal_tipo(event=None):
        if combo_modal_tipo.get().startswith("Valor"):
            lbl_modal_taxa.config(text="Taxa R$:")
        else:
            lbl_modal_taxa.config(text="Taxa %:")

    _sync_lbl_modal_tipo()

    # Resumo
    frame_resumo = tk.Frame(body, bg="#fef3c7", bd=1, relief='solid')
    frame_resumo.pack(fill='x', pady=6)
    inner_resumo = tk.Frame(frame_resumo, bg="#fef3c7")
    inner_resumo.pack(fill='x', padx=10, pady=8)
    lbl_resumo_total = tk.Label(inner_resumo, text="", bg="#fef3c7", font=('Arial', 10, 'bold'), fg="#92400e", anchor='w', justify='left')
    lbl_resumo_total.pack(anchor='w', fill='x')
    lbl_resumo_parcela = tk.Label(inner_resumo, text="", bg="#fef3c7", font=('Arial', 12, 'bold'), fg=CORES["primary"], anchor='w')
    lbl_resumo_parcela.pack(anchor='w', pady=(2, 0))
    tk.Label(inner_resumo, text="As parcelas serão lançadas em Contas a Receber com vencimentos a cada +30 dias.",
             bg="#fef3c7", font=('Arial', 8, 'italic'), fg="#78350f", anchor='w').pack(anchor='w', pady=(4, 0))

    # Preview das parcelas
    frame_tree = tk.LabelFrame(body, text="Preview das Parcelas - Datas de Vencimento", font=('Arial', 10, 'bold'), bg="white")
    frame_tree.pack(fill='both', expand=True, pady=(6, 0))

    cols_preview = ("Parcela", "Vencimento", "Valor", "Status")
    tree_preview = ttk.Treeview(frame_tree, columns=cols_preview, show='headings', height=6)
    for c in cols_preview:
        tree_preview.heading(c, text=c)
    tree_preview.column("Parcela", width=80, anchor='center')
    tree_preview.column("Vencimento", width=120, anchor='center')
    tree_preview.column("Valor", width=140, anchor='center')
    tree_preview.column("Status", width=100, anchor='center')
    tree_preview.pack(fill='both', expand=True, side='left', padx=5, pady=5)
    sb_preview = ttk.Scrollbar(frame_tree, orient='vertical', command=tree_preview.yview)
    tree_preview.configure(yscrollcommand=sb_preview.set)
    sb_preview.pack(side='right', fill='y')

    def calcular_preview():
        try:
            limpar_tree(tree_preview)
            try:
                parc = int(combo_modal_parcelas.get() or 1)
            except Exception:
                parc = 1
            parc = min(12, max(1, parc))
            try:
                taxa = float(str(entry_modal_taxa.get()).replace(",", ".") or 0)
            except Exception:
                taxa = 0
            if taxa < 0:
                taxa = 0
            tipo_m = combo_modal_tipo.get().strip() or "Porcentagem (%)"

            venc_br_modal = entry_modal_venc.get().strip() or hoje_br()
            venc_iso_modal = br_para_iso(venc_br_modal) or hoje_iso()
            try:
                base_date = datetime.strptime(venc_iso_modal, "%Y-%m-%d")
            except Exception:
                base_date = datetime.now()

            total_com_taxa, valor_taxa = calcular_total_com_taxa(total_final, taxa, tipo_m)
            valor_parcela = total_com_taxa / parc if parc > 0 else total_com_taxa

            if tipo_m.startswith("Valor"):
                info_taxa = f"Taxa (valor): + {formatar_moeda(valor_taxa)}"
            else:
                info_taxa = f"Taxa {taxa:g}%: + {formatar_moeda(valor_taxa)}"
            lbl_resumo_total.config(text=f"Venda: {formatar_moeda(total_final)}   |   {info_taxa}   |   Total com taxa: {formatar_moeda(total_com_taxa)}")
            lbl_resumo_parcela.config(text=f"{parc}x de {formatar_moeda(valor_parcela)}")

            for i in range(1, parc + 1):
                venc_parcela = base_date + timedelta(days=30 * (i - 1))
                tree_preview.insert("", "end", values=(f"{i}/{parc}", venc_parcela.strftime("%d/%m/%Y"), formatar_moeda(valor_parcela), "Em Aberto"))

            return parc, taxa, venc_br_modal, total_com_taxa, valor_parcela
        except Exception as e:
            print("Erro preview", e)
            return 1, 0, hoje_br(), total_final, total_final

    def on_change_preview(event=None):
        calcular_preview()

    combo_modal_parcelas.bind("<<ComboboxSelected>>", on_change_preview)
    combo_modal_parcelas.bind("<KeyRelease>", on_change_preview)
    entry_modal_taxa.bind("<KeyRelease>", on_change_preview)
    entry_modal_venc.bind("<KeyRelease>", on_change_preview)
    entry_modal_venc.bind("<FocusOut>", on_change_preview)
    combo_modal_tipo.bind("<<ComboboxSelected>>", lambda e: (_sync_lbl_modal_tipo(), on_change_preview()))

    calcular_preview()

    # ---------------- Ações ----------------
    def _aplicar_no_pdv():
        """Valida e copia a configuração do modal para os campos do PDV.
        Retorna (parc, taxa, venc_br, total_com_taxa, valor_parcela) ou None se inválido."""
        parc, taxa, venc_br_modal, total_com_taxa, valor_parcela = calcular_preview()
        if not br_para_iso(venc_br_modal):
            mostrar_aviso("Data de vencimento inválida! Use dd/mm/aaaa", "Data Inválida")
            try:
                modal.grab_set()
            except Exception:
                pass
            return None
        tipo_m = combo_modal_tipo.get().strip() or "Porcentagem (%)"

        combo_parcelas.set(str(parc))
        entry_taxa.delete(0, tk.END)
        entry_taxa.insert(0, _fmt_num(taxa))
        try:
            combo_tipo_taxa.set(tipo_m)
        except Exception:
            pass
        try:
            lbl_taxa_campo.config(text="Taxa R$:" if tipo_m.startswith("Valor") else "Taxa %:")
        except Exception:
            pass
        entry_venda_venc.delete(0, tk.END)
        entry_venda_venc.insert(0, venc_br_modal)
        # Recalcula carrinho + totais do cartão no PDV
        try:
            atualizar_carrinho_tree()
        except Exception:
            atualizar_calculo_cartao()
        return parc, taxa, venc_br_modal, total_com_taxa, valor_parcela

    def _aplicar_e_fechar():
        """Aplica as parcelas no PDV e fecha o modal — a venda só é finalizada
        quando o usuário clicar no botão FINALIZAR VENDA da tela principal."""
        try:
            res = _aplicar_no_pdv()
            if res is None:
                return
            fechar_modal()
        except Exception as e:
            mostrar_erro(f"Erro ao aplicar parcelas: {e}")

    frame_btn = tk.Frame(footer, bg="#f8fafc")
    frame_btn.pack(fill='x', padx=20, pady=(10, 4))
    tk.Button(frame_btn, text="Cancelar", command=fechar_modal, bg="#64748b", fg="white",
              font=('Arial', 10, 'bold'), bd=0, padx=22, pady=8, cursor='hand2').pack(side='left')
    tk.Button(frame_btn, text="Aplicar Parcelas", command=_aplicar_e_fechar, bg=CORES["success"], fg="white",
              font=('Arial', 11, 'bold'), bd=0, padx=28, pady=8, cursor='hand2').pack(side='right')
    #tk.Label(footer, text="Escolha as parcelas e o acréscimo, clique em Aplicar Parcelas e depois finalize a venda pelo botão principal.", bg="#f8fafc", fg=CORES["text_gray"], font=('Arial', 8)).pack(anchor='w', padx=20, pady=(0, 8))

    modal.bind('<Escape>', lambda e: fechar_modal())
    modal.bind('<Return>', lambda e: _aplicar_e_fechar())
    modal.protocol("WM_DELETE_WINDOW", fechar_modal)
    try:
        combo_modal_parcelas.focus_set()
    except Exception:
        pass
    modal.wait_window()

def on_forma_pagamento_change(event=None):
    try:
        forma = combo_venda_forma.get().strip()
        if forma == "Cartão Crédito":
            abrir_modal_parcelas_cartao()
        atualizar_carrinho_tree()
    except Exception as e:
        print("Erro on_forma", e)

def finalizar_venda():
    global carrinho_venda
    if not carrinho_venda:
        mostrar_aviso("Carrinho vazio! Adicione produtos antes de finalizar.")
        return
    cliente_text = combo_venda_cliente.get().strip()
    if not validar_obrigatorio(cliente_text, "Cliente"):
        combo_venda_cliente.focus()
        return
    cliente_id = None
    cliente_nome = "Consumidor Final"
    if cliente_text and " - " in cliente_text:
        try:
            cid = int(cliente_text.split(" - ")[0])
            if cid !=0:
                cliente_id = cid
                cliente_nome = cliente_text.split(" - ",1)[1]
            else:
                cliente_nome = "Consumidor Final"
        except:
            pass
    forma = combo_venda_forma.get().strip()
    if not validar_obrigatorio(forma, "Forma de Pagamento"):
        combo_venda_forma.focus()
        return
    
    try:
        desconto = float(entry_venda_desc.get().replace(",", ".") or 0)
    except:
        desconto = 0
    total_bruto = sum(i['subtotal'] for i in carrinho_venda)
    total_final = total_bruto - desconto
    if total_final <0:
        total_final = 0
    
    parcelas = 1
    taxa = 0
    total_com_taxa = total_final
    tipo_taxa = "Porcentagem (%)"
    if forma == "Cartão Crédito":
        try:
            parcelas = int(combo_parcelas.get() or 1)
            if parcelas < 1 or parcelas > 12:
                mostrar_aviso("Parcelas deve ser entre 1 e 12!")
                return
        except Exception:
            mostrar_aviso("Parcelas inválidas! Escolha de 1 a 12.")
            return
        try:
            taxa = float(str(entry_taxa.get()).replace(",", ".") or 0)
        except Exception:
            taxa = 0
        try:
            tipo_taxa = combo_tipo_taxa.get().strip() or "Porcentagem (%)"
        except Exception:
            tipo_taxa = "Porcentagem (%)"
        total_com_taxa, _ = calcular_total_com_taxa(total_final, taxa, tipo_taxa)
    
    obs = entry_venda_obs.get().strip()
    venc_br = entry_venda_venc.get().strip()
    if not validar_obrigatorio(venc_br, "Data de Vencimento (1ª parcela)"):
        entry_venda_venc.focus()
        return
    venc_iso = br_para_iso(venc_br) or hoje_iso()
    if not venc_iso:
        mostrar_aviso("Data de vencimento inválida! Use dd/mm/aaaa")
        return
    
    conn = conectar()
    cur = conn.cursor()
    try:
        cur.execute("BEGIN")
        if forma in ("Dinheiro", "PIX", "Cartão Débito"):
            status_venda = "recebido"
        else:
            status_venda = "em_aberto"
        
        cur.execute("""
            INSERT INTO vendas (cliente_id, data, total, desconto, forma_pagamento, status, observacao, parcelas, taxa)
            VALUES (?,?,?,?,?,?,?,?,?)
        """, (cliente_id, hoje_str(), total_com_taxa, desconto, forma, status_venda, obs, parcelas, taxa))
        venda_id = cur.lastrowid
        
        for item in carrinho_venda:
            cur.execute("""
                INSERT INTO venda_itens (venda_id, produto_id, quantidade, preco_unitario, subtotal)
                VALUES (?,?,?,?,?)
            """, (venda_id, item['produto_id'], item['quantidade'], item['preco'], item['subtotal']))
            cur.execute("SELECT estoque_atual FROM produtos WHERE id=?", (item['produto_id'],))
            est_atual = cur.fetchone()[0] or 0
            novo_est = est_atual - item['quantidade']
            cur.execute("UPDATE produtos SET estoque_atual=? WHERE id=?", (novo_est, item['produto_id']))
            cur.execute("INSERT INTO movimentacao_estoque (produto_id,tipo,quantidade,motivo,data,referencia_id) VALUES (?,?,?,?,?,?)",
                        (item['produto_id'], 'venda', item['quantidade'], f"Venda {venda_id}", hoje_str(), venda_id))
        
        if forma in ("Dinheiro", "PIX", "Cartão Débito"):
            descricao_caixa = f"{forma} - {cliente_nome}"
            cur.execute("""
                INSERT INTO caixa (data,tipo,valor,descricao,origem,referencia_id,forma_pagamento)
                VALUES (?,?,?,?,?,?,?)
            """, (hoje_str(), 'entrada', total_com_taxa, descricao_caixa, 'venda', venda_id, forma))
        else:
            valor_parcela = total_com_taxa / parcelas
            base_date = datetime.strptime(venc_iso, "%Y-%m-%d")
            for i in range(1, parcelas+1):
                venc_parcela = base_date + timedelta(days=30*(i-1))
                venc_parcela_iso = venc_parcela.strftime("%Y-%m-%d")
                descricao = f"Cartão Crédito {parcelas}x - Parcela {i}/{parcelas}"
                cur.execute("""
                    INSERT INTO contas_a_receber (cliente_id,venda_id,descricao,valor,vencimento,status,forma_pagamento,parcela_atual,total_parcelas,data_emissao)
                    VALUES (?,?,?,?,?,?,?,?,?,?)
                """, (cliente_id, venda_id, descricao, valor_parcela, venc_parcela_iso, 'em_aberto', forma, i, parcelas, hoje_iso()))
        
        conn.commit()
        mostrar_sucesso(f"Venda {venda_id} finalizada!\n\nTotal: {formatar_moeda(total_com_taxa)}\nForma: {forma}\n{parcelas}x de {formatar_moeda(valor_parcela) if forma=='Cartão Crédito' else ''}", "Venda Finalizada")
        limpar_form_venda()
        listar_vendas()
        listar_estoque()
        listar_produtos()
        listar_caixa()
        listar_contas_receber()
        atualizar_dashboard()
    except Exception as e:
        conn.rollback()
        mostrar_erro(f"Erro ao finalizar venda: {e}")
    finally:
        conn.close()

def listar_vendas():
    # Histórico de vendas foi removido da tela de PDV; mantém função por compatibilidade
    try:
        _ = tree_vendas
    except Exception:
        return
    if tree_vendas is None:
        return
    try:
        filtro = entry_busca_venda.get().strip().lower()
    except Exception:
        filtro = ""
    aplicar_filtro = len(filtro) >= 3
    conn = conectar()
    cur = conn.cursor()
    cur.execute("""
        SELECT v.id, v.data, COALESCE(c.nome,'Avulso'), v.total, v.forma_pagamento, v.status, v.parcelas
        FROM vendas v LEFT JOIN clientes c ON v.cliente_id=c.id
        WHERE v.excluido=0
        ORDER BY v.id DESC LIMIT 2000
    """)
    dados, tags_list = [], []
    for row in cur.fetchall():
        if aplicar_filtro and filtro not in (str(row[0]).lower() + str(row[2]).lower()):
            continue
        id_, data_, cliente, total_, forma, status, parcelas = row
        forma_exib = f"{forma} {parcelas}x" if forma=="Cartão Crédito" and parcelas>1 else forma
        valores = [id_, iso_para_br(data_), cliente, formatar_moeda(total_), forma_exib, formatar_status(status)]
        tags = ()
        if status == 'cancelado':
            tags = ('cancelada',)
        elif status == 'em_atraso':
            tags = ('atraso',)
        elif status == 'em_aberto':
            tags = ('aberto',)
        elif status == 'recebido':
            tags = ('recebido',)
        dados.append(valores)
        tags_list.append(tags)
    conn.close()
    try:
        definir_dados_paginados("vendas", tree_vendas, dados, tags_list)
    except Exception:
        pass

def ver_itens_venda(event=None):
    sel = tree_vendas.selection()
    if not sel:
        return
    venda_id = tree_vendas.item(sel[0])['values'][0]
    conn = conectar()
    cur = conn.cursor()
    cur.execute("""
        SELECT p.nome, vi.quantidade, vi.preco_unitario, vi.subtotal
        FROM venda_itens vi JOIN produtos p ON vi.produto_id=p.id
        WHERE vi.venda_id=?
    """, (venda_id,))
    itens = cur.fetchall()
    cur.execute("SELECT total, forma_pagamento, parcelas, taxa, status FROM vendas WHERE id=?", (venda_id,))
    venda_info = cur.fetchone()
    conn.close()
    texto = f"Itens da Venda {venda_id}:\n\n"
    for it in itens:
        texto += f"- {it[0]} | Qtd: {it[1]} x {formatar_moeda(it[2])} = {formatar_moeda(it[3])}\n"
    if venda_info:
        texto += f"\nTotal: {formatar_moeda(venda_info[0])}\nForma: {venda_info[1]} {venda_info[2]}x\nTaxa: {venda_info[3]}%\nStatus: {formatar_status(venda_info[4])}"
    mostrar_info(texto, f"Venda {venda_id}")

def cancelar_venda():
    sel = tree_vendas.selection()
    if not sel:
        mostrar_aviso("Selecione uma venda para cancelar!")
        return
    vals = tree_vendas.item(sel[0])['values']
    venda_id = vals[0]
    status = vals[5]
    if status == 'cancelado':
        mostrar_aviso("Venda já está cancelada!")
        return
    if not confirmar_moderno("Cancelar Venda", f"Cancelar venda {venda_id}?\n\nIsso irá:\n- Estornar estoque\n- Baixar do caixa (se à vista)\n- Cancelar contas a receber vinculadas\n\nAtualiza em todos os módulos."):
        return
    conn = conectar()
    cur = conn.cursor()
    try:
        cur.execute("BEGIN")
        cur.execute("SELECT total, forma_pagamento FROM vendas WHERE id=? AND excluido=0", (venda_id,))
        venda = cur.fetchone()
        if not venda:
            mostrar_erro("Venda não encontrada!")
            return
        total, forma = venda
        
        cur.execute("SELECT produto_id, quantidade FROM venda_itens WHERE venda_id=?", (venda_id,))
        for prod_id, qtd in cur.fetchall():
            cur.execute("SELECT estoque_atual FROM produtos WHERE id=?", (prod_id,))
            est = cur.fetchone()
            if est:
                novo = (est[0] or 0) + qtd
                cur.execute("UPDATE produtos SET estoque_atual=? WHERE id=?", (novo, prod_id))
                cur.execute("INSERT INTO movimentacao_estoque (produto_id,tipo,quantidade,motivo,data,referencia_id) VALUES (?,?,?,?,?,?)",
                            (prod_id, 'estorno_venda', qtd, f"Cancelamento venda {venda_id}", hoje_str(), venda_id))
        
        if forma in ("Dinheiro", "PIX", "Cartão Débito"):
            cur.execute("SELECT id FROM caixa WHERE origem='venda' AND referencia_id=? AND tipo='entrada' AND excluido=0", (venda_id,))
            caixa_entry = cur.fetchone()
            if caixa_entry:
                cur.execute("""
                    INSERT INTO caixa (data,tipo,valor,descricao,origem,referencia_id,forma_pagamento)
                    VALUES (?,?,?,?,?,?,?)
                """, (hoje_str(), 'saida', total, f"ESTORNO Venda {venda_id} cancelada", 'estorno_venda', venda_id, forma))
        
        cur.execute("UPDATE contas_a_receber SET status='cancelado' WHERE venda_id=? AND excluido=0", (venda_id,))
        cur.execute("UPDATE vendas SET status='cancelado' WHERE id=?", (venda_id,))
        
        conn.commit()
        mostrar_sucesso(f"Venda {venda_id} cancelada!\n\nEstoque estornado e contas atualizadas em todos os módulos.", "Venda Cancelada")
        listar_vendas()
        listar_estoque()
        listar_produtos()
        listar_caixa()
        listar_contas_receber()
        atualizar_dashboard()
    except Exception as e:
        conn.rollback()
        mostrar_erro(f"Erro ao cancelar: {e}")
    finally:
        conn.close()

def excluir_venda_lixeira():
    if not verificar_permissao_exclusao():
        return
    sel = tree_vendas.selection()
    if not sel:
        mostrar_aviso("Selecione uma venda!")
        return
    venda_id = tree_vendas.item(sel[0])['values'][0]
    if not confirmar_moderno("Mover para Lixeira", f"Mover venda {venda_id} para lixeira?\n\nA venda precisa estar cancelada antes."):
        return
    conn = conectar()
    cur = conn.cursor()
    cur.execute("SELECT status FROM vendas WHERE id=?", (venda_id,))
    status = cur.fetchone()
    if status and status[0] != 'cancelado':
        mostrar_aviso("Só é possível excluir vendas CANCELADAS. Cancele primeiro.")
        conn.close()
        return
    try:
        cur.execute("UPDATE vendas SET excluido=1, data_exclusao=?, excluido_por=? WHERE id=?", (hoje_str(), usuario_logado['login'], venda_id))
        conn.commit()
        listar_vendas()
        listar_lixeira_vendas()
        mostrar_sucesso(f"Venda {venda_id} movida para lixeira!", "Movido para Lixeira")
    except Exception as e:
        mostrar_erro(str(e))
    finally:
        conn.close()

# CAIXA
def listar_caixa():
    conn = conectar()
    cur = conn.cursor()
    cur.execute("""
        SELECT cx.id, cx.data, cx.tipo, cx.valor, cx.descricao, cx.origem, cx.forma_pagamento,
               CASE
                 WHEN cx.origem = 'venda' THEN COALESCE((
                     SELECT COALESCE(c.nome, 'Consumidor Final')
                     FROM vendas v LEFT JOIN clientes c ON v.cliente_id = c.id
                     WHERE v.id = cx.referencia_id
                 ), 'Consumidor Final')
                 WHEN cx.origem IN ('conta_receber', 'conta a receber') THEN COALESCE((
                     SELECT COALESCE(c.nome, 'Consumidor Final')
                     FROM contas_a_receber cr LEFT JOIN clientes c ON cr.cliente_id = c.id
                     WHERE cr.id = cx.referencia_id
                 ), '-')
                 WHEN cx.origem IN ('conta_pagar', 'conta a pagar') THEN COALESCE((
                     SELECT COALESCE(f.nome, '-')
                     FROM contas_a_pagar cp LEFT JOIN fornecedores f ON cp.fornecedor_id = f.id
                     WHERE cp.id = cx.referencia_id
                 ), '-')
                 ELSE '-'
               END AS pessoa
        FROM caixa cx
        WHERE cx.excluido = 0
        ORDER BY cx.id DESC
        LIMIT 5000
    """)
    entradas = 0
    saidas = 0
    cur2 = conn.cursor()
    cur2.execute("SELECT COALESCE(SUM(CASE WHEN tipo='entrada' THEN valor ELSE -valor END),0) FROM caixa WHERE excluido=0")
    saldo_total = cur2.fetchone()[0] or 0

    dados, tags_list = [], []
    for row in cur.fetchall():
        valor_fmt = formatar_moeda(row[3])
        if row[2] == 'entrada':
            entradas += row[3] or 0
            tags = ('entrada',)
        else:
            saidas += row[3] or 0
            tags = ('saida',)
        pessoa = row[7] or "-"
        dados.append((
            row[0],
            iso_para_br_data(row[1]) if row[1] else "-",
            row[2],
            valor_fmt,
            row[4] or "-",
            pessoa,
            row[5] or "-",
            row[6] or "-",
            "☐",
        ))
        tags_list.append(tags)
    conn.close()
    definir_dados_paginados("caixa", tree_caixa, dados, tags_list)
    try:
        lbl_caixa_saldo.config(text=formatar_moeda(saldo_total))
        lbl_caixa_info.config(text=f"Total entradas: {formatar_moeda(entradas)} | Total saídas: {formatar_moeda(saidas)}")
    except Exception:
        pass


def lancar_caixa_manual():
    valor_str = entry_caixa_valor.get().strip()
    desc = entry_caixa_desc.get().strip()
    tipo = combo_caixa_tipo.get().strip()
    
    if not validar_obrigatorio(tipo, "Tipo (entrada/saída)"):
        return
    if not validar_obrigatorio(valor_str, "Valor"):
        entry_caixa_valor.focus()
        return
    if not validar_obrigatorio(desc, "Descrição"):
        entry_caixa_desc.focus()
        return
    
    try:
        valor = float(valor_str.replace(",", "."))
    except:
        mostrar_aviso("Valor inválido!")
        return
    forma = combo_caixa_forma.get().strip() or "Manual"
    conn = conectar()
    cur = conn.cursor()
    try:
        cur.execute("INSERT INTO caixa (data,tipo,valor,descricao,origem,forma_pagamento) VALUES (?,?,?,?,?,?)",
                    (hoje_str(), tipo, valor, desc, 'manual', forma))
        conn.commit()
        listar_caixa()
        atualizar_dashboard()
        entry_caixa_valor.delete(0, tk.END)
        entry_caixa_desc.delete(0, tk.END)
        mostrar_sucesso("Lançamento no caixa realizado!", "Caixa")
    except Exception as e:
        mostrar_erro(str(e))
    finally:
        conn.close()

def excluir_lancamento_caixa():
    if not verificar_permissao_exclusao():
        return
    sel = tree_caixa.selection()
    if not sel:
        mostrar_aviso("Selecione um lançamento!")
        return
    caixa_id = tree_caixa.item(sel[0])['values'][0]
    if not confirmar_moderno("Mover para Lixeira", "Mover lançamento para LIXEIRA?"):
        return
    conn = conectar()
    cur = conn.cursor()
    try:
        cur.execute("UPDATE caixa SET excluido=1, data_exclusao=?, excluido_por=? WHERE id=?", (hoje_str(), usuario_logado['login'], caixa_id))
        conn.commit()
        listar_caixa()
        listar_lixeira_caixa()
        atualizar_dashboard()
        mostrar_sucesso("Lançamento movido para lixeira!", "Lixeira")
    except Exception as e:
        mostrar_erro(str(e))
    finally:
        conn.close()

def excluir_lancamentos_caixa_em_massa(ids):
    if not verificar_permissao_exclusao():
        return
    conn = conectar()
    cur = conn.cursor()
    try:
        for caixa_id in ids:
            cur.execute("UPDATE caixa SET excluido=1, data_exclusao=?, excluido_por=? WHERE id=?",
                        (hoje_str(), usuario_logado['login'], caixa_id))
        conn.commit()
        listar_caixa()
        listar_lixeira_caixa()
        atualizar_dashboard()
        mostrar_sucesso(f"{len(ids)} lançamento(s) movido(s) para lixeira!", "Lixeira")
    except Exception as e:
        conn.rollback()
        mostrar_erro(str(e))
    finally:
        conn.close()

# CONTAS A PAGAR
def listar_contas_pagar():
    atualizar_status_atraso()
    conn = conectar()
    cur = conn.cursor()
    
    cur.execute("SELECT COUNT(*), COALESCE(SUM(valor),0) FROM contas_a_pagar WHERE status='em_aberto' AND excluido=0")
    aberto_qtd, aberto_val = cur.fetchone()
    cur.execute("SELECT COUNT(*), COALESCE(SUM(valor),0) FROM contas_a_pagar WHERE status='pago' AND excluido=0")
    pago_qtd, pago_val = cur.fetchone()
    cur.execute("SELECT COUNT(*), COALESCE(SUM(valor),0) FROM contas_a_pagar WHERE status='em_atraso' AND excluido=0")
    atraso_qtd, atraso_val = cur.fetchone()
    cur.execute("SELECT COUNT(*), COALESCE(SUM(valor),0) FROM contas_a_pagar WHERE status='cancelado' AND excluido=0")
    canc_qtd, canc_val = cur.fetchone()
    
    try:
        lbl_cp_aberto.config(text=f"Em Aberto\n{aberto_qtd} contas\n{formatar_moeda(aberto_val)}")
        lbl_cp_pago.config(text=f"Pago\n{pago_qtd} contas\n{formatar_moeda(pago_val)}")
        lbl_cp_atraso.config(text=f"Em Atraso\n{atraso_qtd} contas\n{formatar_moeda(atraso_val)}")
        lbl_cp_cancelado.config(text=f"Cancelado\n{canc_qtd} contas\n{formatar_moeda(canc_val)}")
    except:
        pass
    
    try:
        cur.execute("""
            SELECT cp.id, COALESCE(f.nome,'-'), cp.descricao, cp.valor, cp.vencimento, cp.status
            FROM contas_a_pagar cp LEFT JOIN fornecedores f ON cp.fornecedor_id=f.id
            WHERE cp.excluido=0
            ORDER BY cp.id DESC
        """)
        dados, tags_list = [], []
        for row in cur.fetchall():
            vals = (row[0], row[1], row[2], formatar_moeda(row[3]), iso_para_br(row[4]), formatar_status(row[5]), "☐")
            tags = ()
            if row[5] == 'em_aberto':
                tags = ('aberto',)
            elif row[5] == 'pago':
                tags = ('pago',)
            elif row[5] == 'em_atraso':
                tags = ('atraso',)
            elif row[5] == 'cancelado':
                tags = ('cancelado',)
            dados.append(vals)
            tags_list.append(tags)
        definir_dados_paginados("cp_todos", tree_cp, dados, tags_list)
        for status, tree, chave in [
            ("em_aberto", tree_cp_aberto, "cp_aberto"),
            ("pago", tree_cp_pago, "cp_pago"),
            ("em_atraso", tree_cp_atraso, "cp_atraso"),
            ("cancelado", tree_cp_cancelado, "cp_cancelado"),
        ]:
            try:
                cur.execute("""
                    SELECT cp.id, COALESCE(f.nome,'-'), cp.descricao, cp.valor, cp.vencimento, cp.status
                    FROM contas_a_pagar cp LEFT JOIN fornecedores f ON cp.fornecedor_id=f.id
                    WHERE cp.status=? AND cp.excluido=0
                    ORDER BY cp.id DESC
                """, (status,))
                d = [(row[0], row[1], row[2], formatar_moeda(row[3]), iso_para_br(row[4]), formatar_status(row[5]), "☐") for row in cur.fetchall()]
                definir_dados_paginados(chave, tree, d)
            except Exception:
                pass
    except Exception as e:
        print("Erro listar cp", e)
    conn.close()


def _conta_ja_existe_por_documento(tabela, numero_documento, ignorar_id=None):
    """Retorna True se já existir conta ativa com o mesmo número de documento."""
    if not numero_documento or not str(numero_documento).strip():
        return False
    doc = str(numero_documento).strip()
    conn = conectar()
    cur = conn.cursor()
    try:
        if ignorar_id:
            cur.execute(
                f"SELECT id FROM {tabela} WHERE excluido=0 AND numero_documento=? AND id<>? LIMIT 1",
                (doc, ignorar_id),
            )
        else:
            cur.execute(
                f"SELECT id FROM {tabela} WHERE excluido=0 AND numero_documento=? LIMIT 1",
                (doc,),
            )
        row = cur.fetchone()
        return row is not None
    except Exception:
        return False
    finally:
        conn.close()


def _conta_id_ja_existe(tabela, conta_id):
    """Retorna True se o ID já existir na tabela."""
    if conta_id is None or str(conta_id).strip() == "":
        return False
    try:
        cid = int(str(conta_id).strip())
    except Exception:
        return False
    conn = conectar()
    cur = conn.cursor()
    try:
        cur.execute(f"SELECT id FROM {tabela} WHERE id=? LIMIT 1", (cid,))
        return cur.fetchone() is not None
    except Exception:
        return False
    finally:
        conn.close()

def salvar_conta_pagar():
    forn_text = combo_cp_forn.get().strip()
    desc = entry_cp_desc.get().strip()
    valor_str = entry_cp_valor.get().strip()
    venc_br = entry_cp_venc.get().strip()
    
    if not validar_obrigatorio(forn_text, "Fornecedor"):
        combo_cp_forn.focus()
        return
    if not validar_obrigatorio(desc, "Descrição da Conta"):
        entry_cp_desc.focus()
        return
    if not validar_obrigatorio(valor_str, "Valor da Conta"):
        entry_cp_valor.focus()
        return
    if not validar_obrigatorio(venc_br, "Vencimento"):
        entry_cp_venc.focus()
        return
    
    forn_id = None
    if forn_text and " - " in forn_text:
        try:
            forn_id = int(forn_text.split(" - ")[0])
        except:
            pass
    
    try:
        valor = float(valor_str.replace(",", "."))
    except:
        mostrar_aviso("Valor inválido!")
        return
    venc_iso = br_para_iso(venc_br) or hoje_iso()
    if not venc_iso:
        mostrar_aviso("Data vencimento inválida! Use dd/mm/aaaa")
        return
    
    forma = combo_cp_forma.get().strip() or "Boleto"
    parcelas = 1
    taxa = 0
    if forma == "Cartão Crédito":
        try:
            parcelas = int(combo_cp_parcelas.get() or 1)
        except:
            parcelas = 1
        try:
            taxa = float(entry_cp_taxa.get().replace(",", ".") or 0)
        except:
            taxa = 0
        valor = valor * (1 + taxa/100)
    
    cp_id = entry_cp_id.get().strip()
    
    conn = conectar()
    cur = conn.cursor()
    try:
        if cp_id:
            if not _conta_id_ja_existe("contas_a_pagar", cp_id):
                mostrar_aviso(f"Não existe conta a pagar com ID {cp_id}.")
                return
            cur.execute("UPDATE contas_a_pagar SET fornecedor_id=?, descricao=?, valor=?, vencimento=?, forma_pagamento=? WHERE id=?",
                        (forn_id, desc, valor, venc_iso, forma, cp_id))
        else:
            if forma == "Cartão Crédito" and parcelas > 1:
                valor_parcela = valor / parcelas
                base_date = datetime.strptime(venc_iso, "%Y-%m-%d")
                for i in range(1, parcelas+1):
                    venc_parcela = base_date + timedelta(days=30*(i-1))
                    venc_parcela_iso = venc_parcela.strftime("%Y-%m-%d")
                    desc_parcela = f"{desc} - Parcela {i}/{parcelas}"
                    cur.execute("INSERT INTO contas_a_pagar (fornecedor_id, descricao, valor, vencimento, status, forma_pagamento, parcela_atual, total_parcelas) VALUES (?,?,?,?,?,?,?,?)",
                                (forn_id, desc_parcela, valor_parcela, venc_parcela_iso, 'em_aberto', forma, i, parcelas))
            else:
                cur.execute("INSERT INTO contas_a_pagar (fornecedor_id, descricao, valor, vencimento, status, forma_pagamento) VALUES (?,?,?,?,?,?)",
                            (forn_id, desc, valor, venc_iso, 'em_aberto', forma))
        conn.commit()
        limpar_form_cp()
        listar_contas_pagar()
        atualizar_dashboard()
        mostrar_sucesso(f"Conta a pagar '{desc}' salva!", "Conta Salva")
    except Exception as e:
        mostrar_erro(str(e))
    finally:
        conn.close()

def editar_cp(event=None):
    """Duplo clique: mostra detalhes da conta a pagar."""
    tree_ativo = tree_cp
    try:
        tab_id = notebook_cp.select()
        tab_frame = notebook_cp.nametowidget(tab_id)
        for child in tab_frame.winfo_children():
            if isinstance(child, ttk.Treeview):
                tree_ativo = child
                break
            for subchild in child.winfo_children():
                if isinstance(subchild, ttk.Treeview):
                    tree_ativo = subchild
                    break
    except Exception:
        pass
    sel = tree_ativo.selection()
    if not sel:
        return
    vals = tree_ativo.item(sel[0])['values']
    mostrar_info(
        f"ID: {vals[0]}\nFornecedor: {vals[1]}\nDescrição: {vals[2]}\nValor: {vals[3]}\nVencimento: {vals[4]}\nStatus: {vals[5]}",
        f"Conta a Pagar #{vals[0]}"
    )

def pagar_conta():
    """Paga uma ou várias contas a pagar marcadas no ☐ (ou a linha selecionada)."""
    tree_ativo = _tree_ativa_notebook(notebook_cp, tree_cp)

    ids = obter_ids_selecionados(tree_ativo)
    if not ids:
        sel = tree_ativo.selection()
        if not sel:
            mostrar_aviso("Marque no ☐ uma ou mais contas em aberto para pagar.")
            return
        ids = [tree_ativo.item(s)['values'][0] for s in sel]

    conn = conectar()
    cur = conn.cursor()
    contas = []
    for cp_id in ids:
        cur.execute(
            "SELECT id, valor, status, descricao, forma_pagamento FROM contas_a_pagar WHERE id=? AND excluido=0",
            (cp_id,),
        )
        row = cur.fetchone()
        if not row:
            continue
        cid, valor, status, desc, forma = row
        if status in ('pago', 'cancelado'):
            continue
        contas.append({
            "id": cid,
            "valor": float(valor or 0),
            "desc": desc or "",
            "forma": forma or "Boleto",
            "status": status,
        })

    if not contas:
        conn.close()
        mostrar_aviso("Nenhuma conta válida para pagar.\n(Marque contas Em Aberto / Em Atraso.)")
        return

    total = sum(c["valor"] for c in contas)
    if not confirmar_contas_lote(
        "Confirmar Pagamento",
        contas,
        total,
        acao_label="Pagar contas",
    ):
        conn.close()
        return

    try:
        cur.execute("BEGIN")
        for c in contas:
            cur.execute(
                "UPDATE contas_a_pagar SET status='pago', data_pagamento=? WHERE id=?",
                (hoje_str(), c["id"]),
            )
            cur.execute(
                "INSERT INTO caixa (data,tipo,valor,descricao,origem,referencia_id,forma_pagamento) VALUES (?,?,?,?,?,?,?)",
                (hoje_str(), 'saida', c["valor"], f"Pagamento: {c['desc']}", 'conta_pagar', c["id"], c["forma"]),
            )
        conn.commit()
        mostrar_sucesso(
            f"{len(contas)} conta(s) paga(s)!\n\nTotal: {formatar_moeda(total)}\nBaixa(s) no caixa realizada(s).",
            "Contas Pagas",
        )
        try:
            limpar_selecao_multipla(tree_ativo)
        except Exception:
            pass
        listar_contas_pagar()
        listar_caixa()
        atualizar_dashboard()
    except Exception as e:
        conn.rollback()
        mostrar_erro(str(e))
    finally:
        conn.close()

def _tree_ativa_notebook(notebook, tree_fallback):
    """Retorna a Treeview da aba ativa do notebook (ou fallback)."""
    tree_ativo = tree_fallback
    try:
        tab_id = notebook.select()
        tab_frame = notebook.nametowidget(tab_id)
        for child in tab_frame.winfo_children():
            if isinstance(child, ttk.Treeview):
                return child
            for subchild in child.winfo_children():
                if isinstance(subchild, ttk.Treeview):
                    return subchild
    except Exception:
        pass
    return tree_ativo


def excluir_cp_selecionados():
    """Exclui (move para lixeira) as contas marcadas com ☐ na aba ativa."""
    tree_ativo = _tree_ativa_notebook(notebook_cp, tree_cp)
    ids = obter_ids_selecionados(tree_ativo)
    if not ids:
        mostrar_aviso("Marque ao menos uma conta (clique no ☐) para excluir.")
        return
    if not confirmar_moderno("Excluir selecionados", f"Mover {len(ids)} conta(s) a pagar para a lixeira?"):
        return
    excluir_contas_pagar_em_massa(ids)


def excluir_cr_selecionados():
    """Exclui (move para lixeira) as contas marcadas com ☐ na aba ativa."""
    tree_ativo = _tree_ativa_notebook(notebook_cr, tree_cr)
    ids = obter_ids_selecionados(tree_ativo)
    if not ids:
        mostrar_aviso("Marque ao menos uma conta (clique no ☐) para excluir.")
        return
    if not confirmar_moderno("Excluir selecionados", f"Mover {len(ids)} conta(s) a receber para a lixeira?"):
        return
    excluir_contas_receber_em_massa(ids)


def excluir_cp():
    if not verificar_permissao_exclusao():
        return
    tree_ativo = tree_cp
    try:
        tab_id = notebook_cp.select()
        tab_frame = notebook_cp.nametowidget(tab_id)
        for child in tab_frame.winfo_children():
            if isinstance(child, ttk.Treeview):
                tree_ativo = child
                break
            for subchild in child.winfo_children():
                if isinstance(subchild, ttk.Treeview):
                    tree_ativo = subchild
                    break
    except:
        pass
    sel = tree_ativo.selection()
    if not sel:
        mostrar_aviso("Selecione uma conta!")
        return
    cp_id = tree_ativo.item(sel[0])['values'][0]
    if not confirmar_moderno("Mover para Lixeira", "Mover conta a pagar para LIXEIRA?"):
        return
    conn = conectar()
    cur = conn.cursor()
    try:
        cur.execute("UPDATE contas_a_pagar SET excluido=1, data_exclusao=?, excluido_por=? WHERE id=?", (hoje_str(), usuario_logado['login'], cp_id))
        conn.commit()
        listar_contas_pagar()
        listar_lixeira_cp()
        atualizar_dashboard()
        mostrar_sucesso("Conta movida para lixeira!", "Lixeira")
    except Exception as e:
        mostrar_erro(str(e))
    finally:
        conn.close()

def excluir_contas_pagar_em_massa(ids):
    if not verificar_permissao_exclusao():
        return
    conn = conectar()
    cur = conn.cursor()
    try:
        for cp_id in ids:
            cur.execute("UPDATE contas_a_pagar SET excluido=1, data_exclusao=?, excluido_por=? WHERE id=?",
                        (hoje_str(), usuario_logado['login'], cp_id))
        conn.commit()
        listar_contas_pagar()
        listar_lixeira_cp()
        atualizar_dashboard()
        mostrar_sucesso(f"{len(ids)} conta(s) movida(s) para lixeira!", "Lixeira")
    except Exception as e:
        conn.rollback()
        mostrar_erro(str(e))
    finally:
        conn.close()

def cancelar_cp():
    tree_ativo = tree_cp
    try:
        tab_id = notebook_cp.select()
        tab_frame = notebook_cp.nametowidget(tab_id)
        for child in tab_frame.winfo_children():
            if isinstance(child, ttk.Treeview):
                tree_ativo = child
                break
            for subchild in child.winfo_children():
                if isinstance(subchild, ttk.Treeview):
                    tree_ativo = subchild
                    break
    except:
        pass
    sel = tree_ativo.selection()
    if not sel:
        mostrar_aviso("Selecione uma conta para cancelar!")
        return
    cp_id = tree_ativo.item(sel[0])['values'][0]
    if not confirmar_moderno("Cancelar Conta", f"Cancelar conta {cp_id}?\n\nO status será cancelado e atualizará em todos os módulos."):
        return
    conn = conectar()
    cur = conn.cursor()
    try:
        cur.execute("UPDATE contas_a_pagar SET status='cancelado' WHERE id=? AND excluido=0", (cp_id,))
        conn.commit()
        listar_contas_pagar()
        atualizar_dashboard()
        mostrar_sucesso("Conta cancelada! Atualizado em todos os módulos.", "Cancelado")
    except Exception as e:
        mostrar_erro(str(e))
    finally:
        conn.close()

def limpar_form_cp():
    pass

# CONTAS A RECEBER

def limpar_descricao_conta(desc):
    """Remove 'Consumidor Final' da descrição (já aparece na coluna Cliente)."""
    d = str(desc or "").strip()
    for trecho in (
        " - Consumidor Final",
        " - Consumidor final",
        " – Consumidor Final",
        " — Consumidor Final",
        "Consumidor Final",
        "Consumidor final",
    ):
        d = d.replace(trecho, "")
    d = d.strip(" -–—")
    return d if d else "-"

def listar_contas_receber():
    atualizar_status_atraso()
    conn = conectar()
    cur = conn.cursor()
    
    cur.execute("SELECT COUNT(*), COALESCE(SUM(valor),0) FROM contas_a_receber WHERE status='em_aberto' AND excluido=0")
    aberto_qtd, aberto_val = cur.fetchone()
    cur.execute("SELECT COUNT(*), COALESCE(SUM(valor),0) FROM contas_a_receber WHERE status='recebido' AND excluido=0")
    rec_qtd, rec_val = cur.fetchone()
    cur.execute("SELECT COUNT(*), COALESCE(SUM(valor),0) FROM contas_a_receber WHERE status='em_atraso' AND excluido=0")
    atraso_qtd, atraso_val = cur.fetchone()
    cur.execute("SELECT COUNT(*), COALESCE(SUM(valor),0) FROM contas_a_receber WHERE status='cancelado' AND excluido=0")
    canc_qtd, canc_val = cur.fetchone()
    
    try:
        lbl_cr_aberto.config(text=f"Em Aberto\n{aberto_qtd} contas\n{formatar_moeda(aberto_val)}")
        lbl_cr_recebido.config(text=f"Recebido\n{rec_qtd} contas\n{formatar_moeda(rec_val)}")
        lbl_cr_atraso.config(text=f"Em Atraso\n{atraso_qtd} contas\n{formatar_moeda(atraso_val)}")
        lbl_cr_cancelado.config(text=f"Cancelado\n{canc_qtd} contas\n{formatar_moeda(canc_val)}")
    except:
        pass
    
    try:
        cur.execute("""
            SELECT cr.id,
                   COALESCE(c.nome, CASE WHEN cr.cliente_id IS NULL OR cr.cliente_id=0 THEN 'Consumidor Final' ELSE '-' END),
                   cr.descricao, cr.valor,
                   COALESCE(v.data, cr.data_emissao),
                   cr.vencimento, cr.status, cr.parcela_atual, cr.total_parcelas
            FROM contas_a_receber cr
            LEFT JOIN clientes c ON cr.cliente_id=c.id
            LEFT JOIN vendas v ON cr.venda_id=v.id
            WHERE cr.excluido=0
            ORDER BY cr.id DESC
        """)
        dados, tags_list = [], []
        for row in cur.fetchall():
            parcela_info = f"{row[7]}/{row[8]}" if row[8] and row[8]>1 else "-"
            cliente_nome = row[1] or "Consumidor Final"
            desc_limpa = limpar_descricao_conta(row[2])
            vals = (row[0], cliente_nome, desc_limpa, formatar_moeda(row[3]), iso_para_br_data(row[4]) if row[4] else "-", iso_para_br(row[5]), formatar_status(row[6]), parcela_info, "☐")
            tags = ()
            if row[6] == 'em_aberto':
                tags = ('aberto',)
            elif row[6] == 'recebido':
                tags = ('pago',)
            elif row[6] == 'em_atraso':
                tags = ('atraso',)
            elif row[6] == 'cancelado':
                tags = ('cancelado',)
            dados.append(vals)
            tags_list.append(tags)
        definir_dados_paginados("cr_todos", tree_cr, dados, tags_list)
        for status, tree, chave in [
            ("em_aberto", tree_cr_aberto, "cr_aberto"),
            ("recebido", tree_cr_recebido, "cr_recebido"),
            ("em_atraso", tree_cr_atraso, "cr_atraso"),
            ("cancelado", tree_cr_cancelado, "cr_cancelado"),
        ]:
            try:
                cur.execute("""
                    SELECT cr.id,
                           COALESCE(c.nome, CASE WHEN cr.cliente_id IS NULL OR cr.cliente_id=0 THEN 'Consumidor Final' ELSE '-' END),
                           cr.descricao, cr.valor,
                           COALESCE(v.data, cr.data_emissao),
                           cr.vencimento, cr.status, cr.parcela_atual, cr.total_parcelas
                    FROM contas_a_receber cr
                    LEFT JOIN clientes c ON cr.cliente_id=c.id
                    LEFT JOIN vendas v ON cr.venda_id=v.id
                    WHERE cr.status=? AND cr.excluido=0
                    ORDER BY cr.id DESC
                """, (status,))
                d = []
                for row in cur.fetchall():
                    parcela_info = f"{row[7]}/{row[8]}" if row[8] and row[8]>1 else "-"
                    cliente_nome = row[1] or "Consumidor Final"
                    desc_limpa = limpar_descricao_conta(row[2])
                    d.append((row[0], cliente_nome, desc_limpa, formatar_moeda(row[3]), iso_para_br_data(row[4]) if row[4] else "-", iso_para_br(row[5]), formatar_status(row[6]), parcela_info, "☐"))
                definir_dados_paginados(chave, tree, d)
            except Exception:
                pass
    except Exception as e:
        print("Erro listar cr", e)
    conn.close()

def salvar_conta_receber():
    cli_text = combo_cr_cliente.get().strip()
    desc = entry_cr_desc.get().strip()
    valor_str = entry_cr_valor.get().strip()
    venc_br = entry_cr_venc.get().strip()
    
    if not validar_obrigatorio(cli_text, "Cliente"):
        combo_cr_cliente.focus()
        return
    if not validar_obrigatorio(desc, "Descrição da Conta"):
        entry_cr_desc.focus()
        return
    if not validar_obrigatorio(valor_str, "Valor da Conta"):
        entry_cr_valor.focus()
        return
    if not validar_obrigatorio(venc_br, "Vencimento"):
        entry_cr_venc.focus()
        return
    
    cli_id = None
    if cli_text and " - " in cli_text:
        try:
            cli_id = int(cli_text.split(" - ")[0])
        except:
            pass
    
    try:
        valor = float(valor_str.replace(",", "."))
    except:
        mostrar_aviso("Valor inválido!")
        return
    venc_iso = br_para_iso(venc_br) or hoje_iso()
    if not venc_iso:
        mostrar_aviso("Data vencimento inválida! Use dd/mm/aaaa")
        return
    
    forma = combo_cr_forma.get().strip() or "PIX"
    parcelas = 1
    taxa = 0
    if forma == "Cartão Crédito":
        try:
            parcelas = int(combo_cr_parcelas.get() or 1)
        except:
            parcelas = 1
        try:
            taxa = float(entry_cr_taxa.get().replace(",", ".") or 0)
        except:
            taxa = 0
        valor = valor * (1 + taxa/100)
    
    cr_id = entry_cr_id.get().strip()
    
    conn = conectar()
    cur = conn.cursor()
    try:
        if cr_id:
            if not _conta_id_ja_existe("contas_a_receber", cr_id):
                mostrar_aviso(f"Não existe conta a receber com ID {cr_id}.")
                return
            cur.execute("UPDATE contas_a_receber SET cliente_id=?, descricao=?, valor=?, vencimento=?, forma_pagamento=? WHERE id=?",
                        (cli_id, desc, valor, venc_iso, forma, cr_id))
        else:
            if forma == "Cartão Crédito" and parcelas > 1:
                valor_parcela = valor / parcelas
                base_date = datetime.strptime(venc_iso, "%Y-%m-%d")
                for i in range(1, parcelas+1):
                    venc_parcela = base_date + timedelta(days=30*(i-1))
                    venc_parcela_iso = venc_parcela.strftime("%Y-%m-%d")
                    desc_parcela = f"{desc} - Parcela {i}/{parcelas}"
                    cur.execute("INSERT INTO contas_a_receber (cliente_id, descricao, valor, vencimento, status, forma_pagamento, parcela_atual, total_parcelas) VALUES (?,?,?,?,?,?,?,?)",
                                (cli_id, desc_parcela, valor_parcela, venc_parcela_iso, 'em_aberto', forma, i, parcelas))
            else:
                cur.execute("INSERT INTO contas_a_receber (cliente_id, descricao, valor, vencimento, status, forma_pagamento) VALUES (?,?,?,?,?,?)",
                            (cli_id, desc, valor, venc_iso, 'em_aberto', forma))
        conn.commit()
        limpar_form_cr()
        listar_contas_receber()
        atualizar_dashboard()
        mostrar_sucesso(f"Conta a receber '{desc}' salva!", "Conta Salva")
    except Exception as e:
        mostrar_erro(str(e))
    finally:
        conn.close()

def editar_cr(event=None):
    """Duplo clique: mostra detalhes da conta a receber."""
    tree_ativo = tree_cr
    try:
        tab_id = notebook_cr.select()
        tab_frame = notebook_cr.nametowidget(tab_id)
        for child in tab_frame.winfo_children():
            if isinstance(child, ttk.Treeview):
                tree_ativo = child
                break
            for subchild in child.winfo_children():
                if isinstance(subchild, ttk.Treeview):
                    tree_ativo = subchild
                    break
    except Exception:
        pass
    sel = tree_ativo.selection()
    if not sel:
        return
    vals = tree_ativo.item(sel[0])['values']
    # ID, Cliente, Descrição, Valor, Data Venda, Venc BR, Status, Parcela, Sel
    data_venda = vals[4] if len(vals) > 4 else "-"
    venc = vals[5] if len(vals) > 5 else "-"
    status = vals[6] if len(vals) > 6 else "-"
    parcela = vals[7] if len(vals) > 7 else "-"
    msg = "ID: {0}\nCliente: {1}\nDescrição: {2}\nValor: {3}\nData da venda: {4}\nVencimento: {5}\nStatus: {6}\nParcela: {7}".format(
        vals[0], vals[1], vals[2], vals[3], data_venda, venc, status, parcela
    )
    mostrar_info(msg, "Conta a Receber #" + str(vals[0]))



def receber_conta():
    """Recebe uma ou várias contas a receber marcadas no ☐ (ou a linha selecionada)."""
    tree_ativo = _tree_ativa_notebook(notebook_cr, tree_cr)

    ids = obter_ids_selecionados(tree_ativo)
    if not ids:
        sel = tree_ativo.selection()
        if not sel:
            mostrar_aviso("Marque no ☐ uma ou mais contas em aberto para receber.")
            return
        ids = [tree_ativo.item(s)['values'][0] for s in sel]

    conn = conectar()
    cur = conn.cursor()
    contas = []
    for cr_id in ids:
        cur.execute(
            "SELECT id, valor, status, descricao, forma_pagamento, venda_id FROM contas_a_receber WHERE id=? AND excluido=0",
            (cr_id,),
        )
        row = cur.fetchone()
        if not row:
            continue
        cid, valor, status, desc, forma, venda_id = row
        if status in ('recebido', 'cancelado'):
            continue
        contas.append({
            "id": cid,
            "valor": float(valor or 0),
            "desc": desc or "",
            "forma": forma or "PIX",
            "status": status,
            "venda_id": venda_id,
        })

    if not contas:
        conn.close()
        mostrar_aviso("Nenhuma conta válida para receber.\n(Marque contas Em Aberto / Em Atraso.)")
        return

    total = sum(c["valor"] for c in contas)
    if not confirmar_contas_lote(
        "Confirmar Recebimento",
        contas,
        total,
        acao_label="Receber contas",
    ):
        conn.close()
        return

    try:
        cur.execute("BEGIN")
        vendas_ids = set()
        for c in contas:
            cur.execute(
                "UPDATE contas_a_receber SET status='recebido', data_recebimento=? WHERE id=?",
                (hoje_str(), c["id"]),
            )
            cur.execute(
                "INSERT INTO caixa (data,tipo,valor,descricao,origem,referencia_id,forma_pagamento) VALUES (?,?,?,?,?,?,?)",
                (hoje_str(), 'entrada', c["valor"], f"Recebimento: {c['desc']}", 'conta_receber', c["id"], c["forma"]),
            )
            if c.get("venda_id"):
                vendas_ids.add(c["venda_id"])
        conn.commit()
        for vid in vendas_ids:
            try:
                atualizar_status_venda(vid)
            except Exception:
                pass
        mostrar_sucesso(
            f"{len(contas)} conta(s) recebida(s)!\n\nTotal: {formatar_moeda(total)}\nEntrada(s) no caixa realizada(s).",
            "Contas Recebidas",
        )
        try:
            limpar_selecao_multipla(tree_ativo)
        except Exception:
            pass
        listar_contas_receber()
        listar_caixa()
        listar_vendas()
        atualizar_dashboard()
    except Exception as e:
        conn.rollback()
        mostrar_erro(str(e))
    finally:
        conn.close()



def excluir_cr():
    if not verificar_permissao_exclusao():
        return
    tree_ativo = tree_cr
    try:
        tab_id = notebook_cr.select()
        tab_frame = notebook_cr.nametowidget(tab_id)
        for child in tab_frame.winfo_children():
            if isinstance(child, ttk.Treeview):
                tree_ativo = child
                break
            for subchild in child.winfo_children():
                if isinstance(subchild, ttk.Treeview):
                    tree_ativo = subchild
                    break
    except:
        pass
    sel = tree_ativo.selection()
    if not sel:
        mostrar_aviso("Selecione uma conta!")
        return
    cr_id = tree_ativo.item(sel[0])['values'][0]
    if not confirmar_moderno("Mover para Lixeira", "Mover conta a receber para LIXEIRA?"):
        return
    conn = conectar()
    cur = conn.cursor()
    try:
        cur.execute("UPDATE contas_a_receber SET excluido=1, data_exclusao=?, excluido_por=? WHERE id=?", (hoje_str(), usuario_logado['login'], cr_id))
        conn.commit()
        listar_contas_receber()
        listar_lixeira_cr()
        atualizar_dashboard()
        mostrar_sucesso("Conta movida para lixeira!", "Lixeira")
    except Exception as e:
        mostrar_erro(str(e))
    finally:
        conn.close()

def excluir_contas_receber_em_massa(ids):
    if not verificar_permissao_exclusao():
        return
    conn = conectar()
    cur = conn.cursor()
    try:
        for cr_id in ids:
            cur.execute("UPDATE contas_a_receber SET excluido=1, data_exclusao=?, excluido_por=? WHERE id=?",
                        (hoje_str(), usuario_logado['login'], cr_id))
        conn.commit()
        listar_contas_receber()
        listar_lixeira_cr()
        atualizar_dashboard()
        mostrar_sucesso(f"{len(ids)} conta(s) movida(s) para lixeira!", "Lixeira")
    except Exception as e:
        conn.rollback()
        mostrar_erro(str(e))
    finally:
        conn.close()

def cancelar_cr():
    tree_ativo = tree_cr
    try:
        tab_id = notebook_cr.select()
        tab_frame = notebook_cr.nametowidget(tab_id)
        for child in tab_frame.winfo_children():
            if isinstance(child, ttk.Treeview):
                tree_ativo = child
                break
            for subchild in child.winfo_children():
                if isinstance(subchild, ttk.Treeview):
                    tree_ativo = subchild
                    break
    except:
        pass
    sel = tree_ativo.selection()
    if not sel:
        mostrar_aviso("Selecione uma conta para cancelar!")
        return
    cr_id = tree_ativo.item(sel[0])['values'][0]
    if not confirmar_moderno("Cancelar Conta", f"Cancelar conta {cr_id}?\n\nAtualiza em todos os módulos."):
        return
    conn = conectar()
    cur = conn.cursor()
    try:
        cur.execute("SELECT venda_id FROM contas_a_receber WHERE id=?", (cr_id,))
        venda_id = cur.fetchone()
        venda_id = venda_id[0] if venda_id else None
        cur.execute("UPDATE contas_a_receber SET status='cancelado' WHERE id=? AND excluido=0", (cr_id,))
        conn.commit()
        if venda_id:
            cur = conn.cursor()
            cur.execute("SELECT COUNT(*) FROM contas_a_receber WHERE venda_id=? AND status!='cancelado' AND excluido=0", (venda_id,))
            if cur.fetchone()[0] == 0:
                cur.execute("UPDATE vendas SET status='cancelado' WHERE id=?", (venda_id,))
                conn.commit()
        listar_contas_receber()
        listar_vendas()
        atualizar_dashboard()
        mostrar_sucesso("Conta cancelada! Atualizado em todos os módulos.", "Cancelado")
    except Exception as e:
        mostrar_erro(str(e))
    finally:
        conn.close()

def limpar_form_cr():
    pass

def _add_months(data_iso, months):
    """Soma meses a uma data ISO YYYY-MM-DD"""
    try:
        dt = datetime.strptime(data_iso[:10], "%Y-%m-%d")
        m = dt.month - 1 + months
        y = dt.year + m // 12
        m = m % 12 + 1
        d = min(dt.day, [31, 29 if y % 4 == 0 and (y % 100 != 0 or y % 400 == 0) else 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][m - 1])
        return date(y, m, d).isoformat()
    except Exception:
        return (datetime.strptime(data_iso[:10], "%Y-%m-%d") + timedelta(days=30 * months)).strftime("%Y-%m-%d")

def abrir_modal_conta_pagar():
    """Modal para incluir conta a pagar (única ou mensal)."""
    global root
    modal = tk.Toplevel(root)
    modal.title("Incluir Conta a Pagar")
    modal.geometry("520x520")
    modal.configure(bg="white")
    modal.transient(root)
    modal.grab_set()
    modal.resizable(False, False)
    modal.update_idletasks()
    try:
        x = root.winfo_x() + (root.winfo_width() // 2) - 260
        y = root.winfo_y() + (root.winfo_height() // 2) - 260
        modal.geometry(f"+{x}+{y}")
    except Exception:
        pass

    header = tk.Frame(modal, bg=CORES["danger"], height=56)
    header.pack(fill='x')
    header.pack_propagate(False)
    tk.Label(header, text="📤  Incluir Conta a Pagar", bg=CORES["danger"], fg="white",
             font=('Arial', 13, 'bold')).pack(pady=14)

    body = tk.Frame(modal, bg="white", padx=25, pady=15)
    body.pack(fill='both', expand=True)

    def row_lbl(txt, r):
        tk.Label(body, text=txt, bg="white", fg=CORES["text_dark"], font=('Arial', 9, 'bold'),
                 width=18, anchor='e').grid(row=r, column=0, sticky='e', padx=(0, 8), pady=6)

    row_lbl("Fornecedor*:", 0)
    combo_forn = ttk.Combobox(body, width=32, values=listas_combos.get("fornecedores", []))
    combo_forn.grid(row=0, column=1, sticky='w', pady=6)
    configurar_autocomplete_combo(combo_forn, listas_combos.get("fornecedores", []))

    row_lbl("Nº documento:", 1)
    entry_doc = tk.Entry(body, width=34)
    entry_doc.grid(row=1, column=1, sticky='w', pady=6)

    row_lbl("Data emissão*:", 2)
    entry_emissao = tk.Entry(body, width=18)
    entry_emissao.grid(row=2, column=1, sticky='w', pady=6)
    entry_emissao.insert(0, hoje_br())
    ativar_seletor_data(entry_emissao)

    row_lbl("Vencimento*:", 3)
    entry_venc = tk.Entry(body, width=18)
    entry_venc.grid(row=3, column=1, sticky='w', pady=6)
    entry_venc.insert(0, hoje_br())
    ativar_seletor_data(entry_venc)

    row_lbl("Valor*:", 4)
    entry_valor = tk.Entry(body, width=18)
    entry_valor.grid(row=4, column=1, sticky='w', pady=6)

    row_lbl("Forma pagamento*:", 5)
    combo_forma = ttk.Combobox(body, width=30, values=["Boleto", "PIX", "Dinheiro", "Cartão Débito", "Cartão Crédito", "Transferência"])
    combo_forma.grid(row=5, column=1, sticky='w', pady=6)
    combo_forma.set("Boleto")

    row_lbl("Ocorrência*:", 6)
    combo_ocor = ttk.Combobox(body, width=30, values=["Única", "Mensal"], state="readonly")
    combo_ocor.grid(row=6, column=1, sticky='w', pady=6)
    combo_ocor.set("Única")

    row_lbl("Qtd. meses:", 7)
    combo_meses = ttk.Combobox(body, width=10, values=[str(i) for i in range(1, 37)], state="disabled")
    combo_meses.grid(row=7, column=1, sticky='w', pady=6)
    combo_meses.set("1")
    lbl_info_parc = tk.Label(body, text="", bg="white", fg=CORES["primary"], font=('Arial', 9))
    lbl_info_parc.grid(row=8, column=0, columnspan=2, sticky='w', pady=4)

    def on_ocor_change(event=None):
        if combo_ocor.get() == "Mensal":
            combo_meses.config(state="readonly")
            try:
                m = int(combo_meses.get() or 1)
            except Exception:
                m = 1
            lbl_info_parc.config(text=f"Serão gerados {m} lançamentos com o VALOR INTEGRAL em cada mês (1/{m} … {m}/{m})")
        else:
            combo_meses.config(state="disabled")
            combo_meses.set("1")
            lbl_info_parc.config(text="Registro único (1 lançamento)")

    combo_ocor.bind("<<ComboboxSelected>>", on_ocor_change)
    combo_meses.bind("<<ComboboxSelected>>", on_ocor_change)
    on_ocor_change()

    def salvar():
        forn_text = combo_forn.get().strip()
        doc = entry_doc.get().strip()
        emissao_br = entry_emissao.get().strip()
        venc_br = entry_venc.get().strip()
        valor_str = entry_valor.get().strip()
        forma = combo_forma.get().strip() or "Boleto"
        ocor = combo_ocor.get().strip() or "Única"
        if not validar_obrigatorio(forn_text, "Fornecedor"):
            return
        if not validar_obrigatorio(emissao_br, "Data de emissão"):
            return
        if not validar_obrigatorio(venc_br, "Vencimento"):
            return
        if not validar_obrigatorio(valor_str, "Valor"):
            return
        if not validar_obrigatorio(forma, "Forma de pagamento"):
            return
        forn_id = None
        if " - " in forn_text:
            try:
                forn_id = int(forn_text.split(" - ")[0])
            except Exception:
                pass
        if not forn_id:
            mostrar_aviso("Selecione um fornecedor da lista (digite e escolha).")
            return
        try:
            valor = float(valor_str.replace(",", "."))
            if valor <= 0:
                raise ValueError()
        except Exception:
            mostrar_aviso("Valor inválido!")
            return
        emissao_iso = br_para_iso(emissao_br)
        venc_iso = br_para_iso(venc_br)
        if not emissao_iso or not venc_iso:
            mostrar_aviso("Datas inválidas! Use dd/mm/aaaa")
            return
        try:
            n_meses = int(combo_meses.get() or 1) if ocor == "Mensal" else 1
            if n_meses < 1:
                n_meses = 1
        except Exception:
            n_meses = 1

        desc_base = f"Doc {doc}" if doc else "Conta a pagar"
        if ocor == "Mensal":
            desc_base = f"{desc_base} (Mensal)" if doc else "Conta a pagar (Mensal)"

        # Trava: mesmo nº de documento não pode ser cadastrado de novo
        if doc and _conta_ja_existe_por_documento("contas_a_pagar", doc):
            mostrar_aviso(f"Já existe uma conta a pagar com o nº de documento/ID \"{doc}\".\nNão é permitido cadastrar em duplicidade.")
            return

        conn = conectar()
        cur = conn.cursor()
        try:
            valor_mensal = valor  # mensal = valor integral em cada mês (não divide)
            for i in range(1, n_meses + 1):
                venc_i = venc_iso if i == 1 else _add_months(venc_iso, i - 1)
                desc = desc_base if n_meses == 1 else f"{desc_base} - {i}/{n_meses}"
                cur.execute("""
                    INSERT INTO contas_a_pagar
                    (fornecedor_id, descricao, valor, vencimento, status, forma_pagamento,
                     parcela_atual, total_parcelas, data_emissao, numero_documento)
                    VALUES (?,?,?,?,?,?,?,?,?,?)
                """, (forn_id, desc, valor_mensal, venc_i, 'em_aberto', forma, i, n_meses, emissao_iso, doc or None))
            conn.commit()
            listar_contas_pagar()
            atualizar_dashboard()
            modal.destroy()
            msg = f"Conta a pagar salva!" if n_meses == 1 else f"{n_meses} lançamentos mensais de {formatar_moeda(valor_mensal)} gerados!"
            mostrar_sucesso(msg, "Conta a Pagar")
        except Exception as e:
            mostrar_erro(str(e))
        finally:
            conn.close()

    frame_b = tk.Frame(body, bg="white")
    frame_b.grid(row=9, column=0, columnspan=2, pady=16)
    tk.Button(frame_b, text="💾 Salvar", command=salvar, bg=CORES["success"], fg="white",
              font=('Arial', 10, 'bold'), bd=0, padx=20, pady=8, cursor='hand2').pack(side='left', padx=8)
    tk.Button(frame_b, text="Cancelar", command=modal.destroy, bg="#64748b", fg="white",
              font=('Arial', 10), bd=0, padx=16, pady=8, cursor='hand2').pack(side='left', padx=8)
    modal.bind('<Escape>', lambda e: modal.destroy())
    combo_forn.focus()

def abrir_modal_conta_receber():
    """Modal para incluir conta a receber (única ou mensal)."""
    global root
    modal = tk.Toplevel(root)
    modal.title("Incluir Conta a Receber")
    modal.geometry("520x520")
    modal.configure(bg="white")
    modal.transient(root)
    modal.grab_set()
    modal.resizable(False, False)
    modal.update_idletasks()
    try:
        x = root.winfo_x() + (root.winfo_width() // 2) - 260
        y = root.winfo_y() + (root.winfo_height() // 2) - 260
        modal.geometry(f"+{x}+{y}")
    except Exception:
        pass

    header = tk.Frame(modal, bg=CORES["success"], height=56)
    header.pack(fill='x')
    header.pack_propagate(False)
    tk.Label(header, text="📥  Incluir Conta a Receber", bg=CORES["success"], fg="white",
             font=('Arial', 13, 'bold')).pack(pady=14)

    body = tk.Frame(modal, bg="white", padx=25, pady=15)
    body.pack(fill='both', expand=True)

    def row_lbl(txt, r):
        tk.Label(body, text=txt, bg="white", fg=CORES["text_dark"], font=('Arial', 9, 'bold'),
                 width=18, anchor='e').grid(row=r, column=0, sticky='e', padx=(0, 8), pady=6)

    clientes_vals = [c for c in listas_combos.get("clientes", []) if not str(c).startswith("0 -")]
    row_lbl("Cliente*:", 0)
    combo_cli = ttk.Combobox(body, width=32, values=clientes_vals)
    combo_cli.grid(row=0, column=1, sticky='w', pady=6)
    configurar_autocomplete_combo(combo_cli, clientes_vals)

    row_lbl("Nº documento:", 1)
    entry_doc = tk.Entry(body, width=34)
    entry_doc.grid(row=1, column=1, sticky='w', pady=6)

    row_lbl("Data emissão*:", 2)
    entry_emissao = tk.Entry(body, width=18)
    entry_emissao.grid(row=2, column=1, sticky='w', pady=6)
    entry_emissao.insert(0, hoje_br())
    ativar_seletor_data(entry_emissao)

    row_lbl("Vencimento*:", 3)
    entry_venc = tk.Entry(body, width=18)
    entry_venc.grid(row=3, column=1, sticky='w', pady=6)
    entry_venc.insert(0, hoje_br())
    ativar_seletor_data(entry_venc)

    row_lbl("Valor*:", 4)
    entry_valor = tk.Entry(body, width=18)
    entry_valor.grid(row=4, column=1, sticky='w', pady=6)

    row_lbl("Forma recebimento*:", 5)
    combo_forma = ttk.Combobox(body, width=30, values=["PIX", "Dinheiro", "Boleto", "Cartão Débito", "Cartão Crédito", "Transferência"])
    combo_forma.grid(row=5, column=1, sticky='w', pady=6)
    combo_forma.set("PIX")

    row_lbl("Ocorrência*:", 6)
    combo_ocor = ttk.Combobox(body, width=30, values=["Única", "Mensal"], state="readonly")
    combo_ocor.grid(row=6, column=1, sticky='w', pady=6)
    combo_ocor.set("Única")

    row_lbl("Qtd. meses:", 7)
    combo_meses = ttk.Combobox(body, width=10, values=[str(i) for i in range(1, 37)], state="disabled")
    combo_meses.grid(row=7, column=1, sticky='w', pady=6)
    combo_meses.set("1")
    lbl_info_parc = tk.Label(body, text="", bg="white", fg=CORES["primary"], font=('Arial', 9))
    lbl_info_parc.grid(row=8, column=0, columnspan=2, sticky='w', pady=4)

    def on_ocor_change(event=None):
        if combo_ocor.get() == "Mensal":
            combo_meses.config(state="readonly")
            try:
                m = int(combo_meses.get() or 1)
            except Exception:
                m = 1
            lbl_info_parc.config(text=f"Serão gerados {m} lançamentos com o VALOR INTEGRAL em cada mês (1/{m} … {m}/{m})")
        else:
            combo_meses.config(state="disabled")
            combo_meses.set("1")
            lbl_info_parc.config(text="Registro único (1 lançamento)")

    combo_ocor.bind("<<ComboboxSelected>>", on_ocor_change)
    combo_meses.bind("<<ComboboxSelected>>", on_ocor_change)
    on_ocor_change()

    def salvar():
        cli_text = combo_cli.get().strip()
        doc = entry_doc.get().strip()
        emissao_br = entry_emissao.get().strip()
        venc_br = entry_venc.get().strip()
        valor_str = entry_valor.get().strip()
        forma = combo_forma.get().strip() or "PIX"
        ocor = combo_ocor.get().strip() or "Única"
        if not validar_obrigatorio(cli_text, "Cliente"):
            return
        if not validar_obrigatorio(emissao_br, "Data de emissão"):
            return
        if not validar_obrigatorio(venc_br, "Vencimento"):
            return
        if not validar_obrigatorio(valor_str, "Valor"):
            return
        cli_id = None
        if " - " in cli_text:
            try:
                cli_id = int(cli_text.split(" - ")[0])
            except Exception:
                pass
        if not cli_id:
            mostrar_aviso("Selecione um cliente da lista (digite e escolha).")
            return
        try:
            valor = float(valor_str.replace(",", "."))
            if valor <= 0:
                raise ValueError()
        except Exception:
            mostrar_aviso("Valor inválido!")
            return
        emissao_iso = br_para_iso(emissao_br)
        venc_iso = br_para_iso(venc_br)
        if not emissao_iso or not venc_iso:
            mostrar_aviso("Datas inválidas! Use dd/mm/aaaa")
            return
        try:
            n_meses = int(combo_meses.get() or 1) if ocor == "Mensal" else 1
            if n_meses < 1:
                n_meses = 1
        except Exception:
            n_meses = 1

        desc_base = f"Doc {doc}" if doc else "Conta a receber"
        if ocor == "Mensal":
            desc_base = f"{desc_base} (Mensal)" if doc else "Conta a receber (Mensal)"

        # Trava: mesmo nº de documento não pode ser cadastrado de novo
        if doc and _conta_ja_existe_por_documento("contas_a_receber", doc):
            mostrar_aviso(f"Já existe uma conta a receber com o nº de documento/ID \"{doc}\".\nNão é permitido cadastrar em duplicidade.")
            return

        conn = conectar()
        cur = conn.cursor()
        try:
            valor_mensal = valor  # mensal = valor integral em cada mês (não divide)
            for i in range(1, n_meses + 1):
                venc_i = venc_iso if i == 1 else _add_months(venc_iso, i - 1)
                desc = desc_base if n_meses == 1 else f"{desc_base} - {i}/{n_meses}"
                cur.execute("""
                    INSERT INTO contas_a_receber
                    (cliente_id, descricao, valor, vencimento, status, forma_pagamento,
                     parcela_atual, total_parcelas, data_emissao, numero_documento)
                    VALUES (?,?,?,?,?,?,?,?,?,?)
                """, (cli_id, desc, valor_mensal, venc_i, 'em_aberto', forma, i, n_meses, emissao_iso, doc or None))
            conn.commit()
            listar_contas_receber()
            atualizar_dashboard()
            modal.destroy()
            msg = f"Conta a receber salva!" if n_meses == 1 else f"{n_meses} lançamentos mensais de {formatar_moeda(valor_mensal)} gerados!"
            mostrar_sucesso(msg, "Conta a Receber")
        except Exception as e:
            mostrar_erro(str(e))
        finally:
            conn.close()

    frame_b = tk.Frame(body, bg="white")
    frame_b.grid(row=9, column=0, columnspan=2, pady=16)
    tk.Button(frame_b, text="💾 Salvar", command=salvar, bg=CORES["success"], fg="white",
              font=('Arial', 10, 'bold'), bd=0, padx=20, pady=8, cursor='hand2').pack(side='left', padx=8)
    tk.Button(frame_b, text="Cancelar", command=modal.destroy, bg="#64748b", fg="white",
              font=('Arial', 10), bd=0, padx=16, pady=8, cursor='hand2').pack(side='left', padx=8)
    modal.bind('<Escape>', lambda e: modal.destroy())
    combo_cli.focus()

# RELATÓRIOS
def relatorio_vendas():
    data_ini_br = entry_rel_venda_ini.get().strip()
    data_fim_br = entry_rel_venda_fim.get().strip()
    data_ini_iso = br_para_iso(data_ini_br) if data_ini_br else None
    data_fim_iso = br_para_iso(data_fim_br) if data_fim_br else None
    conn = conectar()
    cur = conn.cursor()
    query = """
        SELECT v.id, v.data, COALESCE(c.nome,'Avulso'), v.total, v.desconto, v.forma_pagamento, v.status, v.parcelas
        FROM vendas v LEFT JOIN clientes c ON v.cliente_id=c.id
        WHERE v.excluido=0
    """
    params = []
    if data_ini_iso:
        query += " AND date(v.data) >= date(?)"
        params.append(data_ini_iso)
    if data_fim_iso:
        query += " AND date(v.data) <= date(?)"
        params.append(data_fim_iso)
    query += " ORDER BY v.id DESC"
    cur.execute(query, params)
    rows = cur.fetchall()
    conn.close()
    if not rows:
        mostrar_aviso("Nenhuma venda no período!")
        return
    colunas = ["ID", "Data", "Cliente", "Total", "Desconto", "Forma Pagamento", "Status", "Parcelas"]
    dados = []
    for r in rows:
        dados.append([r[0], iso_para_br(r[1]), r[2], r[3], r[4], r[5], r[6], r[7]])
    exportar_dados(colunas, dados, f"relatorio_vendas_{hoje_iso()}.xlsx")

def relatorio_estoque():
    conn = conectar()
    cur = conn.cursor()
    cur.execute("""
        SELECT p.id, p.codigo, p.nome, p.preco_custo, p.preco_venda, p.estoque_atual, p.estoque_minimo, COALESCE(f.nome,'-')
        FROM produtos p LEFT JOIN fornecedores f ON p.fornecedor_id=f.id
        WHERE p.excluido=0
        ORDER BY p.id DESC
    """)
    rows = cur.fetchall()
    conn.close()
    colunas = ["ID", "Código", "Nome", "Preço Custo", "Preço Venda", "Estoque Atual", "Estoque Mínimo", "Fornecedor"]
    exportar_dados(colunas, rows, f"relatorio_estoque_{hoje_iso()}.xlsx")

def relatorio_clientes():
    conn = conectar()
    cur = conn.cursor()
    cur.execute("SELECT id, nome, cpf_cnpj, telefone, email, endereco, data_cadastro FROM clientes WHERE excluido=0 ORDER BY id DESC")
    rows = cur.fetchall()
    conn.close()
    colunas = ["ID", "Nome", "CPF/CNPJ", "Telefone", "Email", "Endereço", "Data Cadastro"]
    dados = [[r[0], r[1], r[2], r[3], r[4], r[5], iso_para_br(r[6])] for r in rows]
    exportar_dados(colunas, dados, f"relatorio_clientes_{hoje_iso()}.xlsx")

def relatorio_caixa():
    data_ini_br = entry_rel_caixa_ini.get().strip()
    data_fim_br = entry_rel_caixa_fim.get().strip()
    data_ini_iso = br_para_iso(data_ini_br) if data_ini_br else None
    data_fim_iso = br_para_iso(data_fim_br) if data_fim_br else None
    conn = conectar()
    cur = conn.cursor()
    query = "SELECT id, data, tipo, valor, descricao, origem, forma_pagamento FROM caixa WHERE excluido=0"
    params = []
    if data_ini_iso:
        query += " AND date(data) >= date(?)"
        params.append(data_ini_iso)
    if data_fim_iso:
        query += " AND date(data) <= date(?)"
        params.append(data_fim_iso)
    query += " ORDER BY id DESC"
    cur.execute(query, params)
    rows = cur.fetchall()
    conn.close()
    colunas = ["ID", "Data", "Tipo", "Valor", "Descrição", "Origem", "Forma Pagamento"]
    dados = [[r[0], iso_para_br(r[1]), r[2], r[3], r[4], r[5], r[6]] for r in rows]
    exportar_dados(colunas, dados, f"relatorio_caixa_{hoje_iso()}.xlsx")

def relatorio_contas():
    conn = conectar()
    cur = conn.cursor()
    cur.execute("""
        SELECT 'PAGAR' as tipo, cp.id, COALESCE(f.nome,'-'), cp.descricao, cp.valor, cp.vencimento, cp.status, cp.data_pagamento
        FROM contas_a_pagar cp LEFT JOIN fornecedores f ON cp.fornecedor_id=f.id
        WHERE cp.excluido=0
        UNION ALL
        SELECT 'RECEBER' as tipo, cr.id, COALESCE(c.nome,'-'), cr.descricao, cr.valor, cr.vencimento, cr.status, cr.data_recebimento
        FROM contas_a_receber cr LEFT JOIN clientes c ON cr.cliente_id=c.id
        WHERE cr.excluido=0
        ORDER BY 6
    """)
    rows = cur.fetchall()
    conn.close()
    colunas = ["Tipo", "ID", "Cliente/Fornecedor", "Descrição", "Valor", "Vencimento", "Status", "Data Pag/Rec"]
    dados = [[r[0], r[1], r[2], r[3], r[4], iso_para_br(r[5]), r[6], iso_para_br(r[7])] for r in rows]
    exportar_dados(colunas, dados, f"relatorio_contas_{hoje_iso()}.xlsx")

# LIXEIRA
def listar_lixeira_clientes():
    conn = conectar()
    cur = conn.cursor()
    cur.execute("SELECT id, nome, cpf_cnpj, data_exclusao, excluido_por FROM clientes WHERE excluido=1 ORDER BY data_exclusao DESC")
    dados = [(row[0], row[1], row[2], iso_para_br(row[3]), row[4], "☐") for row in cur.fetchall()]
    conn.close()
    definir_dados_paginados("lix_cli", tree_lixeira_clientes, dados)

def listar_lixeira_fornecedores():
    conn = conectar()
    cur = conn.cursor()
    cur.execute("SELECT id, nome, cnpj, data_exclusao, excluido_por FROM fornecedores WHERE excluido=1 ORDER BY data_exclusao DESC")
    dados = [(row[0], row[1], row[2], iso_para_br(row[3]), row[4], "☐") for row in cur.fetchall()]
    conn.close()
    definir_dados_paginados("lix_forn", tree_lixeira_fornecedores, dados)

def listar_lixeira_produtos():
    conn = conectar()
    cur = conn.cursor()
    cur.execute("SELECT id, codigo, nome, preco_venda, data_exclusao, excluido_por FROM produtos WHERE excluido=1 ORDER BY data_exclusao DESC")
    dados = [(row[0], row[1], row[2], formatar_moeda(row[3]), iso_para_br(row[4]), row[5], "☐") for row in cur.fetchall()]
    conn.close()
    definir_dados_paginados("lix_prod", tree_lixeira_produtos, dados)

def listar_lixeira_vendas():
    conn = conectar()
    cur = conn.cursor()
    cur.execute("SELECT id, data, total, status, data_exclusao, excluido_por FROM vendas WHERE excluido=1 ORDER BY data_exclusao DESC")
    dados = [(row[0], iso_para_br(row[1]), formatar_moeda(row[2]), formatar_status(row[3]), iso_para_br(row[4]), row[5], "☐") for row in cur.fetchall()]
    conn.close()
    definir_dados_paginados("lix_vendas", tree_lixeira_vendas, dados)

def listar_lixeira_caixa():
    conn = conectar()
    cur = conn.cursor()
    cur.execute("SELECT id, data, tipo, valor, descricao, data_exclusao FROM caixa WHERE excluido=1 ORDER BY data_exclusao DESC")
    dados = [(row[0], iso_para_br(row[1]), row[2], formatar_moeda(row[3]), row[4], iso_para_br(row[5]), "☐") for row in cur.fetchall()]
    conn.close()
    definir_dados_paginados("lix_caixa", tree_lixeira_caixa, dados)

def listar_lixeira_cp():
    conn = conectar()
    cur = conn.cursor()
    cur.execute("SELECT id, descricao, valor, vencimento, status, data_exclusao FROM contas_a_pagar WHERE excluido=1 ORDER BY data_exclusao DESC")
    dados = [(row[0], row[1], formatar_moeda(row[2]), iso_para_br(row[3]), formatar_status(row[4]), iso_para_br(row[5]), "☐") for row in cur.fetchall()]
    conn.close()
    definir_dados_paginados("lix_cp", tree_lixeira_cp, dados)

def listar_lixeira_cr():
    conn = conectar()
    cur = conn.cursor()
    cur.execute("SELECT id, descricao, valor, vencimento, status, data_exclusao FROM contas_a_receber WHERE excluido=1 ORDER BY data_exclusao DESC")
    dados = [(row[0], row[1], formatar_moeda(row[2]), iso_para_br(row[3]), formatar_status(row[4]), iso_para_br(row[5]), "☐") for row in cur.fetchall()]
    conn.close()
    definir_dados_paginados("lix_cr", tree_lixeira_cr, dados)


def restaurar_itens_em_massa(tabela, tree):
    ids = obter_ids_selecionados(tree)
    if not ids:
        mostrar_aviso("Marque ao menos um item (clique no ☐) para restaurar.")
        return
    if not confirmar_moderno("Restaurar selecionados", f"Restaurar {len(ids)} item(ns) da lixeira?"):
        return
    conn = conectar()
    cur = conn.cursor()
    try:
        for item_id in ids:
            cur.execute(f"UPDATE {tabela} SET excluido=0, data_exclusao=NULL, excluido_por=NULL WHERE id=?", (item_id,))
        conn.commit()
        mostrar_sucesso(f"{len(ids)} item(ns) restaurado(s)!", "Restaurado")
        listar_todas_lixeiras()
        listar_clientes(); listar_fornecedores(); listar_produtos()
        listar_estoque(); listar_vendas(); listar_caixa()
        listar_contas_pagar(); listar_contas_receber()
        atualizar_combos(); atualizar_dashboard()
    except Exception as e:
        mostrar_erro(str(e))
    finally:
        conn.close()

def excluir_definitivo_em_massa(tabela, tree):
    if not verificar_permissao_exclusao():
        return
    ids = obter_ids_selecionados(tree)
    if not ids:
        mostrar_aviso("Marque ao menos um item (clique no ☐) para excluir.")
        return
    if not confirmar_moderno("EXCLUSÃO DEFINITIVA", f"Apagar DEFINITIVAMENTE {len(ids)} item(ns)? Essa ação NÃO pode ser desfeita!"):
        return
    conn = conectar()
    cur = conn.cursor()
    try:
        for item_id in ids:
            if tabela == "produtos":
                cur.execute("SELECT COUNT(*) FROM venda_itens WHERE produto_id=?", (item_id,))
                if cur.fetchone()[0] > 0:
                    continue
            cur.execute(f"DELETE FROM {tabela} WHERE id=?", (item_id,))
        conn.commit()
        mostrar_sucesso("Itens excluídos definitivamente!", "Excluído")
        listar_todas_lixeiras()
    except Exception as e:
        mostrar_erro(str(e))
    finally:
        conn.close()

def listar_todas_lixeiras():
    listar_lixeira_clientes()
    listar_lixeira_fornecedores()
    listar_lixeira_produtos()
    listar_lixeira_vendas()
    listar_lixeira_caixa()
    listar_lixeira_cp()
    listar_lixeira_cr()

def restaurar_item(tabela, tree):
    sel = tree.selection()
    if not sel:
        mostrar_aviso("Selecione um item para restaurar!")
        return
    item_id = tree.item(sel[0])['values'][0]
    conn = conectar()
    cur = conn.cursor()
    try:
        cur.execute(f"UPDATE {tabela} SET excluido=0, data_exclusao=NULL, excluido_por=NULL WHERE id=?", (item_id,))
        conn.commit()
        mostrar_sucesso("Item restaurado com sucesso!", "Restaurado")
        listar_todas_lixeiras()
        listar_clientes()
        listar_fornecedores()
        listar_produtos()
        listar_estoque()
        listar_vendas()
        listar_caixa()
        listar_contas_pagar()
        listar_contas_receber()
        atualizar_combos()
        atualizar_dashboard()
    except Exception as e:
        mostrar_erro(str(e))
    finally:
        conn.close()

def excluir_definitivo_item(tabela, tree):
    if not verificar_permissao_exclusao():
        return
    sel = tree.selection()
    if not sel:
        mostrar_aviso("Selecione um item!")
        return
    item_id = tree.item(sel[0])['values'][0]
    if not confirmar_moderno("EXCLUSÃO DEFINITIVA", f"⚠️ ATENÇÃO! Excluir DEFINITIVAMENTE o item ID {item_id} da tabela {tabela}?\n\nEssa ação NÃO pode ser desfeita!"):
        return
    if not confirmar_moderno("Confirmar novamente", "Tem certeza ABSOLUTA? O dado será apagado para sempre!"):
        return
    conn = conectar()
    cur = conn.cursor()
    try:
        if tabela == "produtos":
            cur.execute("SELECT COUNT(*) FROM venda_itens WHERE produto_id=?", (item_id,))
            if cur.fetchone()[0] > 0:
                mostrar_erro("Produto com histórico de vendas não pode ser excluído definitivamente!")
                conn.close()
                return
        cur.execute(f"DELETE FROM {tabela} WHERE id=?", (item_id,))
        conn.commit()
        mostrar_sucesso("Item excluído definitivamente!", "Excluído")
        listar_todas_lixeiras()
    except Exception as e:
        mostrar_erro(f"Erro ao excluir: {e}")
    finally:
        conn.close()

# USUÁRIOS
def listar_usuarios():
    if not eh_admin():
        try:
            limpar_tree(tree_usuarios)
        except Exception:
            pass
        return
    conn = conectar()
    cur = conn.cursor()
    cur.execute("SELECT id, nome, login, email, perfil, data_cadastro FROM usuarios WHERE excluido=0 ORDER BY id DESC")
    dados = [(row[0], row[1], row[2], row[3] or "-", row[4], iso_para_br(row[5]), "☐") for row in cur.fetchall()]
    conn.close()
    definir_dados_paginados("usuarios", tree_usuarios, dados)

def salvar_usuario():
    if not eh_admin():
        mostrar_aviso("Somente ADM pode gerenciar usuários!")
        return
    nome = entry_usu_nome.get().strip()
    login = entry_usu_login.get().strip()
    senha = entry_usu_senha.get().strip()
    perfil = combo_usu_perfil.get().strip()
    email = entry_usu_email.get().strip()
    usu_id = entry_usu_id.get().strip()
    
    if not validar_obrigatorio(nome, "Nome do Usuário"):
        return
    if not validar_obrigatorio(login, "Login do Usuário"):
        return
    if not validar_obrigatorio(perfil, "Perfil do Usuário"):
        return
    if not validar_obrigatorio(email, "E-mail do Usuário"):
        return
    if not usu_id and not validar_obrigatorio(senha, "Senha do Usuário"):
        return
    
    conn = conectar()
    cur = conn.cursor()
    try:
        if usu_id:
            if senha:
                cur.execute("UPDATE usuarios SET nome=?, login=?, senha=?, perfil=?, email=? WHERE id=?", (nome, login, hash_senha(senha), perfil, email, usu_id))
            else:
                cur.execute("UPDATE usuarios SET nome=?, login=?, perfil=?, email=? WHERE id=?", (nome, login, perfil, email, usu_id))
        else:
            cur.execute("INSERT INTO usuarios (nome, login, senha, perfil, email) VALUES (?,?,?,?,?)", (nome, login, hash_senha(senha), perfil, email))
        conn.commit()
        limpar_form_usuario()
        listar_usuarios()
        mostrar_sucesso(f"Usuário '{nome}' salvo com sucesso!", "Usuário Salvo")
    except sqlite3.IntegrityError:
        mostrar_erro("Login já existe! Use outro login.")
    except Exception as e:
        mostrar_erro(str(e))
    finally:
        conn.close()

def editar_usuario(event=None):
    if not eh_admin():
        return
    sel = tree_usuarios.selection()
    if not sel:
        return
    usu_id = tree_usuarios.item(sel[0])['values'][0]
    conn = conectar()
    cur = conn.cursor()
    cur.execute("SELECT id, nome, login, perfil, email FROM usuarios WHERE id=?", (usu_id,))
    row = cur.fetchone()
    conn.close()
    if row:
        entry_usu_id.delete(0, tk.END); entry_usu_id.insert(0, row[0])
        entry_usu_nome.delete(0, tk.END); entry_usu_nome.insert(0, row[1])
        entry_usu_login.delete(0, tk.END); entry_usu_login.insert(0, row[2])
        combo_usu_perfil.set(row[3])
        entry_usu_email.delete(0, tk.END); entry_usu_email.insert(0, row[4] or "")
        entry_usu_senha.delete(0, tk.END)

def excluir_usuario():
    if not verificar_permissao_exclusao():
        return
    sel = tree_usuarios.selection()
    if not sel:
        mostrar_aviso("Selecione um usuário!")
        return
    usu_id = tree_usuarios.item(sel[0])['values'][0]
    login = tree_usuarios.item(sel[0])['values'][2]
    if login == usuario_logado['login']:
        mostrar_aviso("Você não pode excluir seu próprio usuário logado!")
        return
    if not confirmar_moderno("Excluir Usuário", f"Excluir usuário {login}?"):
        return
    conn = conectar()
    cur = conn.cursor()
    try:
        cur.execute("UPDATE usuarios SET excluido=1 WHERE id=?", (usu_id,))
        conn.commit()
        listar_usuarios()
        mostrar_sucesso("Usuário excluído!", "Excluído")
    except Exception as e:
        mostrar_erro(str(e))
    finally:
        conn.close()

def excluir_usuarios_em_massa(ids):
    if not verificar_permissao_exclusao():
        return
    conn = conectar()
    cur = conn.cursor()
    try:
        excluidos = 0
        bloqueados = 0
        for usu_id in ids:
            cur.execute("SELECT login FROM usuarios WHERE id=?", (usu_id,))
            row = cur.fetchone()
            if row and row[0] == usuario_logado['login']:
                bloqueados += 1
                continue
            cur.execute("UPDATE usuarios SET excluido=1 WHERE id=?", (usu_id,))
            excluidos += 1
        conn.commit()
        listar_usuarios()
        msg = f"{excluidos} usuário(s) excluído(s)!"
        if bloqueados:
            msg += f"\n({bloqueados} não excluído(s): é o seu próprio usuário logado)"
        mostrar_sucesso(msg, "Excluído")
    except Exception as e:
        conn.rollback()
        mostrar_erro(str(e))
    finally:
        conn.close()

def limpar_form_usuario():
    entry_usu_id.delete(0, tk.END)
    entry_usu_nome.delete(0, tk.END)
    entry_usu_login.delete(0, tk.END)
    entry_usu_senha.delete(0, tk.END)
    entry_usu_email.delete(0, tk.END)
    combo_usu_perfil.set("operador")

# LOGIN - Layout moderno circular (estilo da imagem)
def tela_login():
    init_db()
    
    AZUL = "#1e40af"
    AZUL_BTN = "#1d4ed8"
    AZUL_BTN_HOVER = "#1e3a8a"
    
    login_root = tk.Tk()
    login_root.title("Acesse sua conta - Chaveiro Mestre")
    login_root.geometry("480x560")
    login_root.configure(bg=AZUL)
    login_root.resizable(False, False)
    
    login_root.update_idletasks()
    ws = login_root.winfo_screenwidth()
    hs = login_root.winfo_screenheight()
    x = (ws // 2) - 240
    y = (hs // 2) - 280
    login_root.geometry(f"480x560+{x}+{y}")
    
    # Card retangular branco (sem círculo)
    frame_card = tk.Frame(login_root, bg="white", highlightbackground="#93c5fd", highlightthickness=2)
    frame_card.place(relx=0.5, rely=0.5, anchor="center", width=400, height=500)
    
    frame_inner = tk.Frame(frame_card, bg="white")
    frame_inner.pack(fill="both", expand=True, padx=28, pady=24)
    
    # Logo / Título
    tk.Label(frame_inner, text="🏢", font=("Arial", 28), bg="white", fg=AZUL).pack(pady=(4, 0))
    tk.Label(frame_inner, text="Chaveiro Mestre", font=("Arial", 11, "bold"), bg="white", fg="#1e293b").pack(pady=(0, 6))
    
    tk.Label(frame_inner, text="Acesse sua conta", font=("Arial", 16, "bold"), bg="white", fg="#1e293b").pack(pady=(8, 16))
    
    # Campo usuário
    tk.Label(frame_inner, text="Usuário", font=("Arial", 10), bg="white", fg="#64748b", anchor="w").pack(fill="x")
    
    frame_user = tk.Frame(frame_inner, bg="#f1f5f9", highlightbackground="#cbd5e1", highlightthickness=1)
    frame_user.pack(fill="x", pady=(4, 12), ipady=2)
    
    #tk.Label(frame_user, text="👤", font=("Arial", 12), bg="#f1f5f9", fg="#64748b").pack(side="left", padx=(10, 6))
    entry_login = tk.Entry(frame_user, font=("Arial", 12), bd=0, bg="#f1f5f9", fg="#1e293b", insertbackground="#1e293b")
    entry_login.pack(side="left", fill="x", expand=True, ipady=8, padx=(0, 10))
    
    # Campo senha
    tk.Label(frame_inner, text="Senha", font=("Arial", 10), bg="white", fg="#64748b", anchor="w").pack(fill="x")
    
    frame_pass = tk.Frame(frame_inner, bg="#f1f5f9", highlightbackground="#cbd5e1", highlightthickness=1)
    frame_pass.pack(fill="x", pady=(4, 6), ipady=2)
    
    #tk.Label(frame_pass, text="🔒", font=("Arial", 12), bg="#f1f5f9", fg="#64748b").pack(side="left", padx=(10, 6))
    entry_senha = tk.Entry(frame_pass, font=("Arial", 12), bd=0, bg="#f1f5f9", fg="#1e293b", show="•", insertbackground="#1e293b")
    entry_senha.pack(side="left", fill="x", expand=True, ipady=8)
    
    # Toggle mostrar/ocultar senha
    show_pass = {"on": False}
    def toggle_senha():
        if show_pass["on"]:
            entry_senha.config(show="•")
            btn_eye.config(text="👁")
            show_pass["on"] = False
        else:
            entry_senha.config(show="")
            btn_eye.config(text="🙈")
            show_pass["on"] = True
    
    btn_eye = tk.Button(frame_pass, text="👁", font=("Arial", 11), bg="#f1f5f9", fg="#64748b",
                        bd=0, cursor="hand2", command=toggle_senha, activebackground="#e2e8f0")
    btn_eye.pack(side="right", padx=(4, 8))
    
    # Label de erro
    lbl_erro = tk.Label(frame_inner, text="", bg="white", fg="#ef4444", font=("Arial", 9, "bold"))
    lbl_erro.pack(pady=(2, 4))
    
    def fazer_login(event=None):
        login = entry_login.get().strip()
        senha = entry_senha.get().strip()
        if not login or not senha:
            lbl_erro.config(text="Preencha usuário e senha!")
            return
        conn = conectar()
        cur = conn.cursor()
        cur.execute("SELECT id, nome, login, perfil FROM usuarios WHERE login=? AND senha=? AND excluido=0",
                    (login, hash_senha(senha)))
        user = cur.fetchone()
        conn.close()
        if user:
            global usuario_logado
            usuario_logado = {"id": user[0], "nome": user[1], "login": user[2], "perfil": user[3]}
            login_root.destroy()
            criar_interface()
        else:
            lbl_erro.config(text="Usuário ou senha inválidos!")
            entry_senha.delete(0, tk.END)
    
    # Botão Entrar
    btn_login = tk.Button(frame_inner, text="Entrar", command=fazer_login,
                          bg=AZUL_BTN, fg="white", font=("Arial", 13, "bold"),
                          bd=0, relief="flat", cursor="hand2", activebackground=AZUL_BTN_HOVER,
                          activeforeground="white", height=2)
    btn_login.pack(fill="x", pady=(8, 10))
    
    # Link Esqueceu sua senha?
    def abrir_recuperar_senha():
        modal = tk.Toplevel(login_root)
        modal.title("Recuperar Senha")
        modal.geometry("420x420")
        modal.configure(bg="white")
        modal.transient(login_root)
        modal.grab_set()
        modal.resizable(False, False)
        modal.update_idletasks()
        mx = login_root.winfo_x() + 50
        my = login_root.winfo_y() + 80
        modal.geometry(f"+{mx}+{my}")
        
        tk.Label(modal, text="🔑 Recuperar Senha", font=("Arial", 14, "bold"), bg="white", fg=AZUL).pack(pady=(20, 5))
        tk.Label(modal, text="Informe o e-mail cadastrado no sistema\npara redefinir sua senha.",
                 font=("Arial", 10), bg="white", fg="#64748b", justify="center").pack(pady=(0, 15))
        
        tk.Label(modal, text="E-mail cadastrado:", font=("Arial", 10, "bold"), bg="white", fg="#334155", anchor="w").pack(fill="x", padx=40)
        entry_email_rec = tk.Entry(modal, font=("Arial", 12), bd=1, relief="solid", bg="#f8fafc")
        entry_email_rec.pack(fill="x", padx=40, ipady=6, pady=(4, 10))
        
        tk.Label(modal, text="Nova senha:", font=("Arial", 10, "bold"), bg="white", fg="#334155", anchor="w").pack(fill="x", padx=40)
        entry_nova = tk.Entry(modal, font=("Arial", 12), bd=1, relief="solid", bg="#f8fafc", show="•")
        entry_nova.pack(fill="x", padx=40, ipady=6, pady=(4, 10))
        
        tk.Label(modal, text="Confirmar nova senha:", font=("Arial", 10, "bold"), bg="white", fg="#334155", anchor="w").pack(fill="x", padx=40)
        entry_conf = tk.Entry(modal, font=("Arial", 12), bd=1, relief="solid", bg="#f8fafc", show="•")
        entry_conf.pack(fill="x", padx=40, ipady=6, pady=(4, 8))
        
        lbl_rec_msg = tk.Label(modal, text="", bg="white", fg="#ef4444", font=("Arial", 9, "bold"))
        lbl_rec_msg.pack(pady=4)
        
        def confirmar_recuperacao():
            email = entry_email_rec.get().strip().lower()
            nova = entry_nova.get().strip()
            conf = entry_conf.get().strip()
            if not email:
                lbl_rec_msg.config(text="Informe o e-mail cadastrado!", fg="#ef4444")
                return
            if not nova or len(nova) < 4:
                lbl_rec_msg.config(text="Nova senha deve ter pelo menos 4 caracteres!", fg="#ef4444")
                return
            if nova != conf:
                lbl_rec_msg.config(text="As senhas não coincidem!", fg="#ef4444")
                return
            
            conn = conectar()
            cur = conn.cursor()
            cur.execute("SELECT id, nome, login FROM usuarios WHERE lower(email)=? AND excluido=0", (email,))
            user = cur.fetchone()
            if not user:
                conn.close()
                lbl_rec_msg.config(text="E-mail não encontrado no sistema!", fg="#ef4444")
                return
            
            cur.execute("UPDATE usuarios SET senha=? WHERE id=?", (hash_senha(nova), user[0]))
            conn.commit()
            conn.close()
            
            lbl_rec_msg.config(text=f"✅ Senha alterada com sucesso!\nUsuário: {user[2]}", fg="#10b981")
            entry_login.delete(0, tk.END)
            entry_login.insert(0, user[2])
            modal.after(1800, modal.destroy)
        
        tk.Button(modal, text="Redefinir Senha", command=confirmar_recuperacao,
                  bg=AZUL_BTN, fg="white", font=("Arial", 11, "bold"), bd=0,
                  cursor="hand2", activebackground=AZUL_BTN_HOVER, height=2).pack(fill="x", padx=40, pady=(10, 8))
        
        tk.Button(modal, text="Cancelar", command=modal.destroy,
                  bg="#64748b", fg="white", font=("Arial", 10), bd=0,
                  cursor="hand2", height=1).pack(fill="x", padx=40, pady=(0, 15))
        
        entry_email_rec.focus()
        modal.bind("<Return>", lambda e: confirmar_recuperacao())
        modal.bind("<Escape>", lambda e: modal.destroy())
    
    lbl_esqueci = tk.Label(frame_inner, text="Esqueceu sua senha?", font=("Arial", 10),
                           bg="white", fg=AZUL, cursor="hand2")
    lbl_esqueci.pack(pady=(4, 10))
    lbl_esqueci.bind("<Button-1>", lambda e: abrir_recuperar_senha())

    # Dica discreta
    #tk.Label(frame_inner, text="admin / admin123  •  operador / operador123",font=("Arial", 8), bg="white", fg="#94a3b8").pack(pady=(8, 0))
    
    entry_senha.bind("<Return>", fazer_login)
    entry_login.bind("<Return>", lambda e: entry_senha.focus())
    
    entry_login.focus()
    login_root.mainloop()


def _aplicar_cores_notebook(notebook, cores_lista):
    """Abas coloridas no padrão das abas principais (sem ícones). Oculta as abas nativas do ttk."""
    style = ttk.Style()
    try:
        style.theme_use("clam")
    except Exception:
        pass
    style_nb = f"ColorTabs_{id(notebook)}.TNotebook"
    style_tab = f"ColorTabs_{id(notebook)}.TNotebook.Tab"
    try:
        # notebook sem altura de aba nativa
        style.layout(style_nb, [("TNotebook.client", {"sticky": "nswe"})])
        style.configure(style_nb, background=CORES["bg_light"], borderwidth=0)
        style.layout(style_tab, [])
        notebook.configure(style=style_nb)
    except Exception:
        try:
            notebook.configure(style=style_nb)
        except Exception:
            pass

    parent = notebook.master
    bar = tk.Frame(parent, bg=CORES["bg_light"])
    try:
        bar.pack(fill="x", padx=20, pady=(4, 0), before=notebook)
    except Exception:
        bar.pack(fill="x", padx=20, pady=(4, 0))

    botoes = []
    n = notebook.index("end")
    for i in range(n):
        texto = notebook.tab(i, "text")
        cor = cores_lista[i % len(cores_lista)]

        def _ir(idx=i):
            notebook.select(idx)
            _atualizar_botoes()

        btn = tk.Button(
            bar,
            text=texto,
            bg=cor,
            fg="white",
            font=("Arial", 9, "bold"),
            bd=0,
            relief="flat",
            padx=12,
            pady=6,
            cursor="hand2",
            command=_ir,
        )
        btn.pack(side="left", padx=3, pady=4)
        botoes.append((btn, cor))

    def _atualizar_botoes(event=None):
        try:
            sel = notebook.index(notebook.select())
        except Exception:
            sel = 0
        for i, (btn, cor) in enumerate(botoes):
            if i == sel:
                btn.config(bg=cor, fg="white")
            else:
                btn.config(bg="#e2e8f0", fg=cor)

    notebook.bind("<<NotebookTabChanged>>", _atualizar_botoes)
    _atualizar_botoes()


# INTERFACE
def criar_interface():
    global root, frame_abas, active_tab_key
    global tree_clientes, entry_cli_id, entry_cli_nome, entry_cli_cpf, entry_cli_tel, entry_cli_email, entry_cli_end, entry_cli_numero, entry_cli_bairro, entry_cli_cidade, entry_cli_cep, entry_busca_cliente
    global tree_fornecedores, entry_forn_id, entry_forn_nome, entry_forn_cnpj, entry_forn_tel, entry_forn_email, entry_forn_end, entry_forn_numero, entry_forn_bairro, entry_forn_cidade, entry_forn_cep, entry_busca_forn
    global tree_produtos, entry_prod_id, entry_prod_nome, entry_prod_codigo, entry_prod_desc, entry_prod_custo, entry_prod_venda, entry_prod_estoque, entry_prod_estmin, combo_prod_forn, entry_busca_prod
    global tree_estoque, tree_mov_estoque, combo_est_prod, entry_est_qtd, combo_est_tipo, entry_est_motivo, entry_est_filtro_prod
    global tree_venda_carrinho, tree_vendas, combo_venda_cliente, combo_venda_prod, entry_venda_qtd, entry_venda_desc, combo_venda_forma, entry_venda_obs, entry_venda_venc, lbl_venda_total, lbl_venda_total_titulo, lbl_venda_total_sub, entry_busca_venda
    global tree_caixa, entry_caixa_data, lbl_caixa_saldo, lbl_caixa_info, entry_caixa_valor, combo_caixa_tipo, entry_caixa_desc, combo_caixa_forma
    global tree_cp, tree_cp_aberto, tree_cp_pago, tree_cp_atraso, tree_cp_cancelado, entry_cp_id, combo_cp_forn, entry_cp_desc, entry_cp_valor, entry_cp_venc, combo_cp_forma, combo_cp_status, notebook_cp, lbl_cp_aberto, lbl_cp_pago, lbl_cp_atraso, lbl_cp_cancelado, frame_cp_cartao, combo_cp_parcelas, entry_cp_taxa, lbl_cp_parcela_info
    global tree_cr, tree_cr_aberto, tree_cr_recebido, tree_cr_atraso, tree_cr_cancelado, entry_cr_id, combo_cr_cliente, entry_cr_desc, entry_cr_valor, entry_cr_venc, combo_cr_forma, combo_cr_status, notebook_cr, lbl_cr_aberto, lbl_cr_recebido, lbl_cr_atraso, lbl_cr_cancelado, frame_cr_cartao, combo_cr_parcelas, entry_cr_taxa, lbl_cr_parcela_info
    global entry_rel_venda_ini, entry_rel_venda_fim, entry_rel_caixa_ini, entry_rel_caixa_fim
    global lbl_saldo_val, lbl_vendas_hoje_val, lbl_receber_val, lbl_pagar_val, lbl_estoque_baixo_val, lbl_clientes_val, lbl_produtos_val, tree_dashboard_vendas
    global tree_lixeira_clientes, tree_lixeira_fornecedores, tree_lixeira_produtos, tree_lixeira_vendas, tree_lixeira_caixa, tree_lixeira_cp, tree_lixeira_cr
    global tree_usuarios, entry_usu_id, entry_usu_nome, entry_usu_login, entry_usu_senha, entry_usu_email, combo_usu_perfil
    global telas, open_tabs
    global frame_cartao_credito, combo_parcelas, entry_taxa, combo_tipo_taxa, lbl_taxa_campo, lbl_cartao_total_base, lbl_cartao_total_com_taxa, lbl_cartao_parcela, lbl_cartao_taxa_info, frame_venda_add, frame_venda_botoes, frame_venda_carrinho
    
    root = tk.Tk()
    root.title(f"Chaveiro Mestre - {usuario_logado['nome']} ({usuario_logado['perfil'].upper()})")
    root.geometry("1450x850")
    root.configure(bg=CORES["bg_light"])
    try:
        root.state('zoomed')
    except:
        pass
    
    style = ttk.Style()
    style.theme_use('clam')
    style.configure("Treeview", rowheight=28, font=('Arial', 10), background="white", fieldbackground="white")
    style.configure("Treeview.Heading", font=('Arial', 10, 'bold'), background=CORES["bg_dark"], foreground="white")
    style.map("Treeview", background=[('selected', CORES["primary"])], foreground=[('selected', 'white')])
    
    top_bar = tk.Frame(root, bg=CORES["bg_top"], height=60)
    top_bar.pack(fill='x', side='top')
    top_bar.pack_propagate(False)
    tk.Label(top_bar, text="Chaveiro Mestre", bg=CORES["bg_top"], fg="white", font=('Arial', 14, 'bold')).pack(side='left', padx=20, pady=15)
    frame_user_top = tk.Frame(top_bar, bg=CORES["bg_top"])
    frame_user_top.pack(side='right', padx=20, pady=10)
    perfil_cor = CORES["success"] if eh_admin() else CORES["warning"]
    tk.Label(frame_user_top, text=f"{usuario_logado['nome']}", bg=CORES["bg_top"], fg="white", font=('Arial', 10, 'bold')).pack(side='left', padx=5)
    tk.Label(frame_user_top, text=f"{usuario_logado['perfil'].upper()}", bg=perfil_cor, fg="white", font=('Arial', 8, 'bold'), padx=8, pady=2).pack(side='left', padx=5)
    tk.Label(frame_user_top, text=f"{hoje_br_completo()}", bg=CORES["bg_top"], fg="#94a3b8", font=('Arial', 9)).pack(side='left', padx=15)
    def logout():
        res = perguntar_backup_ao_sair("Sair / Logout")
        if res.get("acao") == "cancel":
            return
        if res.get("acao") == "ok" and res.get("caminho"):
            try:
                mostrar_info("Backup salvo em:\n" + str(res["caminho"]), "Backup")
            except Exception:
                pass
        try:
            root.destroy()
        except Exception:
            pass
        tela_login()

    def ao_fechar_janela():
        res = perguntar_backup_ao_sair("Encerrar sistema")
        if res.get("acao") == "cancel":
            return
        if res.get("acao") == "ok" and res.get("caminho"):
            try:
                mostrar_info("Backup salvo em:\n" + str(res["caminho"]), "Backup")
            except Exception:
                pass
        try:
            root.destroy()
        except Exception:
            pass
    try:
        root.protocol("WM_DELETE_WINDOW", ao_fechar_janela)
    except Exception:
        pass

    tk.Button(frame_user_top, text="Sair", command=logout, bg=CORES["danger"], fg="white", font=('Arial', 9, 'bold'), bd=0, padx=12, pady=4, cursor='hand2').pack(side='left', padx=10)
    
    container = tk.Frame(root, bg=CORES["bg_light"])
    container.pack(fill='both', expand=True)
    
    menu_lateral = tk.Frame(container, bg=CORES["bg_dark"], width=210)
    menu_lateral.pack(side='left', fill='y')
    menu_lateral.pack_propagate(False)
    tk.Label(menu_lateral, text="MENU", bg=CORES["bg_dark"], fg="#94a3b8", font=('Arial', 8, 'bold'), anchor='w', padx=12, pady=8).pack(fill='x')
    scrollable_frame = tk.Frame(menu_lateral, bg=CORES["bg_dark"])
    scrollable_frame.pack(side="left", fill="both", expand=True)
    
    area_direita = tk.Frame(container, bg=CORES["bg_light"])
    area_direita.pack(side='right', fill='both', expand=True)
    
    frame_abas = tk.Frame(area_direita, bg=CORES["bg_tab"], height=45, bd=0)
    frame_abas.pack(fill='x', side='top')
    frame_abas.pack_propagate(False)
    #tk.Label(frame_abas, text="📑 Abas Abertas:", bg=CORES["bg_tab"], fg=CORES["text_gray"], font=('Arial', 9, 'bold')).pack(side='left', padx=10, pady=10)
    
    frame_abas_container = tk.Frame(frame_abas, bg=CORES["bg_tab"])
    frame_abas_container.pack(side='left', fill='both', expand=True, padx=5, pady=5)
    
    frame_conteudo = tk.Frame(area_direita, bg=CORES["bg_light"])
    frame_conteudo.pack(fill='both', expand=True)
    
    telas = {}
    open_tabs = {}
    active_tab_key = None
    
    def criar_tela(nome):
        f = tk.Frame(frame_conteudo, bg=CORES["bg_light"])
        f.grid(row=0, column=0, sticky='nsew')
        frame_conteudo.grid_rowconfigure(0, weight=1)
        frame_conteudo.grid_columnconfigure(0, weight=1)
        telas[nome] = f
        return f
    
    tela_dashboard = criar_tela("dashboard")
    tela_clientes = criar_tela("clientes")
    tela_fornecedores = criar_tela("fornecedores")
    tela_produtos = criar_tela("produtos")
    tela_estoque = criar_tela("estoque")
    tela_vendas = criar_tela("vendas")
    tela_caixa = criar_tela("caixa")
    tela_cp = criar_tela("contas_pagar")
    tela_cr = criar_tela("contas_receber")
    tela_relatorios = criar_tela("relatorios")
    tela_lixeira = criar_tela("excluidos")
    tela_usuarios = criar_tela("usuarios")
    tela_backup = criar_tela("backup")

    # Estilo base (cores por notebook aplicadas em _aplicar_cores_notebook)
    try:
        _style_nb = ttk.Style()
        _style_nb.theme_use("clam")
    except Exception:
        pass
    

    # Cores das abas abertas por módulo
    CORES_ABAS = {
        "dashboard": "#3b82f6",
        "clientes": "#8b5cf6",
        "fornecedores": "#f59e0b",
        "produtos": "#10b981",
        "estoque": "#06b6d4",
        "vendas": "#ec4899",
        "caixa": "#22c55e",
        "contas_pagar": "#ef4444",
        "contas_receber": "#14b8a6",
        "relatorios": "#6366f1",
        "excluidos": "#64748b",
        "usuarios": "#0ea5e9",
        "backup": "#7c3aed",
    }

    def mostrar_tela(nome):
        global active_tab_key
        for t in telas.values():
            t.grid_remove()
        telas[nome].grid()
        active_tab_key = nome
        for key, tab_info in open_tabs.items():
            btn = tab_info['btn']
            btn_close = tab_info.get('btn_close')
            cor = CORES_ABAS.get(key, CORES["primary"])
            if key == nome:
                btn.config(bg=cor, fg="white", relief='flat', bd=0)
                if btn_close:
                    btn_close.config(bg=cor, fg="white")
                tab_info['frame'].config(bg=cor, bd=0, relief='flat')
            else:
                # aba inativa: fundo claro com texto na cor do módulo
                btn.config(bg="#e2e8f0", fg=cor, relief='flat', bd=0)
                if btn_close:
                    btn_close.config(bg="#e2e8f0", fg="#94a3b8")
                tab_info['frame'].config(bg="#e2e8f0", bd=0, relief='flat')
        if nome == "dashboard":
            atualizar_dashboard()
        elif nome == "excluidos":
            listar_todas_lixeiras()
        elif nome == "usuarios":
            listar_usuarios()
        elif nome == "backup":
            try:
                atualizar_lista_backups()
            except Exception:
                pass
    
    def fechar_aba(key):
        if key not in open_tabs:
            return
        open_tabs[key]['frame'].destroy()
        del open_tabs[key]
        if active_tab_key == key:
            if open_tabs:
                ultimo = list(open_tabs.keys())[-1]
                mostrar_tela(ultimo)
            else:
                abrir_aba("dashboard", "Dashboard", "📊")
    
    def abrir_aba(key, titulo, icone=None):
        if key in open_tabs:
            mostrar_tela(key)
            return
        cor = CORES_ABAS.get(key, CORES["primary"])
        tab_frame = tk.Frame(frame_abas_container, bg=cor, bd=0, relief='flat')
        tab_frame.pack(side='left', padx=3, pady=3)
        btn = tk.Button(
            tab_frame,
            text=titulo,
            bg=cor,
            fg="white",
            font=('Arial', 9, 'bold'),
            bd=0,
            relief='flat',
            padx=10,
            pady=3,
            cursor='hand2',
            command=lambda k=key: mostrar_tela(k),
        )
        btn.pack(side='left')
        btn_close = tk.Button(
            tab_frame,
            text="×",
            bg=cor,
            fg="white",
            font=('Arial', 7),
            bd=0,
            relief='flat',
            padx=3,
            pady=1,
            cursor='hand2',
            command=lambda k=key: fechar_aba(k),
        )
        btn_close.pack(side='left', padx=(0, 4), pady=2)
        open_tabs[key] = {"frame": tab_frame, "btn": btn, "btn_close": btn_close, "titulo": titulo, "cor": cor}
        mostrar_tela(key)
    
    botoes_menu = {}
    menu_itens = [
        ("dashboard", "Dashboard", "📊"),
        ("clientes", "Clientes", "👥"),
        ("fornecedores", "Fornecedores", "🏭"),
        ("produtos", "Produtos", "📦"),
        ("estoque", "Estoque", "📊"),
        ("vendas", "Vendas / PDV", "🛒"),
        ("caixa", "Caixa", "💰"),
        ("contas_pagar", "Contas a Pagar", "📤"),
        ("contas_receber", "Contas a Receber", "📥"),
        ("relatorios", "Relatórios", "📑"),
        ("excluidos", "Lixeira", "🗑️"),
        ("backup", "Backup / Restaurar", "💾"),
        ("usuarios", "Usuários", "👤"),
    ]
    
    for key, label, icone in menu_itens:
        if key == "usuarios" and not eh_admin():
            continue
        btn = tk.Button(
            scrollable_frame,
            text=label,
            anchor='w',
            bg=CORES["bg_dark"],
            fg="#cbd5e1",
            font=('Arial', 9, 'bold'),
            bd=0,
            relief='flat',
            padx=12,
            pady=5,
            cursor='hand2',
            activebackground=CORES["bg_dark_hover"],
            activeforeground="white",
            command=lambda k=key, t=label: abrir_aba(k, t),
        )
        btn.pack(fill='x', pady=0, padx=2)
        botoes_menu[key] = btn
        def on_enter(e, b=btn):
            b.config(bg=CORES["bg_dark_hover"])
        def on_leave(e, b=btn):
            b.config(bg=CORES["bg_dark"])
        btn.bind("<Enter>", on_enter)
        btn.bind("<Leave>", on_leave)
    
    # DASHBOARD
    frame_cards = tk.Frame(tela_dashboard, bg=CORES["bg_light"])
    frame_cards.pack(fill='x', padx=20, pady=20)
    def criar_card(parent, titulo, row, col, bg_color, text_color="#1e293b"):
        f = tk.Frame(parent, bg=bg_color, bd=0, relief='flat')
        f.grid(row=row, column=col, padx=10, pady=8, sticky='ew', ipadx=15, ipady=15)
        border = tk.Frame(f, bg=text_color, width=5)
        border.pack(side='left', fill='y')
        inner = tk.Frame(f, bg=bg_color)
        inner.pack(side='left', fill='both', expand=True, padx=12, pady=5)
        tk.Label(inner, text=titulo, bg=bg_color, font=('Arial', 9, 'bold'), fg="#475569", anchor='w').pack(anchor='w', fill='x')
        lbl_val = tk.Label(inner, text="R$ 0,00", bg=bg_color, font=('Arial', 16, 'bold'), fg=text_color, anchor='w')
        lbl_val.pack(anchor='w', pady=(8,0), fill='x')
        parent.grid_columnconfigure(col, weight=1)
        return lbl_val
    lbl_saldo_val = criar_card(frame_cards, "SALDO CAIXA", 0, 0, CORES["card_green"], "#15803d")
    lbl_vendas_hoje_val = criar_card(frame_cards, "VENDAS HOJE", 0, 1, CORES["card_blue"], "#1d4ed8")
    lbl_receber_val = criar_card(frame_cards, "A RECEBER", 0, 2, CORES["card_yellow"], "#b45309")
    lbl_pagar_val = criar_card(frame_cards, "A PAGAR", 0, 3, CORES["card_red"], "#b91c1c")
    frame_cards2 = tk.Frame(tela_dashboard, bg=CORES["bg_light"])
    frame_cards2.pack(fill='x', padx=20, pady=5)
    lbl_estoque_baixo_val = criar_card(frame_cards2, "ESTOQUE BAIXO", 0, 0, CORES["card_orange"], "#c2410c")
    lbl_clientes_val = criar_card(frame_cards2, "CLIENTES", 0, 1, CORES["card_purple"], "#7c3aed")
    lbl_produtos_val = criar_card(frame_cards2, "PRODUTOS", 0, 2, "#ccfbf1", "#0f766e")
    frame_dash_vendas = tk.LabelFrame(tela_dashboard, text="Últimas Vendas", font=('Arial', 11, 'bold'), bg=CORES["bg_white"], fg=CORES["text_dark"], bd=1, relief='solid')
    frame_dash_vendas.pack(fill='both', expand=True, padx=20, pady=10)
    cols_dash = ("ID","Data","Cliente","Total","Forma","Status")
    tree_dashboard_vendas = ttk.Treeview(frame_dash_vendas, columns=cols_dash, show='headings', height=12)
    for c in cols_dash:
        tree_dashboard_vendas.heading(c, text=c)
        tree_dashboard_vendas.column(c, width=120)
    tree_dashboard_vendas.column("Data", width=150)
    tree_dashboard_vendas.column("Cliente", width=200)
    tree_dashboard_vendas.pack(fill='both', expand=True, padx=10, pady=10)
    tk.Button(frame_dash_vendas, text="🔄 Atualizar Dashboard", command=atualizar_dashboard, bg=CORES["primary"], fg="white", font=('Arial', 10, 'bold'), bd=0, padx=15, pady=6).pack(pady=8)
    
    # CLIENTES — labels com largura fixa para alinhamento
    frame_cli_form = tk.LabelFrame(tela_clientes, text=" Cadastro de Cliente", font=('Arial', 11, 'bold'), bg=CORES["bg_white"], padx=15, pady=15)
    frame_cli_form.pack(fill='x', padx=20, pady=10)
    for _c in range(6):
        frame_cli_form.columnconfigure(_c, weight=0)
    def _lbl_cli(txt, r, c, bold=False):
        fnt = ('Arial', 9, 'bold') if bold else ('Arial', 9)
        tk.Label(frame_cli_form, text=txt, bg=CORES["bg_white"], fg=CORES["text_dark"], font=fnt,
                 width=12, anchor='e').grid(row=r, column=c, sticky='e', padx=(6, 4), pady=4)
    # Linha 0
    _lbl_cli("ID:", 0, 0)
    entry_cli_id = tk.Entry(frame_cli_form, width=10)
    entry_cli_id.grid(row=0, column=1, padx=4, pady=4, sticky='w')
    _lbl_cli("Nome*:", 0, 2, True)
    entry_cli_nome = tk.Entry(frame_cli_form, width=30)
    entry_cli_nome.grid(row=0, column=3, padx=4, pady=4, sticky='w')
    _lbl_cli("CPF/CNPJ*:", 0, 4, True)
    entry_cli_cpf = tk.Entry(frame_cli_form, width=20)
    entry_cli_cpf.grid(row=0, column=5, padx=4, pady=4, sticky='w')
    entry_cli_cpf.bind("<KeyRelease>", formatar_cpf_cnpj)
    # Linha 1
    _lbl_cli("Telefone*:", 1, 0, True)
    entry_cli_tel = tk.Entry(frame_cli_form, width=18)
    entry_cli_tel.grid(row=1, column=1, padx=4, pady=4, sticky='w')
    entry_cli_tel.bind("<KeyRelease>", formatar_telefone)
    _lbl_cli("Email:", 1, 2)
    entry_cli_email = tk.Entry(frame_cli_form, width=30)
    entry_cli_email.grid(row=1, column=3, padx=4, pady=4, sticky='w')
    _lbl_cli("CEP:", 1, 4)
    frame_cli_cep = tk.Frame(frame_cli_form, bg=CORES["bg_white"])
    frame_cli_cep.grid(row=1, column=5, padx=4, pady=4, sticky='w')
    entry_cli_cep = tk.Entry(frame_cli_cep, width=12)
    entry_cli_cep.pack(side='left')
    entry_cli_cep.bind("<KeyRelease>", formatar_cep)
    entry_cli_cep.bind("<Return>", lambda e: buscar_cep_cliente())
    tk.Button(frame_cli_cep, text="🔍 CEP", command=buscar_cep_cliente, bg=CORES["primary"], fg="white",
              font=('Arial', 8, 'bold'), bd=0, padx=8, pady=2, cursor='hand2').pack(side='left', padx=(6, 0))
    # Linha 2
    _lbl_cli("Endereço:", 2, 0)
    entry_cli_end = tk.Entry(frame_cli_form, width=30)
    entry_cli_end.grid(row=2, column=1, columnspan=2, padx=4, pady=4, sticky='ew')
    _lbl_cli("Número:", 2, 3)
    entry_cli_numero = tk.Entry(frame_cli_form, width=12)
    entry_cli_numero.grid(row=2, column=4, padx=4, pady=4, sticky='w')
    # Linha 3
    _lbl_cli("Bairro:", 3, 0)
    entry_cli_bairro = tk.Entry(frame_cli_form, width=22)
    entry_cli_bairro.grid(row=3, column=1, padx=4, pady=4, sticky='w')
    _lbl_cli("Cidade:", 3, 2)
    entry_cli_cidade = tk.Entry(frame_cli_form, width=30)
    entry_cli_cidade.grid(row=3, column=3, columnspan=2, padx=4, pady=4, sticky='ew')
    # Botões
    frame_cli_btn = tk.Frame(frame_cli_form, bg=CORES["bg_white"])
    frame_cli_btn.grid(row=4, column=0, columnspan=6, pady=12, sticky='w')
    tk.Button(frame_cli_btn, text="Salvar", command=salvar_cliente, bg=CORES["success"], fg="white", width=14, font=('Arial', 10, 'bold'), bd=0, pady=6).pack(side='left', padx=5)
    tk.Button(frame_cli_btn, text="Limpar", command=limpar_form_cliente, bg="#64748b", fg="white", width=12, bd=0, pady=6).pack(side='left', padx=5)
    frame_cli_busca = tk.Frame(tela_clientes, bg=CORES["bg_light"])
    frame_cli_busca.pack(fill='x', padx=20, pady=5)
    tk.Label(frame_cli_busca, text="🔍 Buscar (digite o cliente):", bg=CORES["bg_light"], font=('Arial', 9, 'bold'), fg=CORES["text_dark"]).pack(side='left')
    entry_busca_cliente = tk.Entry(frame_cli_busca, width=40, font=('Arial', 10))
    entry_busca_cliente.pack(side='left', padx=10, ipady=3)
    frame_cli_tree = tk.Frame(tela_clientes, bg=CORES["bg_white"], bd=1, relief='solid')
    frame_cli_tree.pack(fill='both', expand=True, padx=20, pady=10)
    cols_cli = ("ID", "Nome", "CPF/CNPJ", "Telefone", "Email", "Cidade", "Sel")
    tree_clientes = ttk.Treeview(frame_cli_tree, columns=cols_cli, show='headings')
    for c in cols_cli:
        tree_clientes.heading(c, text=c)
    tree_clientes.column("ID", width=50)
    tree_clientes.column("Nome", width=220)
    tree_clientes.column("Cidade", width=140)
    tree_clientes.column("Sel", width=40, anchor='center', stretch=False)
    tree_clientes.heading("Sel", text="☐")
    tree_clientes.pack(fill='both', expand=True, side='left', padx=5, pady=5)
    sb_cli = ttk.Scrollbar(frame_cli_tree, orient='vertical', command=tree_clientes.yview)
    tree_clientes.configure(yscrollcommand=sb_cli.set)
    sb_cli.pack(side='right', fill='y')
    tree_clientes.bind("<Double-1>", editar_cliente)
    habilitar_selecao_multipla(tree_clientes)
    criar_barra_selecao_multipla(tela_clientes, tree_clientes, excluir_clientes_em_massa).pack(
        fill='x', padx=20, pady=(0, 8), before=frame_cli_tree)
    criar_controle_paginacao(tela_clientes, "clientes", bg=CORES["bg_light"]).pack(
        fill='x', padx=20, pady=(0, 6), before=frame_cli_tree)
    
    # FORNECEDORES — labels alinhadas
    frame_forn_form = tk.LabelFrame(tela_fornecedores, text="Cadastro de Fornecedor", font=('Arial', 11, 'bold'), bg=CORES["bg_white"], padx=15, pady=15)
    frame_forn_form.pack(fill='x', padx=20, pady=10)
    def _lbl_forn(txt, r, c, bold=False):
        fnt = ('Arial', 9, 'bold') if bold else ('Arial', 9)
        tk.Label(frame_forn_form, text=txt, bg=CORES["bg_white"], fg=CORES["text_dark"], font=fnt,
                 width=12, anchor='e').grid(row=r, column=c, sticky='e', padx=(6, 4), pady=4)
    _lbl_forn("ID:", 0, 0)
    entry_forn_id = tk.Entry(frame_forn_form, width=10)
    entry_forn_id.grid(row=0, column=1, padx=4, pady=4, sticky='w')
    _lbl_forn("Nome*:", 0, 2, True)
    entry_forn_nome = tk.Entry(frame_forn_form, width=30)
    entry_forn_nome.grid(row=0, column=3, padx=4, pady=4, sticky='w')
    _lbl_forn("CNPJ/CPF*:", 0, 4, True)
    entry_forn_cnpj = tk.Entry(frame_forn_form, width=20)
    entry_forn_cnpj.grid(row=0, column=5, padx=4, pady=4, sticky='w')
    entry_forn_cnpj.bind("<KeyRelease>", formatar_cpf_cnpj)
    _lbl_forn("Telefone*:", 1, 0, True)
    entry_forn_tel = tk.Entry(frame_forn_form, width=18)
    entry_forn_tel.grid(row=1, column=1, padx=4, pady=4, sticky='w')
    entry_forn_tel.bind("<KeyRelease>", formatar_telefone)
    _lbl_forn("Email:", 1, 2)
    entry_forn_email = tk.Entry(frame_forn_form, width=30)
    entry_forn_email.grid(row=1, column=3, padx=4, pady=4, sticky='w')
    _lbl_forn("CEP:", 1, 4)
    frame_forn_cep = tk.Frame(frame_forn_form, bg=CORES["bg_white"])
    frame_forn_cep.grid(row=1, column=5, padx=4, pady=4, sticky='w')
    entry_forn_cep = tk.Entry(frame_forn_cep, width=12)
    entry_forn_cep.pack(side='left')
    entry_forn_cep.bind("<KeyRelease>", formatar_cep)
    entry_forn_cep.bind("<Return>", lambda e: buscar_cep_fornecedor())
    tk.Button(frame_forn_cep, text="🔍 CEP", command=buscar_cep_fornecedor, bg=CORES["primary"], fg="white",
              font=('Arial', 8, 'bold'), bd=0, padx=8, pady=2, cursor='hand2').pack(side='left', padx=(6, 0))
    _lbl_forn("Endereço:", 2, 0)
    entry_forn_end = tk.Entry(frame_forn_form, width=30)
    entry_forn_end.grid(row=2, column=1, columnspan=2, padx=4, pady=4, sticky='ew')
    _lbl_forn("Número:", 2, 3)
    entry_forn_numero = tk.Entry(frame_forn_form, width=12)
    entry_forn_numero.grid(row=2, column=4, padx=4, pady=4, sticky='w')
    _lbl_forn("Bairro:", 3, 0)
    entry_forn_bairro = tk.Entry(frame_forn_form, width=22)
    entry_forn_bairro.grid(row=3, column=1, padx=4, pady=4, sticky='w')
    _lbl_forn("Cidade:", 3, 2)
    entry_forn_cidade = tk.Entry(frame_forn_form, width=30)
    entry_forn_cidade.grid(row=3, column=3, columnspan=2, padx=4, pady=4, sticky='ew')
    frame_forn_btn = tk.Frame(frame_forn_form, bg=CORES["bg_white"])
    frame_forn_btn.grid(row=4, column=0, columnspan=6, pady=12, sticky='w')
    tk.Button(frame_forn_btn, text="Salvar", command=salvar_fornecedor, bg=CORES["success"], fg="white", width=14, font=('Arial', 10, 'bold'), bd=0, pady=6).pack(side='left', padx=5)
    tk.Button(frame_forn_btn, text="Limpar", command=limpar_form_fornecedor, bg="#64748b", fg="white", width=12, bd=0, pady=6).pack(side='left', padx=5)
    frame_forn_busca = tk.Frame(tela_fornecedores, bg=CORES["bg_light"])
    frame_forn_busca.pack(fill='x', padx=20, pady=5)
    tk.Label(frame_forn_busca, text="🔍 Buscar (digite o fornecedor):", bg=CORES["bg_light"], font=('Arial', 9, 'bold'), fg=CORES["text_dark"]).pack(side='left')
    entry_busca_forn = tk.Entry(frame_forn_busca, width=40, font=('Arial', 10))
    entry_busca_forn.pack(side='left', padx=10, ipady=3)
    frame_forn_tree = tk.Frame(tela_fornecedores, bg=CORES["bg_white"], bd=1, relief='solid')
    frame_forn_tree.pack(fill='both', expand=True, padx=20, pady=10)
    cols_forn = ("ID", "Nome", "CNPJ/CPF", "Telefone", "Email", "Cidade", "Sel")
    tree_fornecedores = ttk.Treeview(frame_forn_tree, columns=cols_forn, show='headings')
    for c in cols_forn:
        tree_fornecedores.heading(c, text=c)
    tree_fornecedores.column("ID", width=50)
    tree_fornecedores.column("Nome", width=220)
    tree_fornecedores.column("Cidade", width=140)
    tree_fornecedores.column("Sel", width=40, anchor='center', stretch=False)
    tree_fornecedores.heading("Sel", text="☐")
    tree_fornecedores.pack(fill='both', expand=True, side='left', padx=5, pady=5)
    sb_forn = ttk.Scrollbar(frame_forn_tree, orient='vertical', command=tree_fornecedores.yview)
    tree_fornecedores.configure(yscrollcommand=sb_forn.set)
    sb_forn.pack(side='right', fill='y')
    tree_fornecedores.bind("<Double-1>", editar_fornecedor)
    habilitar_selecao_multipla(tree_fornecedores)
    criar_barra_selecao_multipla(tela_fornecedores, tree_fornecedores, excluir_fornecedores_em_massa).pack(
        fill='x', padx=20, pady=(0, 8), before=frame_forn_tree)
    criar_controle_paginacao(tela_fornecedores, "fornecedores", bg=CORES["bg_light"]).pack(
        fill='x', padx=20, pady=(0, 6), before=frame_forn_tree)
    
    # PRODUTOS
    frame_prod_form = tk.LabelFrame(tela_produtos, text="Cadastro de Produto", font=('Arial', 11, 'bold'), bg=CORES["bg_white"], padx=15, pady=15)
    frame_prod_form.pack(fill='x', padx=20, pady=10)
    def _lbl_prod(txt, r, c, bold=False):
        fnt = ('Arial', 9, 'bold') if bold else ('Arial', 9)
        tk.Label(frame_prod_form, text=txt, bg=CORES["bg_white"], fg=CORES["text_dark"], font=fnt,
                 width=12, anchor='e').grid(row=r, column=c, sticky='e', padx=(6, 4), pady=4)
    _lbl_prod("ID:", 0, 0)
    entry_prod_id = tk.Entry(frame_prod_form, width=10)
    entry_prod_id.grid(row=0, column=1, padx=4, pady=4, sticky='w')
    _lbl_prod("Código*:", 0, 2, True)
    entry_prod_codigo = tk.Entry(frame_prod_form, width=14)
    entry_prod_codigo.grid(row=0, column=3, padx=4, pady=4, sticky='w')
    _lbl_prod("Nome*:", 0, 4, True)
    entry_prod_nome = tk.Entry(frame_prod_form, width=32)
    entry_prod_nome.grid(row=0, column=5, columnspan=2, padx=4, pady=4, sticky='w')
    _lbl_prod("Custo*:", 1, 0, True)
    entry_prod_custo = tk.Entry(frame_prod_form, width=12)
    entry_prod_custo.grid(row=1, column=1, padx=4, pady=4, sticky='w')
    _lbl_prod("Venda*:", 1, 2, True)
    entry_prod_venda = tk.Entry(frame_prod_form, width=12)
    entry_prod_venda.grid(row=1, column=3, padx=4, pady=4, sticky='w')
    _lbl_prod("Estoque*:", 1, 4, True)
    entry_prod_estoque = tk.Entry(frame_prod_form, width=12)
    entry_prod_estoque.grid(row=1, column=5, padx=4, pady=4, sticky='w')
    _lbl_prod("Est.Min:", 1, 6)
    entry_prod_estmin = tk.Entry(frame_prod_form, width=10)
    entry_prod_estmin.grid(row=1, column=7, padx=4, pady=4, sticky='w')
    entry_prod_estmin.insert(0, "5")
    _lbl_prod("Fornecedor*:", 2, 0, True)
    combo_prod_forn = ttk.Combobox(frame_prod_form, width=28)
    combo_prod_forn.grid(row=2, column=1, columnspan=2, padx=4, pady=4, sticky='w')
    _lbl_prod("Descrição:", 2, 3)
    entry_prod_desc = tk.Entry(frame_prod_form, width=40)
    entry_prod_desc.grid(row=2, column=4, columnspan=4, padx=4, pady=4, sticky='w')
    frame_prod_btn = tk.Frame(frame_prod_form, bg=CORES["bg_white"])
    frame_prod_btn.grid(row=3, column=0, columnspan=8, pady=12, sticky='w')
    tk.Button(frame_prod_btn, text="Salvar", command=salvar_produto, bg=CORES["success"], fg="white", width=14, font=('Arial', 10, 'bold'), bd=0, pady=6).pack(side='left', padx=5)
    tk.Button(frame_prod_btn, text="Limpar", command=limpar_form_produto, bg="#64748b", fg="white", width=12, bd=0, pady=6).pack(side='left', padx=5)
    frame_prod_busca = tk.Frame(tela_produtos, bg=CORES["bg_light"])
    frame_prod_busca.pack(fill='x', padx=20, pady=5)
    tk.Label(frame_prod_busca, text="🔍 Buscar (digite o produto):", bg=CORES["bg_light"], font=('Arial', 9, 'bold'), fg=CORES["text_dark"]).pack(side='left')
    entry_busca_prod = tk.Entry(frame_prod_busca, width=40, font=('Arial', 10))
    entry_busca_prod.pack(side='left', padx=10, ipady=3)
    frame_prod_tree = tk.Frame(tela_produtos, bg=CORES["bg_white"], bd=1, relief='solid')
    frame_prod_tree.pack(fill='both', expand=True, padx=20, pady=10)
    cols_prod = ("ID", "Código", "Nome", "Preço Venda", "Estoque", "Fornecedor", "Sel")
    tree_produtos = ttk.Treeview(frame_prod_tree, columns=cols_prod, show='headings')
    for c in cols_prod:
        tree_produtos.heading(c, text=c)
    tree_produtos.column("ID", width=40)
    tree_produtos.column("Código", width=100)
    tree_produtos.column("Nome", width=280)
    tree_produtos.column("Preço Venda", width=100)
    tree_produtos.column("Estoque", width=80)
    tree_produtos.column("Sel", width=40, anchor='center', stretch=False)
    tree_produtos.heading("Sel", text="☐")
    tree_produtos.pack(fill='both', expand=True, side='left', padx=5, pady=5)
    sb_prod = ttk.Scrollbar(frame_prod_tree, orient='vertical', command=tree_produtos.yview)
    tree_produtos.configure(yscrollcommand=sb_prod.set)
    sb_prod.pack(side='right', fill='y')
    tree_produtos.bind("<Double-1>", editar_produto)
    habilitar_selecao_multipla(tree_produtos)
    criar_barra_selecao_multipla(tela_produtos, tree_produtos, excluir_produtos_em_massa).pack(
        fill='x', padx=20, pady=(0, 8), before=frame_prod_tree)
    criar_controle_paginacao(tela_produtos, "produtos", bg=CORES["bg_light"]).pack(
        fill='x', padx=20, pady=(0, 6), before=frame_prod_tree)
    tree_produtos.tag_configure('baixo', background='#fef3c7')
    tree_produtos.tag_configure('zerado', background='#fee2e2')
    
    # ESTOQUE
    frame_est_top = tk.Frame(tela_estoque, bg=CORES["bg_light"])
    frame_est_top.pack(fill='x', padx=20, pady=10)
    frame_est_mov = tk.LabelFrame(frame_est_top, text="Movimentação Manual de Estoque", font=('Arial', 11, 'bold'), bg=CORES["bg_white"], padx=15, pady=15)
    frame_est_mov.pack(side='left', fill='x', expand=True, padx=5)
    tk.Label(frame_est_mov, text="Produto* (digite para buscar):", bg=CORES["bg_white"], font=('Arial', 9, 'bold'), fg=CORES["text_dark"]).grid(row=0,column=0, sticky='w')
    combo_est_prod = ttk.Combobox(frame_est_mov, width=40)
    combo_est_prod.grid(row=0,column=1, padx=5, pady=4)
    tk.Label(frame_est_mov, text="Tipo*:", bg=CORES["bg_white"], fg=CORES["text_dark"]).grid(row=0,column=2, sticky='w')
    combo_est_tipo = ttk.Combobox(frame_est_mov, width=12, values=["entrada","saida","ajuste"])
    combo_est_tipo.grid(row=0,column=3, padx=5, pady=4)
    combo_est_tipo.set("entrada")
    tk.Label(frame_est_mov, text="Qtd*:", bg=CORES["bg_white"], fg=CORES["text_dark"]).grid(row=1,column=0, sticky='w')
    entry_est_qtd = tk.Entry(frame_est_mov, width=12)
    entry_est_qtd.grid(row=1,column=1, padx=5, pady=4, sticky='w')
    tk.Label(frame_est_mov, text="Motivo:", bg=CORES["bg_white"]).grid(row=1,column=2, sticky='w')
    entry_est_motivo = tk.Entry(frame_est_mov, width=28)
    entry_est_motivo.grid(row=1,column=3, padx=5, pady=4)
    tk.Button(frame_est_mov, text="✅ Movimentar", command=movimentar_estoque, bg=CORES["primary"], fg="white", font=('Arial', 9, 'bold'), bd=0, padx=12, pady=5).grid(row=1,column=4, padx=10)
    frame_est_filtro = tk.Frame(tela_estoque, bg=CORES["bg_light"])
    frame_est_filtro.pack(fill='x', padx=20, pady=5)
    tk.Label(frame_est_filtro, text="🔍 Filtrar movimentação", bg=CORES["bg_light"], font=('Arial', 9, 'bold'), fg=CORES["text_dark"]).pack(side='left')
    entry_est_filtro_prod = tk.Entry(frame_est_filtro, width=30)
    entry_est_filtro_prod.pack(side='left', padx=10, ipady=3)
    paned_est = tk.PanedWindow(tela_estoque, orient='vertical', bg=CORES["bg_light"])
    paned_est.pack(fill='both', expand=True, padx=20, pady=5)
    frame_est_tree = tk.LabelFrame(paned_est, text="Saldo Atual de Estoque", font=('Arial', 10, 'bold'), bg=CORES["bg_white"])
    paned_est.add(frame_est_tree, height=250)
    cols_est = ("ID", "Código", "Nome", "Atual", "Mínimo", "Status")
    tree_estoque = ttk.Treeview(frame_est_tree, columns=cols_est, show='headings')
    for c in cols_est:
        tree_estoque.heading(c, text=c)
    tree_estoque.column("ID", width=40)
    tree_estoque.column("Código", width=100)
    tree_estoque.column("Nome", width=280)
    tree_estoque.pack(fill='both', expand=True, side='left', padx=5, pady=5)
    sb_est = ttk.Scrollbar(frame_est_tree, orient='vertical', command=tree_estoque.yview)
    tree_estoque.configure(yscrollcommand=sb_est.set)
    sb_est.pack(side='right', fill='y')
    tree_estoque.tag_configure('baixo', background='#fef3c7')
    tree_estoque.tag_configure('zerado', background='#fee2e2')
    criar_controle_paginacao(frame_est_tree, "estoque", bg=CORES["bg_white"]).pack(fill='x', padx=5, pady=(0, 4))
    frame_mov_tree = tk.LabelFrame(paned_est, text="Histórico de Movimentações (200 últimas)", font=('Arial', 10, 'bold'), bg=CORES["bg_white"])
    paned_est.add(frame_mov_tree, height=250)
    cols_mov = ("Data (BR)", "Produto","Tipo","Quantidade","Motivo")
    tree_mov_estoque = ttk.Treeview(frame_mov_tree, columns=cols_mov, show='headings')
    for c in cols_mov:
        tree_mov_estoque.heading(c, text=c)
    tree_mov_estoque.column("Data (BR)", width=150)
    tree_mov_estoque.column("Produto", width=200)
    tree_mov_estoque.pack(fill='both', expand=True, side='left', padx=5, pady=5)
    sb_mov = ttk.Scrollbar(frame_mov_tree, orient='vertical', command=tree_mov_estoque.yview)
    tree_mov_estoque.configure(yscrollcommand=sb_mov.set)
    sb_mov.pack(side='right', fill='y')
    criar_controle_paginacao(frame_mov_tree, "mov_estoque", bg=CORES["bg_white"]).pack(fill='x', padx=5, pady=(0, 4))
    
    # VENDAS - DADOS DA VENDA
    frame_venda_top = tk.Frame(tela_vendas, bg=CORES["bg_light"])
    frame_venda_top.pack(fill="x", padx=20, pady=10)
    frame_venda_cli = tk.LabelFrame(frame_venda_top, text="Dados da Venda", font=("Arial", 11, "bold"), bg=CORES["bg_white"], padx=15, pady=12)
    frame_venda_cli.pack(fill="x", expand=True, padx=5)
    tk.Label(frame_venda_cli, text="Cliente:", bg=CORES["bg_white"], font=("Arial",9,"bold"), fg=CORES["text_dark"]).grid(row=0,column=0,sticky="w")
    combo_venda_cliente = ttk.Combobox(frame_venda_cli, width=28); combo_venda_cliente.grid(row=0,column=1,padx=5,pady=5); combo_venda_cliente.set("0 - Consumidor Final")
    tk.Label(frame_venda_cli, text="Produto:", bg=CORES["bg_white"], font=("Arial",9,"bold"), fg=CORES["text_dark"]).grid(row=0,column=2,sticky="w")
    combo_venda_prod = ttk.Combobox(frame_venda_cli, width=38); combo_venda_prod.grid(row=0,column=3,padx=5,pady=5)
    tk.Label(frame_venda_cli, text="Qtde:", bg=CORES["bg_white"], font=("Arial",9,"bold"), fg=CORES["text_dark"]).grid(row=0,column=4,sticky="w")
    entry_venda_qtd = tk.Entry(frame_venda_cli, width=7); entry_venda_qtd.grid(row=0,column=5,padx=5,pady=5); entry_venda_qtd.insert(0,"1")
    tk.Label(frame_venda_cli, text="Desconto R$:", bg=CORES["bg_white"], font=("Arial",9,"bold"), fg=CORES["text_dark"]).grid(row=1,column=0,sticky="w")
    entry_venda_desc = tk.Entry(frame_venda_cli, width=12); entry_venda_desc.grid(row=1,column=1,padx=5,pady=5,sticky="w"); entry_venda_desc.insert(0,"0")
    entry_venda_desc.bind("<KeyRelease>", lambda e: atualizar_carrinho_tree())
    lbl_venda_forma = tk.Label(frame_venda_cli, text="Forma de Pagamento:", bg=CORES["bg_white"], font=("Arial",9,"bold"), fg=CORES["text_dark"])
    lbl_venda_forma.grid(row=1,column=2,sticky="w")
    combo_venda_forma = ttk.Combobox(frame_venda_cli, width=20, values=["Dinheiro","PIX","Cartão Débito","Cartão Crédito"], state="readonly")
    combo_venda_forma.grid(row=1,column=3,padx=5,pady=5,sticky="w"); combo_venda_forma.set("Dinheiro")
    combo_venda_forma.bind("<<ComboboxSelected>>", on_forma_pagamento_change)
    # Escondido por padrão — só aparece ao clicar no botão "Forma de Pagamento"
    lbl_venda_forma.grid_remove()
    combo_venda_forma.grid_remove()
    # Campos internos, não exibidos no PDV.
    entry_venda_obs = tk.Entry(root)
    entry_venda_venc = tk.Entry(root); entry_venda_venc.insert(0, hoje_br())
    frame_venda_total = tk.Frame(frame_venda_top, bg=CORES["card_green"], bd=0); frame_venda_total.pack(side="right", padx=15, ipadx=20, ipady=10)
    lbl_venda_total_titulo = tk.Label(frame_venda_total, text="TOTAL DA VENDA", bg=CORES["card_green"], font=("Arial",10,"bold"), fg="#15803d"); lbl_venda_total_titulo.pack()
    lbl_venda_total = tk.Label(frame_venda_total, text="R$ 0,00", bg=CORES["card_green"], font=("Arial",20,"bold"), fg="#15803d"); lbl_venda_total.pack()
    lbl_venda_total_sub = tk.Label(frame_venda_total, text="", bg=CORES["card_green"], font=("Arial",8), fg="#166534"); lbl_venda_total_sub.pack()
    frame_venda_add = tk.Frame(tela_vendas, bg=CORES["bg_white"]); frame_venda_add.pack(fill="x", padx=20, pady=3)
    # Frame Cartão Crédito - criado aqui, antes de ser usado - COM MODAL
    frame_cartao_credito = tk.LabelFrame(
        tela_vendas,
        text="Cartão de Crédito — Parcelas + Taxa (% ou R$)",
        font=('Arial', 11, 'bold'),
        bg="#fef3c7",
        padx=15,
        pady=12,
    )

    # Linha 0: parcelas + tipo taxa + valor taxa
    tk.Label(frame_cartao_credito, text="Parcelas (1–12)*:", bg="#fef3c7", font=('Arial', 9, 'bold'), fg=CORES["text_dark"]).grid(row=0, column=0, sticky='w', pady=3)
    combo_parcelas = ttk.Combobox(frame_cartao_credito, width=6, values=[str(i) for i in range(1, 13)])
    combo_parcelas.grid(row=0, column=1, padx=5, pady=3, sticky='w')
    combo_parcelas.set("1")
    combo_parcelas.bind("<<ComboboxSelected>>", lambda e: atualizar_carrinho_tree())
    combo_parcelas.bind("<KeyRelease>", lambda e: atualizar_carrinho_tree())

    tk.Label(frame_cartao_credito, text="Tipo da taxa*:", bg="#fef3c7", font=('Arial', 9, 'bold')).grid(row=0, column=2, sticky='w', padx=(12, 4))
    combo_tipo_taxa = ttk.Combobox(frame_cartao_credito, width=16, values=["Porcentagem (%)", "Valor (R$)"], state="readonly")
    combo_tipo_taxa.grid(row=0, column=3, padx=4, pady=3, sticky='w')
    combo_tipo_taxa.set("Porcentagem (%)")

    lbl_taxa_campo = tk.Label(frame_cartao_credito, text="Taxa %:", bg="#fef3c7", font=('Arial', 9, 'bold'))
    lbl_taxa_campo.grid(row=0, column=4, sticky='w', padx=(12, 4))
    entry_taxa = tk.Entry(frame_cartao_credito, width=10)
    entry_taxa.grid(row=0, column=5, padx=4, pady=3, sticky='w')
    entry_taxa.insert(0, "0")
    entry_taxa.bind("<KeyRelease>", lambda e: atualizar_carrinho_tree())

    def _on_tipo_taxa_change(event=None):
        tipo = combo_tipo_taxa.get().strip()
        if tipo.startswith("Valor"):
            lbl_taxa_campo.config(text="Taxa R$:")
        else:
            lbl_taxa_campo.config(text="Taxa %:")
        atualizar_carrinho_tree()

    combo_tipo_taxa.bind("<<ComboboxSelected>>", _on_tipo_taxa_change)

    # Linha 1: totais
    tk.Label(frame_cartao_credito, text="Total da venda:", bg="#fef3c7", font=('Arial', 9, 'bold')).grid(row=1, column=0, sticky='w', pady=(8, 2))
    lbl_cartao_total_base = tk.Label(frame_cartao_credito, text="R$ 0,00", bg="#fef3c7", font=('Arial', 11, 'bold'), fg=CORES["text_dark"])
    lbl_cartao_total_base.grid(row=1, column=1, sticky='w', padx=5)

    tk.Label(frame_cartao_credito, text="TOTAL (venda + taxa):", bg="#fef3c7", font=('Arial', 9, 'bold'), fg=CORES["danger"]).grid(row=1, column=2, sticky='w', padx=(12, 4))
    lbl_cartao_total_com_taxa = tk.Label(frame_cartao_credito, text="R$ 0,00", bg="#fef3c7", font=('Arial', 12, 'bold'), fg=CORES["danger"])
    lbl_cartao_total_com_taxa.grid(row=1, column=3, sticky='w', padx=4)

    tk.Label(frame_cartao_credito, text="Valor da parcela:", bg="#fef3c7", font=('Arial', 9, 'bold')).grid(row=1, column=4, sticky='w', padx=(12, 4))
    lbl_cartao_parcela = tk.Label(frame_cartao_credito, text="R$ 0,00", bg="#fef3c7", font=('Arial', 12, 'bold'), fg=CORES["primary"])
    lbl_cartao_parcela.grid(row=1, column=5, columnspan=2, sticky='w', padx=4)

    lbl_cartao_taxa_info = tk.Label(frame_cartao_credito, text="", bg="#fef3c7", fg="#92400e", font=('Arial', 9, 'italic'))
    lbl_cartao_taxa_info.grid(row=2, column=0, columnspan=7, sticky='w', pady=(6, 0))

    tk.Label(
        frame_cartao_credito,
        text="Com Cartão de Crédito, as parcelas são lançadas em Contas a Receber (vencimento a cada +30 dias).",
        bg="#fef3c7",
        fg="#92400e",
        font=('Arial', 8, 'italic'),
    ).grid(row=3, column=0, columnspan=7, sticky='w', pady=(4, 0))
    # Não faz pack agora — só aparece ao escolher Cartão Crédito
    
# Adicionar/Remover ficam logo ACIMA do Carrinho
    frame_venda_botoes_topo = tk.Frame(tela_vendas, bg=CORES["bg_white"])
    frame_venda_botoes_topo.pack(fill="x", padx=20, pady=(6, 0))

    LARGURA_BOTAO = 22

    tk.Button(
        frame_venda_botoes_topo, text="Adicionar", command=adicionar_item_venda,
        bg=CORES["primary"], fg="white", font=('Arial', 10, 'bold'),
        width=LARGURA_BOTAO, bd=0, pady=8
    ).grid(row=0, column=0, sticky="w", padx=(0, 6))

    tk.Button(
        frame_venda_botoes_topo, text="Remover", command=remover_item_venda,
        bg=CORES["danger"], fg="white", font=('Arial', 10, 'bold'),
        width=LARGURA_BOTAO, bd=0, pady=8
    ).grid(row=0, column=1, sticky="w", padx=6)

    def _abrir_opcoes_forma_pagamento():
        """Abre um modal com as formas de pagamento disponíveis (o campo antigo
        continua sempre escondido — não volta a aparecer do lado do Desconto)."""
        modal_fp = tk.Toplevel(root)
        modal_fp.title("Forma de Pagamento")
        modal_fp.configure(bg="white")
        modal_fp.transient(root)
        modal_fp.grab_set()
        modal_fp.resizable(False, False)
        largura, altura = 360, 340
        try:
            modal_fp.update_idletasks()
            x = root.winfo_x() + (root.winfo_width() // 2) - (largura // 2)
            y = root.winfo_y() + (root.winfo_height() // 2) - (altura // 2)
            modal_fp.geometry(f"{largura}x{altura}+{x}+{y}")
        except Exception:
            modal_fp.geometry(f"{largura}x{altura}")

        def _fechar_modal_fp():
            try:
                modal_fp.grab_release()
            except Exception:
                pass
            modal_fp.destroy()

        tk.Label(modal_fp, text="Escolha a forma de pagamento", bg="white",
                 font=('Arial', 12, 'bold'), fg=CORES["text_dark"]).pack(pady=(20, 4))
        tk.Label(modal_fp, text=f"Forma atual: {combo_venda_forma.get()}", bg="white",
                 font=('Arial', 9), fg=CORES["text_gray"]).pack(pady=(0, 12))

        def _escolher(forma):
            combo_venda_forma.set(forma)
            _fechar_modal_fp()
            on_forma_pagamento_change()

        opcoes = [("Dinheiro", CORES["primary"]), ("PIX", CORES["primary"]),
                  ("Cartão Débito", CORES["info"]), ("Cartão Crédito", CORES["info"])]
        for nome, cor in opcoes:
            tk.Button(modal_fp, text=nome, command=lambda n=nome: _escolher(n),
                      bg=cor, fg="white", font=('Arial', 11, 'bold'),
                      bd=0, pady=10, width=22, cursor='hand2').pack(pady=6)

        tk.Button(modal_fp, text="Cancelar", command=_fechar_modal_fp,
                  bg="#64748b", fg="white", font=('Arial', 9), bd=0, padx=14, pady=6,
                  cursor='hand2').pack(pady=(14, 0))

        modal_fp.bind('<Escape>', lambda e: _fechar_modal_fp())
        modal_fp.protocol("WM_DELETE_WINDOW", _fechar_modal_fp)

    # Forma de Pagamento e Finalizar Venda ficam no canto inferior esquerdo
    frame_venda_botoes = tk.Frame(tela_vendas, bg=CORES["bg_white"])
    frame_venda_botoes.pack(side='bottom', anchor='sw', fill='x', padx=20, pady=10)

    tk.Button(
        frame_venda_botoes, text="Forma de Pagamento", command=_abrir_opcoes_forma_pagamento,
        bg=CORES["info"], fg="white", font=('Arial', 10, 'bold'),
        width=LARGURA_BOTAO, bd=0, pady=8
    ).grid(row=0, column=0, sticky="w", padx=(0, 6))

    tk.Button(
        frame_venda_botoes, text="Finalizar Venda", command=finalizar_venda,
        bg=CORES["success"], fg="white", font=('Arial', 10, 'bold'),
        width=LARGURA_BOTAO, bd=0, pady=8
    ).grid(row=0, column=1, sticky="w", padx=6)
    
    
    frame_venda_carrinho = tk.LabelFrame(tela_vendas, text="Carrinho", font=('Arial', 10, 'bold'), bg=CORES["bg_white"])
    frame_venda_carrinho.pack(fill='both', expand=True, padx=20, pady=5)
    cols_vc = ("ID","Produto","Qtd","Preço Unit","Desconto","Acréscimo","Subtotal")
    tree_venda_carrinho = ttk.Treeview(frame_venda_carrinho, columns=cols_vc, show='headings', height=6)
    for c in cols_vc:
        tree_venda_carrinho.heading(c, text=c)
    tree_venda_carrinho.column("ID", width=50)
    tree_venda_carrinho.column("Produto", width=220)
    tree_venda_carrinho.column("Qtd", width=55)
    tree_venda_carrinho.column("Preço Unit", width=95)
    tree_venda_carrinho.column("Desconto", width=90)
    tree_venda_carrinho.column("Acréscimo", width=90)
    tree_venda_carrinho.column("Subtotal", width=110)
    tree_venda_carrinho.pack(fill='both', expand=True, side='left', padx=5, pady=5)
    sb_vc = ttk.Scrollbar(frame_venda_carrinho, orient='vertical', command=tree_venda_carrinho.yview)
    tree_venda_carrinho.configure(yscrollcommand=sb_vc.set)
    sb_vc.pack(side='right', fill='y')
    
    # Histórico de vendas, busca, ver itens, cancelar e lixeira removidos do PDV
    tree_vendas = None
    entry_busca_venda = None
    
    # CAIXA
    frame_caixa_top = tk.Frame(tela_caixa, bg=CORES["bg_light"])
    frame_caixa_top.pack(fill='x', padx=20, pady=10)
    frame_caixa_saldo = tk.Frame(frame_caixa_top, bg=CORES["card_green"], bd=0, relief='flat')
    frame_caixa_saldo.pack(side='left', padx=10, ipadx=25, ipady=15)
    tk.Label(frame_caixa_saldo, text="SALDO ATUAL DO CAIXA", bg=CORES["card_green"], font=('Arial', 11, 'bold'), fg="#15803d").pack()
    lbl_caixa_saldo = tk.Label(frame_caixa_saldo, text="R$ 0,00", bg=CORES["card_green"], font=('Arial', 22, 'bold'), fg="#15803d")
    lbl_caixa_saldo.pack()
    lbl_caixa_info = tk.Label(frame_caixa_top, text="", font=('Arial', 10), bg=CORES["bg_light"], justify='left', fg=CORES["text_dark"])
    lbl_caixa_info.pack(side='left', padx=20)
    entry_caixa_data = None  # filtro por data removido
    frame_caixa_manual = tk.LabelFrame(tela_caixa, text="Lançamento Manual", font=('Arial', 11, 'bold'), bg=CORES["bg_white"], padx=15, pady=12)
    frame_caixa_manual.pack(fill='x', padx=20, pady=5)
    tk.Label(frame_caixa_manual, text="Tipo*:", bg=CORES["bg_white"], font=('Arial', 9, 'bold'), fg=CORES["text_dark"]).grid(row=0,column=0, sticky='w')
    combo_caixa_tipo = ttk.Combobox(frame_caixa_manual, width=12, values=["entrada","saida"])
    combo_caixa_tipo.grid(row=0,column=1, padx=5, pady=4)
    combo_caixa_tipo.set("entrada")
    tk.Label(frame_caixa_manual, text="Valor*:", bg=CORES["bg_white"], fg=CORES["text_dark"]).grid(row=0,column=2, sticky='w')
    entry_caixa_valor = tk.Entry(frame_caixa_manual, width=12)
    entry_caixa_valor.grid(row=0,column=3, padx=5, pady=4)
    tk.Label(frame_caixa_manual, text="Descrição*:", bg=CORES["bg_white"], fg=CORES["text_dark"]).grid(row=0,column=4, sticky='w')
    entry_caixa_desc = tk.Entry(frame_caixa_manual, width=30)
    entry_caixa_desc.grid(row=0,column=5, padx=5, pady=4)
    tk.Label(frame_caixa_manual, text="Forma:", bg=CORES["bg_white"]).grid(row=0,column=6, sticky='w')
    combo_caixa_forma = ttk.Combobox(frame_caixa_manual, width=14, values=["Dinheiro","PIX","Cartão","Transferência","Manual"])
    combo_caixa_forma.grid(row=0,column=7, padx=5, pady=4)
    combo_caixa_forma.set("Dinheiro")
    tk.Button(frame_caixa_manual, text="Lançar", command=lancar_caixa_manual, bg=CORES["primary"], fg="white", font=('Arial', 9, 'bold'), bd=0, padx=12, pady=5).grid(row=0,column=8, padx=10)
    frame_caixa_tree = tk.Frame(tela_caixa, bg=CORES["bg_white"], bd=1, relief='solid')
    frame_caixa_tree.pack(fill='both', expand=True, padx=20, pady=10)
    cols_caixa = ("ID","Data BR","Tipo","Valor","Descrição","Cliente/Fornecedor","Origem","Forma","Sel")
    tree_caixa = ttk.Treeview(frame_caixa_tree, columns=cols_caixa, show='headings')
    for c in cols_caixa:
        tree_caixa.heading(c, text=c)
    tree_caixa.column("ID", width=50)
    tree_caixa.column("Data BR", width=130)
    tree_caixa.column("Tipo", width=80)
    tree_caixa.column("Valor", width=110)
    tree_caixa.column("Descrição", width=220)
    tree_caixa.column("Cliente/Fornecedor", width=160)
    tree_caixa.column("Sel", width=40, anchor='center', stretch=False)
    tree_caixa.heading("Sel", text="☐")
    tree_caixa.pack(fill='both', expand=True, side='left', padx=5, pady=5)
    sb_caixa = ttk.Scrollbar(frame_caixa_tree, orient='vertical', command=tree_caixa.yview)
    tree_caixa.configure(yscrollcommand=sb_caixa.set)
    sb_caixa.pack(side='right', fill='y')
    habilitar_selecao_multipla(tree_caixa)
    criar_barra_selecao_multipla(tela_caixa, tree_caixa, excluir_lancamentos_caixa_em_massa).pack(
        fill='x', padx=20, pady=(0, 8), before=frame_caixa_tree)
    criar_controle_paginacao(tela_caixa, "caixa", bg=CORES["bg_light"]).pack(
        fill='x', padx=20, pady=(0, 6), before=frame_caixa_tree)
    tree_caixa.tag_configure('entrada', background='#dcfce7')
    tree_caixa.tag_configure('saida', background='#fee2e2')
    
    # CONTAS A PAGAR COM SUB-ABAS + PARCELAS CARTÃO CRÉDITO
    frame_cp_cards = tk.Frame(tela_cp, bg=CORES["bg_light"])
    frame_cp_cards.pack(fill='x', padx=20, pady=10)
    def criar_card_conta(parent, titulo, bg, fg):
        f = tk.Frame(parent, bg=bg, bd=1, relief='solid')
        f.pack(side='left', padx=8, pady=5, fill='x', expand=True, ipadx=10, ipady=10)
        lbl = tk.Label(f, text=titulo, bg=bg, fg=fg, font=('Arial', 10, 'bold'), justify='center')
        lbl.pack()
        return lbl
    lbl_cp_aberto = criar_card_conta(frame_cp_cards, "Em Aberto\n0 contas\nR$ 0,00", "#dbeafe", "#1e40af")
    lbl_cp_pago = criar_card_conta(frame_cp_cards, "Pago\n0 contas\nR$ 0,00", "#dcfce7", "#166534")
    lbl_cp_atraso = criar_card_conta(frame_cp_cards, "Em Atraso\n0 contas\nR$ 0,00", "#fef3c7", "#92400e")
    lbl_cp_cancelado = criar_card_conta(frame_cp_cards, "Cancelado\n0 contas\nR$ 0,00", "#fee2e2", "#991b1b")
    
    # Barra superior: incluir conta
    frame_cp_top = tk.Frame(tela_cp, bg=CORES["bg_light"])
    frame_cp_top.pack(fill='x', padx=20, pady=(0, 5))
    tk.Button(frame_cp_top, text="➕ Incluir Conta a Pagar", command=abrir_modal_conta_pagar,
              bg=CORES["success"], fg="white", font=('Arial', 10, 'bold'), bd=0, padx=16, pady=8, cursor='hand2').pack(side='left')
    # Dummies ocultos (compatibilidade)
    _dummy_cp = tk.Frame(tela_cp)
    combo_cp_forn = ttk.Combobox(_dummy_cp)
    entry_cp_id = tk.Entry(_dummy_cp)
    entry_cp_desc = tk.Entry(_dummy_cp)
    entry_cp_valor = tk.Entry(_dummy_cp)
    entry_cp_venc = tk.Entry(_dummy_cp)
    combo_cp_forma = ttk.Combobox(_dummy_cp)
    combo_cp_parcelas = ttk.Combobox(_dummy_cp)
    entry_cp_taxa = tk.Entry(_dummy_cp)
    frame_cp_cartao = tk.Frame(_dummy_cp)
    lbl_cp_parcela_info = tk.Label(_dummy_cp)
    
    frame_cp_acoes = tk.Frame(tela_cp, bg=CORES["bg_light"])
    frame_cp_acoes.pack(fill='x', side='bottom', padx=20, pady=10)
    tk.Button(frame_cp_acoes, text="Pagar selecionadas", command=pagar_conta, bg=CORES["primary"], fg="white",
              font=('Arial', 10, 'bold'), bd=0, padx=14, pady=8, cursor='hand2').pack(side='left', padx=(0, 8))
    tk.Button(frame_cp_acoes, text="Cancelar", command=cancelar_cp, bg=CORES["warning"], fg="white",
              font=('Arial', 10, 'bold'), bd=0, padx=14, pady=8, cursor='hand2').pack(side='left', padx=8)
    tk.Button(frame_cp_acoes, text="Excluir selecionados", command=excluir_cp_selecionados, bg=CORES["danger"], fg="white",
              font=('Arial', 10, 'bold'), bd=0, padx=14, pady=8, cursor='hand2').pack(side='left', padx=8)

    notebook_cp = ttk.Notebook(tela_cp)
    notebook_cp.pack(fill='both', expand=True, padx=20, pady=5)

    cols_cp = ("ID", "Fornecedor", "Descrição", "Valor", "Venc BR", "Status", "Sel")
    larguras_cp = {"ID": 50, "Fornecedor": 180, "Descrição": 220, "Valor": 100, "Venc BR": 100, "Status": 100}

    def _criar_aba_cp(texto_aba, chave_pag, tree_var_name):
        tab = tk.Frame(notebook_cp, bg=CORES["bg_white"])
        notebook_cp.add(tab, text=texto_aba)
        # barra multi-select no topo
        frame_tree = tk.Frame(tab, bg=CORES["bg_white"])
        tree = ttk.Treeview(frame_tree, columns=cols_cp, show='headings')
        _configurar_tree_com_checkbox(tree, cols_cp, larguras_cp)
        tree.pack(fill='both', expand=True, side='left')
        sb = ttk.Scrollbar(frame_tree, orient='vertical', command=tree.yview)
        tree.configure(yscrollcommand=sb.set)
        sb.pack(side='right', fill='y')
        sbh = ttk.Scrollbar(tab, orient='horizontal', command=tree.xview)
        tree.configure(xscrollcommand=sbh.set)
        tree.bind("<Double-1>", editar_cp)
        tree.tag_configure('aberto', background='#dbeafe')
        tree.tag_configure('pago', background='#dcfce7')
        tree.tag_configure('atraso', background='#fef3c7')
        tree.tag_configure('cancelado', background='#fee2e2')
        # pack order: barra, paginação, tree, scrollbar horizontal
        criar_barra_selecao_multipla(tab, tree, excluir_contas_pagar_em_massa, mostrar_excluir=False).pack(fill='x', padx=5, pady=(6, 2))
        criar_controle_paginacao(tab, chave_pag, bg=CORES["bg_white"]).pack(fill='x', padx=5, pady=(0, 2))
        frame_tree.pack(fill='both', expand=True, padx=5, pady=2)
        sbh.pack(fill='x', padx=5, pady=(0, 4))
        globals()[tree_var_name] = tree
        return tree

    tree_cp = _criar_aba_cp("Todas", "cp_todos", "tree_cp")
    tree_cp_aberto = _criar_aba_cp("Em Aberto", "cp_aberto", "tree_cp_aberto")
    tree_cp_pago = _criar_aba_cp("Pago", "cp_pago", "tree_cp_pago")
    tree_cp_atraso = _criar_aba_cp("Em Atraso", "cp_atraso", "tree_cp_atraso")
    tree_cp_cancelado = _criar_aba_cp("Cancelado", "cp_cancelado", "tree_cp_cancelado")
    try:
        _aplicar_cores_notebook(
            notebook_cp,
            ["#6366f1", "#3b82f6", "#10b981", "#f59e0b", "#ef4444"],
        )
    except Exception:
        pass

    # CONTAS A RECEBER
    frame_cr_cards = tk.Frame(tela_cr, bg=CORES["bg_light"])
    frame_cr_cards.pack(fill='x', padx=20, pady=10)
    lbl_cr_aberto = criar_card_conta(frame_cr_cards, "Em Aberto\n0 contas\nR$ 0,00", "#dbeafe", "#1e40af")
    lbl_cr_recebido = criar_card_conta(frame_cr_cards, "Recebido\n0 contas\nR$ 0,00", "#dcfce7", "#166534")
    lbl_cr_atraso = criar_card_conta(frame_cr_cards, "Em Atraso\n0 contas\nR$ 0,00", "#fef3c7", "#92400e")
    lbl_cr_cancelado = criar_card_conta(frame_cr_cards, "Cancelado\n0 contas\nR$ 0,00", "#fee2e2", "#991b1b")
    
    frame_cr_top = tk.Frame(tela_cr, bg=CORES["bg_light"])
    frame_cr_top.pack(fill='x', padx=20, pady=(0, 5))
    tk.Button(frame_cr_top, text="➕ Incluir Conta a Receber", command=abrir_modal_conta_receber,
              bg=CORES["success"], fg="white", font=('Arial', 10, 'bold'), bd=0, padx=16, pady=8, cursor='hand2').pack(side='left')
    # Dummies ocultos
    _dummy_cr = tk.Frame(tela_cr)
    combo_cr_cliente = ttk.Combobox(_dummy_cr)
    entry_cr_id = tk.Entry(_dummy_cr)
    entry_cr_desc = tk.Entry(_dummy_cr)
    entry_cr_valor = tk.Entry(_dummy_cr)
    entry_cr_venc = tk.Entry(_dummy_cr)
    combo_cr_forma = ttk.Combobox(_dummy_cr)
    combo_cr_parcelas = ttk.Combobox(_dummy_cr)
    entry_cr_taxa = tk.Entry(_dummy_cr)
    frame_cr_cartao = tk.Frame(_dummy_cr)
    lbl_cr_parcela_info = tk.Label(_dummy_cr)
    
    frame_cr_acoes = tk.Frame(tela_cr, bg=CORES["bg_light"])
    frame_cr_acoes.pack(fill='x', side='bottom', padx=20, pady=10)
    tk.Button(frame_cr_acoes, text="Receber selecionadas", command=receber_conta, bg=CORES["primary"], fg="white",
              font=('Arial', 10, 'bold'), bd=0, padx=14, pady=8, cursor='hand2').pack(side='left', padx=(0, 8))
    tk.Button(frame_cr_acoes, text="Cancelar", command=cancelar_cr, bg=CORES["warning"], fg="white",
              font=('Arial', 10, 'bold'), bd=0, padx=14, pady=8, cursor='hand2').pack(side='left', padx=8)
    tk.Button(frame_cr_acoes, text="Excluir selecionados", command=excluir_cr_selecionados, bg=CORES["danger"], fg="white",
              font=('Arial', 10, 'bold'), bd=0, padx=14, pady=8, cursor='hand2').pack(side='left', padx=8)

    notebook_cr = ttk.Notebook(tela_cr)
    notebook_cr.pack(fill='both', expand=True, padx=20, pady=5)

    cols_cr = ("ID", "Cliente", "Descrição", "Valor", "Data Venda", "Venc BR", "Status", "Parcela", "Sel")
    larguras_cr = {"ID": 50, "Cliente": 140, "Descrição": 180, "Valor": 90, "Data Venda": 95, "Venc BR": 95, "Status": 95, "Parcela": 70}

    def _criar_aba_cr(texto_aba, chave_pag, tree_var_name):
        tab = tk.Frame(notebook_cr, bg=CORES["bg_white"])
        notebook_cr.add(tab, text=texto_aba)
        frame_tree = tk.Frame(tab, bg=CORES["bg_white"])
        tree = ttk.Treeview(frame_tree, columns=cols_cr, show='headings')
        _configurar_tree_com_checkbox(tree, cols_cr, larguras_cr)
        tree.pack(fill='both', expand=True, side='left')
        sb = ttk.Scrollbar(frame_tree, orient='vertical', command=tree.yview)
        tree.configure(yscrollcommand=sb.set)
        sb.pack(side='right', fill='y')
        sbh = ttk.Scrollbar(tab, orient='horizontal', command=tree.xview)
        tree.configure(xscrollcommand=sbh.set)
        tree.bind("<Double-1>", editar_cr)
        tree.tag_configure('aberto', background='#dbeafe')
        tree.tag_configure('pago', background='#dcfce7')
        tree.tag_configure('atraso', background='#fef3c7')
        tree.tag_configure('cancelado', background='#fee2e2')
        criar_barra_selecao_multipla(tab, tree, excluir_contas_receber_em_massa, mostrar_excluir=False).pack(fill='x', padx=5, pady=(6, 2))
        criar_controle_paginacao(tab, chave_pag, bg=CORES["bg_white"]).pack(fill='x', padx=5, pady=(0, 2))
        frame_tree.pack(fill='both', expand=True, padx=5, pady=2)
        sbh.pack(fill='x', padx=5, pady=(0, 4))
        globals()[tree_var_name] = tree
        return tree

    tree_cr = _criar_aba_cr("Todas", "cr_todos", "tree_cr")
    tree_cr_aberto = _criar_aba_cr("Em Aberto", "cr_aberto", "tree_cr_aberto")
    tree_cr_recebido = _criar_aba_cr("Recebido", "cr_recebido", "tree_cr_recebido")
    tree_cr_atraso = _criar_aba_cr("Em Atraso", "cr_atraso", "tree_cr_atraso")
    tree_cr_cancelado = _criar_aba_cr("Cancelado", "cr_cancelado", "tree_cr_cancelado")
    try:
        _aplicar_cores_notebook(
            notebook_cr,
            ["#6366f1", "#3b82f6", "#10b981", "#f59e0b", "#ef4444"],
        )
    except Exception:
        pass


    # RELATÓRIOS
    frame_rel = tk.Frame(tela_relatorios, bg=CORES["bg_light"], padx=20, pady=20)
    frame_rel.pack(fill='both', expand=True)
    tk.Label(frame_rel, text="📑 EXPORTAR RELATÓRIOS - Datas no padrão BR dd/mm/aaaa", font=('Arial', 14, 'bold'), bg=CORES["bg_light"], fg=CORES["text_dark"]).grid(row=0,column=0,columnspan=3, pady=15)
    frame_rel_venda = tk.LabelFrame(frame_rel, text="📊 Relatório de Vendas", font=('Arial', 11, 'bold'), bg=CORES["bg_white"], padx=15, pady=15)
    frame_rel_venda.grid(row=1,column=0, padx=10, pady=10, sticky='ew')
    tk.Label(frame_rel_venda, text="Data Inicial dd/mm/aaaa:", bg=CORES["bg_white"]).grid(row=0,column=0, sticky='w', pady=3)
    entry_rel_venda_ini = tk.Entry(frame_rel_venda, width=15)
    entry_rel_venda_ini.grid(row=0,column=1, padx=5, pady=3)
    ativar_seletor_data(entry_rel_venda_ini)
    tk.Label(frame_rel_venda, text="Data Final dd/mm/aaaa:", bg=CORES["bg_white"]).grid(row=1,column=0, sticky='w', pady=3)
    entry_rel_venda_fim = tk.Entry(frame_rel_venda, width=15)
    entry_rel_venda_fim.grid(row=1,column=1, padx=5, pady=3)
    ativar_seletor_data(entry_rel_venda_fim)
    tk.Button(frame_rel_venda, text="📊 Exportar Vendas", command=relatorio_vendas, bg=CORES["primary"], fg="white", width=20, font=('Arial', 10, 'bold'), bd=0, pady=6).grid(row=2,column=0,columnspan=2, pady=12)
    frame_rel_est = tk.LabelFrame(frame_rel, text="📦 Relatório de Estoque", font=('Arial', 11, 'bold'), bg=CORES["bg_white"], padx=15, pady=15)
    frame_rel_est.grid(row=1,column=1, padx=10, pady=10, sticky='ew')
    tk.Label(frame_rel_est, text="Exporta todos os produtos\ncom estoque atual, custo, venda.", bg=CORES["bg_white"], justify='left').grid(row=0,column=0, pady=5)
    tk.Button(frame_rel_est, text="📦 Exportar Estoque", command=relatorio_estoque, bg=CORES["success"], fg="white", width=20, font=('Arial', 10, 'bold'), bd=0, pady=6).grid(row=1,column=0, pady=12)
    frame_rel_cli = tk.LabelFrame(frame_rel, text="👥 Relatório de Clientes", font=('Arial', 11, 'bold'), bg=CORES["bg_white"], padx=15, pady=15)
    frame_rel_cli.grid(row=1,column=2, padx=10, pady=10, sticky='ew')
    tk.Label(frame_rel_cli, text="Exporta lista completa\nde clientes cadastrados.", bg=CORES["bg_white"], justify='left').grid(row=0,column=0, pady=5)
    tk.Button(frame_rel_cli, text="👥 Exportar Clientes", command=relatorio_clientes, bg=CORES["purple"], fg="white", width=20, font=('Arial', 10, 'bold'), bd=0, pady=6).grid(row=1,column=0, pady=12)
    frame_rel_caixa = tk.LabelFrame(frame_rel, text="💰 Relatório de Caixa", font=('Arial', 11, 'bold'), bg=CORES["bg_white"], padx=15, pady=15)
    frame_rel_caixa.grid(row=2,column=0, padx=10, pady=10, sticky='ew')
    tk.Label(frame_rel_caixa, text="Data Inicial dd/mm/aaaa:", bg=CORES["bg_white"]).grid(row=0,column=0, sticky='w', pady=3)
    entry_rel_caixa_ini = tk.Entry(frame_rel_caixa, width=15)
    entry_rel_caixa_ini.grid(row=0,column=1, padx=5, pady=3)
    ativar_seletor_data(entry_rel_caixa_ini)
    tk.Label(frame_rel_caixa, text="Data Final dd/mm/aaaa:", bg=CORES["bg_white"]).grid(row=1,column=0, sticky='w', pady=3)
    entry_rel_caixa_fim = tk.Entry(frame_rel_caixa, width=15)
    entry_rel_caixa_fim.grid(row=1,column=1, padx=5, pady=3)
    ativar_seletor_data(entry_rel_caixa_fim)
    tk.Button(frame_rel_caixa, text="💰 Exportar Caixa", command=relatorio_caixa, bg="#f97316", fg="white", width=20, font=('Arial', 10, 'bold'), bd=0, pady=6).grid(row=2,column=0,columnspan=2, pady=12)
    frame_rel_contas = tk.LabelFrame(frame_rel, text="📑 Relatório Contas", font=('Arial', 11, 'bold'), bg=CORES["bg_white"], padx=15, pady=15)
    frame_rel_contas.grid(row=2,column=1, padx=10, pady=10, sticky='ew')
    tk.Label(frame_rel_contas, text="Exporta todas as contas\na pagar e a receber.", bg=CORES["bg_white"], justify='left').grid(row=0,column=0, pady=5)
    tk.Button(frame_rel_contas, text="📑 Exportar Contas", command=relatorio_contas, bg=CORES["danger"], fg="white", width=20, font=('Arial', 10, 'bold'), bd=0, pady=6).grid(row=1,column=0, pady=12)
    frame_info = tk.LabelFrame(frame_rel, text="ℹ️ Informações", font=('Arial', 11, 'bold'), bg=CORES["bg_white"], padx=15, pady=15)
    frame_info.grid(row=2,column=2, padx=10, pady=10, sticky='ew')
    info_text = f"""
Sistema v3.1 Corrigido
Nome do Sistema: Sitema de Gestão
Autor: Robson Souza
Ano: 2026
 
Todos os direitos reservados.
Este código foi desenvolvido por Robson Souza e não pode ser 
copiado, distribuído ou modificado sem autorização prévia.
"""
    tk.Label(frame_info, text=info_text, bg=CORES["bg_white"], justify='left', font=('Arial', 9), fg=CORES["text_dark"]).pack()
    
    # LIXEIRA
    frame_lixeira_top = tk.Frame(tela_lixeira, bg=CORES["bg_light"])
    frame_lixeira_top.pack(fill='x', padx=20, pady=10)
    tk.Label(frame_lixeira_top, text="Lixeira - Itens Excluídos", font=('Arial', 14, 'bold'), bg=CORES["bg_light"], fg=CORES["danger"]).pack(side='left')
    notebook_lixeira = ttk.Notebook(tela_lixeira)
    notebook_lixeira.pack(fill='both', expand=True, padx=20, pady=10)

    def _montar_aba_lixeira(titulo, chave_pag, tabela, cols, tree_attr_name):
        tab = tk.Frame(notebook_lixeira, bg=CORES["bg_white"])
        notebook_lixeira.add(tab, text=titulo)
        # tree + barra multi (Restaurar / Excluir definitivo ficam só na barra "Marcar todos")
        frame_tree = tk.Frame(tab, bg=CORES["bg_white"])
        frame_tree.pack(fill='both', expand=True, padx=10, pady=5)
        cols_com_sel = tuple(list(cols) + ["Sel"])
        tree = ttk.Treeview(frame_tree, columns=cols_com_sel, show='headings')
        for c in cols_com_sel:
            tree.heading(c, text=c)
        tree.column("Sel", width=40, anchor='center', stretch=False)
        tree.heading("Sel", text="☐")
        tree.pack(fill='both', expand=True, side='left')
        sb = ttk.Scrollbar(frame_tree, orient='vertical', command=tree.yview)
        tree.configure(yscrollcommand=sb.set)
        sb.pack(side='right', fill='y')
        globals()[tree_attr_name] = tree
        habilitar_selecao_multipla(tree)
        # Botões na parte inferior esquerda
        criar_barra_lixeira(tab, tree, tabela).pack(side='bottom', fill='x', padx=10, pady=8, anchor='w')
        return tree

    cols_lix = ("ID", "Nome", "CPF/CNPJ", "Data Exclusão BR", "Excluído Por")
    tree_lixeira_clientes = _montar_aba_lixeira("Clientes", "lix_cli", "clientes", cols_lix, "tree_lixeira_clientes")
    tree_lixeira_fornecedores = _montar_aba_lixeira("Fornecedores", "lix_forn", "fornecedores", cols_lix, "tree_lixeira_fornecedores")
    cols_lix_prod = ("ID", "Código", "Nome", "Preço", "Data Exclusão BR", "Excluído Por")
    tree_lixeira_produtos = _montar_aba_lixeira("Produtos", "lix_prod", "produtos", cols_lix_prod, "tree_lixeira_produtos")
    cols_lix_venda = ("ID", "Data BR", "Total", "Status", "Data Exclusão BR", "Excluído Por")
    tree_lixeira_vendas = _montar_aba_lixeira("Vendas", "lix_vendas", "vendas", cols_lix_venda, "tree_lixeira_vendas")
    cols_lix_caixa = ("ID", "Data BR", "Tipo", "Valor", "Descrição", "Data Exclusão BR")
    tree_lixeira_caixa = _montar_aba_lixeira("Caixa", "lix_caixa", "caixa", cols_lix_caixa, "tree_lixeira_caixa")
    cols_lix_conta = ("ID", "Descrição", "Valor", "Venc BR", "Status", "Data Exclusão BR")
    tree_lixeira_cp = _montar_aba_lixeira("Contas Pagar", "lix_cp", "contas_a_pagar", cols_lix_conta, "tree_lixeira_cp")
    tree_lixeira_cr = _montar_aba_lixeira("Contas Receber", "lix_cr", "contas_a_receber", cols_lix_conta, "tree_lixeira_cr")

    # Cores das abas da lixeira (mesma lógica das abas principais)
    try:
        _aplicar_cores_notebook(
            notebook_lixeira,
            ["#8b5cf6", "#f59e0b", "#10b981", "#ec4899", "#22c55e", "#ef4444", "#14b8a6"],
        )
    except Exception:
        pass


    # BACKUP / RESTAURAÇÃO
    frame_bk_top = tk.Frame(tela_backup, bg=CORES["bg_light"])
    frame_bk_top.pack(fill='x', padx=20, pady=12)
    tk.Label(
        frame_bk_top,
        text="Backup e Restauração do Sistema",
        font=('Arial', 14, 'bold'),
        bg=CORES["bg_light"],
        fg=CORES["text_dark"],
    ).pack(side='left')
    lbl_bk_status = tk.Label(frame_bk_top, text="", font=('Arial', 9), bg=CORES["bg_light"], fg=CORES["text_gray"])
    lbl_bk_status.pack(side='right')

    frame_bk_info = tk.LabelFrame(
        tela_backup,
        text="ℹ️ Como funciona",
        font=('Arial', 10, 'bold'),
        bg=CORES["bg_white"],
        padx=12,
        pady=10,
    )
    frame_bk_info.pack(fill='x', padx=20, pady=5)
    tk.Label(
        frame_bk_info,
        text=(
            "• Gerar backup agora: salva uma cópia do banco na pasta backups/  •  Exportar: salva a cópia onde você escolher\n"
            "• Restaurar / Importar: substitui os dados atuais pelos do backup e ATUALIZA todas as telas automaticamente\n"
            "• Antes de qualquer restauração é gerada uma cópia de segurança (antes_restore_...) dos dados atuais\n"
            "• Atualizar dados: recarrega todas as listas, combos e o dashboard a partir do banco de dados"
        ),
        bg=CORES["bg_white"],
        fg=CORES["text_dark"],
        justify='left',
        font=('Arial', 9),
    ).pack(anchor='w')

    def _set_status_bk(texto):
        try:
            lbl_bk_status.config(text=texto)
        except Exception:
            pass

    def _backup_manual():
        caminho = fazer_backup(silencioso=False)
        atualizar_lista_backups()
        if caminho:
            _set_status_bk(f"Último backup: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
        return caminho

    def _exportar_backup():
        garantir_pasta_backup()
        destino = filedialog.asksaveasfilename(
            parent=root,
            title="Salvar backup como...",
            defaultextension=".db",
            filetypes=[("Banco SQLite", "*.db"), ("Todos", "*.*")],
            initialfile=gerar_nome_backup(),
            initialdir=BACKUP_DIR,
        )
        if not destino:
            return
        if fazer_backup(destino=destino, silencioso=False):
            atualizar_lista_backups()
            _set_status_bk(f"Último backup: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")

    def _pos_restauracao(caminho):
        atualizar_lista_backups()
        _set_status_bk(f"Restaurado de {os.path.basename(caminho)} em {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
        # Mostra o dashboard já com os dados restaurados
        try:
            abrir_aba("dashboard", "Dashboard", "📊")
        except Exception:
            pass

    def _restaurar_selecionado():
        sel = tree_backups.selection()
        if not sel:
            mostrar_aviso("Selecione um backup na lista.")
            return
        caminho = tree_backups.item(sel[0])["values"][3]
        if restaurar_backup(caminho):
            _pos_restauracao(caminho)

    def _restaurar_arquivo():
        garantir_pasta_backup()
        caminho = filedialog.askopenfilename(
            parent=root,
            title="Importar / restaurar backup (.db)",
            filetypes=[("Banco SQLite", "*.db"), ("Todos", "*.*")],
        )
        if not caminho:
            return
        if restaurar_backup(caminho):
            _pos_restauracao(caminho)

    def _atualizar_dados_manual():
        ok = recarregar_dados_sistema(limpar_formularios=False)
        atualizar_lista_backups()
        if ok:
            _set_status_bk(f"Dados atualizados em {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
            mostrar_sucesso("Todas as telas foram atualizadas com os dados do banco.", "Dados atualizados")
        else:
            mostrar_aviso("Algumas telas não puderam ser atualizadas.\nFaça logout e login para recarregar tudo.", "Atualização parcial")

    def _abrir_pasta_backups():
        garantir_pasta_backup()
        pasta = BACKUP_DIR
        try:
            if os.name == "nt":
                os.startfile(pasta)
            elif sys.platform == "darwin":
                os.system(f'open "{pasta}"')
            else:
                os.system(f'xdg-open "{pasta}" >/dev/null 2>&1 &')
        except Exception:
            mostrar_info(f"Pasta de backups:\n{pasta}", "Backups")

    frame_bk_acoes = tk.Frame(tela_backup, bg=CORES["bg_light"])
    frame_bk_acoes.pack(fill='x', padx=20, pady=8)
    tk.Button(frame_bk_acoes, text="Gerar backup agora", command=_backup_manual,
              bg=CORES["success"], fg="white", font=('Arial', 10, 'bold'), bd=0, padx=14, pady=8, cursor='hand2').pack(side='left', padx=(0, 8))
    tk.Button(frame_bk_acoes, text="Exportar backup", command=_exportar_backup,
              bg=CORES["primary"], fg="white", font=('Arial', 10, 'bold'), bd=0, padx=14, pady=8, cursor='hand2').pack(side='left', padx=4)
    tk.Button(frame_bk_acoes, text="Importar / Restaurar de arquivo", command=_restaurar_arquivo,
              bg=CORES["warning"], fg="white", font=('Arial', 10, 'bold'), bd=0, padx=14, pady=8, cursor='hand2').pack(side='left', padx=4)
    tk.Button(frame_bk_acoes, text="Atualizar dados", command=_atualizar_dados_manual,
              bg=CORES["purple"], fg="white", font=('Arial', 10, 'bold'), bd=0, padx=14, pady=8, cursor='hand2').pack(side='right')

    frame_bk_tree = tk.LabelFrame(
        tela_backup,
        text="📋 Backups salvos na pasta local",
        font=('Arial', 10, 'bold'),
        bg=CORES["bg_white"],
    )
    frame_bk_tree.pack(fill='both', expand=True, padx=20, pady=10)

    cols_bk = ("Arquivo", "Data", "Tamanho", "Caminho")
    tree_backups = ttk.Treeview(frame_bk_tree, columns=cols_bk, show='headings')
    for c in cols_bk:
        tree_backups.heading(c, text=c)
    tree_backups.column("Arquivo", width=280)
    tree_backups.column("Data", width=150)
    tree_backups.column("Tamanho", width=100)
    tree_backups.column("Caminho", width=320)
    tree_backups.pack(fill='both', expand=True, side='left', padx=5, pady=5)
    sb_bk = ttk.Scrollbar(frame_bk_tree, orient='vertical', command=tree_backups.yview)
    tree_backups.configure(yscrollcommand=sb_bk.set)
    sb_bk.pack(side='right', fill='y')

    def atualizar_lista_backups():
        try:
            limpar_tree(tree_backups)
        except Exception:
            return
        for nome, mtime, tamanho, caminho in listar_backups_locais():
            if tamanho >= 1024 * 1024:
                tam_txt = f"{tamanho / (1024*1024):.2f} MB"
            else:
                tam_txt = f"{tamanho / 1024:.1f} KB"
            tree_backups.insert("", "end", values=(nome, mtime, tam_txt, caminho))

    # Permite que restaurar_backup()/recarregar_dados_sistema() atualizem esta lista
    globals()["_hook_atualizar_lista_backups"] = atualizar_lista_backups

    tree_backups.bind("<Double-1>", lambda e: _restaurar_selecionado())

    frame_bk_bottom = tk.Frame(tela_backup, bg=CORES["bg_light"])
    frame_bk_bottom.pack(fill='x', padx=20, pady=(0, 12))
    tk.Button(
        frame_bk_bottom,
        text="Restaurar selecionado",
        command=_restaurar_selecionado,
        bg=CORES["danger"],
        fg="white",
        font=('Arial', 10, 'bold'),
        bd=0,
        padx=14,
        pady=8,
        cursor='hand2',
    ).pack(side='left')
    tk.Button(
        frame_bk_bottom,
        text="Atualizar lista",
        command=atualizar_lista_backups,
        bg="#64748b",
        fg="white",
        font=('Arial', 10, 'bold'),
        bd=0,
        padx=14,
        pady=8,
        cursor='hand2',
    ).pack(side='left', padx=8)
    tk.Button(
        frame_bk_bottom,
        text="Abrir pasta de backups",
        command=_abrir_pasta_backups,
        bg="white",
        fg=CORES["text_dark"],
        font=('Arial', 10),
        bd=1,
        relief='solid',
        padx=14,
        pady=7,
        cursor='hand2',
    ).pack(side='right')
    atualizar_lista_backups()


    # USUÁRIOS
    frame_usu_form = tk.LabelFrame(tela_usuarios, text="Cadastro de Usuário (Somente ADM)", font=('Arial', 11, 'bold'), bg=CORES["bg_white"], padx=15, pady=15)
    frame_usu_form.pack(fill='x', padx=20, pady=10)
    # Labels alinhadas (coluna 0) + campos (coluna 1 e 3)
    lbl_u = dict(bg=CORES["bg_white"], font=('Arial', 9, 'bold'), fg=CORES["text_dark"], anchor='e')
    tk.Label(frame_usu_form, text="ID:", width=12, **lbl_u).grid(row=0, column=0, sticky='e', padx=(0, 6), pady=5)
    entry_usu_id = tk.Entry(frame_usu_form, width=10)
    entry_usu_id.grid(row=0, column=1, sticky='w', padx=5, pady=5)
    tk.Label(frame_usu_form, text="Nome*:", width=12, **lbl_u).grid(row=0, column=2, sticky='e', padx=(12, 6), pady=5)
    entry_usu_nome = tk.Entry(frame_usu_form, width=28)
    entry_usu_nome.grid(row=0, column=3, sticky='w', padx=5, pady=5)
    tk.Label(frame_usu_form, text="Login*:", width=12, **lbl_u).grid(row=1, column=0, sticky='e', padx=(0, 6), pady=5)
    entry_usu_login = tk.Entry(frame_usu_form, width=18)
    entry_usu_login.grid(row=1, column=1, sticky='w', padx=5, pady=5)
    tk.Label(frame_usu_form, text="Senha*:", width=12, **lbl_u).grid(row=1, column=2, sticky='e', padx=(12, 6), pady=5)
    entry_usu_senha = tk.Entry(frame_usu_form, width=18, show="•")
    entry_usu_senha.grid(row=1, column=3, sticky='w', padx=5, pady=5)
    tk.Label(frame_usu_form, text="Perfil*:", width=12, **lbl_u).grid(row=2, column=0, sticky='e', padx=(0, 6), pady=5)
    combo_usu_perfil = ttk.Combobox(frame_usu_form, width=16, values=["admin", "operador"])
    combo_usu_perfil.grid(row=2, column=1, sticky='w', padx=5, pady=5)
    combo_usu_perfil.set("operador")
    tk.Label(frame_usu_form, text="E-mail*:", width=12, **lbl_u).grid(row=2, column=2, sticky='e', padx=(12, 6), pady=5)
    entry_usu_email = tk.Entry(frame_usu_form, width=28)
    entry_usu_email.grid(row=2, column=3, sticky='w', padx=5, pady=5)
    tk.Label(
        frame_usu_form,
        text="Deixe a senha em branco para não alterar ao editar • E-mail usado para recuperação de senha",
        bg=CORES["bg_white"], fg=CORES["text_gray"], font=('Arial', 8),
    ).grid(row=3, column=0, columnspan=4, sticky='w', padx=5, pady=(4, 2))
    frame_usu_btn = tk.Frame(frame_usu_form, bg=CORES["bg_white"])
    frame_usu_btn.grid(row=4, column=0, columnspan=4, sticky='w', pady=12)
    tk.Button(frame_usu_btn, text="Salvar", command=salvar_usuario, bg=CORES["success"], fg="white", width=14, font=('Arial', 10, 'bold'), bd=0, pady=6).pack(side='left', padx=(0, 8)), 
    tk.Button(frame_usu_btn, text="Limpar", command=limpar_form_usuario, bg="#64748b", fg="white", width=12, bd=0, pady=6).pack(side='left', padx=4)
    frame_usu_tree = tk.Frame(tela_usuarios, bg=CORES["bg_white"], bd=1, relief='solid')
    frame_usu_tree.pack(fill='both', expand=True, padx=20, pady=10)
    cols_usu = ("ID","Nome","Login","E-mail","Perfil","Data Cadastro BR","Sel")
    tree_usuarios = ttk.Treeview(frame_usu_tree, columns=cols_usu, show='headings')
    for c in cols_usu:
        tree_usuarios.heading(c, text=c)
    tree_usuarios.column("ID", width=50)
    tree_usuarios.column("Nome", width=160)
    tree_usuarios.column("Login", width=100)
    tree_usuarios.column("E-mail", width=180)
    tree_usuarios.column("Sel", width=40, anchor='center', stretch=False)
    tree_usuarios.heading("Sel", text="☐")
    tree_usuarios.pack(fill='both', expand=True, side='left', padx=5, pady=5)
    sb_usu = ttk.Scrollbar(frame_usu_tree, orient='vertical', command=tree_usuarios.yview)
    tree_usuarios.configure(yscrollcommand=sb_usu.set)
    sb_usu.pack(side='right', fill='y')
    tree_usuarios.bind("<Double-1>", editar_usuario)
    habilitar_selecao_multipla(tree_usuarios)
    criar_barra_selecao_multipla(tela_usuarios, tree_usuarios, excluir_usuarios_em_massa).pack(
        fill='x', padx=20, pady=(0, 8), before=frame_usu_tree)

    criar_controle_paginacao(tela_usuarios, "usuarios", bg=CORES["bg_light"]).pack(
        fill='x', padx=20, pady=(0, 6), before=frame_usu_tree)
    
    def _opcoes_clientes(texto):
        conn = conectar()
        cur = conn.cursor()
        cur.execute("SELECT id, nome, cpf_cnpj, telefone, cidade FROM clientes WHERE excluido=0 ORDER BY id DESC")
        rows = cur.fetchall()
        conn.close()
        t = texto.lower()
        out = []
        for r in rows:
            label = f"{r[0]} - {r[1]}" + (f" | {r[2]}" if r[2] else "") + (f" | {r[4]}" if r[4] else "")
            if t in label.lower() or t in str(r[1]).lower() or t in str(r[2] or "").lower() or t in str(r[3] or "").lower():
                out.append(label)
        return out

    def _opcoes_fornecedores(texto):
        conn = conectar()
        cur = conn.cursor()
        cur.execute("SELECT id, nome, cnpj, telefone, cidade FROM fornecedores WHERE excluido=0 ORDER BY id DESC")
        rows = cur.fetchall()
        conn.close()
        t = texto.lower()
        out = []
        for r in rows:
            label = f"{r[0]} - {r[1]}" + (f" | {r[2]}" if r[2] else "") + (f" | {r[4]}" if r[4] else "")
            if t in label.lower() or t in str(r[1]).lower() or t in str(r[2] or "").lower():
                out.append(label)
        return out

    def _opcoes_produtos(texto):
        conn = conectar()
        cur = conn.cursor()
        cur.execute("SELECT id, codigo, nome FROM produtos WHERE excluido=0 ORDER BY id DESC")
        rows = cur.fetchall()
        conn.close()
        t = texto.lower()
        out = []
        for r in rows:
            label = f"{r[0]} - {r[2]} ({r[1] or '-'})"
            if t in label.lower() or t in str(r[1] or "").lower() or t in str(r[2]).lower():
                out.append(label)
        return out

    def _opcoes_vendas(texto):
        conn = conectar()
        cur = conn.cursor()
        cur.execute("""
            SELECT v.id, COALESCE(c.nome,'Avulso'), v.total, v.status
            FROM vendas v LEFT JOIN clientes c ON v.cliente_id=c.id
            WHERE v.excluido=0 ORDER BY v.id DESC LIMIT 200
        """)
        rows = cur.fetchall()
        conn.close()
        t = texto.lower()
        out = []
        for r in rows:
            label = f"Venda #{r[0]} - {r[1]} | {formatar_moeda(r[2])} | {r[3]}"
            if t in label.lower() or t in str(r[0]) or t in str(r[1]).lower():
                out.append(label)
        return out

    def _opcoes_mov_estoque(texto):
        conn = conectar()
        cur = conn.cursor()
        cur.execute("SELECT id, codigo, nome FROM produtos WHERE excluido=0 ORDER BY id DESC")
        rows = cur.fetchall()
        conn.close()
        t = texto.lower()
        return [f"{r[2]} ({r[1] or '-'})" for r in rows if t in str(r[1] or "").lower() or t in str(r[2]).lower()]

    def _ao_escolher_cliente(valor):
        # Se veio no formato "ID - Nome | ...", extrai só o nome para filtrar
        if " - " in valor:
            try:
                nome = valor.split(" - ", 1)[1].split(" | ")[0].strip()
                entry_busca_cliente.delete(0, tk.END)
                entry_busca_cliente.insert(0, nome)
            except Exception:
                pass
        listar_clientes()

    def _ao_escolher_forn(valor):
        if " - " in valor:
            try:
                nome = valor.split(" - ", 1)[1].split(" | ")[0].strip()
                entry_busca_forn.delete(0, tk.END)
                entry_busca_forn.insert(0, nome)
            except Exception:
                pass
        listar_fornecedores()

    def _ao_escolher_prod(valor):
        if " - " in valor:
            try:
                # "ID - Nome (codigo)"
                parte = valor.split(" - ", 1)[1]
                nome = parte.split(" (")[0].strip()
                entry_busca_prod.delete(0, tk.END)
                entry_busca_prod.insert(0, nome)
            except Exception:
                pass
        listar_produtos()

    def _ao_escolher_venda(valor):
        # Preferência: filtrar pelo ID da venda
        if valor.startswith("Venda #"):
            try:
                vid = valor.split("#", 1)[1].split(" ")[0].strip()
                entry_busca_venda.delete(0, tk.END)
                entry_busca_venda.insert(0, vid)
            except Exception:
                pass
        listar_vendas()

    def _ao_escolher_mov(valor):
        # "Nome (codigo)"
        nome = valor.split(" (")[0].strip() if " (" in valor else valor
        entry_est_filtro_prod.delete(0, tk.END)
        entry_est_filtro_prod.insert(0, nome)
        listar_mov_estoque()

    def configurar_buscas():
        try:
            configurar_busca_com_opcoes(entry_busca_cliente, _opcoes_clientes, _ao_escolher_cliente, min_chars=3)
            configurar_busca_com_opcoes(entry_busca_forn, _opcoes_fornecedores, _ao_escolher_forn, min_chars=3)
            configurar_busca_com_opcoes(entry_busca_prod, _opcoes_produtos, _ao_escolher_prod, min_chars=3)
            configurar_busca_com_opcoes(entry_est_filtro_prod, _opcoes_mov_estoque, _ao_escolher_mov, min_chars=3)
            pass  # filtro por data do caixa removido
        except Exception as e:
            print("Erro configurar buscas:", e)
    
    init_db()
    atualizar_combos()
    listar_clientes()
    listar_fornecedores()
    listar_produtos()
    listar_estoque()
    listar_vendas()
    listar_caixa()
    listar_contas_pagar()
    listar_contas_receber()
    listar_todas_lixeiras()
    listar_usuarios()
    atualizar_dashboard()
    configurar_buscas()
    
    abrir_aba("dashboard", "Dashboard", "📊")
    
    root.mainloop()

if __name__ == "__main__":
    tela_login()
