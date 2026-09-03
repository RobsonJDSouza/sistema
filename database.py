#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Módulo de banco de dados do Sistema de Gestão.
Arquivo físico: sistema.db (mesma pasta deste arquivo / do programa principal).
"""

import os
import sqlite3
import hashlib
from datetime import date

# Caminho do arquivo SQLite (ao lado deste módulo)
DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "sistema.db")


def conectar():
    """Abre conexão SQLite com foreign keys ativadas."""
    conn = sqlite3.connect(DB_PATH)
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def hash_senha(senha):
    """Hash SHA-256 da senha."""
    return hashlib.sha256(str(senha).encode()).hexdigest()


def _hoje_iso():
    return date.today().isoformat()


def migrar_tabela_check(nome_tabela, create_sql_novo):
    """Migra tabela antiga com CHECK restritivo para nova sem CHECK."""
    conn = conectar()
    cur = conn.cursor()
    try:
        cur.execute("SELECT sql FROM sqlite_master WHERE type='table' AND name=?", (nome_tabela,))
        row = cur.fetchone()
        if row and row[0]:
            sql_old = row[0].lower()
            if "check" in sql_old and ("pendente" in sql_old or "concluida" in sql_old):
                print(f"🔧 Migrando tabela {nome_tabela} - removendo CHECK antigo...")
                cur.execute(f"ALTER TABLE {nome_tabela} RENAME TO {nome_tabela}_old")
                cur.execute(create_sql_novo)
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
        except Exception:
            pass
    conn.close()
    return False


def init_db():
    """Cria tabelas, aplica migrações e usuários padrão."""
    conn = conectar()
    cur = conn.cursor()

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
                except Exception:
                    pass
    conn.commit()

    try:
        cur.execute("UPDATE vendas SET status='em_aberto' WHERE status='concluida' AND excluido=0")
        cur.execute("UPDATE vendas SET status='em_aberto' WHERE status='pendente' AND excluido=0")
        cur.execute("UPDATE contas_a_receber SET status='em_aberto' WHERE status='pendente' AND excluido=0")
        cur.execute("UPDATE contas_a_pagar SET status='em_aberto' WHERE status='pendente' AND excluido=0")
        cur.execute("UPDATE contas_a_receber SET status='em_atraso' WHERE status='vencido' AND excluido=0")
        cur.execute("UPDATE contas_a_pagar SET status='em_atraso' WHERE status='vencido' AND excluido=0")
        cur.execute("UPDATE contas_a_receber SET status='recebido' WHERE status='pago' AND excluido=0")
        conn.commit()
    except Exception:
        pass

    try:
        hoje = _hoje_iso()
        cur.execute("UPDATE contas_a_receber SET status='em_atraso' WHERE status='em_aberto' AND date(vencimento) < date(?) AND excluido=0", (hoje,))
        cur.execute("UPDATE contas_a_pagar SET status='em_atraso' WHERE status='em_aberto' AND date(vencimento) < date(?) AND excluido=0", (hoje,))
        conn.commit()
    except Exception:
        pass

    cur.execute("SELECT COUNT(*) FROM usuarios WHERE excluido=0")
    if cur.fetchone()[0] == 0:
        cur.execute("INSERT INTO usuarios (nome, login, senha, perfil, email) VALUES (?,?,?,?,?)",
                    ("Administrador", "admin", hash_senha("admin123"), "admin", "admin@sistema.local"))
        cur.execute("INSERT INTO usuarios (nome, login, senha, perfil, email) VALUES (?,?,?,?,?)",
                    ("Operador", "operador", hash_senha("operador123"), "operador", "operador@sistema.local"))
    else:
        try:
            cur.execute("UPDATE usuarios SET email='admin@sistema.local' WHERE login='admin' AND (email IS NULL OR email='')")
            cur.execute("UPDATE usuarios SET email='operador@sistema.local' WHERE login='operador' AND (email IS NULL OR email='')")
        except Exception:
            pass

    conn.commit()
    conn.close()
