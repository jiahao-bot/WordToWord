import sqlite3
import hashlib
import datetime
import pandas as pd
import os
# 引入 streamlit 来读取云端配置
import streamlit as st
from dotenv import load_dotenv

load_dotenv()

# 定义一个读取配置的函数：优先读云端 Secrets，读不到再读本地 .env
def get_config(key, default_value):
    try:
        # 尝试从 Streamlit Secrets 读取
        return st.secrets[key]
    except:
        # 如果报错（比如在本地且没有配置 secrets.toml），则读取环境变量
        return os.getenv(key, default_value)

DB_FILE = get_config("DB_NAME", "wordtoword.db")
ADMIN_USER = get_config("ADMIN_USERNAME", "admin")
ADMIN_PASS = get_config("ADMIN_PASSWORD", "admin123")

def init_db():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute(
        '''CREATE TABLE IF NOT EXISTS users (username TEXT PRIMARY KEY, password TEXT, role TEXT, created_at TEXT)''')
    c.execute(
        '''CREATE TABLE IF NOT EXISTS logs (id INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT, action TEXT, timestamp TEXT)''')
    c.execute(
        '''CREATE TABLE IF NOT EXISTS feedback (id INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT, content TEXT, rating INTEGER, timestamp TEXT)''')

    # 使用环境变量中的用户名
    c.execute("SELECT * FROM users WHERE username=?", (ADMIN_USER,))
    if not c.fetchone():
        # 使用环境变量中的密码
        pwd_hash = hashlib.sha256(ADMIN_PASS.encode()).hexdigest()
        c.execute("INSERT INTO users VALUES (?, ?, ?, ?)",
                  (ADMIN_USER, pwd_hash, 'admin', datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
    conn.commit()
    conn.close()


def login_user(username, password):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    pwd_hash = hashlib.sha256(password.encode()).hexdigest()
    c.execute("SELECT role FROM users WHERE username=? AND password=?", (username, pwd_hash))
    res = c.fetchone()
    conn.close()
    return res[0] if res else None


def register_user(username, password):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("SELECT * FROM users WHERE username=?", (username,))
    if c.fetchone():
        conn.close()
        return False
    pwd_hash = hashlib.sha256(password.encode()).hexdigest()
    c.execute("INSERT INTO users VALUES (?, ?, ?, ?)",
              (username, pwd_hash, 'user', datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
    conn.commit()
    conn.close()
    return True


def log_action(username, action):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("INSERT INTO logs (username, action, timestamp) VALUES (?, ?, ?)",
              (username, action, datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
    conn.commit()
    conn.close()


def submit_feedback(username, content, rating):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("INSERT INTO feedback (username, content, rating, timestamp) VALUES (?, ?, ?, ?)",
              (username, content, rating, datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
    conn.commit()
    conn.close()


def get_admin_data():
    conn = sqlite3.connect(DB_FILE)
    users = pd.read_sql("SELECT username, role, created_at FROM users", conn)
    logs = pd.read_sql("SELECT * FROM logs ORDER BY timestamp DESC LIMIT 50", conn)
    fb = pd.read_sql("SELECT * FROM feedback ORDER BY timestamp DESC", conn)
    conn.close()
    return users, logs, fb
