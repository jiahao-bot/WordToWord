import sqlite3
import hashlib
import datetime
import pandas as pd
import os
import streamlit as st
from dotenv import load_dotenv
import json

load_dotenv()


# 定义配置读取
def get_config(key, default_value):
    try:
        return st.secrets[key]
    except:
        return os.getenv(key, default_value)


DB_FILE = get_config("DB_NAME", "wordtoword.db")
ADMIN_USER = get_config("ADMIN_USERNAME", "admin")
ADMIN_PASS = get_config("ADMIN_PASSWORD", "admin123")


def init_db():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    # 用户表
    c.execute(
        '''CREATE TABLE IF NOT EXISTS users (username TEXT PRIMARY KEY, password TEXT, role TEXT, created_at TEXT)''')
    # 日志表
    c.execute(
        '''CREATE TABLE IF NOT EXISTS logs (id INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT, action TEXT, timestamp TEXT)''')
    # 反馈表
    c.execute(
        '''CREATE TABLE IF NOT EXISTS feedback (id INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT, content TEXT, rating INTEGER, timestamp TEXT)''')

    # 【新增 1】用户配置表 (用于记忆 API Key 等设置)
    c.execute('''CREATE TABLE IF NOT EXISTS user_config (username TEXT PRIMARY KEY, api_key TEXT, updated_at TEXT)''')
    # 【新增 1.1】模型配置表 (用于记忆模型相关设置)
    c.execute(
        '''CREATE TABLE IF NOT EXISTS user_model_config (username TEXT PRIMARY KEY, config_json TEXT, updated_at TEXT)''')

    # 【新增 2】用户档案表 (用于记忆上传过的简历/文档内容)
    c.execute(
        '''CREATE TABLE IF NOT EXISTS profiles (id INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT, profile_name TEXT, content_text TEXT, created_at TEXT)''')

    # 初始化管理员
    c.execute("SELECT * FROM users WHERE username=?", (ADMIN_USER,))
    if not c.fetchone():
        pwd_hash = hashlib.sha256(ADMIN_PASS.encode()).hexdigest()
        c.execute("INSERT INTO users VALUES (?, ?, ?, ?)",
                  (ADMIN_USER, pwd_hash, 'admin', datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
    conn.commit()
    conn.close()


# --- 用户认证 ---
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


# --- 配置记忆 (API Key) ---
def save_user_apikey(username, api_key):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    # 插入或更新
    c.execute("INSERT OR REPLACE INTO user_config (username, api_key, updated_at) VALUES (?, ?, ?)",
              (username, api_key, timestamp))
    conn.commit()
    conn.close()


def get_user_apikey(username):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("SELECT api_key FROM user_config WHERE username=?", (username,))
    res = c.fetchone()
    conn.close()
    return res[0] if res else ""


# --- 模型配置记忆 ---
def save_user_model_config(username, config):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    c.execute("INSERT OR REPLACE INTO user_model_config (username, config_json, updated_at) VALUES (?, ?, ?)",
              (username, json.dumps(config, ensure_ascii=False), timestamp))
    conn.commit()
    conn.close()


def get_user_model_config(username):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("SELECT config_json FROM user_model_config WHERE username=?", (username,))
    res = c.fetchone()
    conn.close()
    if not res:
        return {}
    try:
        return json.loads(res[0])
    except json.JSONDecodeError:
        return {}


# --- 档案记忆 (简历内容) ---
def save_profile(username, profile_name, content_text):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    # 检查是否已存在同名档案，存在则更新
    c.execute("SELECT id FROM profiles WHERE username=? AND profile_name=?", (username, profile_name))
    exist = c.fetchone()
    if exist:
        c.execute("UPDATE profiles SET content_text=?, created_at=? WHERE id=?", (content_text, timestamp, exist[0]))
    else:
        c.execute("INSERT INTO profiles (username, profile_name, content_text, created_at) VALUES (?, ?, ?, ?)",
                  (username, profile_name, content_text, timestamp))
    conn.commit()
    conn.close()


def get_user_profiles(username):
    conn = sqlite3.connect(DB_FILE)
    # 返回 profile_name 列表
    df = pd.read_sql(
        "SELECT profile_name, content_text, created_at FROM profiles WHERE username=? ORDER BY created_at DESC", conn,
        params=(username,))
    conn.close()
    return df


def delete_profile(username, profile_name):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("DELETE FROM profiles WHERE username=? AND profile_name=?", (username, profile_name))
    conn.commit()
    conn.close()


# --- 日志与反馈 ---
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
