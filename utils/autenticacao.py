import streamlit as st
import json
from datetime import datetime, timedelta
import firebase_admin
from firebase_admin import credentials, auth
import pyrebase
from firebase_config import firebaseConfig

# Inicializar Firebase Admin SDK (apenas uma vez)
if not firebase_admin._apps:
    try:
        # Tentar obter credenciais do Streamlit Secrets
        if 'service_account' in st.secrets:
            service_account_info = st.secrets["service_account"]
            cred = credentials.Certificate(service_account_info)
        else:
            # Fallback para arquivo JSON local (modo de desenvolvimento)
            try:
                cred = credentials.Certificate('serviceAccountKey.json')
            except:
                st.error("Não foi possível carregar as credenciais do Firebase. Verificar configuração.")
                
        firebase_admin.initialize_app(cred)
    except Exception as e:
        # Ativar modo de demonstração automaticamente em caso de erros
        st.session_state.demo_mode = True

# Inicializar Pyrebase com a configuração importada
try:
    firebase = pyrebase.initialize_app(firebaseConfig)
    auth_firebase = firebase.auth()
except Exception as e:
    # Modo silencioso de erro - irá permitir o modo de demonstração
    pass

def verificar_senha():
    """Retorna `True` se a senha estiver correta ou se estiver no modo de demonstração."""
    # Se já autenticado, retorna True
    if "user_authenticated" in st.session_state and st.session_state.user_authenticated:
        return True
    
    # Se estiver em modo de demonstração, retorna True
    if "demo_mode" in st.session_state and st.session_state.demo_mode:
        return True

    # Título de login com logo
    st.markdown(
        """
        <div style="text-align: center; margin-bottom: 30px;">
            <h1 style="color: #5A713D;">Escutaris HSE Analytics</h1>
            <p style="color: #666; font-size: 18px;">Análise de Fatores Psicossociais no Trabalho</p>
        </div>
        """,
        unsafe_allow_html=True
    )
    
    # Formulário de login
    with st.form("login_form"):
        st.markdown("<h3>Login</h3>", unsafe_allow_html=True)
        email = st.text_input("Email")
        senha = st.text_input("Senha", type="password")
        
        col1, col2 = st.columns([1, 1])
        with col1:
            submitted = st.form_submit_button("Login", use_container_width=True)
        with col2:
            # Modo de demonstração para facilitar testes
            demo_mode = st.form_submit_button("Modo Demonstração", use_container_width=True)
    
    # Verificar credenciais do Firebase
    if submitted:
        try:
            # Tente autenticar com Firebase
            user = auth_firebase.sign_in_with_email_and_password(email, senha)
            
            # Armazenar token e timestamp para sessão segura
            st.session_state.user_authenticated = True
            st.session_state.user_info = user
            st.session_state.auth_time = datetime.now()
            st.session_state.auth_expiry = datetime.now() + timedelta(hours=24)  # Expiração de 24h
            
            st.success("Login realizado com sucesso!")
            return True
        except Exception as e:
            st.error("Email ou senha incorretos. Por favor, tente novamente.")
            st.info("Se você não tem uma conta, entre em contato conosco.")
            return False
    
    # Modo de demonstração para testes
    if demo_mode:
        st.session_state.user_authenticated = True
        st.session_state.demo_mode = True
        st.session_state.auth_time = datetime.now()
        st.session_state.auth_expiry = datetime.now() + timedelta(hours=1)  # Expiração mais curta no modo demo
        
        st.warning("Você está usando o modo de demonstração. Alguns recursos podem estar limitados e a sessão expira em 1 hora.")
        return True
    
    # Verificar timeout de sessão
    if "auth_expiry" in st.session_state and datetime.now() > st.session_state.auth_expiry:
        # Limpar dados de autenticação se expirados
        st.session_state.user_authenticated = False
        if "user_info" in st.session_state:
            del st.session_state.user_info
        if "auth_time" in st.session_state:
            del st.session_state.auth_time
        if "auth_expiry" in st.session_state:
            del st.session_state.auth_expiry
            
        st.error("Sua sessão expirou. Por favor, faça login novamente.")
        return False
    
    # Informações de rodapé
    st.markdown(
        """
        <div style="position: fixed; bottom: 20px; text-align: center; width: 100%;">
            <p style="color: #666; font-size: 12px;">
                © 2023 Escutaris | Para acesso, entre em contato: 
                <a href="mailto:contato@escutaris.com.br">contato@escutaris.com.br</a>
            </p>
        </div>
        """,
        unsafe_allow_html=True
    )
        
    return False
