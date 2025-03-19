import streamlit as st
import json
import firebase_admin
from firebase_admin import credentials, auth
import pyrebase

# Inicializar Firebase Admin SDK (apenas uma vez)
if not firebase_admin._apps:
    try:
        # Obter credenciais do Streamlit Secrets
        service_account_info = st.secrets["service_account"]
        
        # Convertendo o dicionário para o formato esperado pelo Firebase Admin
        cred = credentials.Certificate(service_account_info)
        firebase_admin.initialize_app(cred)
        
        print("Firebase Admin SDK inicializado com sucesso")
    except Exception as e:
        st.error(f"Erro ao inicializar Firebase Admin: {e}")
        print(f"Erro ao inicializar Firebase Admin: {e}")

# Configuração do Firebase para autenticação client-side
firebase_config = {
    "apiKey": st.secrets["firebase"]["api_key"],
    "authDomain": st.secrets["firebase"]["auth_domain"],
    "projectId": st.secrets["firebase"]["project_id"],
    "storageBucket": st.secrets["firebase"]["storage_bucket"],
    "messagingSenderId": st.secrets["firebase"]["messaging_sender_id"],
    "appId": st.secrets["firebase"]["app_id"],
    "databaseURL": st.secrets["firebase"]["database_url"]
}

# Inicializar Pyrebase
firebase = pyrebase.initialize_app(firebase_config)
auth_firebase = firebase.auth()

def verificar_senha():
    """Verifica a autenticação do usuário via Firebase."""
    if "user_authenticated" in st.session_state and st.session_state.user_authenticated:
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
            st.session_state.user_authenticated = True
            st.session_state.user_info = user
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
        st.warning("Você está usando o modo de demonstração. Alguns recursos podem estar limitados.")
        return True
    
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
