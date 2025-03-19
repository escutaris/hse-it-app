import streamlit as st
import pyrebase
import os
import firebase_admin
from firebase_admin import credentials, auth
from dotenv import load_dotenv

# Carregar variáveis de ambiente
load_dotenv()

# Configuração do Firebase
firebaseConfig = {
  "apiKey": os.getenv("FIREBASE_API_KEY"),
  "authDomain": os.getenv("FIREBASE_AUTH_DOMAIN"),
  "projectId": os.getenv("FIREBASE_PROJECT_ID"),
  "storageBucket": os.getenv("FIREBASE_STORAGE_BUCKET"),
  "messagingSenderId": os.getenv("FIREBASE_MESSAGING_SENDER_ID"),
  "appId": os.getenv("FIREBASE_APP_ID"),
  "measurementId": os.getenv("FIREBASE_MEASUREMENT_ID"),
  "databaseURL": os.getenv("FIREBASE_DATABASE_URL", "")
}

# Inicializar Firebase Admin SDK (apenas uma vez)
if not firebase_admin._apps:
    try:
        # Tentar carregar o arquivo de credenciais
        cred = credentials.Certificate("serviceAccountKey.json")
        firebase_admin.initialize_app(cred)
    except Exception as e:
        st.error(f"Erro ao inicializar Firebase Admin: {e}")
        st.info("Verifique se o arquivo serviceAccountKey.json está na raiz do projeto")

# Inicializar Firebase para autenticação client-side
firebase = pyrebase.initialize_app(firebaseConfig)
auth_firebase = firebase.auth()

def verificar_senha():
    """Verifica a autenticação do usuário via Firebase."""
    if "user_authenticated" in st.session_state and st.session_state.user_authenticated:
        return True

    # Adicionar logo/imagem
    st.image("https://raw.githubusercontent.com/username/hse-app/main/logo.png", width=200)
    
    # Formulário de login
    with st.form("login_form"):
        st.title("Escutaris HSE Analytics")
        st.markdown("Entre com suas credenciais para acessar o sistema")
        
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
            return True
        except Exception as e:
            st.error("Email ou senha incorretos. Por favor, tente novamente.")
            st.info("Se você não tem uma conta, entre em contato conosco.")
            return False
    
    # Modo de demonstração para testes
    if demo_mode:
        st.session_state.user_authenticated = True
        st.session_state.demo_mode = True
        return True
        
    return False
