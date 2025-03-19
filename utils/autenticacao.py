import streamlit as st
import pyrebase
from firebase_config import firebaseConfig

# Inicializar Firebase
firebase = pyrebase.initialize_app(firebaseConfig)
auth = firebase.auth()

def verificar_senha():
    """Verifica a autenticação do usuário via Firebase."""
    if "user_authenticated" in st.session_state and st.session_state.user_authenticated:
        return True

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
            user = auth.sign_in_with_email_and_password(email, senha)
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
