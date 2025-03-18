import streamlit as st

def verificar_senha():
    """Retorna `True` se a senha estiver correta."""
    if "password_correct" not in st.session_state:
        st.session_state.password_correct = False

    if st.session_state.password_correct:
        return True

    # Formulário de login
    with st.form("login_form"):
        st.text_input("Usuário", key="username")
        st.text_input("Senha", type="password", key="password")
        submitted = st.form_submit_button("Login")

    if submitted:
        if st.session_state.password == "senha123" and st.session_state.username == "admin":
            st.session_state.password_correct = True
            return True
        else:
            st.error("Usuário ou senha incorretos")
            return False
    return False
