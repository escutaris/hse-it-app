import streamlit as st
import json
import os
import hashlib
from datetime import datetime, timedelta

# Arquivo para armazenar usuários (pode ser substituído por um DB posteriormente)
USERS_FILE = "data/usuarios.json"

# Garantir que o diretório existe
os.makedirs(os.path.dirname(USERS_FILE), exist_ok=True)

# Inicializar arquivo de usuários se não existir
if not os.path.exists(USERS_FILE):
    default_users = {
        "admin@escutaris.com.br": {
            "senha_hash": hashlib.sha256("senha123".encode()).hexdigest(),
            "nome": "Administrador",
            "empresa": "Escutaris",
            "plano": "admin",
            "validade": "2099-12-31"
        },
        "demo@exemplo.com": {
            "senha_hash": hashlib.sha256("demo123".encode()).hexdigest(),
            "nome": "Usuário Demo",
            "empresa": "Demo Ltda",
            "plano": "demo",
            "validade": "2099-12-31"
        }
    }
    with open(USERS_FILE, 'w') as f:
        json.dump(default_users, f, indent=4)

def carregar_usuarios():
    """Carrega a lista de usuários do arquivo JSON"""
    try:
        with open(USERS_FILE, 'r') as f:
            return json.load(f)
    except Exception as e:
        st.error(f"Erro ao carregar usuários: {str(e)}")
        return {}

def adicionar_usuario(email, senha, nome, empresa, plano="basico", validade_dias=30):
    """Adiciona um novo usuário ao sistema"""
    usuarios = carregar_usuarios()
    
    # Calcular data de validade
    hoje = datetime.now()
    validade = (hoje + timedelta(days=validade_dias)).strftime("%Y-%m-%d")
    
    # Criar hash da senha
    senha_hash = hashlib.sha256(senha.encode()).hexdigest()
    
    # Adicionar usuário
    usuarios[email] = {
        "senha_hash": senha_hash,
        "nome": nome,
        "empresa": empresa,
        "plano": plano,
        "validade": validade
    }
    
    # Salvar de volta no arquivo
    with open(USERS_FILE, 'w') as f:
        json.dump(usuarios, f, indent=4)
    
    return True

def verificar_senha(email, senha):
    """Verifica se as credenciais são válidas e se a licença está ativa"""
    usuarios = carregar_usuarios()
    
    if email not in usuarios:
        return False
    
    usuario = usuarios[email]
    senha_hash = hashlib.sha256(senha.encode()).hexdigest()
    
    if senha_hash != usuario["senha_hash"]:
        return False
    
    # Verificar validade
    try:
        validade = datetime.strptime(usuario["validade"], "%Y-%m-%d").date()
        hoje = datetime.now().date()
        
        if hoje > validade:
            return "expirado"
    except Exception:
        # Se houver erro no formato de data, passar adiante
        pass
    
    return usuario

def verificar_autenticacao():
    """Função principal para verificar se usuário está autenticado"""
    # Para compatibilidade com código existente
    return verificar_senha()

def verificar_senha():
    """Retorna `True` se a senha estiver correta ou se estiver no modo de demonstração."""
    # Se já autenticado, retorna True
    if "user_authenticated" in st.session_state and st.session_state.user_authenticated:
        # Verificar se a sessão não expirou
        if "auth_expiry" in st.session_state:
            if datetime.now() < st.session_state.auth_expiry:
                return True
            else:
                # Limpar sessão expirada
                st.session_state.user_authenticated = False
                st.error("Sua sessão expirou. Por favor, faça login novamente.")
    
    # Interface de login
    st.markdown(
        """
        <div style="text-align: center; margin-bottom: 30px;">
            <h1 style="color: #5A713D;">Escutaris HSE Analytics</h1>
            <p style="color: #666; font-size: 18px;">Análise de Fatores Psicossociais no Trabalho</p>
        </div>
        """,
        unsafe_allow_html=True
    )
    
    # Tabs para Login e Modo Demo
    tab1, tab2 = st.tabs(["Login", "Demonstração"])
    
    with tab1:
        with st.form("login_form"):
            email = st.text_input("Email")
            senha = st.text_input("Senha", type="password")
            submitted = st.form_submit_button("Entrar", use_container_width=True)
            
            if submitted:
                resultado = verificar_senha(email, senha)
                
                if resultado == "expirado":
                    st.error("Sua licença expirou. Entre em contato com a Escutaris para renovação.")
                    st.info("Email: contato@escutaris.com.br")
                    return False
                
                if resultado:
                    # Login bem-sucedido
                    st.session_state.user_authenticated = True
                    st.session_state.user_info = resultado
                    st.session_state.auth_time = datetime.now()
                    st.session_state.auth_expiry = datetime.now() + timedelta(hours=8)
                    st.success(f"Bem-vindo, {resultado['nome']} da {resultado['empresa']}!")
                    return True
                else:
                    st.error("Email ou senha incorretos.")
    
    with tab2:
        st.info("O modo de demonstração permite testar as funcionalidades do aplicativo com dados de exemplo.")
        if st.button("Acessar modo demonstração", use_container_width=True):
            st.session_state.user_authenticated = True
            st.session_state.demo_mode = True
            st.session_state.user_info = {
                "nome": "Usuário Demo",
                "empresa": "Demonstração",
                "plano": "demo"
            }
            st.session_state.auth_time = datetime.now()
            st.session_state.auth_expiry = datetime.now() + timedelta(hours=1)
            st.success("Modo demonstração ativado! Você tem acesso por 1 hora.")
            st.warning("No modo de demonstração, algumas funcionalidades podem estar limitadas.")
            return True
    
    # Informações de contato
    st.markdown(
        """
        <div style="text-align: center; margin-top: 30px;">
            <p style="color: #666;">Ainda não tem acesso? Entre em contato:</p>
            <p style="color: #5A713D; font-weight: bold;">contato@escutaris.com.br | (75) 98221-7557</p>
        </div>
        """,
        unsafe_allow_html=True
    )
    
    return False

# Função para página de admin (opcional, pode ser implementada mais tarde)
def pagina_admin():
    """Página de administração para gerenciar usuários"""
    if not st.session_state.get("user_authenticated") or st.session_state.get("user_info", {}).get("plano") != "admin":
        st.error("Acesso não autorizado")
        return
    
    st.title("Administração de Usuários")
    
    # Carregar usuários
    usuarios = carregar_usuarios()
    
    # Interface para adicionar usuário
    with st.expander("Adicionar Novo Usuário"):
        with st.form("novo_usuario"):
            col1, col2 = st.columns(2)
            with col1:
                novo_email = st.text_input("Email")
                nova_senha = st.text_input("Senha", type="password")
                nova_empresa = st.text_input("Empresa")
            with col2:
                novo_nome = st.text_input("Nome")
                novo_plano = st.selectbox("Plano", ["basico", "premium", "admin"])
                validade_dias = st.number_input("Validade (dias)", min_value=1, value=30)
            
            adicionar = st.form_submit_button("Adicionar Usuário")
            
            if adicionar:
                if not novo_email or not nova_senha or not novo_nome or not nova_empresa:
                    st.error("Todos os campos são obrigatórios")
                else:
                    sucesso = adicionar_usuario(novo_email, nova_senha, novo_nome, nova_empresa, novo_plano, validade_dias)
                    if sucesso:
                        st.success(f"Usuário {novo_email} adicionado com sucesso!")
                        st.experimental_rerun()  # Atualizar a página
    
    # Listar usuários
    st.subheader("Usuários Cadastrados")
    for email, info in usuarios.items():
        with st.expander(f"{email} - {info['empresa']}"):
            st.write(f"**Nome:** {info['nome']}")
            st.write(f"**Plano:** {info['plano']}")
            st.write(f"**Validade:** {info['validade']}")
            
            if st.button(f"Excluir {email}", key=f"del_{email}"):
                # Remover usuário
                del usuarios[email]
                with open(USERS_FILE, 'w') as f:
                    json.dump(usuarios, f, indent=4)
                st.success(f"Usuário {email} removido!")
                st.experimental_rerun()
