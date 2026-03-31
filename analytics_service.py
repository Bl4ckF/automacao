import sqlite3
import pandas as pd
from datetime import datetime

DB_NAME = "banco.db"

ERROS_PADRAO = [
    ("CÂMERA", "Câmera virada"),
    ("CÂMERA", "Câmera é tampada"),
    ("CÂMERA", "Eletricista fica na frente da câmera intencionalmente"),
    ("CÂMERA", "Lente suja"),
    ("CÂMERA", "Gravação pausada"),
    ("CÂMERA", "Gravação sem áudio"),
    ("CÂMERA", "Câmera trocada"),
    ("CÂMERA", "Equipe se afasta da câmera"),
    ("CÂMERA", "Equipe não aparece na gravação"),
    ("CÂMERA", "Má posicionamento da câmera"),
    ("CÂMERA", "Derrubou a câmera"),
    ("CÂMERA", "Atividade não foi gravada"),
    ("CÂMERA", "Não foi gravado a atividade por completo"),
    ("CÂMERA", "Maior parte da atividade não foi gravado"),
    ("CÂMERA", "Não foi gravado o preenchimento da APR"),
    ("CÂMERA", "Não foi gravado o início da atividade"),
    ("CÂMERA", "Não foi gravado o término da atividade"),

    ("VEÍCULO", "Subiu na lateral da carroceria"),
    ("VEÍCULO", "Subiu na traseira da carroceria"),
    ("VEÍCULO", "Subiu na grade do caminhão"),
    ("VEÍCULO", "Eletricista na carroceria com carga em movimentação"),
    ("VEÍCULO", "Eletricista de baixo de carga suspensa"),
    ("VEÍCULO", "Eletricista sobe no poste suspenso"),
    ("VEÍCULO", "Eletricista sobe na broca de perfuração"),
    ("VEÍCULO", "Eletricista senta do poste suspenso"),
    ("VEÍCULO", "Calço não colocado"),
    ("VEÍCULO", "Uso inadequado dos estabilizadores"),

    ("ESTRUTURA", "Subiu na estrutura do poste"),
    ("ESTRUTURA", "Subiu nas cruzetas"),
    ("ESTRUTURA", "Subiu no transformador"),
    ("ESTRUTURA", "Subiu no cabo de transmissão"),
    ("ESTRUTURA", "Subiu no cabo de comunicação"),
    ("ESTRUTURA", "Subiu no telhado"),
    ("ESTRUTURA", "Subiu no braço da iluminária pública"),
    ("ESTRUTURA", "Sentou nas cruzetas"),
    ("ESTRUTURA", "Sentou no transformador"),
    ("ESTRUTURA", "Sentou no cabo de transmissão"),
    ("ESTRUTURA", "Sentou no cabo de comunicação"),
    
    ("EQUIPE", "Jogou lixo no chão"),
    ("EQUIPE", "Eletricista sobe na escada com a linha de vida solta"),
    ("EQUIPE", "Equipe apoia o poste suspenso na escada"),
    ("EQUIPE", "Içamento inadequado"),
    ("EQUIPE", "Descida inadequada"),
    ("EQUIPE", "Subida inadequada"),
    ("EQUIPE", "Escada não foi segurada na subida do eletricista"),
    ("EQUIPE", "Uso inadequado do cinto de segurança"),
    ("EQUIPE", "Uso inadequado das ferramentas"),
    ("EQUIPE", "Uso inadequado dos EPI's"),
    ("EQUIPE", "Área isolada inadequada"),
    ("EQUIPE", "Área não foi isolada"),
    ("EQUIPE", "Eletricista passa por cima da área isolada"),
    ("EQUIPE", "Jogou os cones no chão"),
    ("EQUIPE", "Jogou Ferramentas no chão"),
    ('EQUIPE', "Jogou objeto/ferramenta de cima do poste"),
    ("EQUIPE", "jogou ferramenta/objeto do chão para o eletricista no poste"),
    ('EQUIPE', "Deixou ferramenta/objeto nos cabos de transmissão/comunicação"),
    ("EQUIPE", "O Guardião não está supervisionando o eletricista no poste"),
    ("EQUIPE", "Eletricista está de maneira inadequada na escada"),

    ("PERCURSO", "Acima da velocidade permitida"),
    ("PERCURSO", "Ultrapassagem indevida"),
    ("PERCURSO", "Não sinalizou a presença do veículo"),
    ("PERCURSO", "Estacionou em local proibido"),
    ("PERCURSO", "Não utilizou o cinto de segurança"),
    ("PERCURSO", "Furou o sinal vermelho"),
    ("PERCURSO", "Não respeitou a sinalização de trânsito"),
    
    ("TERCEIRO", "Pedestre entrou na área isolada"),
    ("TERCEIRO", "Terceiro sobe na estrutura do poste"),
    ("TERCEIRO", "Terceiro sobe no transformador"),
    ("TERCEIRO", "Terceiro realiza atividade não autorizada no poste"),  
]

_cache_equipes = None
_cache_datas = None

def limpar_cache():
    global _cache_equipes, _cache_datas
    _cache_equipes = None
    _cache_datas = None 

def init_db():
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    
    # Criar tabela de relatorios
    c.execute("""
        CREATE TABLE IF NOT EXISTS relatorios (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            equipe TEXT,
            data TEXT,
            status TEXT DEFAULT 'Não Conforme',
            data_criacao TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)
    
    # Criar tabela de erros
    c.execute("""
        CREATE TABLE IF NOT EXISTS erros (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            relatorio_id INTEGER,
            descricao TEXT,
            FOREIGN KEY (relatorio_id) REFERENCES relatorios(id)
        )
    """)
    
    # Criar tabela de erros_padrao
    c.execute("""
        CREATE TABLE IF NOT EXISTS erros_padrao (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            categoria TEXT,
            descricao TEXT UNIQUE
        )
    """)
    
    # Inserir erros padrão se a tabela estiver vazia
    c.execute("SELECT COUNT(*) FROM erros_padrao")
    if c.fetchone()[0] == 0:
        c.executemany(
            "INSERT OR IGNORE INTO erros_padrao (categoria, descricao) VALUES (?, ?)",
            ERROS_PADRAO
        )
    
    # Criar índices
    c.execute("CREATE INDEX IF NOT EXISTS idx_relatorios_equipe ON relatorios(equipe)")
    c.execute("CREATE INDEX IF NOT EXISTS idx_relatorios_data ON relatorios(data)")
    c.execute("CREATE INDEX IF NOT EXISTS idx_relatorios_status ON relatorios(status)")
    c.execute("CREATE INDEX IF NOT EXISTS idx_erros_relatorio_id ON erros(relatorio_id)")

    conn.commit()
    conn.close()
    return atualizar_esquema_banco()

def atualizar_nome_equipe(relatorio_id, novo_nome_equipe):
    """Atualiza o nome da equipe de um relatório"""
    if not novo_nome_equipe or not novo_nome_equipe.strip():
        return False
    
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    
    try:
        c.execute(
            "UPDATE relatorios SET equipe = ? WHERE id = ?",
            (novo_nome_equipe.strip(), relatorio_id)
        )
        conn.commit()
        alterado = c.rowcount > 0
        return alterado
    except Exception as e:
        print(f"Erro ao atualizar nome da equipe: {e}")
        return False
    finally:
        limpar_cache()
        conn.close()
    
def excluir_relatorio(relatorio_id):
    """Exclui um relatório e seus erros"""
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    
    try:
        c.execute("DELETE FROM erros WHERE relatorio_id = ?", (relatorio_id,))
        c.execute("DELETE FROM relatorios WHERE id = ?", (relatorio_id,))
        conn.commit()
        excluido = c.rowcount > 0
        return excluido
    except Exception as e:
        print(f"Erro ao excluir relatório: {e}")
        return False
    finally:
        limpar_cache()
        conn.close()
    
def atualizar_status_relatorio(relatorio_id, novo_status):
    """Atualiza o status de um relatório"""
    if novo_status not in ["Conforme", "Não Conforme"]:
        return False
    
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    
    try:
        c.execute(
            "UPDATE relatorios SET status = ? WHERE id = ?",
            (novo_status, relatorio_id)
        )
        conn.commit()
        alterado = c.rowcount > 0
        return alterado
    except Exception as e:
        print(f"Erro ao atualizar status: {e}")
        return False
    finally:
        limpar_cache()
        conn.close()
    
def listar_erros():
    """Lista todos os erros padrão"""
    conn = sqlite3.connect(DB_NAME)
    df = pd.read_sql(
        "SELECT categoria, descricao FROM erros_padrao ORDER BY categoria, descricao",
        conn
    )
    conn.close()
    return df

def salvar_relatorio(equipe, data, erros, status="Não Conforme"):
    relatorio_id = None
    conn = None
    try:
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()

        if status not in ["Conforme", "Não Conforme"]:
            status = "Não Conforme"

        # Inserir relatório
        c.execute(
            "INSERT INTO relatorios (equipe, data, status) VALUES (?, ?, ?)",
            (equipe, data, status)
        )
        
        relatorio_id = c.lastrowid

        # Inserir erros
        for desc in erros:
            c.execute(
                "INSERT INTO erros (relatorio_id, descricao) VALUES (?, ?)",
                (relatorio_id, desc)
            )

        conn.commit()
        print(f"✅ Relatório salvo: ID={relatorio_id}")
        
    except Exception as e:
        print(f"❌ Erro ao salvar relatório: {e}")
        if conn:
            conn.rollback()
        raise e
        
    finally:
        if conn:
            conn.close()
    
    # Exportar para Excel unificado (um arquivo com duas abas)
    try:
        from exportador_completo import exportar_dados_para_excel_fixo
        
        resultado = exportar_dados_para_excel_fixo()
        if resultado and resultado.get('sucesso'):
            print("✅ Arquivo Excel unificado atualizado com sucesso!")
            print(f"   • Arquivo: {resultado.get('arquivo', 'N/A')}")
            print(f"   • Aba ERROS  : {resultado.get('total_erros', 0)} registros")
            print(f"   • Aba STATUS : {resultado.get('total_status', 0)} registros")
        else:
            erro = resultado.get('erro', 'Erro desconhecido') if resultado else 'Resultado vazio'
            print(f"⚠️ Falha na exportação unificada: {erro}")
            
    except ImportError:
        print("⚠️ Módulo exportador_completo não encontrado. Exportação não realizada.")
    except Exception as e:
        print(f"⚠️ Erro na exportação unificada: {e}")
        import traceback
        traceback.print_exc()
    
    limpar_cache()
    return relatorio_id

def erros_por_descricao():
    """Conta erros por descrição"""
    conn = sqlite3.connect(DB_NAME)
    df = pd.read_sql("""
        SELECT descricao, COUNT(*) as total
        FROM erros
        GROUP BY descricao
        ORDER BY total DESC
    """, conn)
    conn.close()
    return df

def erros_por_equipe():
    """Conta erros por equipe"""
    conn = sqlite3.connect(DB_NAME)
    df = pd.read_sql("""
        SELECT r.equipe, COUNT(e.id) as total
        FROM erros e
        JOIN relatorios r ON r.id = e.relatorio_id
        GROUP BY r.equipe
        ORDER BY total DESC
    """, conn)
    conn.close()
    return df

def listar_relatorios(filtro_equipe="", filtro_data="", filtro_erro="", filtro_status=""):
    """Lista relatórios com filtros"""
    conn = sqlite3.connect(DB_NAME)

    query = """
        SELECT 
            r.id,
            r.equipe,
            r.data,
            r.status,
            r.data_criacao,
            GROUP_CONCAT(e.descricao, '; ') as erros,
            COUNT(e.id) as total_erros
        FROM relatorios r
        LEFT JOIN erros e ON r.id = e.relatorio_id
        WHERE 1=1
    """

    params = []

    if filtro_equipe:
        query += " AND r.equipe LIKE ?"
        params.append(f"%{filtro_equipe}%")

    if filtro_data:
        query += " AND r.data = ?"
        params.append(filtro_data)
    
    if filtro_status:
        query += " AND r.status = ?"
        params.append(filtro_status)

    query += """
        GROUP BY r.id, r.equipe, r.data, r.status, r.data_criacao
        ORDER BY r.data_criacao DESC
    """

    df = pd.read_sql(query, conn, params=params)
    conn.close()

    if filtro_erro:
        mask = df['erros'].str.contains(filtro_erro, case=False, na=False)
        df = df[mask]

    return df

def obter_detalhes_relatorio(relatorio_id):
    """Obtém detalhes de um relatório específico"""
    conn = sqlite3.connect(DB_NAME)
    
    df_relatorio = pd.read_sql("""
        SELECT * FROM relatorios WHERE id = ?
    """, conn, params=(relatorio_id,))
    
    df_erros = pd.read_sql("""
        SELECT descricao FROM erros 
        WHERE relatorio_id = ?
        ORDER BY id
    """, conn, params=(relatorio_id,))
    
    conn.close()
    return df_relatorio, df_erros

def listar_datas_disponiveis(usar_cache=True):
    global _cache_datas
    if usar_cache and _cache_datas is not None:
        return _cache_datas
    conn = sqlite3.connect(DB_NAME)
    df = pd.read_sql("SELECT DISTINCT data FROM relatorios ORDER BY data DESC", conn)
    conn.close()
    _cache_datas = df['data'].tolist()
    return _cache_datas

def listar_equipes_disponiveis(usar_cache=True):
    global _cache_equipes
    if usar_cache and _cache_equipes is not None:
        return _cache_equipes
    conn = sqlite3.connect(DB_NAME)
    df = pd.read_sql("SELECT DISTINCT equipe FROM relatorios ORDER BY equipe", conn)
    conn.close()
    _cache_equipes = df['equipe'].tolist()
    return _cache_equipes

def listar_status_disponiveis():
    """Lista status disponíveis"""
    conn = sqlite3.connect(DB_NAME)
    df = pd.read_sql("""
        SELECT DISTINCT status 
        FROM relatorios 
        ORDER BY status
    """, conn)
    conn.close()
    return df['status'].tolist()

def estatisticas_por_status():
    """Estatísticas de relatórios por status"""
    conn = sqlite3.connect(DB_NAME)
    df = pd.read_sql("""
        SELECT 
            status,
            COUNT(*) as quantidade,
            ROUND(COUNT(*) * 100.0 / (SELECT COUNT(*) FROM relatorios), 2) as percentual
        FROM relatorios
        GROUP BY status
        ORDER BY quantidade DESC
    """, conn)
    conn.close()
    return df

def atualizar_esquema_banco():
    """Atualiza o esquema do banco de dados se necessário"""
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    
    try:
        # Verificar se a coluna 'status' existe na tabela relatorios
        c.execute("PRAGMA table_info(relatorios)")
        colunas = c.fetchall()
        colunas_existentes = [col[1] for col in colunas]
        
        # Se não existe a coluna status, adicionar
        if 'status' not in colunas_existentes:
            print("➕ Adicionando coluna 'status' à tabela relatorios...")
            try:
                c.execute("ALTER TABLE relatorios ADD COLUMN status TEXT DEFAULT 'Não Conforme'")
                print("✅ Coluna 'status' adicionada com sucesso")
            except Exception as e:
                print(f"⚠️ Erro ao adicionar coluna: {e}")
                print("🔄 Recriando tabela...")
                
                c.execute("""
                    CREATE TABLE relatorios_nova (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        equipe TEXT,
                        data TEXT,
                        status TEXT DEFAULT 'Não Conforme',
                        data_criacao TIMESTAMP
                    )
                """)
                
                c.execute("""
                    INSERT INTO relatorios_nova (id, equipe, data, data_criacao, status)
                    SELECT id, equipe, data, data_criacao, 'Não Conforme'
                    FROM relatorios
                """)
                
                c.execute("DROP TABLE relatorios")
                c.execute("ALTER TABLE relatorios_nova RENAME TO relatorios")
                print("✅ Tabela recriada com coluna 'status'")
        
        conn.commit()
        return True
        
    except Exception as e:
        print(f"❌ Erro ao atualizar esquema: {e}")
        conn.rollback()
        return False
        
    finally:
        conn.close()

# Função de compatibilidade para exportação (caso seja chamada externamente)
def exportar_dados_para_excel_fixo():
    """Função de compatibilidade que chama a exportação unificada."""
    try:
        from exportador_completo import exportar_dados_para_excel_fixo as export_unificado
        return export_unificado()
    except ImportError:
        print("❌ Módulo exportador_completo não encontrado.")
        return {'sucesso': False, 'erro': 'Módulo exportador_completo não encontrado'}

def sincronizar_erros_padrao():
    """Sincroniza a tabela erros_padrao com a lista ERROS_PADRAO definida no código.
    - Insere novos erros.
    - Remove erros que não estão mais na lista (desde que não estejam vinculados a nenhum relatório).
    - Atualiza categoria se necessário.
    """
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    
    # Obter todos os erros atuais no banco
    c.execute("SELECT categoria, descricao FROM erros_padrao")
    erros_atuais = {desc: cat for cat, desc in c.fetchall()}
    
    # Converter ERROS_PADRAO para dicionário {descricao: categoria}
    novos_erros = {desc: cat for cat, desc in ERROS_PADRAO}
    
    # 1. Inserir novos erros
    for desc, cat in novos_erros.items():
        if desc not in erros_atuais:
            c.execute("INSERT INTO erros_padrao (categoria, descricao) VALUES (?, ?)", (cat, desc))
            print(f"➕ Erro adicionado: {cat} - {desc}")
    
    # 2. Atualizar categoria de erros existentes (se mudou)
    for desc, cat in novos_erros.items():
        if desc in erros_atuais and erros_atuais[desc] != cat:
            c.execute("UPDATE erros_padrao SET categoria = ? WHERE descricao = ?", (cat, desc))
            print(f"✏️ Categoria atualizada: {desc} -> {cat}")
    
    # 3. Remover erros que não estão mais na lista (somente se não usados em nenhum relatório)
    for desc in erros_atuais:
        if desc not in novos_erros:
            # Verificar se o erro já foi usado em algum relatório
            c.execute("SELECT COUNT(*) FROM erros WHERE descricao = ?", (desc,))
            count = c.fetchone()[0]
            if count == 0:
                c.execute("DELETE FROM erros_padrao WHERE descricao = ?", (desc,))
                print(f"🗑️ Erro removido (não utilizado): {desc}")
            else:
                print(f"⚠️ Erro não removido pois está em {count} relatório(s): {desc}")
    
    conn.commit()
    conn.close()
    print("✅ Sincronização de erros padrão concluída.")

__all__ = [
    'init_db',
    'listar_erros',
    'salvar_relatorio',
    'erros_por_descricao',
    'erros_por_equipe',
    'listar_relatorios',
    'obter_detalhes_relatorio',
    'excluir_relatorio',
    'listar_datas_disponiveis',
    'listar_equipes_disponiveis',
    'atualizar_status_relatorio',
    'listar_status_disponiveis',
    'estatisticas_por_status',
    'atualizar_nome_equipe',
    'exportar_dados_para_excel_fixo',
    'sincronizar_erros_padrao',
]