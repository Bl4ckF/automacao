# updater.py
import os
import sys
import time
import shutil
import subprocess
import tempfile
import requests
import asyncio

class GitHubUpdater:
    def __init__(self, github_repo, target_file, current_version, force_kill=False):
        self.github_repo = github_repo          # ex: "Bl4ckF/automacao"
        self.target_file = target_file          # ex: "app.exe"
        self.current_version = current_version  # ex: "1.0.0"
        self.force_kill = force_kill

    async def update(self):
        try:
            # URL correta da API
            api_url = f"https://api.github.com/repos/Bl4ckF/automacao/releases/latest"
            headers = {'Accept': 'application/vnd.github+json'}
            
            # Se o repositório for privado, descomente a linha abaixo e adicione seu token
            # headers['Authorization'] = 'token SEU_TOKEN_AQUI'

            response = requests.get(api_url, headers=headers, timeout=10)
            response.raise_for_status()
            release_data = response.json()

            latest_version = release_data.get("tag_name", "").lstrip('v')
            if not latest_version:
                print("Nenhuma tag de versão encontrada na release.")
                return False

            if latest_version <= self.current_version:
                print(f"Versão já atualizada: {self.current_version} == {latest_version}")
                return False

            # Procura o asset .exe
            asset_url = None
            for asset in release_data.get("assets", []):
                if asset["name"] == self.target_file:
                    asset_url = asset["browser_download_url"]
                    break

            if not asset_url:
                print(f"Arquivo '{self.target_file}' não encontrado nos assets da release.")
                return False

            print(f"Baixando nova versão {latest_version}...")
            with tempfile.NamedTemporaryFile(delete=False, suffix=".exe") as tmp_file:
                download_response = requests.get(asset_url, stream=True)
                download_response.raise_for_status()
                for chunk in download_response.iter_content(chunk_size=8192):
                    tmp_file.write(chunk)
                novo_exe = tmp_file.name
            print("Download concluído.")

            # Caminho do executável atual
            if getattr(sys, 'frozen', False):
                exe_atual = sys.executable
            else:
                exe_atual = os.path.abspath(self.target_file)

            # Cria script temporário para substituição
            aux_script = self._criar_script_substituicao(novo_exe, exe_atual)
            subprocess.Popen([sys.executable, aux_script])
            print("Aplicando atualização...")
            sys.exit(0)

        except Exception as e:
            print(f"Erro durante atualização: {e}")
            return False

    def _criar_script_substituicao(self, novo_exe, alvo_exe):
        script_content = f'''import os, time, shutil, subprocess, sys
time.sleep(3)
try:
    if os.path.exists(r"{alvo_exe}"):
        os.remove(r"{alvo_exe}")
    shutil.move(r"{novo_exe}", r"{alvo_exe}")
    subprocess.Popen([r"{alvo_exe}"])
except Exception as e:
    with open(r"C:\\temp\\updater_error.log", "w") as f:
        f.write(str(e))
'''
        fd, script_path = tempfile.mkstemp(suffix=".py", text=True)
        with os.fdopen(fd, 'w') as f:
            f.write(script_content)
        return script_path