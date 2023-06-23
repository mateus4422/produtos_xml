from git import Repo

# Caminho para o repositório local
repo_path = r'C:\Users\Mateus Ramos\PycharmProjects\Python\Conversão de Codigo'

# Inicializar o repositório
repo = Repo(repo_path)

# Adicionar os arquivos para o commit
repo.index.add(['con', 'arquivo2.txt'])

# Fazer o commit com uma mensagem
commit_message = 'Adicionando arquivos'
repo.index.commit(commit_message)

# Fazer o push para o repositório remoto
origin = repo.remote('origin')
origin.push()
