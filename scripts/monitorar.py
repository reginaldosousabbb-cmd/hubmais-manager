"""
HUBMAIS MANAGER — Monitor automático
=======================================
Fica de olho nas pastas dados/tma e dados/resultados. Quando você joga um
relatório novo lá (baixado do sistema), ele detecta sozinho, espera o arquivo
terminar de ser gravado, e roda atualizar_base.py + metricas.py automaticamente.

USO:
    python scripts/monitorar.py

Deixa essa janela aberta rodando em segundo plano. Pra rodar sem precisar
deixar o terminal aberto o tempo todo, veja a opção "Tarefa Agendada do
Windows" no fim deste arquivo (comentário).

Dependência (instala uma vez):
    pip install watchdog
"""

import subprocess
import sys
import time
from pathlib import Path

try:
    from watchdog.observers import Observer
    from watchdog.events import FileSystemEventHandler
except ImportError:
    print("❌ Falta instalar a biblioteca 'watchdog'. Rode:")
    print("   pip install watchdog")
    sys.exit(1)

PASTA_PROJETO = Path(__file__).resolve().parent.parent
PASTA_TMA = PASTA_PROJETO / "dados" / "tma"
PASTA_RESULTADOS = PASTA_PROJETO / "dados" / "resultados"
SCRIPT_ATUALIZAR = PASTA_PROJETO / "scripts" / "atualizar_base.py"
SCRIPT_METRICAS = PASTA_PROJETO / "scripts" / "metricas.py"
SCRIPT_SITE = PASTA_PROJETO / "scripts" / "exportar_site.py"

# Tempo de espera após detectar o arquivo, pra garantir que o Excel/Windows
# terminou de gravar (evita ler um .xlsx pela metade e dar erro).
SEGUNDOS_DEBOUNCE = 8


def arquivo_terminou_de_gravar(caminho: Path, tentativas: int = 6, intervalo: float = 1.0) -> bool:
    """Confere se o tamanho do arquivo parou de mudar (indício de que a gravação terminou)."""
    tamanho_anterior = -1
    for _ in range(tentativas):
        if not caminho.exists():
            return False
        tamanho_atual = caminho.stat().st_size
        if tamanho_atual == tamanho_anterior and tamanho_atual > 0:
            return True
        tamanho_anterior = tamanho_atual
        time.sleep(intervalo)
    return False


def rodar_pipeline():
    print("\n" + "=" * 60)
    print("🔄 Novo relatório detectado — atualizando base e painel...")
    print("=" * 60)
    try:
        subprocess.run([sys.executable, str(SCRIPT_ATUALIZAR)], check=True)
        subprocess.run([sys.executable, str(SCRIPT_METRICAS)], check=True)
        subprocess.run([sys.executable, str(SCRIPT_SITE)], check=True)
        print("\n✅ Tudo atualizado (base, painel Excel e dashboard do site).\n")
    except subprocess.CalledProcessError as erro:
        print(f"\n❌ Deu erro ao atualizar (código {erro.returncode}). Veja a mensagem acima.\n")


class HandlerRelatorioNovo(FileSystemEventHandler):
    def __init__(self):
        self._ultimo_disparo = 0

    def _tratar_evento(self, caminho_str: str):
        caminho = Path(caminho_str)
        if caminho.suffix.lower() != ".xlsx" or caminho.name.startswith("~$"):
            return  # ignora arquivos temporários do Excel (~$arquivo.xlsx)

        print(f"📂 Detectado: {caminho.name}")
        if not arquivo_terminou_de_gravar(caminho):
            print("   ⚠️  Não deu pra confirmar que terminou de gravar, tentando mesmo assim.")

        # Evita disparar duas vezes pro mesmo arquivo em sequência rápida
        agora = time.time()
        if agora - self._ultimo_disparo < SEGUNDOS_DEBOUNCE:
            return
        self._ultimo_disparo = agora

        time.sleep(2)  # pequena folga extra
        rodar_pipeline()

    def on_created(self, event):
        if not event.is_directory:
            self._tratar_evento(event.src_path)

    def on_modified(self, event):
        if not event.is_directory:
            self._tratar_evento(event.src_path)


if __name__ == "__main__":
    print("=" * 60)
    print("👀 HUBMAIS MANAGER — Monitor automático rodando")
    print(f"   Vigiando: {PASTA_TMA}")
    print(f"   Vigiando: {PASTA_RESULTADOS}")
    print("   (Ctrl+C pra parar)")
    print("=" * 60)

    handler = HandlerRelatorioNovo()
    observer = Observer()
    observer.schedule(handler, str(PASTA_TMA), recursive=False)
    observer.schedule(handler, str(PASTA_RESULTADOS), recursive=False)
    observer.start()

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()


# ==========================================================================
# ALTERNATIVA SEM DEIXAR TERMINAL ABERTO: Tarefa Agendada do Windows
# ==========================================================================
# Se preferir não deixar essa janela rodando, dá pra usar o Agendador de
# Tarefas do Windows pra chamar atualizar_base.py + metricas.py de tempos em
# tempos (ex.: a cada 15 min), em vez de reagir na hora. Passos:
#
# 1. Abra "Agendador de Tarefas" (Task Scheduler)
# 2. Criar Tarefa Básica > nome "HUBMAIS Atualizar Base"
# 3. Disparador: Diariamente, repetir a cada 15 minutos, durante o expediente
# 4. Ação: Iniciar um programa
#      Programa: python
#      Argumentos: scripts\atualizar_base.py
#      Iniciar em: C:\Users\PCUSER\Downloads\REGINALDO\HUBMAIS_MANAGER
# 5. Repita criando uma 2ª tarefa igual, só que pra scripts\metricas.py
#
# Essa opção não reage na hora que o arquivo cai (espera até 15 min), mas não
# depende de janela aberta. O monitorar.py (watchdog) reage na hora, mas
# precisa ficar rodando.
