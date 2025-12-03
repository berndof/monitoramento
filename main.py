import datetime
import fnmatch
import logging
import os
import subprocess
import time

import psutil
import win32api
import win32gui
import win32process
from psutil import Process

DEBUG = True
logging.basicConfig(level=logging.DEBUG if DEBUG else logging.INFO)
logger = logging.getLogger("main")

TARGET_PROCESS = "msedge.exe"
START_TIME = datetime.datetime.now()
TIMEOUT = 10  # Segundos

ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
START_SCRIPT_PATH = os.path.join(ROOT_DIR, "open_browser.ps1")

WINDOW_TITLE_TO_MONITOR = {
    "*Grafana": 1,
    "NOC SCC*": 0,
}


def get_process_name_from_hwnd(hwnd):
    """Obtém o PID e o nome do processo a partir do handle da janela (hwnd)."""
    try:
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        process: Process = psutil.Process(pid)
        return pid, process.name()
    except (psutil.NoSuchProcess, psutil.AccessDenied):
        return None, "Desconhecido"


def move_window_to_monitor(hwnd, monitor_index, title):
    """Calcula a posição e move a janela para o monitor especificado."""
    monitors = win32api.EnumDisplayMonitors()
    if monitor_index >= len(monitors):
        exc = Exception(
            f"Monitor {monitor_index} não existe. Apenas {len(monitors)} monitores detectados."
        )
        logger.exception(exc)
        raise exc

    hmonitor = monitors[monitor_index][0]
    mi = win32api.GetMonitorInfo(hmonitor)

    # Coordenadas do monitor de destino (Monitor Rect)
    left, top, right, bottom = mi["Monitor"]
    width, height = right - left, bottom - top

    # Move e redimensiona a janela para ocupar todo o monitor
    win32gui.MoveWindow(hwnd, left, top, width, height, True)

    # Garante que a janela esteja visível e maximizada (SW_SHOWMAXIMIZE = 3, SW_MAXIMIZE = 3, SW_SHOWNORMAL = 1)
    # O 9 (SW_RESTORE) restaura para o estado normal (se minimizada) e a ativa. Usando 3 para maximizar se necessário.
    win32gui.ShowWindow(hwnd, 3)  # SW_SHOWMAXIMIZE = 3

    # Traz a janela para o primeiro plano
    win32gui.SetForegroundWindow(hwnd)

    logger.info(f"Janela '{title}' movida para Monitor {monitor_index} ({left},{top})")


def enum_handler(hwnd, result):
    """Função de callback para EnumWindows. Filtra janelas visíveis do TARGET_PROCESS."""
    if win32gui.IsWindowVisible(hwnd):
        title = win32gui.GetWindowText(hwnd)
        if title:
            pid, pname = get_process_name_from_hwnd(hwnd)
            if pname.lower() == TARGET_PROCESS.lower():
                result.append((hwnd, pid, pname, title))


def title_matches_any_pattern(title: str, patterns: list[str]) -> bool:
    """Verifica se o título corresponde a qualquer um dos padrões usando fnmatch."""
    return any(fnmatch.fnmatchcase(title, pattern) for pattern in patterns)


def get_all_target_windows():
    """Enumera todas as janelas que pertencem ao TARGET_PROCESS."""
    windows = []
    win32gui.EnumWindows(enum_handler, windows)
    return windows


def check_all_expected_windows_opened():
    """Verifica se todas as janelas esperadas estão abertas e retorna-as se sim."""
    all_target_windows = get_all_target_windows()

    # 1. Identificar quais títulos esperados já foram encontrados
    found_patterns = set()
    for _, _, _, title in all_target_windows:
        for pattern in WINDOW_TITLE_TO_MONITOR.keys():
            if fnmatch.fnmatchcase(title, pattern):
                found_patterns.add(pattern)

    expected_patterns = set(WINDOW_TITLE_TO_MONITOR.keys())

    # 2. Verificar se todos os padrões esperados foram encontrados
    if found_patterns == expected_patterns:
        logger.debug("Todas as janelas esperadas foram abertas.")
        return all_target_windows
    else:
        missing_patterns = expected_patterns - found_patterns
        logger.debug(f"Ainda faltam as janelas com os padrões: {missing_patterns}")
        return None


def wait_windows():
    """Espera até que todas as janelas esperadas sejam abertas ou ocorra um timeout."""
    start_time = time.time()

    while time.time() - start_time < TIMEOUT:
        windows = check_all_expected_windows_opened()
        if windows:
            logger.info("Todas as janelas encontradas.")
            return windows
        time.sleep(0.5)  # Espera 500ms antes de tentar novamente

    logger.exception(f"Timeout de {TIMEOUT}s: Nem todas as janelas apareceram.")
    raise TimeoutError("Uma janela esperada não foi encontrada após o timeout.")


def get_monitors():
    """Retorna a lista completa de todos os monitores detectados."""
    monitors = win32api.EnumDisplayMonitors()

    if not monitors:
        exc = Exception("Nenhum monitor detectado.")
        logger.exception(exc)
        raise exc

    for i, monitor in enumerate(monitors):
        monitor_info = win32api.GetMonitorInfo(monitor[0])
        logger.debug(
            f"Monitor {i}: {monitor_info['Monitor']} (Trabalho: {monitor_info['Work']})"
        )

    # CORREÇÃO: Removido o 'return' que estava dentro do loop.
    return monitors


def main():
    logger.info("Iniciando o script de gerenciamento de janelas.")

    # 1. Tenta encontrar janelas já abertas
    windows = check_all_expected_windows_opened()

    # 2. Se as janelas não foram encontradas, executa o script Powershell e espera
    if not windows:
        logger.info(
            f"Janelas não encontradas. Abrindo browser via {START_SCRIPT_PATH}..."
        )
        try:
            # TODO Comando com 'start-process' para Powershell pode ser mais limpo
            subprocess.run(
                [
                    "powershell",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-File",
                    START_SCRIPT_PATH,
                ],
                check=True,
            )
            logger.debug(
                f"Comando Powershell executado. Esperando {TIMEOUT}s pelas janelas..."
            )
            windows = wait_windows()
        except TimeoutError:
            return
        except Exception as e:
            logger.exception(f"Erro ao executar o script Powershell: {e}")
            return

    # 3. Obtém informações dos monitores
    monitors = get_monitors()
    num_monitors = len(monitors)

    if len(WINDOW_TITLE_TO_MONITOR) > num_monitors:
        logger.warning(
            f"Mais janelas para posicionar ({len(WINDOW_TITLE_TO_MONITOR)}) do que monitores detectados ({num_monitors})."
        )

    # 4. Move as janelas encontradas
    moved_titles = set()

    for hwnd, pid, pname, title in windows:
        for title_pattern, monitor_index in WINDOW_TITLE_TO_MONITOR.items():
            if fnmatch.fnmatchcase(title, title_pattern) and title not in moved_titles:
                try:
                    move_window_to_monitor(hwnd, monitor_index, title)
                    moved_titles.add(title)  # Marca a janela como movida
                    break
                except Exception as e:
                    logger.error(
                        f"Não foi possível mover a janela '{title}' para o monitor {monitor_index}: {e}"
                    )
        else:
            logger.warning(
                f"Janela '{title}' não corresponde a nenhum padrão em WINDOW_TITLE_TO_MONITOR e não será movida."
            )


if __name__ == "__main__":
    main()
