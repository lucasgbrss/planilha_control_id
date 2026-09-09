from datetime import datetime
import json
import os
from pathlib import Path


APP_CONFIG_DIR = "ControlIDReader"


def default_config_path():
    """Retorna o caminho padrao do config.json do app."""
    appdata = Path(os.environ.get("APPDATA", Path.home() / "AppData" / "Roaming"))
    return appdata / APP_CONFIG_DIR / "config.json"


def load_config(config_path=None):
    """Carrega configuracoes salvas. Retorna dict vazio quando nao houver config valida."""
    path = Path(config_path) if config_path else default_config_path()
    if not path.exists():
        return {}

    return parse_config_content(path.read_text(encoding="utf-8"))


def parse_config_content(content):
    """Converte conteudo JSON em dict de config, tolerando conteudo invalido."""
    content = (content or "").strip()
    if not content:
        return {}

    try:
        data = json.loads(content)
    except json.JSONDecodeError:
        return {}

    return data if isinstance(data, dict) else {}


def save_config(config, config_path=None):
    """Salva configuracoes em JSON e retorna o caminho usado."""
    path = Path(config_path) if config_path else default_config_path()
    path.parent.mkdir(parents=True, exist_ok=True)

    path.write_text(serialize_config(config), encoding="utf-8")
    return path


def serialize_config(config):
    """Serializa configuracoes para JSON, adicionando data de execucao quando ausente."""
    payload = dict(config)
    payload.setdefault("data_ultima_execucao", datetime.now().isoformat())
    return json.dumps(payload, indent=2, ensure_ascii=False)
