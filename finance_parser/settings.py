from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any
import copy
import re

try:
    import yaml
except ImportError as exc:  # pragma: no cover - dependency guidance path
    raise ImportError(
        "PyYAML is required for settings support. Install dependencies with: "
        "python -m pip install -r requirements.txt"
    ) from exc


PROJECT_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_EXAMPLE_SETTINGS_PATH = PROJECT_ROOT / "config" / "settings.example.yaml"
DEFAULT_USER_SETTINGS_PATH = PROJECT_ROOT / "config" / "settings.yaml"


def _deep_merge(base: dict[str, Any], override: dict[str, Any]) -> dict[str, Any]:
    """
    Recursively merge override values into base without mutating inputs.
    """
    result = copy.deepcopy(base)

    for key, value in override.items():
        if (
            key in result
            and isinstance(result[key], dict)
            and isinstance(value, dict)
        ):
            result[key] = _deep_merge(result[key], value)
        else:
            result[key] = copy.deepcopy(value)

    return result


def _load_yaml_file(path: Path) -> dict[str, Any]:
    if not path.exists():
        return {}

    with path.open("r", encoding="utf-8") as handle:
        data = yaml.safe_load(handle) or {}

    if not isinstance(data, dict):
        raise ValueError(f"Settings file must contain a YAML mapping at top level: {path}")

    return data


def _normalise_key(value: object) -> str:
    return "" if value is None else str(value).strip().upper()


@dataclass(frozen=True)
class AppSettings:
    """
    Loaded application settings.

    This class is intentionally small for the first configuration commit.
    Later parser code can progressively migrate hardcoded personal assumptions
    to these helper methods.
    """

    data: dict[str, Any]
    project_root: Path = PROJECT_ROOT

    @classmethod
    def load(
        cls,
        settings_path: Path | None = None,
        example_path: Path | None = None,
    ) -> "AppSettings":
        example_path = example_path or DEFAULT_EXAMPLE_SETTINGS_PATH
        settings_path = settings_path or DEFAULT_USER_SETTINGS_PATH

        example = _load_yaml_file(example_path)
        user = _load_yaml_file(settings_path)

        merged = _deep_merge(example, user)
        settings = cls(data=merged)
        settings.validate()
        return settings

    def validate(self) -> None:
        required_top_level = ["project", "paths", "budgeting", "investments"]
        missing = [key for key in required_top_level if key not in self.data]
        if missing:
            raise ValueError(f"Missing top-level settings section(s): {', '.join(missing)}")

        required_paths = [
            "budgeting_input",
            "budgeting_output",
            "budgeting_rules",
            "investment_input",
            "investment_output",
            "investment_rules",
        ]
        paths = self.data.get("paths", {})
        missing_paths = [key for key in required_paths if key not in paths]
        if missing_paths:
            raise ValueError(f"Missing settings.paths value(s): {', '.join(missing_paths)}")

    def get(self, *keys: str, default: Any = None) -> Any:
        current: Any = self.data
        for key in keys:
            if not isinstance(current, dict) or key not in current:
                return default
            current = current[key]
        return current

    def path(self, key: str) -> Path:
        raw = self.get("paths", key)
        if raw is None:
            raise KeyError(f"Missing path setting: paths.{key}")

        path = Path(str(raw)).expanduser()
        if path.is_absolute():
            return path
        return self.project_root / path

    def canonical_source_bank(self, source_bank: object) -> str:
        bank = _normalise_key(source_bank)
        aliases = self.get("budgeting", "source_bank_aliases", default={}) or {}
        aliases_norm = {_normalise_key(k): _normalise_key(v) for k, v in aliases.items()}
        return aliases_norm.get(bank, bank)

    def fixed_source_account(self, source_bank: object) -> str | None:
        bank = self.canonical_source_bank(source_bank)
        config = self.get("budgeting", "source_account_inference", bank, default={}) or {}
        value = config.get("fixed_source_account")
        return _normalise_key(value) if value else None

    def filename_source_account(self, source_bank: object, filename: str | Path) -> str | None:
        bank = self.canonical_source_bank(source_bank)
        config = self.get("budgeting", "source_account_inference", bank, default={}) or {}
        prefixes = config.get("filename_prefixes", {}) or {}

        stem = Path(filename).stem.upper()
        tokens = [token for token in re.split(r"[^A-ZÅÄÖ0-9]+", stem) if token]

        for prefix, account in prefixes.items():
            prefix_norm = _normalise_key(prefix)
            if prefix_norm in tokens or stem.startswith(prefix_norm):
                return _normalise_key(account)

        return None

    def default_include_for_source_bank(self, source_bank: object) -> str:
        bank = self.canonical_source_bank(source_bank)
        bank_config = self.get("budgeting", "source_account_inference", bank, default={}) or {}

        if "default_include" in bank_config:
            return _normalise_key(bank_config["default_include"])

        return _normalise_key(self.get("budgeting", "default_include", default="YES"))

    def broker_default(self, broker: object, key: str, default: Any = None) -> Any:
        broker_key = _normalise_key(broker)
        return self.get("investments", "broker_defaults", broker_key, key, default=default)

    def _portfolio_setting(self, section: str, broker: object, portfolio: object) -> str:
        """
        Look up a per-(Broker, Portfolio) investment setting.

        The config value for a broker may be a plain string (applies to every
        portfolio under that broker) or a mapping from Portfolio to value,
        with an optional "default" key as a broker-wide fallback. This mirrors
        how Owner is deliberately kept out of parser code on the budgeting
        side - which account/portfolio belongs to which person is a
        deployment-specific fact that belongs in settings, not in
        InstrumentMaster (content-based rules) or parser code.
        """
        broker_key = _normalise_key(broker)
        portfolio_key = _normalise_key(portfolio)
        config = self.get("investments", section, broker_key, default=None)

        if config is None:
            return ""
        if isinstance(config, str):
            return _normalise_key(config)
        if isinstance(config, dict):
            normalised = {_normalise_key(k): v for k, v in config.items()}
            if portfolio_key in normalised:
                return _normalise_key(normalised[portfolio_key])
            if "DEFAULT" in normalised:
                return _normalise_key(normalised["DEFAULT"])
        return ""

    def portfolio_owner(self, broker: object, portfolio: object) -> str:
        return self._portfolio_setting("portfolio_owners", broker, portfolio)

    def portfolio_type(self, broker: object, portfolio: object) -> str:
        return self._portfolio_setting("portfolio_types", broker, portfolio)


_SETTINGS_CACHE: AppSettings | None = None


def get_settings() -> AppSettings:
    global _SETTINGS_CACHE
    if _SETTINGS_CACHE is None:
        _SETTINGS_CACHE = AppSettings.load()
    return _SETTINGS_CACHE


def reload_settings() -> AppSettings:
    global _SETTINGS_CACHE
    _SETTINGS_CACHE = AppSettings.load()
    return _SETTINGS_CACHE
