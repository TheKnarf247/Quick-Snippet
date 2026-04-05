import configparser
from pathlib import Path


class ConfigManager:
    def __init__(self, config_file: str):
        self.config_file = Path(config_file)
        self.config = configparser.ConfigParser()

    def load(self) -> configparser.ConfigParser:
        self.config.read(self.config_file, encoding="utf-8")
        return self.config

    def save(self) -> None:
        with open(self.config_file, "w", encoding="utf-8") as fh:
            self.config.write(fh)