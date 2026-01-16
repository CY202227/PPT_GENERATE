# Copyright 2024 PPT Generate Project
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#     http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.

"""Logging utilities for the PPT Generate project."""

import logging
import sys
from typing import Optional


class Logger:
    """A logger that follows Google style logging practices."""

    _logger: Optional[logging.Logger] = None
    _is_configured = False

    @classmethod
    def get_logger(cls, name: str = "ppt_generate") -> logging.Logger:
        """Get or create a logger instance.

        Args:
            name: The name of the logger, defaults to "ppt_generate".

        Returns:
            A configured logger instance.
        """
        if cls._logger is None:
            cls._logger = logging.getLogger(name)
            if not cls._is_configured:
                cls._configure_logger()
                cls._is_configured = True

        return cls._logger

    @classmethod
    def _configure_logger(cls) -> None:
        """Configure the logger with Google-style formatting and handlers."""
        if cls._logger is None:
            return

        # Remove any existing handlers
        for handler in cls._logger.handlers[:]:
            cls._logger.removeHandler(handler)

        # Set log level
        cls._logger.setLevel(logging.INFO)

        # Create console handler with Google-style formatter
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setLevel(logging.INFO)

        # Google-style log format: [LEVEL] timestamp file:line message
        formatter = logging.Formatter(
            '[%(levelname)s] %(asctime)s %(filename)s:%(lineno)d %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        console_handler.setFormatter(formatter)

        cls._logger.addHandler(console_handler)

        # Prevent duplicate logs from parent loggers
        cls._logger.propagate = False

    @classmethod
    def set_level(cls, level: str) -> None:
        """Set the logging level.

        Args:
            level: The logging level (DEBUG, INFO, WARNING, ERROR, CRITICAL).
        """
        if cls._logger is None:
            cls.get_logger()

        level_map = {
            'DEBUG': logging.DEBUG,
            'INFO': logging.INFO,
            'WARNING': logging.WARNING,
            'ERROR': logging.ERROR,
            'CRITICAL': logging.CRITICAL,
        }

        if level.upper() in level_map:
            cls._logger.setLevel(level_map[level.upper()])
            for handler in cls._logger.handlers:
                handler.setLevel(level_map[level.upper()])

    @classmethod
    def add_file_handler(cls, log_file: str, level: str = "INFO") -> None:
        """Add a file handler to the logger.

        Args:
            log_file: Path to the log file.
            level: The logging level for the file handler.
        """
        if cls._logger is None:
            cls.get_logger()

        file_handler = logging.FileHandler(log_file, encoding='utf-8')
        file_handler.setLevel(getattr(logging, level.upper()))

        formatter = logging.Formatter(
            '[%(levelname)s] %(asctime)s %(filename)s:%(lineno)d %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        file_handler.setFormatter(formatter)

        cls._logger.addHandler(file_handler)