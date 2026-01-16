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

"""Configuration management for the PPT Generate project."""

import os
from pathlib import Path
from typing import Optional
from dataclasses import dataclass
from dotenv import load_dotenv

from .logging import Logger


# Load environment variables from .env file
load_dotenv()


logger = Logger.get_logger(__name__)


@dataclass
class SearxngConfig:
    """Configuration for Searxng search service."""
    url: str = "http://47.100.251.249:8888/search"
    timeout: int = 30
    max_concurrent: int = 10
    default_count: int = 10


@dataclass
class ImageGenConfig:
    """Configuration for image generation service."""
    url: str = "http://47.100.251.249:38089/generate_image"
    timeout: int = 60
    default_denoising_strength: float = 1.0
    default_height: int = 512
    default_width: int = 512
    default_num_inference_steps: int = 30


@dataclass
class CrawlerConfig:
    """Configuration for web crawling."""
    timeout: int = 30
    max_concurrent: int = 5
    user_agent_mode: str = "random"
    enable_stealth: bool = True


@dataclass
class GeneratorConfig:
    """Configuration for PPT generation."""
    model_name: str = "gpt-4"  # Default LLM model
    temperature: float = 0.7
    max_tokens: int = 4000
    max_slides_per_request: int = 5  # Limit slides per LLM call to avoid timeout
    api_key: Optional[str] = None
    base_url: Optional[str] = None


@dataclass
class ExportConfig:
    """Configuration for PPT export."""
    marp_cli_path: Optional[str] = None
    pandoc_path: Optional[str] = None
    theme_dir: str = "templates/themes"
    template_dir: str = "templates"


@dataclass
class MCPConfig:
    """Configuration for MCP server."""
    host: str = "localhost"
    port: int = 8000
    max_workers: int = 4


class Config:
    """Main configuration class that loads from environment variables."""

    def __init__(self) -> None:
        """Initialize configuration from environment variables."""
        self._load_config()

    def _load_config(self) -> None:
        """Load configuration from environment variables."""
        # Searxng configuration
        self.searxng = SearxngConfig(
            url=os.getenv("SEARXNG_URL", SearxngConfig.url),
            timeout=int(os.getenv("SEARXNG_TIMEOUT", SearxngConfig.timeout)),
            max_concurrent=int(os.getenv("SEARXNG_MAX_CONCURRENT", SearxngConfig.max_concurrent)),
            default_count=int(os.getenv("SEARXNG_DEFAULT_COUNT", SearxngConfig.default_count)),
        )

        # Image generation configuration
        self.image_gen = ImageGenConfig(
            url=os.getenv("IMAGE_GEN_URL", ImageGenConfig.url),
            timeout=int(os.getenv("IMAGE_GEN_TIMEOUT", ImageGenConfig.timeout)),
            default_denoising_strength=float(os.getenv("IMAGE_GEN_DENOISING_STRENGTH", ImageGenConfig.default_denoising_strength)),
            default_height=int(os.getenv("IMAGE_GEN_HEIGHT", ImageGenConfig.default_height)),
            default_width=int(os.getenv("IMAGE_GEN_WIDTH", ImageGenConfig.default_width)),
            default_num_inference_steps=int(os.getenv("IMAGE_GEN_STEPS", ImageGenConfig.default_num_inference_steps)),
        )

        # Crawler configuration
        self.crawler = CrawlerConfig(
            timeout=int(os.getenv("CRAWLER_TIMEOUT", CrawlerConfig.timeout)),
            max_concurrent=int(os.getenv("CRAWLER_MAX_CONCURRENT", CrawlerConfig.max_concurrent)),
            user_agent_mode=os.getenv("CRAWLER_USER_AGENT_MODE", CrawlerConfig.user_agent_mode),
            enable_stealth=os.getenv("CRAWLER_ENABLE_STEALTH", str(CrawlerConfig.enable_stealth)).lower() == "true",
        )

        # Generator configuration
        # Support multiple environment variable names for flexibility
        model_name = os.getenv("LLM_MODEL_NAME") or os.getenv("GENERATOR_MODEL") or GeneratorConfig.model_name
        api_key = os.getenv("LLM_API_KEY") or os.getenv("OPENAI_API_KEY")
        base_url = os.getenv("LLM_BASE_URL") or os.getenv("OPENAI_BASE_URL")

        self.generator = GeneratorConfig(
            model_name=model_name,
            temperature=float(os.getenv("GENERATOR_TEMPERATURE", GeneratorConfig.temperature)),
            max_tokens=int(os.getenv("GENERATOR_MAX_TOKENS", GeneratorConfig.max_tokens)),
            max_slides_per_request=int(os.getenv("GENERATOR_MAX_SLIDES_PER_REQUEST", GeneratorConfig.max_slides_per_request)),
            api_key=api_key,
            base_url=base_url,
        )

        # Export configuration
        self.export = ExportConfig(
            marp_cli_path=os.getenv("MARP_CLI_PATH"),
            pandoc_path=os.getenv("PANDOC_PATH"),
            theme_dir=os.getenv("THEME_DIR", ExportConfig.theme_dir),
            template_dir=os.getenv("TEMPLATE_DIR", ExportConfig.template_dir),
        )

        # MCP configuration
        self.mcp = MCPConfig(
            host=os.getenv("MCP_HOST", MCPConfig.host),
            port=int(os.getenv("MCP_PORT", MCPConfig.port)),
            max_workers=int(os.getenv("MCP_MAX_WORKERS", MCPConfig.max_workers)),
        )

        # General configuration
        self.log_level = os.getenv("LOG_LEVEL", "INFO")
        self.log_file = os.getenv("LOG_FILE")

        # Project paths
        self.project_root = Path(__file__).parent.parent.parent
        self.src_dir = self.project_root / "src"
        self.templates_dir = self.project_root / "templates"
        self.web_dir = self.project_root / "web"

    def setup_logging(self) -> None:
        """Setup logging based on configuration."""
        Logger.set_level(self.log_level)
        if self.log_file:
            Logger.add_file_handler(self.log_file, self.log_level)
        logger.info("Logging configured successfully")


# Global configuration instance
config = Config()