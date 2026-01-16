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

"""Export service for converting Marp markdown to HTML and PPTX formats."""

import os
import subprocess
import tempfile
import shutil
from pathlib import Path
from typing import Optional, Dict, Any, Union

from src.utils.config import config
from src.utils.logging import Logger


logger = Logger.get_logger(__name__)


class ExportService:
    """Service for exporting PPT content to various formats using Marp and Pandoc."""

    def __init__(self) -> None:
        """Initialize the export service."""
        self.marp_cli_path = self._find_marp_cli()
        self.pandoc_path = self._find_pandoc()

        logger.info(f"Export service initialized - Marp: {self.marp_cli_path}, Pandoc: {self.pandoc_path}")

    def _find_marp_cli(self) -> Optional[str]:
        """Find marp-cli executable path."""
        # Check configured path first
        if config.export.marp_cli_path and Path(config.export.marp_cli_path).exists():
            return config.export.marp_cli_path

        # Try common installation paths
        common_paths = [
            "marp",  # If in PATH
            "/usr/local/bin/marp",
            "/usr/bin/marp",
            "npx marp",  # Use npx if available
        ]

        for path in common_paths:
            if self._check_command_exists(path):
                return path

        logger.warning("marp-cli not found. HTML/PPTX export may not work.")
        return None

    def _find_pandoc(self) -> Optional[str]:
        """Find pandoc executable path."""
        # Check configured path first
        if config.export.pandoc_path and Path(config.export.pandoc_path).exists():
            return config.export.pandoc_path

        # Try common paths
        common_paths = [
            "pandoc",  # If in PATH
            "/usr/local/bin/pandoc",
            "/usr/bin/pandoc",
        ]

        for path in common_paths:
            if self._check_command_exists(path):
                return path

        logger.warning("pandoc not found. PPTX export via pandoc may not work.")
        return None

    def _check_command_exists(self, command: str) -> bool:
        """Check if a command exists and is executable."""
        try:
            # If it's a multi-part command like 'npx marp'
            if " " in command:
                cmd_parts = command.split()
                # Check if the primary executable exists
                if not shutil.which(cmd_parts[0]):
                    return False
                
                # Try running with --version to be sure
                result = subprocess.run(
                    cmd_parts + ["--version"],
                    capture_output=True,
                    text=True,
                    timeout=5,
                    shell=True if os.name == 'nt' else False
                )
                return result.returncode == 0
            
            # For single commands
            return shutil.which(command) is not None
        except (subprocess.TimeoutExpired, FileNotFoundError, Exception):
            return False

    def export_to_html(
        self,
        marp_markdown: str,
        output_path: Union[str, Path],
        theme: str = "default",
        options: Optional[Dict[str, Any]] = None
    ) -> bool:
        """
        Export Marp markdown to HTML format.

        Args:
            marp_markdown: The Marp markdown content
            output_path: Path to save the HTML file
            theme: Marp theme to use
            options: Additional export options

        Returns:
            True if export successful, False otherwise
        """
        if not self.marp_cli_path:
            logger.error("marp-cli not available for HTML export")
            return False

        output_path = Path(output_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)

        logger.info(f"Exporting to HTML: {output_path}")

        # Create temporary markdown file
        with tempfile.NamedTemporaryFile(mode='w', suffix='.md', delete=False, encoding='utf-8') as temp_file:
            temp_file.write(marp_markdown)
            temp_md_path = temp_file.name

        try:
            # Build marp command
            cmd = [self.marp_cli_path]

            # Add theme if specified
            if theme != "default":
                theme_path = self._get_theme_path(theme)
                if theme_path:
                    cmd.extend(["--theme", str(theme_path)])

            # Add output options
            cmd.extend([
                temp_md_path,
                "--html",
                "--output", str(output_path)
            ])

            # Add additional options
            if options:
                for key, value in options.items():
                    if isinstance(value, bool) and value:
                        cmd.append(f"--{key}")
                    elif not isinstance(value, bool):
                        cmd.extend([f"--{key}", str(value)])

            # Execute command
            logger.debug(f"Running command: {' '.join(cmd)}")
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=60,
                shell=True if os.name == 'nt' else False
            )

            if result.returncode == 0:
                logger.info(f"HTML export successful: {output_path}")
                return True
            else:
                logger.error(f"HTML export failed: {result.stderr}")
                return False

        except subprocess.TimeoutExpired:
            logger.error("HTML export timed out")
            return False
        except Exception as e:
            logger.error(f"HTML export error: {str(e)}")
            return False
        finally:
            # Clean up temporary file
            try:
                os.unlink(temp_md_path)
            except OSError:
                pass

    def export_to_pptx(
        self,
        marp_markdown: str,
        output_path: Union[str, Path],
        method: str = "marp",
        options: Optional[Dict[str, Any]] = None
    ) -> bool:
        """
        Export Marp markdown to PPTX format.

        Args:
            marp_markdown: The Marp markdown content
            output_path: Path to save the PPTX file
            method: Export method ("marp" or "pandoc")
            options: Additional export options

        Returns:
            True if export successful, False otherwise
        """
        if method == "marp":
            return self._export_pptx_via_marp(marp_markdown, output_path, options)
        elif method == "pandoc":
            return self._export_pptx_via_pandoc(marp_markdown, output_path, options)
        else:
            logger.error(f"Unknown export method: {method}")
            return False

    def _export_pptx_via_marp(
        self,
        marp_markdown: str,
        output_path: Union[str, Path],
        options: Optional[Dict[str, Any]] = None
    ) -> bool:
        """Export to PPTX using marp-cli."""
        if not self.marp_cli_path:
            logger.error("marp-cli not available for PPTX export")
            return False

        output_path = Path(output_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)

        logger.info(f"Exporting to PPTX via Marp: {output_path}")

        # Create temporary markdown file
        with tempfile.NamedTemporaryFile(mode='w', suffix='.md', delete=False, encoding='utf-8') as temp_file:
            temp_file.write(marp_markdown)
            temp_md_path = temp_file.name

        try:
            # Build marp command for PPTX export
            cmd = [
                self.marp_cli_path,
                temp_md_path,
                "--pptx",
                "--output", str(output_path)
            ]

            # Add options
            if options:
                for key, value in options.items():
                    if isinstance(value, bool) and value:
                        cmd.append(f"--{key}")
                    elif not isinstance(value, bool):
                        cmd.extend([f"--{key}", str(value)])

            logger.debug(f"Running command: {' '.join(cmd)}")
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=120,  # PPTX export can take longer
                shell=True if os.name == 'nt' else False
            )

            if result.returncode == 0:
                logger.info(f"PPTX export via Marp successful: {output_path}")
                return True
            else:
                logger.error(f"PPTX export via Marp failed: {result.stderr}")
                return False

        except subprocess.TimeoutExpired:
            logger.error("PPTX export via Marp timed out")
            return False
        except Exception as e:
            logger.error(f"PPTX export via Marp error: {str(e)}")
            return False
        finally:
            # Clean up temporary file
            try:
                os.unlink(temp_md_path)
            except OSError:
                pass

    def _export_pptx_via_pandoc(
        self,
        marp_markdown: str,
        output_path: Union[str, Path],
        options: Optional[Dict[str, Any]] = None
    ) -> bool:
        """Export to PPTX using pandoc."""
        if not self.pandoc_path:
            logger.error("pandoc not available for PPTX export")
            return False

        output_path = Path(output_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)

        logger.info(f"Exporting to PPTX via Pandoc: {output_path}")

        # Create temporary markdown file
        with tempfile.NamedTemporaryFile(mode='w', suffix='.md', delete=False, encoding='utf-8') as temp_file:
            temp_file.write(marp_markdown)
            temp_md_path = temp_file.name

        try:
            # Build pandoc command
            cmd = [
                self.pandoc_path,
                temp_md_path,
                "-o", str(output_path),
                "--to", "pptx",
                "--from", "markdown+smart"
            ]

            # Add reference document if available
            template_path = self._get_pptx_template()
            if template_path:
                cmd.extend(["--reference-doc", str(template_path)])

            # Add options
            if options:
                for key, value in options.items():
                    if isinstance(value, bool) and value:
                        cmd.append(f"--{key}")
                    elif not isinstance(value, bool):
                        cmd.extend([f"--{key}", str(value)])

            logger.debug(f"Running command: {' '.join(cmd)}")
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=120,
                shell=True if os.name == 'nt' else False
            )

            if result.returncode == 0:
                logger.info(f"PPTX export via Pandoc successful: {output_path}")
                return True
            else:
                logger.error(f"PPTX export via Pandoc failed: {result.stderr}")
                return False

        except subprocess.TimeoutExpired:
            logger.error("PPTX export via Pandoc timed out")
            return False
        except Exception as e:
            logger.error(f"PPTX export via Pandoc error: {str(e)}")
            return False
        finally:
            # Clean up temporary file
            try:
                os.unlink(temp_md_path)
            except OSError:
                pass

    def preview_html(
        self,
        marp_markdown: str,
        theme: str = "default",
        options: Optional[Dict[str, Any]] = None
    ) -> Optional[str]:
        """
        Generate HTML preview content without saving to file.

        Args:
            marp_markdown: The Marp markdown content
            theme: Marp theme to use
            options: Additional export options

        Returns:
            HTML content as string, or None if failed
        """
        if not self.marp_cli_path:
            logger.error("marp-cli not available for HTML preview")
            return None

        logger.info("Generating HTML preview")

        # Create temporary files
        with tempfile.NamedTemporaryFile(mode='w', suffix='.md', delete=False, encoding='utf-8') as temp_md:
            temp_md.write(marp_markdown)
            temp_md_path = temp_md.name

        with tempfile.NamedTemporaryFile(suffix='.html', delete=False) as temp_html:
            temp_html_path = temp_html.name

        try:
            # Export to temporary HTML file
            success = self.export_to_html(
                marp_markdown=marp_markdown,
                output_path=temp_html_path,
                theme=theme,
                options=options
            )

            if success:
                # Read the HTML content
                with open(temp_html_path, 'r', encoding='utf-8') as f:
                    html_content = f.read()
                logger.info("HTML preview generated successfully")
                return html_content
            else:
                return None

        except Exception as e:
            logger.error(f"HTML preview generation error: {str(e)}")
            return None
        finally:
            # Clean up temporary files
            for temp_path in [temp_md_path, temp_html_path]:
                try:
                    os.unlink(temp_path)
                except OSError:
                    pass

    def _get_theme_path(self, theme_name: str) -> Optional[Path]:
        """Get the path to a Marp theme file."""
        theme_dir = Path(config.export.theme_dir)
        if not theme_dir.exists():
            return None

        # Look for theme files
        theme_extensions = ['.css', '.scss', '.sass']
        for ext in theme_extensions:
            theme_path = theme_dir / f"{theme_name}{ext}"
            if theme_path.exists():
                return theme_path

        logger.warning(f"Theme '{theme_name}' not found in {theme_dir}")
        return None

    def _get_pptx_template(self) -> Optional[Path]:
        """Get the path to a PPTX template file."""
        template_dir = Path(config.export.template_dir)
        if not template_dir.exists():
            return None

        # Look for template files
        template_patterns = ['template.pptx', 'default.pptx', '*.pptx']
        for pattern in template_patterns:
            if '*' in pattern:
                matches = list(template_dir.glob(pattern))
                if matches:
                    return matches[0]
            else:
                template_path = template_dir / pattern
                if template_path.exists():
                    return template_path

        logger.debug(f"No PPTX template found in {template_dir}")
        return None

    def get_available_themes(self) -> list[str]:
        """Get list of available Marp themes."""
        theme_dir = Path(config.export.theme_dir)
        if not theme_dir.exists():
            return ["default"]

        themes = []
        theme_extensions = ['.css', '.scss', '.sass']

        for item in theme_dir.iterdir():
            if item.is_file():
                for ext in theme_extensions:
                    if item.name.endswith(ext):
                        theme_name = item.name[:-len(ext)]
                        themes.append(theme_name)
                        break

        return themes if themes else ["default"]

    def validate_export_capabilities(self) -> Dict[str, bool]:
        """Check which export capabilities are available."""
        return {
            "html_export": self.marp_cli_path is not None,
            "pptx_via_marp": self.marp_cli_path is not None,
            "pptx_via_pandoc": self.pandoc_path is not None,
            "html_preview": self.marp_cli_path is not None,
        }