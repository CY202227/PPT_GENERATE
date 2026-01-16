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

"""MCP server implementation for PPT generation tools."""

import json
from typing import Any, List, Optional
from pathlib import Path

# FastMCP imports
from fastmcp import FastMCP

from src.utils.config import config
from src.utils.logging import Logger
from src.core.generator import GeneratorService
from src.core.exporter import ExportService


logger = Logger.get_logger(__name__)


class PPTGenerateServer:
    """MCP Server for PPT generation tools."""

    def __init__(self) -> None:
        """Initialize the MCP server."""
        self.app = FastMCP("ppt-generate")
        self.generator = GeneratorService()
        self.exporter = ExportService()

        self._setup_tools()
        logger.info("PPT Generate MCP Server initialized")

    def _setup_tools(self) -> None:
        """Setup MCP tools."""
        # Tool: Complete PPT Generation Expert
        @self.app.tool()
        async def create_ppt(
            topic: str,
            num_slides: int = 10,
            style: str = "formal",
            target_audience: str = "general",
            research_report: Optional[str] = None,
            template_pptx_path: Optional[str] = None,
            export_format: str = "pptx"
        ) -> str:
            """
            Complete PPT generation expert. Handles everything from outline to final export.

            Args:
                topic: The presentation topic (e.g., "Future of AI in Medicine")
                num_slides: Desired number of slides (default 10)
                style: Built-in style preset ("formal" for corporate, "cyberpunk" for tech, etc.)
                target_audience: The audience the PPT is tailored for
                research_report: Optional long-form research report to base the PPT on (strongly recommended)
                template_pptx_path: Optional path to an existing PPTX file to extract styles/colors/fonts from
                export_format: Final output format ("pptx" or "html")

            Returns:
                JSON string with the result, including Marp document and exported file path.
            """
            logger.info(f"MCP create_ppt called for topic: {topic}")

            try:
                custom_css = None
                # 1. Extract template if provided
                if template_pptx_path and Path(template_pptx_path).exists():
                    logger.info(f"Extracting template from: {template_pptx_path}")
                    custom_css = self.generator.convert_pptx_template(template_pptx_path)

                # 2. Generate complete presentation
                presentation = await self.generator.generate_presentation(
                    topic=topic,
                    style=style,
                    num_slides=num_slides,
                    target_audience=target_audience,
                    research_report=research_report,
                    custom_css=custom_css,
                    generate_images=False # Disabled background image generation
                )

                # 3. Export to requested format
                output_dir = Path("outputs")
                output_dir.mkdir(exist_ok=True)
                
                safe_topic = "".join(x for x in topic if x.isalnum() or x in " -_").strip()[:30]
                filename = f"{safe_topic}_{export_format}.{export_format}"
                output_path = output_dir / filename

                success = False
                if export_format.lower() == "pptx":
                    success = self.exporter.export_to_pptx(
                        marp_markdown=presentation.marp_document,
                        output_path=str(output_path)
                    )
                else:
                    success = self.exporter.export_to_html(
                        marp_markdown=presentation.marp_document,
                        output_path=str(output_path),
                        theme="default"
                    )

                # 4. Prepare result
                slides_data = []
                for slide in presentation.slides:
                    slides_data.append({
                        "slide_number": slide.slide_number,
                        "title": slide.title,
                        "content": slide.content,
                        "notes": slide.notes
                    })

                return json.dumps({
                    "success": True,
                    "topic": presentation.title,
                    "exported_file": str(output_path) if success else None,
                    "export_format": export_format,
                    "marp_document": presentation.marp_document,
                    "outline": {
                        "title": presentation.outline.title,
                        "estimated_slides": presentation.outline.estimated_slides,
                        "key_points": presentation.outline.key_points
                    },
                    "slides": slides_data,
                    "metadata": presentation.metadata
                }, ensure_ascii=False, indent=2)

            except Exception as e:
                logger.error(f"MCP create_ppt failed: {str(e)}")
                return json.dumps({
                    "success": False,
                    "error": str(e),
                    "topic": topic
                }, ensure_ascii=False)

    def run(self) -> None:
        """Run the MCP server."""
        logger.info(f"Starting PPT Generate MCP Server on {config.mcp.host}:{config.mcp.port}")

        # Run the FastMCP server
        self.app.run(
            host=config.mcp.host,
            port=config.mcp.port
        )


def main() -> None:
    """Main entry point for the MCP server."""
    server = PPTGenerateServer()
    server.run()


if __name__ == "__main__":
    main()