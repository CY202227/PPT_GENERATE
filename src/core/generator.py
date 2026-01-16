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

"""PPT generation service with LLM integration."""

import asyncio
import json
import os
from typing import Dict, Any, List, Optional, Union
from dataclasses import dataclass
from openai import AsyncOpenAI

from src.utils.config import config
from src.utils.logging import Logger
from src.core.template_converter import PPTXTemplateConverter


logger = Logger.get_logger(__name__)


@dataclass
class PPTOutline:
    """PPT outline structure."""
    title: str
    sections: List[Dict[str, Any]]
    estimated_slides: int
    target_audience: str
    key_points: List[str]


@dataclass
class SlideContent:
    """Individual slide content structure."""
    slide_number: int
    title: str
    content: str
    marp_markdown: str
    background_image: Optional[str] = None
    notes: Optional[str] = None
    layout_type: str = "content"
    needs_image: bool = False
    image_prompt: Optional[str] = None


@dataclass
class PPTContent:
    """Complete PPT content structure."""
    title: str
    outline: PPTOutline
    slides: List[SlideContent]
    marp_document: str
    metadata: Dict[str, Any]


class GeneratorService:
    """Service for generating PPT content using LLM."""

    def __init__(self) -> None:
        """Initialize the generator service."""
        self.template_converter = PPTXTemplateConverter()

        # Initialize LLM client
        self.llm_client = None
        self._setup_llm_client()

    def _setup_llm_client(self) -> None:
        """Setup LLM client based on configuration."""
        api_key = config.generator.api_key
        base_url = config.generator.base_url
        
        if not api_key:
            logger.warning("LLM API key not found in environment variables (checked LLM_API_KEY, OPENAI_API_KEY). PPT generation will fail.")
            self.llm_client = None
            return

        self.llm_client = AsyncOpenAI(
            api_key=api_key,
            base_url=base_url
        )
        logger.info(f"LLM client setup with model: {config.generator.model_name}, Base URL: {base_url or 'OpenAI Default'}")

    async def generate_presentation(
        self,
        topic: str,
        style: str = "modern",
        num_slides: int = 10,
        target_audience: str = "general",
        research_data: Optional[List[Dict[str, Any]]] = None,
        research_report: Optional[str] = None,
        template_path: Optional[str] = None,
        custom_css: Optional[str] = None,
        generate_images: bool = False
    ) -> PPTContent:
        """
        Generate a complete presentation.

        Args:
            topic: The presentation topic
            style: Presentation style (modern, corporate, academic, etc.)
            num_slides: Target number of slides
            target_audience: Target audience for the presentation
            research_data: Additional research data from web/file sources
            research_report: A pre-synthesized long-form research report (Markdown)
            template_path: Path to PPT template file
            custom_css: Custom CSS for Marp styling
            generate_images: Whether to generate background images

        Returns:
            Complete PPT content structure
        """
        logger.info(f"Starting presentation generation: {topic} ({num_slides} slides, {style} style)")

        # Step 1: Generate outline
        outline = await self._generate_outline(
            topic=topic,
            style=style,
            num_slides=num_slides,
            target_audience=target_audience,
            research_data=research_data,
            research_report=research_report
        )

        # Step 2: Generate individual slides
        slides = await self._generate_slides(
            outline=outline,
            research_data=research_data,
            research_report=research_report,
            max_slides_per_batch=config.generator.max_slides_per_request
        )

        # Step 3: Compile Marp document
        marp_document = self._compile_marp_document(
            title=topic,
            slides=slides,
            style=style,
            template_path=template_path,
            custom_css=custom_css
        )

        # Step 4: Create metadata
        metadata = {
            "topic": topic,
            "style": style,
            "target_audience": target_audience,
            "total_slides": len(slides),
            "estimated_duration": self._estimate_duration(len(slides)),
            "generated_at": "2024-01-01T00:00:00Z",  # Placeholder
            "generator_version": "1.0.0"
        }

        ppt_content = PPTContent(
            title=topic,
            outline=outline,
            slides=slides,
            marp_document=marp_document,
            metadata=metadata
        )

        logger.info(f"Presentation generation completed: {len(slides)} slides generated")
        return ppt_content

    def convert_pptx_template(self, template_path: str) -> str:
        """
        Convert PPTX template to Marp CSS.

        Args:
            template_path: Path to the PPTX template file

        Returns:
            CSS string for Marp
        """
        try:
            logger.info(f"Converting PPTX template: {template_path}")
            css = self.template_converter.convert_pptx_to_marp_css(template_path)
            logger.info("Template conversion completed")
            return css
        except Exception as e:
            logger.error(f"Template conversion failed: {str(e)}")
            raise

    async def _generate_outline(
        self,
        topic: str,
        style: str,
        num_slides: int,
        target_audience: str,
        research_data: Optional[List[Dict[str, Any]]] = None,
        research_report: Optional[str] = None
    ) -> PPTOutline:
        """Generate presentation outline using LLM."""
        logger.info("Generating presentation outline")

        # Prepare context from research data or report
        context = self._prepare_research_context(research_data, research_report)

        # Create prompt for outline generation
        prompt = self._build_outline_prompt(
            topic=topic,
            style=style,
            num_slides=num_slides,
            target_audience=target_audience,
            context=context,
            is_report_mode=bool(research_report)
        )

        # Call LLM to generate outline
        outline_data = await self._call_llm_for_outline(prompt)

        # Parse and validate outline
        outline = self._parse_outline_response(outline_data, topic, target_audience)

        logger.info(f"Outline generated: {len(outline.sections)} sections, {outline.estimated_slides} slides")
        return outline

    async def _generate_slides(
        self,
        outline: PPTOutline,
        research_data: Optional[List[Dict[str, Any]]] = None,
        research_report: Optional[str] = None,
        max_slides_per_batch: int = 5
    ) -> List[SlideContent]:
        """Generate content for individual slides."""
        logger.info(f"Generating slide content for {outline.estimated_slides} slides")

        slides = []
        slide_number = 1
        
        # Batch by total slides requested, not just sections
        current_batch_sections = []
        current_batch_slides_count = 0
        
        for section in outline.sections:
            section_slides = section.get('slides', 1)
            
            # If adding this section exceeds max slides per batch, process current batch first
            if current_batch_sections and current_batch_slides_count + section_slides > max_slides_per_batch:
                batch_slides = await self._generate_slide_batch(
                    batch_sections=current_batch_sections,
                    outline=outline,
                    research_data=research_data,
                    research_report=research_report,
                    start_slide_number=slide_number
                )
                slides.extend(batch_slides)
                slide_number += len(batch_slides)
                
                # Reset batch
                current_batch_sections = []
                current_batch_slides_count = 0
                await asyncio.sleep(0.5)
            
            current_batch_sections.append(section)
            current_batch_slides_count += section_slides
            
        # Process remaining sections
        if current_batch_sections:
            batch_slides = await self._generate_slide_batch(
                batch_sections=current_batch_sections,
                outline=outline,
                research_data=research_data,
                research_report=research_report,
                start_slide_number=slide_number
            )
            slides.extend(batch_slides)
            slide_number += len(batch_slides)

        logger.info(f"Generated {len(slides)} slides in {len(outline.sections)} sections")
        return slides

    async def _generate_slide_batch(
        self,
        batch_sections: List[Dict[str, Any]],
        outline: PPTOutline,
        research_data: Optional[List[Dict[str, Any]]] = None,
        research_report: Optional[str] = None,
        start_slide_number: int = 1
    ) -> List[SlideContent]:
        """Generate a batch of slides for given sections."""
        context = self._prepare_research_context(research_data, research_report)

        prompt = self._build_slides_prompt(
            sections=batch_sections,
            outline=outline,
            context=context,
            start_slide_number=start_slide_number,
            is_report_mode=bool(research_report)
        )

        # Call LLM to generate slides
        slides_data = await self._call_llm_for_slides(prompt)

        # Parse slide responses
        slides = self._parse_slides_response(slides_data, start_slide_number)

        return slides

    def _compile_marp_document(
        self,
        title: str,
        slides: List[SlideContent],
        style: str,
        template_path: Optional[str] = None,
        custom_css: Optional[str] = None
    ) -> str:
        """Compile slides into a complete Marp markdown document."""
        logger.info("Compiling Marp document")

        # Marp document header
        marp_header = f"""---
marp: true
title: {title}
theme: {style}
paginate: true
header: '{title}'
footer: 'Generated by PPT Generate'
"""

        # Add custom CSS if provided
        if custom_css:
            marp_header += f"\n---\n\n<style>\n{custom_css}\n</style>\n\n---\n"
        else:
            marp_header += "\n---\n"

        # Add content slides
        document = marp_header
        for i, slide in enumerate(slides):
            document += slide.marp_markdown
            if i < len(slides) - 1:
                document += "\n\n---\n\n"

        logger.info(f"Marp document compiled: {len(slides)} content slides")
        return document

    async def synthesize_research_report(self, topic: str, research_data: List[Dict[str, Any]]) -> str:
        """
        Synthesize raw research data into a comprehensive Markdown report.

        Args:
            topic: The research topic
            research_data: List of raw crawl/search results

        Returns:
            A long-form structured Markdown report
        """
        if not research_data:
            logger.warning("No research data to synthesize")
            return f"# {topic} 研究报告\n\n未找到相关研究内容。"

        logger.info(f"Synthesizing research report for: {topic}")

        # Prepare context from all available data
        raw_context = ""
        for i, item in enumerate(research_data):
            content = item.get("content", item.get("summary", ""))
            raw_context += f"### Source {i+1}: {item.get('title')}\nURL: {item.get('url')}\nContent: {content[:1500]}\n\n"

        prompt = f"""You are a professional research analyst. Your task is to synthesize the following raw research data into a comprehensive, structured, and insightful research report about "{topic}".

RAW DATA:
{raw_context}

REPORT REQUIREMENTS:
1. Format: Professional Markdown
2. Language: Match the language of the topic (Chinese for Chinese topics)
3. Structure:
   - Executive Summary
   - Background & Context
   - Key Findings (with data/evidence if available)
   - Technical/Thematic Analysis
   - Case Studies or Examples
   - Future Trends & Recommendations
   - References (list the URLs provided in raw data)
4. Style: Analytical, objective, and detailed.
5. Length: Aim for 1500-2500 characters.

The report should serve as the definitive "source of truth" for creating a high-quality presentation later."""

        try:
            response = await self.llm_client.chat.completions.create(
                model=config.generator.model_name,
                messages=[
                    {"role": "system", "content": "You are a professional research analyst expert in synthesizing complex information."},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.5, # Lower temperature for factual synthesis
                max_tokens=3000
            )
            
            report = response.choices[0].message.content
            logger.info(f"Research report synthesized successfully ({len(report)} characters)")
            return report
        except Exception as e:
            logger.error(f"Failed to synthesize research report: {str(e)}")
            return f"# {topic} 研究报告\n\n合成报告时出错：{str(e)}"

    def _prepare_research_context(
        self, 
        research_data: Optional[List[Dict[str, Any]]] = None,
        research_report: Optional[str] = None
    ) -> str:
        """
        Prepare research context for LLM. 
        Prioritizes the structured research_report if available.
        """
        if research_report:
            logger.info("Using pre-synthesized research report as context")
            return f"DETAILED RESEARCH REPORT (Source of Truth):\n\n{research_report}"

        if not research_data:
            return ""

        context_parts = []
        for item in research_data[:5]:  # Limit to avoid context overflow
            title = item.get('title', 'Unknown Source')

            # Prefer full content from deep research, fallback to summary
            if "content" in item and item["content"]:
                # Deep research provides full markdown content
                content = item["content"][:2000]
                context_parts.append(f"Source: {title} ({item.get('url', 'N/A')})\n{content}")
            elif "full_text" in item:
                # Uploaded file content
                content = item["full_text"][:1500]
                context_parts.append(f"Document: {item.get('file_name', title)}\n{content}")
            elif "summary" in item:
                # Search result summary
                content = item["summary"][:500]
                context_parts.append(f"Summary: {title}\n{content}")

        context = "\n\n".join(context_parts)
        logger.info(f"Prepared research context from {len(context_parts)} raw sources, total length: {len(context)} characters")
        return context

    def _build_outline_prompt(
        self,
        topic: str,
        style: str,
        num_slides: int,
        target_audience: str,
        context: str,
        is_report_mode: bool = False
    ) -> str:
        """Build prompt for outline generation."""
        mode_instruction = ""
        if is_report_mode:
            mode_instruction = "IMPORTANT: You are in REPORT MODE. A detailed research report has been provided. Your outline MUST strictly follow the structure and key findings of the report."
        else:
            mode_instruction = "You are in DATA MODE. Synthesize the most relevant information from provided research fragments to create a cohesive outline."

        prompt = f"""Create a detailed outline for a {style} style presentation on the topic: "{topic}"

Target audience: {target_audience}
Target number of slides: {num_slides}

{mode_instruction}

{f'Research Context:\n{context}\n' if context else ''}

Please provide a JSON response with the following structure:
{{
  "title": "Presentation Title",
  "sections": [
    {{
      "title": "Title Slide",
      "slides": 1,
      "key_points": ["Presentation Title", "Subtitle/Presenter"],
      "content_focus": "A professional opening slide",
      "needs_image": false,
      "layout_type": "cover",
      "image_prompt": null
    }},
    {{
      "title": "Section Title",
      "slides": 3,
      "key_points": ["Point 1", "Point 2", "Point 3"],
      "content_focus": "Brief description of what this section covers",
      "needs_image": true,
      "layout_type": "content",
      "image_prompt": "Description of what the image should show"
    }}
  ],
  "estimated_total_slides": {num_slides},
  "presentation_goal": "Brief statement of the presentation's main objective"
}}

Ensure the outline is logical, flows well, covers the topic comprehensively."""

        return prompt

    def _build_slides_prompt(
        self,
        sections: List[Dict[str, Any]],
        outline: PPTOutline,
        context: str,
        start_slide_number: int,
        is_report_mode: bool = False
    ) -> str:
        """Build prompt for slide content generation."""
        mode_instruction = ""
        if is_report_mode:
            mode_instruction = "IMPORTANT: You are in REPORT MODE. A detailed research report has been provided. The slides you generate MUST strictly reflect the content, data, and insights from that report."
        else:
            mode_instruction = "You are in DATA MODE. Use the provided research fragments to generate high-quality, informative slides."

        sections_text = "\n".join([
            f"- {section['title']}: {section.get('content_focus', '')} (Target: {section.get('slides', 1)} slides, Layout: {section.get('layout_type', 'content')})"
            for section in sections
        ])

        prompt = f"""Generate detailed content for the following presentation sections.
Presentation topic: {outline.title}
Style: Modern professional
Target audience: {outline.target_audience}

{mode_instruction}

Sections to generate:
{sections_text}

{f'Research Context:\n{context}\n' if context else ''}

CRITICAL REQUIREMENTS:
1. You MUST create the exact number of slides requested for each section.
2. Use the specified layout_type to format the content appropriately:
   - "cover": Centered title and minimal text
   - "content": Standard bullet point slides
   - "split": Side-by-side content
   - "quote": Centered quote layout
3. Use proper Marp markdown formatting

Return a JSON object with a "slides" key containing an array of slide objects:
{{
  "slides": [
    {{
      "slide_number": {start_slide_number},
      "title": "Slide Title",
      "content": "Main content summary",
      "marp_markdown": "<!-- _class: content -->\\n# Slide Title\\n\\n- Point 1\\n- Point 2",
      "section": "Section Name",
      "key_takeaway": "Main point",
      "layout_type": "content",
      "needs_image": false
    }}
  ]
}}
"""

        return prompt

    async def _call_llm_for_outline(self, prompt: str) -> Dict[str, Any]:
        """Call LLM to generate presentation outline."""
        if not self.llm_client:
            logger.error("LLM client not initialized - API key is likely missing")
            raise ValueError("LLM client not initialized.")

        logger.info(f"Calling LLM ({config.generator.model_name}) for outline generation")
        
        try:
            response = await self.llm_client.chat.completions.create(
                model=config.generator.model_name,
                messages=[
                    {"role": "system", "content": "You are a professional presentation expert. You output only valid JSON."},
                    {"role": "user", "content": prompt}
                ],
                response_format={"type": "json_object"},
                temperature=config.generator.temperature,
                max_tokens=config.generator.max_tokens
            )
            
            content = response.choices[0].message.content
            return json.loads(content)
        except Exception as e:
            logger.error(f"Error calling LLM for outline: {str(e)}")
            raise

    async def _call_llm_for_slides(self, prompt: str) -> List[Dict[str, Any]]:
        """Call LLM to generate slide content."""
        if not self.llm_client:
            logger.error("LLM client not initialized - API key is likely missing")
            raise ValueError("LLM client not initialized.")

        logger.info(f"Calling LLM ({config.generator.model_name}) for slides generation")
        
        try:
            response = await self.llm_client.chat.completions.create(
                model=config.generator.model_name,
                messages=[
                    {"role": "system", "content": "You are a professional presentation expert. You output only valid JSON."},
                    {"role": "user", "content": prompt}
                ],
                response_format={"type": "json_object"},
                temperature=config.generator.temperature,
                max_tokens=config.generator.max_tokens
            )
            
            content = response.choices[0].message.content
            data = json.loads(content)
            
            if isinstance(data, dict) and "slides" in data:
                return data["slides"]
            elif isinstance(data, list):
                return data
            elif isinstance(data, dict):
                for key in ["data", "content"]:
                    if key in data and isinstance(data[key], list):
                        return data[key]
                return [data]
            
            return []
        except Exception as e:
            logger.error(f"Error calling LLM for slides: {str(e)}")
            raise

    def _parse_outline_response(self, response: Dict[str, Any], topic: str, target_audience: str) -> PPTOutline:
        """Parse LLM response into PPTOutline structure."""
        sections = response.get("sections", [])
        estimated_slides = response.get("estimated_total_slides", len(sections) * 2)

        key_points = []
        for section in sections:
            key_points.extend(section.get("key_points", []))

        return PPTOutline(
            title=response.get("title", topic),
            sections=sections,
            estimated_slides=estimated_slides,
            target_audience=target_audience,
            key_points=key_points
        )

    def _parse_slides_response(self, response: List[Dict[str, Any]], start_slide_number: int) -> List[SlideContent]:
        """Parse LLM response into SlideContent structures."""
        slides = []
        for i, slide_data in enumerate(response):
            slide = SlideContent(
                slide_number=start_slide_number + i,
                title=slide_data.get("title", f"Slide {start_slide_number + i}"),
                content=slide_data.get("content", ""),
                marp_markdown=slide_data.get("marp_markdown", ""),
                notes=slide_data.get("key_takeaway")
            )

            # Add additional metadata
            slide.layout_type = slide_data.get("layout_type", "content")
            slide.needs_image = slide_data.get("needs_image", False)
            slide.image_prompt = slide_data.get("image_prompt")

            slides.append(slide)

        return slides

    def _estimate_duration(self, num_slides: int) -> str:
        """Estimate presentation duration based on slide count."""
        min_minutes = num_slides * 1
        max_minutes = num_slides * 2
        return f"{min_minutes}-{max_minutes} minutes"
