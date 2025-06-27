import os
import sys
import xml.etree.ElementTree as ET
from xml.dom import minidom
import base64
import mimetypes
from typing import List, Dict, Optional
import json
from openai import OpenAI

# --- 1. 新增: 智能分析代理 (LLM交互) ---
class SmartArtAgent:
    """
    一个与大语言模型（LLM）交互的代理，负责将非结构化的自然语言文本，
    转换为结构化的图表数据。
    """
    def __init__(self, api_key: str, base_url: str):
        try:
            self.client = OpenAI(api_key=api_key, base_url=base_url)
        except Exception as e:
            raise RuntimeError(f"初始化OpenAI客户端时出错: {e}")

    def analyze_text(self, input_paragraph: str) -> List[Dict]:
        """
        分析文本，调用LLM，并返回一个包含所有适用图表数据的JSON数组。
        """
        prompt = f"""
你是一位顶级的业务分析师。你的任务是分析给定的文本，并将其从多个维度进行结构化，以便进行可视化。

**你有以下几种数据"schema"可以选择:**
1.  `Flowchart`: 用于表示有先后顺序的步骤或流程。数据是一个字符串数组。
    - 格式: {{"schema": "Flowchart", "data": ["步骤一", "步骤二", "步骤三"]}}
2.  `KeyPoints`: 用于总结没有强顺序关系的要点和详情。数据是一个对象数组，每个对象包含key和detail。
    - 格式: {{"schema": "KeyPoints", "data": [{{"key": "要点1", "detail": "要点1的详情"}}]}}

**任务要求:**
1.  **分析输入文本**:
    ```text
    {input_paragraph}
    ```
2.  **找出所有适用的Schema**: 如果文本同时包含流程和要点，你应该同时提供 `Flowchart` 和 `KeyPoints` 两种schema的分析结果。
3.  **提取数据**: 为每一个你选择的schema，从文本中提取并精炼信息。
4.  **严格的JSON数组输出**: 你的回答必须是一个单一的JSON数组，其中包含你分析出的所有图表对象。不要添加任何解释或额外的文本。
"""
        messages = [
            {"role": "system", "content": "You are a business analyst that structures information into a specific JSON format based on predefined schemas. You should return an array of all applicable schemas."},
            {"role": "user", "content": prompt}
        ]

        print("正在向模型请求，进行多维度智能分析和数据提取...")
        try:
            response = self.client.chat.completions.create(
                model="Qwen-72B", 
                messages=messages,
                temperature=0.0,
                response_format={"type": "json_object"} # 确保返回的是JSON对象
            )
            content = response.choices[0].message.content
            # LLM可能返回一个包含数组的JSON对象，例如 {"charts": [...]}, 我们需要从中提取出数组
            parsed_json = json.loads(content)
            
            # 兼容几种可能的返回格式
            if isinstance(parsed_json, list):
                return parsed_json
            elif isinstance(parsed_json, dict):
                # 寻找字典中第一个是列表的值
                for key, value in parsed_json.items():
                    if isinstance(value, list):
                        return value
            
            raise ValueError("LLM返回的JSON中未找到预期的数组格式")

        except json.JSONDecodeError:
            raise ValueError("解析LLM返回的JSON时失败。")
        except Exception as e:
            raise RuntimeError(f"与LLM API交互时出错: {e}")


# --- 2. 核心 SmartArt 处理逻辑 (渲染引擎) ---
class SmartArtProcessor:
    """
    一个高内聚的类，独立负责处理SmartArt生成的所有逻辑。
    仿照 Napkin AI 的后台服务思路，这个类通过更抽象的参数（如 chart_type, style_options）
    而非具体的文件路径来驱动，使其更像一个可配置的图形生成服务。
    """
    def __init__(self, chart_type: str = 'flowchart_vertical', style_options: Optional[Dict] = None, output_dir: str = 'tmp', filename_base: str = 'smart_art_output'):
        self.output_dir = output_dir
        self.filename_base = filename_base
        self.chart_type = chart_type
        self.style_options = style_options if style_options is not None else {}
        self.template_root = self._load_template()

    def _load_template(self) -> ET.Element:
        """根据 chart_type 加载内部管理的XML模板。"""
        template_path = self._get_template_path_from_type()
        try:
            tree = ET.parse(template_path)
            return tree.getroot()
        except (ET.ParseError, FileNotFoundError) as e:
            raise RuntimeError(f"加载或解析XML模板 '{template_path}' 失败: {e}")

    def _get_template_path_from_type(self) -> str:
        """根据图表类型返回模板路径。这里是未来扩展不同图表类型的地方。"""
        if self.chart_type == 'flowchart_vertical':
            return './templates/flowchart_template.xml'
        # 未来可以轻松扩展，例如:
        # elif self.chart_type == 'cycle_diagram':
        #     return './templates/cycle_template.xml'
        else:
            raise ValueError(f"不支持的图表类型: '{self.chart_type}'")

    def _calculate_text_layout(self, text: str, font_size: int, max_width: int) -> dict:
        """
        根据可用宽度和字体大小，智能计算文本布局（换行和行数）。
        能更好地处理中英文混合文本。
        """
        if not text or max_width <= 0 or font_size <= 0:
            return {'wrapped_lines': [text] if text else [], 'line_count': 1}

        lines = []
        current_line = ""
        current_width = 0

        def get_char_width(char: str) -> float:
            # 简单启发式：中日韩等全角字符宽度约等于字体大小，其他（如拉丁字母）约等于一半。
            return font_size if '\u4e00' <= char <= '\u9fff' else font_size * 0.6

        for char in text:
            char_w = get_char_width(char)
            if current_width + char_w > max_width and current_line:
                lines.append(current_line)
                current_line = char
                current_width = char_w
            else:
                current_line += char
                current_width += char_w
        
        if current_line:
            lines.append(current_line)

        return {'wrapped_lines': lines or ([text] if text else []), 'line_count': len(lines) or 1}

    def _fill_template_and_get_xml(self, texts: list, image_mapping: Optional[dict]) -> str:
        if image_mapping is None: image_mapping = {}
        
        # 使用从构造函数加载的模板副本进行操作，避免修改原始模板
        root = ET.fromstring(ET.tostring(self.template_root))

        # 在填充前，预先从模板中解析出布局的关键参数
        def_root = root.find("definition")
        if def_root is None: raise ValueError("XML模板缺少 <definition> 块。")
        layout_style = def_root.find("layout")
        if layout_style is None: raise ValueError("模板 <definition> 块中缺少 <layout> 定义。")
        img_txt_bg_style = def_root.find("image_text_background")
        if img_txt_bg_style is None: raise ValueError("模板 <definition> 块中缺少 <image_text_background> 定义。")

        BOX_WIDTH = int(self.style_options.get("box_width", layout_style.get("box_width", "300")))
        IMG_BG_X_PADDING = int(self.style_options.get("image_bg_x_padding", img_txt_bg_style.get("x_padding", "5")))
        text_area_width = BOX_WIDTH - (IMG_BG_X_PADDING * 4) # 文本区域的有效宽度
        base_box_height = 60

        nodes_in_template = root.findall("node")
        
        for i, text_item in enumerate(texts):
            if i >= len(nodes_in_template): break
            node = nodes_in_template[i]

            base_font_size, min_font_size = 18, 10
            font_size = max(min_font_size, base_font_size - len(text_item) // 6)

            # 使用统一的布局计算函数
            layout_info = self._calculate_text_layout(text_item, font_size, text_area_width)
            num_lines = layout_info['line_count']
            box_height = max(base_box_height, num_lines * font_size * 1.8)

            if (text_elem := node.find("text")) is not None: text_elem.text = text_item
            if (style_elem := node.find("style")) is not None:
                style_elem.set("font_size", str(font_size))
                style_elem.set("box_height", str(int(box_height)))
            if (image_elem := node.find("image")) is not None:
                image_path = image_mapping.get(f"Image{i+1}") or image_mapping.get(text_item)
                image_elem.text = image_path if image_path else ""

        for i in range(len(texts), len(nodes_in_template)):
            root.remove(nodes_in_template[i])
            
        rough_string = ET.tostring(root, 'utf-8')
        reparsed = minidom.parseString(rough_string)
        return reparsed.toprettyxml(indent="  ")

    def _encode_image_to_base64(self, image_path: str) -> str | None:
        try:
            mime_type, _ = mimetypes.guess_type(image_path)
            if not mime_type or not mime_type.startswith('image'): return None
            with open(image_path, "rb") as f:
                return f"data:{mime_type};base64,{base64.b64encode(f.read()).decode('utf-8')}"
        except (FileNotFoundError, TypeError): return None
        
    def _render_flowchart_from_xml(self, root: ET.Element) -> str:
        def_root = root.find("definition")
        if def_root is None: raise ValueError("XML模板缺少 <definition> 块。")

        # --- 从XML模板获取默认样式 ---
        layout_style = def_root.find("layout")
        box_style = def_root.find("box_style")
        img_txt_bg_style = def_root.find("image_text_background")
        text_style = def_root.find("text_style")
        arrow_style = def_root.find("arrow_style")

        # --- 检查关键定义块是否存在 ---
        if layout_style is None: raise ValueError("模板 <definition> 块中缺少 <layout> 定义。")
        if box_style is None: raise ValueError("模板 <definition> 块中缺少 <box_style> 定义。")
        if img_txt_bg_style is None: raise ValueError("模板 <definition> 块中缺少 <image_text_background> 定义。")
        if text_style is None: raise ValueError("模板 <definition> 块中缺少 <text_style> 定义。")
        if arrow_style is None: raise ValueError("模板 <definition> 块中缺少 <arrow_style> 定义。")

        # --- 合并默认样式和用户自定义样式 (style_options) ---
        # 用户提供的 style_options 具有更高优先级
        s = self.style_options
        
        BOX_WIDTH = int(s.get("box_width", layout_style.get("box_width", "300")))
        V_GAP = int(s.get("vertical_gap", layout_style.get("vertical_gap", "40")))
        
        DEFAULT_BOX_HEIGHT = s.get("default_box_height", box_style.get("default_height", "80"))
        BOX_FILL = s.get("box_fill", box_style.get("fill", "#e3f2fd"))
        BOX_STROKE = s.get("box_stroke", box_style.get("stroke", "#90caf9"))
        BOX_STROKE_WIDTH = s.get("box_stroke_width", box_style.get("stroke_width", "2"))
        BOX_RX = s.get("box_corner_radius", box_style.get("corner_radius", "5"))

        IMG_BG_FILL = s.get("image_bg_fill", img_txt_bg_style.get("fill", "white"))
        IMG_BG_OPACITY = s.get("image_bg_opacity", img_txt_bg_style.get("opacity", "0.8"))
        IMG_BG_X_PADDING = int(s.get("image_bg_x_padding", img_txt_bg_style.get("x_padding", "5")))
        IMG_BG_RX = s.get("image_bg_corner_radius", img_txt_bg_style.get("corner_radius", "3"))
        
        DEFAULT_FONT_SIZE = s.get("default_font_size", text_style.get("default_size", "15"))
        TEXT_FONT_FAMILY = s.get("font_family", text_style.get("font_family", "sans-serif"))
        TEXT_FILL = s.get("text_fill", text_style.get("fill", "#1e88e5"))
        TEXT_FONT_WEIGHT = s.get("font_weight", text_style.get("font_weight", "bold"))

        ARROW_FILL = s.get("arrow_fill", arrow_style.get("fill", "#90caf9"))
        ARROW_STROKE = s.get("arrow_stroke", arrow_style.get("stroke", "#90caf9"))
        ARROW_STROKE_WIDTH = s.get("arrow_stroke_width", arrow_style.get("stroke_width", "2"))
        ARROW_HEAD_SIZE = int(s.get("arrow_head_size", arrow_style.get("head_size", "6")))
        ARROW_MARGIN_BOTTOM = int(s.get("arrow_margin_bottom", arrow_style.get("head_margin_bottom", "10")))

        nodes = root.findall("node")
        box_heights = [int(node.find("style").get("box_height", DEFAULT_BOX_HEIGHT)) for node in nodes]
        svg_height = sum(box_heights) + max(0, len(nodes) - 1) * V_GAP
        svg_defs, svg_elements, image_patterns, p_counter = [], [], {}, 0

        for i, node in enumerate(nodes):
            if (img_elem := node.find("image")) is not None and img_elem.text and img_elem.text not in image_patterns:
                pattern_id = f"bg-p-{p_counter}"
                image_patterns[img_elem.text] = pattern_id
                p_counter += 1
                if encoded_image := self._encode_image_to_base64(img_elem.text):
                    svg_defs.append(f'<pattern id="{pattern_id}" patternUnits="userSpaceOnUse" width="{BOX_WIDTH}" height="{box_heights[i]}"><image href="{encoded_image}" x="0" y="0" width="{BOX_WIDTH}" height="{box_heights[i]}" preserveAspectRatio="xMidYMid slice"/></pattern>')
        
        current_y = 0
        for i, node in enumerate(nodes):
            text, style = node.find("text").text or "", node.find("style")
            img_elem = node.find("image")
            box_h, font_s = box_heights[i], int(style.get("font_size", DEFAULT_FONT_SIZE))
            
            fill_attr = f"fill:url(#{image_patterns[img_elem.text]})" if img_elem is not None and img_elem.text and img_elem.text in image_patterns else f"fill:{BOX_FILL}"
            rect_style = f"{fill_attr};stroke:{BOX_STROKE};stroke-width:{BOX_STROKE_WIDTH};"
            svg_elements.append(f'<rect x="0" y="{current_y}" width="{BOX_WIDTH}" height="{box_h}" rx="{BOX_RX}" ry="{BOX_RX}" style="{rect_style}"/>')
            
            if img_elem is not None and img_elem.text:
                 text_bg_style = f"fill:{IMG_BG_FILL};fill-opacity:{IMG_BG_OPACITY};"
                 bg_width = BOX_WIDTH - (IMG_BG_X_PADDING * 2)
                 svg_elements.append(f'<rect x="{IMG_BG_X_PADDING}" y="{current_y + box_h/4}" width="{bg_width}" height="{box_h/2}" style="{text_bg_style}" rx="{IMG_BG_RX}"/>')
            
            # --- Text Wrapping Logic ---
            text_element_style = f'font-family="{TEXT_FONT_FAMILY}" font-size="{font_s}" fill="{TEXT_FILL}" font-weight="{TEXT_FONT_WEIGHT}"'
            line_height = font_s * 1.4

            # 再次使用统一的布局计算函数，确保渲染与计算时逻辑一致
            text_area_width = BOX_WIDTH - (IMG_BG_X_PADDING * 4)
            layout_info = self._calculate_text_layout(text, font_s, text_area_width)
            wrapped_lines = layout_info['wrapped_lines']
            num_lines = layout_info['line_count']

            total_text_height = (num_lines -1) * line_height + font_s
            block_top_y = current_y + (box_h - total_text_height) / 2

            text_container = f'<text {text_element_style} text-anchor="middle">'
            for j, line in enumerate(wrapped_lines):
                tspan_y = block_top_y + (j * line_height)
                text_container += f'<tspan x="{BOX_WIDTH/2}" y="{tspan_y}" dominant-baseline="hanging">{line}</tspan>'
            text_container += '</text>'
            svg_elements.append(text_container)
            
            if i < len(nodes) - 1:
                arrow_y1, arrow_y2, arrow_x = current_y + box_h, current_y + box_h + V_GAP, BOX_WIDTH/2
                line_style = f"stroke:{ARROW_STROKE};stroke-width:{ARROW_STROKE_WIDTH}"
                arrow_line_end_y = arrow_y2 - ARROW_MARGIN_BOTTOM
                svg_elements.append(f'<line x1="{arrow_x}" y1="{arrow_y1}" x2="{arrow_x}" y2="{arrow_line_end_y}" style="{line_style}"/>')
                h = ARROW_HEAD_SIZE
                svg_elements.append(f'<polygon points="{arrow_x-h},{arrow_line_end_y} {arrow_x+h},{arrow_line_end_y} {arrow_x},{arrow_y2}" style="fill:{ARROW_FILL};"/>')

            current_y += box_h + V_GAP

        defs_str = f'<defs>{"".join(svg_defs)}</defs>' if svg_defs else ""
        return f'<svg xmlns="http://www.w3.org/2000/svg" width="{BOX_WIDTH}" height="{int(svg_height)}">{defs_str}{"".join(svg_elements)}</svg>'

    def _render_xml_to_svg(self, xml_content: str) -> str:
        try:
            root = ET.fromstring(xml_content)
            # 未来可以根据 chart_type 或 XML 中的 type 属性分发到不同的渲染函数
            if self.chart_type == 'flowchart_vertical' and root.attrib.get("type") == "Flowchart":
                return self._render_flowchart_from_xml(root)
            raise ValueError(f"不支持的图表类型: '{self.chart_type}' 或 XML type 属性不匹配。")
        except ET.ParseError as e:
            raise ValueError(f"解析XML时出错: {e}") from e

    def process(self, texts: List[str], images: Optional[Dict[str, str]] = None) -> Dict[str, str]:
        """
        高级公共方法，编排整个图表生成流程。
        现在它不接受 template_path，而是使用在实例化时定义的 chart_type。
        """
        if images is None: images = {}
        if not os.path.exists(self.output_dir):
            os.makedirs(self.output_dir)
            
        try:
            xml_content = self._fill_template_and_get_xml(
                texts=texts,
                image_mapping=images
            )
            
            xml_file_path = os.path.join(self.output_dir, f"{self.filename_base}.xml")
            with open(xml_file_path, 'w', encoding='utf-8') as f:
                f.write(xml_content)
            
            svg_content = self._render_xml_to_svg(xml_content)
            
            svg_file_path = os.path.join(self.output_dir, f"{self.filename_base}.svg")
            with open(svg_file_path, 'w', encoding='utf-8') as f:
                f.write(svg_content)
            
            return {
                "status": "success",
                "xml_file_path": os.path.abspath(xml_file_path),
                "svg_file_path": os.path.abspath(svg_file_path)
            }
            
        except (RuntimeError, ValueError) as e:
            error_message = f"生成SmartArt图表时出错: {e}"
            print(error_message, file=sys.stderr)
            return {
                "status": "error",
                "message": error_message
            }

# --- 3. 简化的外部接口函数 (渲染服务) ---
def generate_smartart(
    texts: List[str], 
    chart_type: str = 'flowchart_vertical',
    images: Optional[Dict[str, str]] = None,
    style_options: Optional[Dict] = None,
    output_dir: str = 'tmp',
    filename_base: str = 'smart_art_output'
) -> Dict[str, str]:
    """
    根据给定的文本列表、图表类型和样式选项，生成一个SmartArt流程图。
    这是一个更高级别的外部接口，将实现细节委托给 SmartArtProcessor 类，
    使其更像一个服务调用。
    
    Args:
        texts (List[str]): 流程图步骤的文本列表。
        chart_type (str): 要生成的图表类型 (例如, 'flowchart_vertical')。
        images (Dict[str, str], optional): 图片映射字典。 Defaults to None.
        style_options (Dict, optional): 用于覆盖默认模板样式的字典。 Defaults to None.
        output_dir (str): 保存生成文件的目录。
        filename_base (str): 生成文件的基础名（不含扩展名）。
    
    Returns:
        Dict[str, str]: 包含状态和文件路径的结果字典。
    """
    processor = SmartArtProcessor(
        chart_type=chart_type,
        style_options=style_options,
        output_dir=output_dir,
        filename_base=filename_base
    )
    return processor.process(texts, images)

# --- 4. 新增: 端到端的主函数 ---
def create_visuals_from_text(
    input_text: str,
    api_key: str,
    base_url: str,
    output_dir: str = 'tmp',
    style_options: Optional[Dict] = None,
    images: Optional[Dict[str, str]] = None
) -> Dict[str, list]:
    """
    接收自然语言文本，智能分析并生成所有适用的图表。
    这是新的、最高级别的端到端入口点。
    """
    print("--- 开始智能可视化流程 ---")
    results = {"status": "success", "generated_files": [], "errors": []}
    
    # 1. 调用Agent进行智能分析
    try:
        agent = SmartArtAgent(api_key=api_key, base_url=base_url)
        structured_data_list = agent.analyze_text(input_text)
        print(f"智能分析完成，提取出 {len(structured_data_list)} 个可视化图表。")
    except (RuntimeError, ValueError) as e:
        error_msg = f"智能分析阶段失败: {e}"
        print(error_msg, file=sys.stderr)
        results["status"] = "error"
        results["errors"].append(error_msg)
        return results

    # 2. 遍历分析结果，调用渲染服务生成图表
    for i, chart_data in enumerate(structured_data_list):
        schema = chart_data.get("schema")
        data = chart_data.get("data")
        
        if not schema or not data:
            print(f"警告: 第 {i+1} 个分析结果缺少 schema 或 data，已跳过。")
            continue

        print(f"\n正在生成第 {i+1}/{len(structured_data_list)} 个图表，类型: {schema}...")
        
        # 根据schema分发到不同的渲染逻辑
        if schema == 'Flowchart':
            if isinstance(data, list):
                # 为每个图表生成唯一的文件名
                filename = f"flowchart_{i+1}"
                render_result = generate_smartart(
                    texts=data,
                    chart_type='flowchart_vertical',
                    style_options=style_options,
                    output_dir=output_dir,
                    filename_base=filename,
                    images=images
                )
                if render_result.get("status") == "success":
                    print(f"✅ 流程图生成成功: {render_result.get('svg_file_path')}")
                    results["generated_files"].append(render_result)
                else:
                    error_msg = f"流程图渲染失败: {render_result.get('message')}"
                    print(f"❌ {error_msg}", file=sys.stderr)
                    results["errors"].append(error_msg)
            else:
                results["errors"].append(f"Flowchart schema 的 data 格式应为列表，但得到 {type(data)}。")

        elif schema == 'KeyPoints':
            # 当前脚本主要处理流程图，KeyPoints可以作为未来扩展点
            print(f"ℹ️  识别到 KeyPoints schema，但当前版本主要渲染流程图，暂不生成文件。")
            # 这里可以添加生成KeyPoints卡片、列表等的逻辑
            pass
        
        else:
            print(f"警告: 未知的 schema 类型 '{schema}'，已跳过。")

    if results["errors"]:
        results["status"] = "partial_success" if results["generated_files"] else "error"

    print("\n--- 智能可视化流程结束 ---")
    return results


# --- 5. 主程序入口 / 测试 ---
if __name__ == '__main__':
    """
    此块演示从一段原始文本开始，端到端地生成SmartArt图表。
    """
    # --- 配置 ---
    # !!重要!! 请替换为您自己的API密钥和端点URL
    API_KEY = os.getenv("API_KEY")
    BASE_URL = os.getenv("BASE_URL")

    # 1. 定义原始输入文本
    input_paragraph = """
    我们项目的启动流程是这样的：首先，产品经理会根据市场调研和用户反馈来提出初步的需求。
    接着，开发团队、测试团队和产品经理会一起召开一个需求评审会议。
    关于评审，主要关注几个核心要点。第一是技术可行性，评估现有架构能否支持。第二是资源匹配，看我们的人力是否足够。
    评审通过后，UI设计师会先出设计稿，然后开发团队会根据设计稿和最终确定的需求文档，同步进行编码工作。
    """
    
    # 2. 定义图片映射 (可选)
    # 使用 "Image1", "Image2" 等键来为特定的框（第1个，第2个...）指定背景图片。
    example_images = {
        "Image1": r"C:\Users\CHENQIMING\Pictures\20230406163343262212222.png"
    }

    # 3. 定义自定义样式 (可选)
    custom_styles = {
        "box_fill": "#e3f2fd",       # 框体填充色 (淡蓝色)
        "box_stroke": "#90caf9",     # 框体边框色 (蓝色)
        "text_fill": "#1565c0",      # 文字颜色 (深蓝色)
        "arrow_fill": "#90caf9",     # 箭头填充色
        "arrow_stroke": "#90caf9",   # 箭头边框色
        "font_family": "Microsoft YaHei", # 字体
        "box_corner_radius": "8"     # 框体圆角
    }
    
    # 4. 调用端到端的主函数
    final_result = create_visuals_from_text(
        input_text=input_paragraph,
        api_key=API_KEY,
        base_url=BASE_URL,
        style_options=custom_styles,
        images=example_images
    )
    
    # 5. 打印最终结果
    print("\n--- 最终执行结果 ---")
    print(json.dumps(final_result, indent=4, ensure_ascii=False))

    # 6. 检查模板目录是否存在
    if not os.path.exists('./templates'):
        print("\n警告: 'templates' 目录不存在。", file=sys.stderr)
        print("请创建一个 'templates' 目录，并将 'flowchart_template.xml' 放入其中。", file=sys.stderr)