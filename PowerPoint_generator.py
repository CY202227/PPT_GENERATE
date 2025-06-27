import os
import json
import logging
import uuid
import shutil
import zipfile
import xml.etree.ElementTree as ET
from dotenv import load_dotenv
from openai import OpenAI
from typing import Dict, List, Any, Tuple
import traceback

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# 加载环境变量
load_dotenv()

class PPTGenerator:
    
    def __init__(self, api_key: str, base_url: str):
        try:
            self.client = OpenAI(api_key=api_key, base_url=base_url)
            self.template_description = None
            self.outline = None
            logger.info("OpenAI客户端初始化成功")
        except Exception as e:
            logger.error(f"初始化OpenAI客户端失败: {e}")
            raise
    
    def load_template_description(self, template_description_path: str) -> bool:
        """加载模板描述文件"""
        try:
            with open(template_description_path, 'r', encoding='utf-8') as f:
                self.template_description = json.load(f)
            logger.info("成功加载模板描述文件")
            return True
        except Exception as e:
            logger.error(f"加载模板描述文件失败: {e}")
            return False
    
    def load_outline(self, outline_path: str) -> bool:
        """加载大纲文件"""
        try:
            with open(outline_path, 'r', encoding='utf-8') as f:
                self.outline = json.load(f)
            logger.info("成功加载大纲文件")
            return True
        except Exception as e:
            logger.error(f"加载大纲文件失败: {e}")
            return False
    
    def should_skip_slide(self, slide_number: int, total_content_pages: int) -> bool:
        """判断是否应该跳过某个幻灯片的处理"""
        # 获取结束页的编号（总页数）
        total_slides = 3 + total_content_pages  # 首页 + 目录页 + 内容页数 + 结束页
        
        # 如果是结束页，跳过处理
        if slide_number == total_slides:
            return True
        
        return False
    
    def detect_slide_type(self, slide_xml_path: str) -> Tuple[bool, bool]:
        """
        检测幻灯片是否包含SmartArt或图表
        
        Args:
            slide_xml_path: 幻灯片XML文件的路径
            
        Returns:
            Tuple[bool, bool]: (has_smartart, has_chart)
        """
        try:
            tree = ET.parse(slide_xml_path)
            root = tree.getroot()
            
            # 定义命名空间
            namespaces = {
                'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
                'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
                'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
                'dgm': 'http://schemas.openxmlformats.org/drawingml/2006/diagram',
                'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
            }
            
            # 检查SmartArt
            has_smartart = False
            # 通过dgm:relIds标签
            if root.findall('.//dgm:relIds', namespaces):
                has_smartart = True
            # 通过graphicData的URI
            elif root.findall(".//a:graphicData[@uri='http://schemas.openxmlformats.org/drawingml/2006/diagram']", namespaces):
                has_smartart = True
            # 通过OLE对象
            elif root.findall(".//p:oleObj", namespaces):
                for oleObj in root.findall(".//p:oleObj", namespaces):
                    progId = oleObj.get('progId', '')
                    if 'SmartArt' in progId:
                        has_smartart = True
                        break
            
            # 检查图表
            has_chart = False
            # 直接的图表标签
            if root.findall('.//c:chart', namespaces):
                has_chart = True
            # 通过graphicData的URI
            elif root.findall(".//a:graphicData[@uri='http://schemas.openxmlformats.org/drawingml/2006/chart']", namespaces):
                has_chart = True
            # 通过OLE对象
            elif root.findall(".//p:oleObj", namespaces):
                for oleObj in root.findall(".//p:oleObj", namespaces):
                    progId = oleObj.get('progId', '')
                    if 'Chart' in progId:
                        has_chart = True
                        break
            
            # 如果发现图表，进一步确认是否有关联的图表文件
            if has_chart:
                # 检查关系文件
                slide_dir = os.path.dirname(slide_xml_path)
                rels_path = os.path.join(slide_dir, '_rels', os.path.basename(slide_xml_path) + '.rels')
                
                if os.path.exists(rels_path):
                    rels_tree = ET.parse(rels_path)
                    rels_root = rels_tree.getroot()
                    # 查找是否有指向图表的关系
                    chart_rels = rels_root.findall(".//*[@Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart']", namespaces)
                    has_chart = len(chart_rels) > 0
            
            logger.info(f"幻灯片类型检测结果 - SmartArt: {has_smartart}, Chart: {has_chart}")
            return has_smartart, has_chart
            
        except ET.ParseError as e:
            logger.error(f"解析XML文件失败: {e}")
            return False, False
        except Exception as e:
            logger.error(f"检测幻灯片类型时发生错误: {e}")
            return False, False
    
    def extract_pptx(self, pptx_path: str, extract_dir: str) -> bool:
        """解压PPTX文件"""
        try:
            with zipfile.ZipFile(pptx_path, 'r') as zip_ref:
                zip_ref.extractall(extract_dir)
            logger.info(f"成功解压PPTX到: {extract_dir}")
            return True
        except Exception as e:
            logger.error(f"解压PPTX失败: {e}")
            return False

    def create_pptx(self, source_dir: str, output_path: str) -> bool:
        """将文件夹重新打包为PPTX"""
        try:
            with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                for root, _, files in os.walk(source_dir):
                    for file in files:
                        file_path = os.path.join(root, file)
                        arcname = os.path.relpath(file_path, source_dir)
                        zipf.write(file_path, arcname)
            logger.info(f"成功创建PPTX: {output_path}")
            return True
        except Exception as e:
            logger.error(f"创建PPTX失败: {e}")
            return False
    
    def update_chart_text(self, slide_xml_path: str, text_content: str) -> bool:
        """
        更新图表中的文字内容
        
        Args:
            slide_xml_path: 幻灯片XML文件的路径
            text_content: 新的文本内容
            
        Returns:
            bool: 是否成功更新
        """
        try:
            tree = ET.parse(slide_xml_path)
            root = tree.getroot()
            
            # 定义命名空间
            namespaces = {
                'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
                'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
                'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
                'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
            }
            
            # 查找所有图表引用
            chart_refs = root.findall('.//c:chart', namespaces)
            if not chart_refs:
                return False
                
            # 获取幻灯片所在目录
            slide_dir = os.path.dirname(slide_xml_path)
            ppt_dir = os.path.dirname(slide_dir)
            
            for chart_ref in chart_refs:
                # 获取图表关系ID
                rid = chart_ref.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                if not rid:
                    continue
                    
                # 读取关系文件找到图表XML
                rels_path = os.path.join(slide_dir, '_rels', os.path.basename(slide_xml_path) + '.rels')
                if not os.path.exists(rels_path):
                    continue
                    
                rels_tree = ET.parse(rels_path)
                rels_root = rels_tree.getroot()
                
                # 查找图表文件路径
                chart_rel = rels_root.find(f".//Relationship[@Id='{rid}']")
                if chart_rel is None:
                    continue
                    
                chart_path = chart_rel.get('Target')
                if not chart_path:
                    continue
                    
                # 转换为绝对路径
                if not os.path.isabs(chart_path):
                    chart_path = os.path.join(ppt_dir, chart_path.lstrip('/'))
                
                # 更新图表文本
                if os.path.exists(chart_path):
                    chart_tree = ET.parse(chart_path)
                    chart_root = chart_tree.getroot()
                    
                    # 获取图表中的所有文本
                    text_elements = []
                    node_texts = []
                    
                    # 标题
                    title = chart_root.find('.//c:title//c:v', namespaces)
                    if title is not None and title.text:
                        text_elements.append(title)
                        node_texts.append(title.text.strip())
                    
                    # 图例
                    legends = chart_root.findall('.//c:legend//c:v', namespaces)
                    for legend in legends:
                        if legend.text:
                            text_elements.append(legend)
                            node_texts.append(legend.text.strip())
                    
                    # 轴标题和标签
                    axes = chart_root.findall('.//c:axis//c:v', namespaces)
                    for axis in axes:
                        if axis.text:
                            text_elements.append(axis)
                            node_texts.append(axis.text.strip())
                    
                    # 数据标签
                    data_labels = chart_root.findall('.//c:dLbls//c:v', namespaces)
                    for label in data_labels:
                        if label.text:
                            text_elements.append(label)
                            node_texts.append(label.text.strip())
                    
                    if text_elements:
                        # 构建提示词
                        prompt = f"""
你是PPT图表内容生成专家。请根据以下内容生成新的图表文本。

原有文本:
{node_texts}

新的内容:
{text_content}

要求:
1. 必须生成与原文本数量完全相同的文本
2. 每段文本长度要接近原文本
3. 保持原有文本的功能（如果是标题就生成标题，如果是图例就生成图例等）
4. 内容要符合图表展示风格

请返回JSON格式:
{{
  "texts": {node_texts}  // 数组长度必须是{len(node_texts)}
}}
"""
                        # 调用API生成新内容
                        response = self.client.chat.completions.create(
                            model="Qwen-72B",
                            messages=[{"role": "user", "content": prompt}],
                            response_format={"type": "json_object"},
                            temperature=0.7,
                            max_tokens=1000
                        )
                        
                        result = json.loads(response.choices[0].message.content)
                        new_texts = result.get("texts", [])
                        
                        # 确保生成的文本数量正确
                        if len(new_texts) != len(text_elements):
                            logger.warning(f"生成的文本数量({len(new_texts)})与原文本数量({len(text_elements)})不匹配")
                            continue
                        
                        # 更新图表中的文本
                        for elem, new_text in zip(text_elements, new_texts):
                            old_text = elem.text
                            elem.text = new_text
                            logger.info(f"更新图表文本: '{old_text[:30]}...' -> '{new_text[:30]}...'")
                        
                        # 保存修改后的图表XML
                        chart_tree.write(chart_path, encoding='UTF-8', xml_declaration=True)
                        logger.info(f"成功更新图表内容: {chart_path}")
            
            return True
            
        except Exception as e:
            logger.error(f"更新图表文本失败: {e}")
            logger.error(traceback.format_exc())
            return False

    def get_slide_type_and_elements(self, slide_number: int) -> Tuple[str, Dict]:
        """获取幻灯片类型和对应的文本元素定义"""
        try:
            if slide_number == 1:
                return "title_slide", self.template_description["slide01"]["text_elements"]
            elif slide_number == 2:
                return "content_page", self.template_description["slide02"]["text_elements"]
            else:
                content_page_number = slide_number - 2  # 减去首页和目录页
                if content_page_number % 3 == 1:  # 章节标题页
                    return "chapter_title", self.template_description["chapter_title_slides"]["text_elements"]
                else:  # 内容页
                    return "content_page", self.template_description["content_slides"]["text_elements"]
        except Exception as e:
            logger.error(f"获取幻灯片类型和元素定义失败: {e}")
            return "unknown", {}

    def get_slide_content(self, slide_number: int, total_content_pages: int) -> Dict[str, Any]:
        """根据幻灯片编号获取对应的内容"""
        try:
            slide_type, text_elements = self.get_slide_type_and_elements(slide_number)
            
            # 首页 (slide1)
            if slide_type == "title_slide":
                return {
                    str(text_elements["main_title"]["id"]): self.outline.get("main_title", ""),
                    str(text_elements["subtitle"]["id"]): self.outline.get("subtitle", ""),  # 添加副标题
                    str(text_elements["speaker"]["id"]): self.outline.get("speaker_name", "")
                }
            
            # 目录页 (slide2)
            elif slide_type == "content_page" and slide_number == 2:
                content = {
                    str(text_elements["main_title"]["id"]): "目录",
                    str(text_elements["subtitle"]["id"]): ""  # 副标题可以为空
                }
                
                # 根据总页数计算章节数
                total_chapters = (total_content_pages + 2) // 3  # 每章3页：标题页+2个内容页
                chapters = []
                for i in range(total_chapters):
                    chapter_key = f"chapter{str(i + 1).zfill(2)}_title"
                    chapters.append(self.outline.get(chapter_key, ""))
                
                # 获取目录项的ID范围
                id_range = text_elements["chapters"]["id_range"].split("-")
                start_id, end_id = int(id_range[0]), int(id_range[1])
                
                for i, chapter in enumerate(chapters):
                    if start_id + i <= end_id:
                        content[str(start_id + i)] = chapter
                
                return content
            
            # 章节标题页
            elif slide_type == "chapter_title":
                content_page_number = slide_number - 2
                chapter_index = (content_page_number - 1) // 3
                chapter_key = f"chapter{str(chapter_index + 1).zfill(2)}_title"
                
                return {
                    str(text_elements["chapter_number"]["id"]): f"{chapter_index + 1}",  # 保留序号
                    str(text_elements["chapter_title"]["id"]): self.outline.get(chapter_key, "")
                }
            
            # 内容页
            elif slide_type == "content_page" and slide_number > 2:
                content_page_number = slide_number - 2
                chapter_index = (content_page_number - 1) // 3
                page_index = (content_page_number - 1) % 3
                
                sections_key = f"sections{chapter_index + 1}"
                sections = self.outline.get(sections_key, [{}])
                
                # 确保sections是列表且有内容
                if not isinstance(sections, list) or not sections:
                    sections = [{}]
                
                section_data = sections[0]
                content = {
                    str(text_elements["title"]["id"]): self.outline.get(f"chapter{str(chapter_index + 1).zfill(2)}_title", ""),
                    str(text_elements["spacing"]["id"]): ""  # 空白文本框保持为空
                }
                
                # 获取正文内容的ID
                body_id = str(text_elements["body"]["id"])
                
                # 根据页面索引获取对应的section内容
                section_contents = []
                start_section = page_index * 3
                for i in range(start_section, min(start_section + 3, len(section_data))):
                    section_key = f"section{str(i + 1).zfill(2)}"
                    if section_key in section_data:
                        section_title = section_data[section_key].get("title", "")
                        section_content = section_data[section_key].get("content", "")
                        section_contents.append(f"{section_title}\n{section_content}")
                
                content[body_id] = "\n\n".join(section_contents)
                
                return content
            
            return {}
            
        except Exception as e:
            logger.error(f"获取幻灯片内容失败: {e}")
            return {}

    def generate_content_for_slide_type(self, slide_xml_path: str, text_content: str) -> Dict[str, Any]:
        """根据页面内容生成相应的新内容"""
        try:
            # 获取幻灯片编号
            slide_number = int(os.path.basename(slide_xml_path).replace("slide", "").replace(".xml", ""))
            content_pages = int(self.outline.get("page", 0))
            
            # 获取该幻灯片的内容
            slide_content = self.get_slide_content(slide_number, content_pages)
            if not slide_content:
                return {"texts": []}
            
            # 首先检测幻灯片类型
            has_smartart, has_chart = self.detect_slide_type(slide_xml_path)
            
            # 解析XML文件
            tree = ET.parse(slide_xml_path)
            root = tree.getroot()
            
            # 定义命名空间
            namespaces = {
                'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
                'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
                'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
                'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
                'dgm': 'http://schemas.openxmlformats.org/drawingml/2006/diagram'
            }
            
            # 存储需要更新的文本节点和对应的新内容
            updates = []
            
            # 遍历所有文本框
            for sp in root.findall('.//p:sp', namespaces):
                # 获取文本框ID
                cnvPr = sp.find('.//p:cNvPr', namespaces)
                if cnvPr is not None:
                    shape_id = cnvPr.get('id')
                    shape_name = cnvPr.get('name')
                    logger.info(f"处理文本框: ID={shape_id}, Name={shape_name}")
                    
                    # 获取文本节点
                    txBody = sp.find('.//p:txBody', namespaces)
                    if txBody is not None:
                        # 遍历所有段落
                        for p in txBody.findall('.//a:p', namespaces):
                            # 检查是否包含数字序号
                            has_number = False
                            number_text = None
                            number_node = None
                            
                            # 先检查是否有纯数字文本
                            for r in p.findall('.//a:r', namespaces):
                                t = r.find('a:t', namespaces)
                                if t is not None and t.text and t.text.strip().isdigit():
                                    has_number = True
                                    number_text = t.text
                                    number_node = t
                                    break
                            
                            # 如果文本框ID在slide_content中有对应内容，使用该内容
                            if shape_id in slide_content:
                                new_text = slide_content[shape_id]
                                # 如果有数字序号，保留序号
                                if has_number:
                                    # 更新除了序号以外的所有文本节点
                                    for r in p.findall('.//a:r', namespaces):
                                        t = r.find('a:t', namespaces)
                                        if t is not None and t is not number_node:
                                            updates.append((t, new_text))
                                else:
                                    # 更新所有文本节点
                                    for r in p.findall('.//a:r', namespaces):
                                        t = r.find('a:t', namespaces)
                                        if t is not None:
                                            updates.append((t, new_text))
                            # 否则，对所有非序号文本生成新内容
                            else:
                                # 收集所有非序号文本
                                text_nodes = []
                                full_text = ""
                                for r in p.findall('.//a:r', namespaces):
                                    t = r.find('a:t', namespaces)
                                    if t is not None and t is not number_node and t.text and t.text.strip():
                                        text_nodes.append(t)
                                        full_text += t.text
                                
                                if text_nodes:
                                    # 构建提示词
                                    prompt = f"""
请根据以下要求生成新的文本内容：

原文本：{full_text}
新的上下文：{text_content}

要求：
1. 保持与原文本长度相近
2. 内容要符合新的上下文
3. 保持专业性和正式性
4. 如果原文本是标题，生成新的标题
5. 如果原文本是正文，生成新的正文

请直接返回生成的文本，不要包含任何其他内容。
"""
                                    # 调用API生成新内容
                                    response = self.client.chat.completions.create(
                                        model="Qwen-72B",
                                        messages=[{"role": "user", "content": prompt}],
                                        temperature=0.7,
                                        max_tokens=500
                                    )
                                    
                                    new_text = response.choices[0].message.content.strip()
                                    # 将新文本分配给所有非序号文本节点
                                    for t in text_nodes:
                                        updates.append((t, new_text))
            
            # 如果是图表，处理图表文本
            if has_chart:
                logger.info("检测到图表，处理图表文本")
                self.update_chart_text(slide_xml_path, json.dumps(slide_content, ensure_ascii=False))
            
            # 如果是SmartArt，跳过内容生成
            if has_smartart:
                logger.info("检测到SmartArt，跳过内容生成")
                return {"texts": []}
            
            # 更新所有需要更新的文本节点
            for node, new_text in updates:
                old_text = node.text if node.text else ""
                node.text = new_text
                logger.info(f"更新文本: '{old_text[:30]}...' -> '{new_text[:30]}...'")
            
            # 保存修改后的XML
            tree.write(slide_xml_path, encoding='UTF-8', xml_declaration=True)
            logger.info("成功更新幻灯片内容")
            
            return {"texts": [update[1] for update in updates]}
            
        except Exception as e:
            logger.error(f"生成内容失败: {e}")
            logger.error(traceback.format_exc())
            return {"texts": []}
    
    def smart_template_generation(self, template_path: str, output_path: str, text_content: str) -> bool:
        """智能模板生成：复制模板并修改内容"""
        try:
            if not self.template_description or not self.outline:
                logger.error("模板描述或大纲未加载")
                return False
            
            # 获取内容页数
            content_pages = int(self.outline.get("page", 0))
            if content_pages <= 0:
                logger.error("大纲中未指定有效的页数")
                return False
            
            # 1. 创建临时目录
            temp_dir = f"temp_{uuid.uuid4().hex}"
            os.makedirs(temp_dir, exist_ok=True)
            
            # 2. 复制模板文件
            temp_pptx = os.path.join(temp_dir, "temp.pptx")
            shutil.copy2(template_path, temp_pptx)
            
            # 3. 解压PPTX
            extract_dir = os.path.join(temp_dir, "extracted")
            if not self.extract_pptx(temp_pptx, extract_dir):
                return False
            
            # 4. 修改幻灯片内容
            slides_dir = os.path.join(extract_dir, "ppt", "slides")
            if os.path.exists(slides_dir):
                # 计算总页数（首页 + 目录页 + 内容页数）
                total_slides = 2 + content_pages
                
                # 遍历所有幻灯片XML文件
                slide_files = sorted(os.listdir(slides_dir))
                for slide_file in slide_files:
                    if slide_file.endswith(".xml"):
                        slide_number = int(slide_file.replace("slide", "").replace(".xml", ""))
                        
                        # 如果超出了总页数，跳过处理
                        if slide_number > total_slides:
                            logger.info(f"跳过处理超出范围的幻灯片 {slide_number}")
                            continue
                        
                        slide_path = os.path.join(slides_dir, slide_file)
                        # 为每个幻灯片生成新内容
                        self.generate_content_for_slide_type(slide_path, text_content)
            
            # 5. 重新打包为PPTX
            if not self.create_pptx(extract_dir, output_path):
                return False
            
            # 6. 清理临时文件
            shutil.rmtree(temp_dir)
            logger.info("清理临时文件完成")
            
            logger.info(f"PPT生成完成，保存到: {os.path.abspath(output_path)}")
            return True
            
        except Exception as e:
            logger.error(f"PPT生成失败: {e}")
            logger.error(traceback.format_exc())
            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir)
            return False
    
    def test_api_connection(self) -> bool:
        """测试API连接"""
        try:
            logger.info("测试API连接...")
            response = self.client.chat.completions.create(
                model="Qwen-72B",
                messages=[{"role": "user", "content": "你好"}],
                max_tokens=10
            )
            logger.info("API连接测试成功")
            return True
        except Exception as e:
            logger.error(f"API连接测试失败: {e}")
            return False

def main():
    """主函数"""
    # 从环境变量获取配置
    API_KEY = os.getenv("API_KEY")
    BASE_URL = os.getenv("BASE_URL")
    
    if not API_KEY or not BASE_URL:
        print("错误：请在.env文件中设置API_KEY和BASE_URL")
        print(API_KEY)
        print(BASE_URL)
        return
    
    TEMPLATE_PATH = 'template.pptx'
    OUTPUT_PATH = 'new_generated_ppt.pptx'
    TEMPLATE_DESCRIPTION_PATH = 'template_description.json'
    OUTLINE_PATH = 'template_outline_01.json'
    
    # 测试文本
    test_content = """
日前，为切实保障航空运行安全，民航局发布紧急通知，自6月28日起禁止旅客携带没有3C标识、3C标识不清晰、被召回型号或批次的充电宝乘坐境内航班。"""
    
    try:
        print("=" * 50)
        
        # 创建生成器
        generator = PPTGenerator(api_key=API_KEY, base_url=BASE_URL)
        
        # 测试API连接
        if not generator.test_api_connection():
            print("\nAPI连接失败，请检查配置")
            return
        
        # 加载模板描述和大纲
        if not generator.load_template_description(TEMPLATE_DESCRIPTION_PATH):
            print("\n加载模板描述失败")
            return
            
        if not generator.load_outline(OUTLINE_PATH):
            print("\n加载大纲失败")
            return
        
        # 智能模板生成
        if os.path.exists(TEMPLATE_PATH):
            print(f"\n发现模板文件: {TEMPLATE_PATH}")
            
            success = generator.smart_template_generation(
                template_path=TEMPLATE_PATH,
                output_path=OUTPUT_PATH,
                text_content=test_content
            )
            
        else:
            print(f"\n未发现模板文件")
            return
        
        if success:
            print(f"输出文件: {os.path.abspath(OUTPUT_PATH)}")
        else:
            print(f"\n PPT生成失败")
    
    except Exception as e:
        print(f"\n程序执行失败: {e}")

if __name__ == '__main__':
    main() 