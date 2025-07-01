import os
import json
import logging
import uuid
import shutil
import zipfile
import xml.etree.ElementTree as ET

from openai import OpenAI
from typing import Dict, Any, Tuple
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



class PPTGenerator:
    
    def __init__(self, api_key: str, base_url: str):
        try:
            self.client = OpenAI(api_key=api_key, base_url=base_url)
            logger.info("OpenAI客户端初始化成功")
        except Exception as e:
            logger.error(f"初始化OpenAI客户端失败: {e}")
            raise
    
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
                'dgm': 'http://schemas.openxmlformats.org/drawingml/2006/diagram'
            }
            
            # 检查SmartArt
            # SmartArt在PPT XML中通常由 <dgm:relIds> 标签标识
            has_smartart = len(root.findall('.//dgm:relIds', namespaces)) > 0
            
            # 检查图表
            # 图表在PPT XML中通常由 <c:chart> 标签标识
            has_chart = len(root.findall('.//c:chart', namespaces)) > 0
            
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
                            new_texts = new_texts[:len(text_elements)]
                        
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

    def load_outline(self, outline_path: str) -> Dict[str, Any]:
        """加载大纲文件"""
        try:
            with open(outline_path, 'r', encoding='utf-8') as f:
                outline = json.load(f)
            logger.info("成功加载大纲文件")
            return outline
        except Exception as e:
            logger.error(f"加载大纲文件失败: {e}")
            return {}

    def generate_content_for_slide_type(self, slide_xml_path: str, text_content: str, outline: Dict[str, Any] = None) -> Dict[str, Any]:
        """根据页面内容生成相应的新内容"""
        try:
            # 首先检测幻灯片类型
            has_smartart, has_chart = self.detect_slide_type(slide_xml_path)
            
            # 如果是图表，更新图表文本
            if has_chart:
                logger.info("检测到图表，开始更新图表文本")
                self.update_chart_text(slide_xml_path, text_content)
            
            # 如果是SmartArt，则跳过内容生成
            if has_smartart:
                logger.info("检测到SmartArt，跳过内容生成")
                return {"texts": []}
            
            # 解析XML文件
            tree = ET.parse(slide_xml_path)
            root = tree.getroot()
            
            # 定义命名空间
            namespaces = {
                'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
                'p': 'http://schemas.openxmlformats.org/presentationml/2006/main'
            }
            
            # 获取所有非空文本节点
            text_nodes = []
            node_texts = []
            for node in root.findall('.//a:t', namespaces):
                if node.text and node.text.strip():
                    text_nodes.append(node)
                    node_texts.append(node.text.strip())
            
            if not text_nodes:
                logger.info("该页面没有找到文本内容")
                return {"texts": []}
            
            # 构建提示词
            outline_prompt = ""
            if outline:
                outline_prompt = f"""
参考大纲：
{json.dumps(outline, ensure_ascii=False, indent=2)}

请注意：
1. 这个大纲仅作为内容生成的参考和指导
2. 不要直接复制大纲中的内容
3. 根据大纲的整体主题和结构，生成符合当前幻灯片位置和样式的内容
4. 保持内容的连贯性和逻辑性
"""

            prompt = f"""
你是PPT内容生成专家。请根据以下内容生成新的PPT内容。

原有文本框内容:
{node_texts}

新的内容主题:
{text_content}

{outline_prompt if outline else ""}

要求:
1. 必须生成与原文本数量完全相同的文本
2. 每段文本长度要接近原文本
3. 内容要符合PPT展示风格
4. 如果原文本是标题，生成的也要是标题风格
5. 如果原文本是正文，生成的也要是正文风格
6. 根据原文本的位置和样式，生成合适的内容

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
            if len(new_texts) != len(text_nodes):
                logger.warning(f" 生成的文本数量({len(new_texts)})与原文本数量({len(text_nodes)})不匹配")
                return {"texts": []}
            
            # 更新XML中的文本
            for node, new_text in zip(text_nodes, new_texts):
                old_text = node.text
                node.text = new_text
                logger.info(f" 更新文本: '{old_text[:30]}...' -> '{new_text[:30]}...'")
            
            # 保存修改后的XML
            tree.write(slide_xml_path, encoding='UTF-8', xml_declaration=True)
            logger.info(f"成功更新幻灯片内容")
            return result
            
        except Exception as e:
            logger.error(f" 生成内容失败: {e}")
            return {"texts": []}
    
    def smart_template_generation(self, template_path: str, output_path: str, text_content: str, outline_path: str = None) -> bool:
        """智能模板生成：复制模板并修改内容"""
        try:
            # 1. 加载大纲（如果提供）
            outline = None
            if outline_path:
                outline = self.load_outline(outline_path)
                if not outline:
                    return False
            
            # 2. 创建临时目录
            temp_dir = f"temp_{uuid.uuid4().hex}"
            os.makedirs(temp_dir, exist_ok=True)
            
            # 3. 复制模板文件
            temp_pptx = os.path.join(temp_dir, "temp.pptx")
            shutil.copy2(template_path, temp_pptx)
            
            # 4. 解压PPTX
            extract_dir = os.path.join(temp_dir, "extracted")
            if not self.extract_pptx(temp_pptx, extract_dir):
                return False
            
            # 5. 修改幻灯片内容
            slides_dir = os.path.join(extract_dir, "ppt", "slides")
            if os.path.exists(slides_dir):
                # 遍历所有幻灯片XML文件
                for slide_file in sorted(os.listdir(slides_dir)):
                    if slide_file.endswith(".xml"):
                        slide_path = os.path.join(slides_dir, slide_file)
                        # 为每个幻灯片生成新内容
                        self.generate_content_for_slide_type(slide_path, text_content, outline)
            
            # 6. 重新打包为PPTX
            if not self.create_pptx(extract_dir, output_path):
                return False
            
            # 7. 清理临时文件
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
    # 配置参数
    API_KEY = os.getenv("API_KEY")
    BASE_URL = os.getenv("BASE_URL")
    TEMPLATE_PATH = 'template_dfsj.pptx'
    OUTPUT_PATH = 'new_generated_ppt.pptx'
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
        
        # 智能模板生成
        if os.path.exists(TEMPLATE_PATH):
            print(f"\n发现模板文件: {TEMPLATE_PATH}")
            
            # 检查是否使用大纲模式
            if os.path.exists(OUTLINE_PATH):
                print(f"\n使用大纲模式，大纲文件: {OUTLINE_PATH}")
                success = generator.smart_template_generation(
                    template_path=TEMPLATE_PATH,
                    output_path=OUTPUT_PATH,
                    text_content=test_content,
                    outline_path=OUTLINE_PATH
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