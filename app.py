import os
import sys
import webbrowser

from smart_art_generator import SmartArtGenerator

# --- 1. 配置 ---
# 建议从环境变量或配置文件中读取，这里为方便演示而硬编码
API_KEY = "sk-123456789"
BASE_URL = "http://180.153.21.76:17009/v1"

# 输入的文本段落
INPUT_TEXT = "我们项目的启动流程是这样的：首先，产品经理提出需求。接着，开发团队会对需求进行评审。关于评审，主要关注技术可行性和资源匹配两个核心要点。评审通过后，UI和开发将同步进行工作。"

# --- 图像映射: 将流程图中的特定文本映射到本地图像路径 ---
# 支持按位置 ("Image1", "Image2", ...) 或按文本内容进行映射
IMAGE_MAPPING = {
    "Image1": r"C:\Users\CHENQIMING\Pictures\20230406163343262212222.png"
}

# XML 模板和输出路径
TEMPLATE_PATH = 'flowchart_template.xml'
OUTPUT_DIR = 'tmp'
OUTPUT_FILENAME_BASE = 'smart_art'

def main():
    """
    主函数，执行从文本分析到基于模板生成XML，再到渲染SVG的整个流程。
    """
    print("--- 基于模板的智能艺术图表生成开始 ---")
    
    # --- 2. 初始化生成器 ---
    try:
        generator = SmartArtGenerator(api_key=API_KEY, base_url=BASE_URL)
    except (ValueError, RuntimeError) as e:
        print(f"[错误] 初始化失败: {e}", file=sys.stderr)
        sys.exit(1)
        
    # --- 3. 分析文本并使用模板生成所有适用的XML ---
    try:
        xml_charts = generator.process_text_with_template(
            INPUT_TEXT, 
            template_path=TEMPLATE_PATH, 
            image_mapping=IMAGE_MAPPING
        )
        if not xml_charts:
            print("\n[警告] 未能从文本中生成任何图表。")
            sys.exit(0)
    except (RuntimeError, ValueError) as e:
        print(f"[错误] 处理过程中发生错误: {e}", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"[错误] 发生未知错误: {e}", file=sys.stderr)
        sys.exit(1)

    # --- 4. 保存XML和渲染的SVG图表 ---
    try:
        if not os.path.exists(OUTPUT_DIR):
            os.makedirs(OUTPUT_DIR)
        
        saved_svg_files = []
        for chart_name, xml_content in xml_charts.items():
            # 保存填充后的XML
            xml_file_path = os.path.join(OUTPUT_DIR, f"{OUTPUT_FILENAME_BASE}_{chart_name.lower()}.xml")
            with open(xml_file_path, 'w', encoding='utf-8') as f:
                f.write(xml_content)
            print(f"\n[完成] {chart_name} XML已基于模板生成并保存到: {os.path.abspath(xml_file_path)}")
            
            # 从XML渲染SVG
            svg_content = generator.render_xml_to_svg(xml_content)
            svg_file_path = os.path.join(OUTPUT_DIR, f"{OUTPUT_FILENAME_BASE}_{chart_name.lower()}.svg")
            with open(svg_file_path, 'w', encoding='utf-8') as f:
                f.write(svg_content)
            print(f"[完成] {chart_name} SVG图像已渲染并保存到: {os.path.abspath(svg_file_path)}")
            saved_svg_files.append(svg_file_path)

        # 在浏览器中打开第一个生成的SVG文件
        if saved_svg_files:
            webbrowser.open('file://' + os.path.realpath(saved_svg_files[0]))
            print(f"\n已在默认浏览器中打开: {os.path.basename(saved_svg_files[0])}")

    except (IOError, ValueError) as e:
        print(f"[错误] 无法写入文件或渲染SVG: {e}", file=sys.stderr)
        sys.exit(1)

if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')
    main() 