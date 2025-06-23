
"""
Module để tạo PowerPoint presentations từ dữ liệu structured
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from datetime import datetime
from typing import Dict, List, Optional, Any
from io import BytesIO
import json
import logging

# Setup logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class PowerPointGenerator:
    """
    Class chính để tạo PowerPoint presentations
    """
    
    def __init__(self):
        self.presentation = None
        self.slide_layouts = None
        self.current_template = "education"
        
        # Template configurations
        self.templates = {
            "education": {
                "primary_color": "#2E86AB",
                "secondary_color": "#A23B72", 
                "background_color": "#F18F01",
                "text_color": "#0F0F0F",
                "font_size": {
                    "title": 32,
                    "subtitle": 24,
                    "content": 18,
                    "caption": 14
                }
            },
            "business": {
                "primary_color": "#1565C0",
                "secondary_color": "#FFA726",
                "background_color": "#E3F2FD", 
                "text_color": "#263238",
                "font_size": {
                    "title": 36,
                    "subtitle": 28,
                    "content": 20,
                    "caption": 16
                }
            }
        }
    
    def create_new_presentation(self, template: str = "education") -> bool:
        """
        Tạo presentation mới
        
        Args:
            template (str): Tên template muốn sử dụng
            
        Returns:
            bool: True nếu tạo thành công
        """
        try:
            self.presentation = Presentation()
            self.slide_layouts = self.presentation.slide_layouts
            self.current_template = template
            
            logger.info(f"Created new presentation with template: {template}")
            return True
            
        except Exception as e:
            logger.error(f"Error creating presentation: {str(e)}")
            return False
    
    def add_title_slide(self, title: str, subtitle: str = "", author: str = "") -> bool:
        """
        Thêm slide tiêu đề
        
        Args:
            title (str): Tiêu đề chính
            subtitle (str): Phụ đề
            author (str): Tác giả
            
        Returns:
            bool: True nếu thêm thành công
        """
        try:
            slide_layout = self.slide_layouts[0]  # Title slide layout
            slide = self.presentation.slides.add_slide(slide_layout)
            
            # Set title
            title_shape = slide.shapes.title
            title_shape.text = title
            self._apply_title_formatting(title_shape)
            
            # Set subtitle
            if subtitle and len(slide.placeholders) > 1:
                subtitle_shape = slide.placeholders[1]
                if author:
                    subtitle_shape.text = f"{subtitle}\n\n{author}"
                else:
                    subtitle_shape.text = subtitle
                self._apply_subtitle_formatting(subtitle_shape)
            
            logger.info(f"Added title slide: {title}")
            return True
            
        except Exception as e:
            logger.error(f"Error adding title slide: {str(e)}")
            return False
    
    def add_content_slide(self, title: str, content: List[str], slide_type: str = "bullet") -> bool:
        """
        Thêm slide nội dung
        
        Args:
            title (str): Tiêu đề slide
            content (List[str]): Danh sách nội dung
            slide_type (str): Loại slide (bullet, numbered, paragraph)
            
        Returns:
            bool: True nếu thêm thành công
        """
        try:
            slide_layout = self.slide_layouts[1]  # Content layout
            slide = self.presentation.slides.add_slide(slide_layout)
            
            # Set title
            title_shape = slide.shapes.title
            title_shape.text = title
            self._apply_content_title_formatting(title_shape)
            
            # Set content
            content_shape = slide.placeholders[1]
            text_frame = content_shape.text_frame
            text_frame.clear()
            
            # Add content based on type
            if slide_type == "bullet":
                self._add_bullet_content(text_frame, content)
            elif slide_type == "numbered":
                self._add_numbered_content(text_frame, content)
            else:  # paragraph
                self._add_paragraph_content(text_frame, content)
            
            logger.info(f"Added content slide: {title}")
            return True
            
        except Exception as e:
            logger.error(f"Error adding content slide: {str(e)}")
            return False
    
    def add_two_column_slide(self, title: str, left_content: List[str], right_content: List[str]) -> bool:
        """
        Thêm slide 2 cột
        
        Args:
            title (str): Tiêu đề slide
            left_content (List[str]): Nội dung cột trái
            right_content (List[str]): Nội dung cột phải
            
        Returns:
            bool: True nếu thêm thành công
        """
        try:
            slide_layout = self.slide_layouts[3]  # Two content layout
            slide = self.presentation.slides.add_slide(slide_layout)
            
            # Set title
            title_shape = slide.shapes.title
            title_shape.text = title
            self._apply_content_title_formatting(title_shape)
            
            # Left column
            left_shape = slide.placeholders[1]
            left_frame = left_shape.text_frame
            left_frame.clear()
            self._add_bullet_content(left_frame, left_content)
            
            # Right column
            right_shape = slide.placeholders[2]
            right_frame = right_shape.text_frame
            right_frame.clear()
            self._add_bullet_content(right_frame, right_content)
            
            logger.info(f"Added two-column slide: {title}")
            return True
            
        except Exception as e:
            logger.error(f"Error adding two-column slide: {str(e)}")
            return False
    
    def add_image_slide(self, title: str, image_path: str, caption: str = "") -> bool:
        """
        Thêm slide với hình ảnh
        
        Args:
            title (str): Tiêu đề slide
            image_path (str): Đường dẫn đến hình ảnh
            caption (str): Chú thích hình ảnh
            
        Returns:
            bool: True nếu thêm thành công
        """
        try:
            slide_layout = self.slide_layouts[6]  # Blank layout
            slide = self.presentation.slides.add_slide(slide_layout)
            
            # Add title
            title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(9), Inches(1))
            title_frame = title_box.text_frame
            title_frame.text = title
            self._apply_content_title_formatting(title_box)
            
            # Add image
            try:
                img = slide.shapes.add_picture(
                    image_path,
                    Inches(1),
                    Inches(1.5),
                    width=Inches(8)
                )
                
                # Add caption
                if caption:
                    caption_box = slide.shapes.add_textbox(
                        Inches(1),
                        img.top + img.height + Inches(0.2),
                        Inches(8),
                        Inches(0.5)
                    )
                    caption_frame = caption_box.text_frame
                    caption_frame.text = caption
                    self._apply_caption_formatting(caption_box)
                
                logger.info(f"Added image slide: {title}")
                return True
                
            except Exception as img_error:
                logger.error(f"Error adding image: {img_error}")
                # Add placeholder text instead
                placeholder_box = slide.shapes.add_textbox(
                    Inches(1), Inches(2), Inches(8), Inches(4)
                )
                placeholder_frame = placeholder_box.text_frame
                placeholder_frame.text = f"[Hình ảnh: {image_path}]\n{caption}"
                return True
                
        except Exception as e:
            logger.error(f"Error adding image slide: {str(e)}")
            return False
    
    def add_chart_slide(self, title: str, chart_data: Dict[str, Any]) -> bool:
        """
        Thêm slide với biểu đồ
        
        Args:
            title (str): Tiêu đề slide
            chart_data (Dict): Dữ liệu biểu đồ
                {
                    "type": "column|line|pie|bar",
                    "categories": ["A", "B", "C"],
                    "series": {
                        "Series 1": [1, 2, 3],
                        "Series 2": [4, 5, 6]
                    }
                }
                
        Returns:
            bool: True nếu thêm thành công
        """
        try:
            slide_layout = self.slide_layouts[6]  # Blank layout
            slide = self.presentation.slides.add_slide(slide_layout)
            
            # Add title
            title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(9), Inches(1))
            title_frame = title_box.text_frame
            title_frame.text = title
            self._apply_content_title_formatting(title_box)
            
            # Prepare chart data
            chart_data_obj = CategoryChartData()
            chart_data_obj.categories = chart_data.get('categories', [])
            
            for series_name, values in chart_data.get('series', {}).items():
                chart_data_obj.add_series(series_name, values)
            
            # Chart type mapping
            chart_types = {
                'column': XL_CHART_TYPE.COLUMN_CLUSTERED,
                'line': XL_CHART_TYPE.LINE,
                'pie': XL_CHART_TYPE.PIE,
                'bar': XL_CHART_TYPE.BAR_CLUSTERED
            }
            
            chart_type = chart_types.get(chart_data.get('type', 'column'), XL_CHART_TYPE.COLUMN_CLUSTERED)
            
            # Add chart
            chart = slide.shapes.add_chart(
                chart_type,
                Inches(1), Inches(2), Inches(8), Inches(5),
                chart_data_obj
            ).chart
            
            chart.has_legend = True
            chart.legend.position = 2  # Right
            
            logger.info(f"Added chart slide: {title}")
            return True
            
        except Exception as e:
            logger.error(f"Error adding chart slide: {str(e)}")
            return False
    
    def add_table_slide(self, title: str, table_data: List[List[str]], has_header: bool = True) -> bool:
        """
        Thêm slide với bảng
        
        Args:
            title (str): Tiêu đề slide
            table_data (List[List[str]]): Dữ liệu bảng
            has_header (bool): Có header row không
            
        Returns:
            bool: True nếu thêm thành công
        """
        try:
            slide_layout = self.slide_layouts[6]  # Blank layout
            slide = self.presentation.slides.add_slide(slide_layout)
            
            # Add title
            title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(9), Inches(1))
            title_frame = title_box.text_frame
            title_frame.text = title
            self._apply_content_title_formatting(title_box)
            
            # Create table
            if table_data:
                rows = len(table_data)
                cols = len(table_data[0]) if table_data else 0
                
                if rows > 0 and cols > 0:
                    table = slide.shapes.add_table(
                        rows, cols,
                        Inches(1), Inches(2),
                        Inches(8), Inches(4)
                    ).table
                    
                    # Fill table with data
                    for i, row_data in enumerate(table_data):
                        for j, cell_data in enumerate(row_data):
                            cell = table.cell(i, j)
                            cell.text = str(cell_data)
                            
                            # Style header row
                            if i == 0 and has_header:
                                self._apply_table_header_formatting(cell)
                            else:
                                self._apply_table_cell_formatting(cell)
            
            logger.info(f"Added table slide: {title}")
            return True
            
        except Exception as e:
            logger.error(f"Error adding table slide: {str(e)}")
            return False
    
    def add_conclusion_slide(self, title: str = "Kết luận", points: List[str] = None) -> bool:
        """
        Thêm slide kết luận
        
        Args:
            title (str): Tiêu đề slide
            points (List[str]): Các điểm kết luận
            
        Returns:
            bool: True nếu thêm thành công
        """
        try:
            if points is None:
                points = ["Cảm ơn các bạn đã lắng nghe!", "Có câu hỏi nào không?"]
            
            return self.add_content_slide(title, points, "bullet")
            
        except Exception as e:
            logger.error(f"Error adding conclusion slide: {str(e)}")
            return False
    
    def create_from_structured_data(self, presentation_data: Dict[str, Any]) -> bool:
        """
        Tạo presentation từ dữ liệu có cấu trúc
        
        Args:
            presentation_data (Dict): Dữ liệu presentation
                {
                    "title": "Presentation Title",
                    "subtitle": "Subtitle",
                    "author": "Author Name",
                    "template": "education",
                    "slides": [
                        {
                            "type": "content|two_column|image|chart|table",
                            "title": "Slide Title",
                            "content": [...],
                            "extra_data": {...}
                        }
                    ]
                }
                
        Returns:
            bool: True nếu tạo thành công
        """
        try:
            # Create new presentation
            template = presentation_data.get('template', 'education')
            if not self.create_new_presentation(template):
                return False
            
            # Add title slide
            self.add_title_slide(
                presentation_data.get('title', 'Presentation'),
                presentation_data.get('subtitle', ''),
                presentation_data.get('author', '')
            )
            
            # Add content slides
            for slide_data in presentation_data.get('slides', []):
                slide_type = slide_data.get('type', 'content')
                title = slide_data.get('title', '')
                
                if slide_type == 'content':
                    content = slide_data.get('content', [])
                    self.add_content_slide(title, content)
                    
                elif slide_type == 'two_column':
                    left_content = slide_data.get('left_content', [])
                    right_content = slide_data.get('right_content', [])
                    self.add_two_column_slide(title, left_content, right_content)
                    
                elif slide_type == 'image':
                    image_path = slide_data.get('image_path', '')
                    caption = slide_data.get('caption', '')
                    self.add_image_slide(title, image_path, caption)
                    
                elif slide_type == 'chart':
                    chart_data = slide_data.get('chart_data', {})
                    self.add_chart_slide(title, chart_data)
                    
                elif slide_type == 'table':
                    table_data = slide_data.get('table_data', [])
                    has_header = slide_data.get('has_header', True)
                    self.add_table_slide(title, table_data, has_header)
            
            # Add conclusion slide if specified
            if presentation_data.get('add_conclusion', True):
                conclusion_points = presentation_data.get('conclusion_points', None)
                self.add_conclusion_slide(points=conclusion_points)
            
            logger.info("Successfully created presentation from structured data")
            return True
            
        except Exception as e:
            logger.error(f"Error creating presentation from data: {str(e)}")
            return False
    
    def save_to_file(self, filename: str) -> bool:
        """
        Lưu presentation vào file
        
        Args:
            filename (str): Tên file để lưu
            
        Returns:
            bool: True nếu lưu thành công
        """
        try:
            if self.presentation is None:
                logger.error("No presentation to save")
                return False
                
            self.presentation.save(filename)
            logger.info(f"Presentation saved to: {filename}")
            return True
            
        except Exception as e:
            logger.error(f"Error saving presentation: {str(e)}")
            return False
    
    def save_to_buffer(self) -> Optional[BytesIO]:
        """
        Lưu presentation vào BytesIO buffer để download
        
        Returns:
            BytesIO: Buffer chứa file PowerPoint hoặc None nếu lỗi
        """
        try:
            if self.presentation is None:
                logger.error("No presentation to save")
                return None
                
            buffer = BytesIO()
            self.presentation.save(buffer)
            buffer.seek(0)
            
            logger.info("Presentation saved to buffer")
            return buffer
            
        except Exception as e:
            logger.error(f"Error saving presentation to buffer: {str(e)}")
            return None
    
    def get_slide_count(self) -> int:
        """
        Lấy số lượng slides
        
        Returns:
            int: Số lượng slides
        """
        if self.presentation is None:
            return 0
        return len(self.presentation.slides)
    
    # Private formatting methods
    def _apply_title_formatting(self, shape):
        """Apply formatting cho title slide"""
        template = self.templates[self.current_template]
    
        for paragraph in shape.text_frame.paragraphs:
            paragraph.font.size = Pt(template['font_size']['title'])
            paragraph.font.bold = True
            # FIX: Sử dụng RGBColor với từng component
            color_hex = template['primary_color'].replace('#', '')
            r = int(color_hex[0:2], 16)
            g = int(color_hex[2:4], 16) 
            b = int(color_hex[4:6], 16)
            paragraph.font.color.rgb = RGBColor(r, g, b)
            paragraph.alignment = PP_ALIGN.CENTER
    
    def _apply_subtitle_formatting(self, shape):
        """Apply formatting cho subtitle"""
        template = self.templates[self.current_template]
        
        for paragraph in shape.text_frame.paragraphs:
            paragraph.font.size = Pt(template['font_size']['subtitle'])
            # FIX: Sử dụng RGBColor với từng component
            color_hex = template['secondary_color'].replace('#', '')
            r = int(color_hex[0:2], 16)
            g = int(color_hex[2:4], 16)
            b = int(color_hex[4:6], 16)
            paragraph.font.color.rgb = RGBColor(r, g, b)
            paragraph.alignment = PP_ALIGN.CENTER
    
    def _apply_content_title_formatting(self, shape):
        """Apply formatting cho content title"""
        template = self.templates[self.current_template]
        
        for paragraph in shape.text_frame.paragraphs:
            paragraph.font.size = Pt(template['font_size']['subtitle'])
            paragraph.font.bold = True
            # FIX: Sử dụng RGBColor với từng component
            color_hex = template['primary_color'].replace('#', '')
            r = int(color_hex[0:2], 16)
            g = int(color_hex[2:4], 16)
            b = int(color_hex[4:6], 16)
            paragraph.font.color.rgb = RGBColor(r, g, b)
    
    def _apply_caption_formatting(self, shape):
        """Apply formatting cho caption"""
        template = self.templates[self.current_template]
        
        for paragraph in shape.text_frame.paragraphs:
            paragraph.font.size = Pt(template['font_size']['caption'])
            paragraph.font.italic = True
            # FIX: Sử dụng RGBColor với từng component
            color_hex = template['text_color'].replace('#', '')
            r = int(color_hex[0:2], 16)
            g = int(color_hex[2:4], 16)
            b = int(color_hex[4:6], 16)
            paragraph.font.color.rgb = RGBColor(r, g, b)
            paragraph.alignment = PP_ALIGN.CENTER

    def _apply_table_header_formatting(self, cell):
        """Apply formatting cho table header"""
        template = self.templates[self.current_template]
        
        cell.fill.solid()
        # FIX: Sử dụng RGBColor với từng component cho fill
        color_hex = template['primary_color'].replace('#', '')
        r = int(color_hex[0:2], 16)
        g = int(color_hex[2:4], 16)
        b = int(color_hex[4:6], 16)
        cell.fill.fore_color.rgb = RGBColor(r, g, b)
        
        for paragraph in cell.text_frame.paragraphs:
            paragraph.font.color.rgb = RGBColor(255, 255, 255)
            paragraph.font.bold = True
            paragraph.font.size = Pt(template['font_size']['content'])

    def _apply_table_cell_formatting(self, cell):
        """Apply formatting cho table cell"""
        template = self.templates[self.current_template]
        
        for paragraph in cell.text_frame.paragraphs:
            paragraph.font.size = Pt(template['font_size']['content'])
            # FIX: Sử dụng RGBColor với từng component
            color_hex = template['text_color'].replace('#', '')
            r = int(color_hex[0:2], 16)
            g = int(color_hex[2:4], 16)
            b = int(color_hex[4:6], 16)
            paragraph.font.color.rgb = RGBColor(r, g, b)

    def _add_bullet_content(self, text_frame, content):
        """Add bullet point content"""
        template = self.templates[self.current_template]
        
        for i, item in enumerate(content):
            if i == 0:
                p = text_frame.paragraphs[0]
            else:
                p = text_frame.add_paragraph()
            
            p.text = str(item)
            p.level = 0
            p.font.size = Pt(template['font_size']['content'])
            # FIX: Sử dụng RGBColor với từng component
            color_hex = template['text_color'].replace('#', '')
            r = int(color_hex[0:2], 16)
            g = int(color_hex[2:4], 16)
            b = int(color_hex[4:6], 16)
            p.font.color.rgb = RGBColor(r, g, b)

    def _add_numbered_content(self, text_frame, content):
        """Add numbered content"""
        template = self.templates[self.current_template]
        
        for i, item in enumerate(content):
            if i == 0:
                p = text_frame.paragraphs[0]
            else:
                p = text_frame.add_paragraph()
            
            p.text = f"{i + 1}. {str(item)}"
            p.level = 0
            p.font.size = Pt(template['font_size']['content'])
            # FIX: Sử dụng RGBColor với từng component
            color_hex = template['text_color'].replace('#', '')
            r = int(color_hex[0:2], 16)
            g = int(color_hex[2:4], 16)
            b = int(color_hex[4:6], 16)
            p.font.color.rgb = RGBColor(r, g, b)

    def _add_paragraph_content(self, text_frame, content):
        """Add paragraph content"""
        template = self.templates[self.current_template]
        
        # Join all content into a single paragraph
        full_text = '\n\n'.join(str(item) for item in content)
        
        p = text_frame.paragraphs[0]
        p.text = full_text
        p.font.size = Pt(template['font_size']['content'])
        # FIX: Sử dụng RGBColor với từng component
        color_hex = template['text_color'].replace('#', '')
        r = int(color_hex[0:2], 16)
        g = int(color_hex[2:4], 16)
        b = int(color_hex[4:6], 16)
        p.font.color.rgb = RGBColor(r, g, b)

# Utility functions
def create_presentation_from_json(json_file: str) -> Optional[PowerPointGenerator]:
    """
    Tạo presentation từ file JSON
    
    Args:
        json_file (str): Đường dẫn đến file JSON
        
    Returns:
        PowerPointGenerator: Instance đã tạo presentation hoặc None nếu lỗi
    """
    try:
        with open(json_file, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        generator = PowerPointGenerator()
        if generator.create_from_structured_data(data):
            return generator
        return None
        
    except Exception as e:
        logger.error(f"Error creating presentation from JSON: {str(e)}")
        return None

def create_sample_presentation() -> PowerPointGenerator:
    """
    Tạo một presentation mẫu để test
    
    Returns:
        PowerPointGenerator: Instance với presentation mẫu
    """
    sample_data = {
        "title": "Bài Giảng Mẫu",
        "subtitle": "Được tạo bởi AI PowerPoint Generator",
        "author": "AI Assistant",
        "template": "education",
        "slides": [
            {
                "type": "content",
                "title": "Mục tiêu bài học",
                "content": [
                    "Hiểu được khái niệm cơ bản",
                    "Vận dụng kiến thức vào thực tế",
                    "Phát triển tư duy logic"
                ]
            },
            {
                "type": "two_column",
                "title": "So sánh",
                "left_content": [
                    "Ưu điểm:",
                    "• Dễ hiểu",
                    "• Thực tế",
                    "• Hiệu quả"
                ],
                "right_content": [
                    "Nhược điểm:",
                    "• Phức tạp",
                    "• Cần thời gian",
                    "• Yêu cầu kiên nhẫn"
                ]
            },
            {
                "type": "chart",
                "title": "Thống kê kết quả",
                "chart_data": {
                    "type": "column",
                    "categories": ["Q1", "Q2", "Q3", "Q4"],
                    "series": {
                        "Điểm số": [7.5, 8.0, 8.5, 9.0],
                        "Tham gia": [85, 90, 95, 98]
                    }
                }
            }
        ],
        "add_conclusion": True,
        "conclusion_points": [
            "Đã hoàn thành mục tiêu bài học",
            "Học sinh tích cực tham gia",
            "Kết quả đạt được như mong đợi"
        ]
    }
    
    generator = PowerPointGenerator()
    generator.create_from_structured_data(sample_data)
    return generator


# Test the module
if __name__ == "__main__":
    # Test tạo presentation mẫu
    print("Testing PowerPoint Generator...")
    
    generator = create_sample_presentation()
    
    if generator.save_to_file("test_presentation.pptx"):
        print("✅ Test successful! Created test_presentation.pptx")
        print(f"Slides created: {generator.get_slide_count()}")
    else:
        print("❌ Test failed!")