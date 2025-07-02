# theme_system.py - Hệ thống Theme Đẹp
"""
Hệ thống theme và layout hiện đại cho AI PowerPoint Generator
Inspired by professional presentation designs
"""

from typing import Dict, Any
from pptx.dml.color import RGBColor
from pptx.util import Pt, Inches

class ModernThemeSystem:
    """Hệ thống theme hiện đại với màu sắc và layout đẹp"""
    
    def __init__(self):
        self.themes = self._init_themes()
        self.icons = self._init_icons()
        self.layouts = self._init_layouts()
        
    def _init_themes(self) -> Dict[str, Dict[str, Any]]:
        """Khởi tạo các theme hiện đại"""
        return {
            "python_modern": {
                "name": "Python Modern",
                "description": "Theme hiện đại cho lập trình Python",
                "colors": {
                    "primary": "#3776AB",      # Python Blue
                    "secondary": "#FFD43B",    # Python Yellow  
                    "accent": "#306998",       # Dark Python Blue
                    "background": "#F8F9FA",   # Light Gray
                    "text": "#2C3E50",         # Dark Blue Gray
                    "success": "#27AE60",      # Green
                    "warning": "#F39C12",      # Orange
                    "error": "#E74C3C"         # Red
                },
                "gradients": {
                    "primary": "linear-gradient(135deg, #3776AB 0%, #306998 100%)",
                    "secondary": "linear-gradient(135deg, #FFD43B 0%, #F39C12 100%)",
                    "background": "linear-gradient(135deg, #F8F9FA 0%, #E9ECEF 100%)"
                },
                "fonts": {
                    "title": {"size": 36, "weight": "bold", "family": "Segoe UI"},
                    "subtitle": {"size": 24, "weight": "semibold", "family": "Segoe UI"}, 
                    "content": {"size": 18, "weight": "normal", "family": "Segoe UI"},
                    "caption": {"size": 14, "weight": "light", "family": "Segoe UI"}
                }
            },
            
            "tech_gradient": {
                "name": "Tech Gradient", 
                "description": "Theme gradient hiện đại cho công nghệ",
                "colors": {
                    "primary": "#667EEA",      # Blue Purple
                    "secondary": "#764BA2",    # Purple
                    "accent": "#F093FB",       # Pink Purple
                    "background": "#FFFFFF",   # White
                    "text": "#2D3748",         # Dark Gray
                    "success": "#48BB78",      # Green
                    "warning": "#ED8936",      # Orange
                    "error": "#F56565"         # Red
                },
                "gradients": {
                    "primary": "linear-gradient(135deg, #667EEA 0%, #764BA2 100%)",
                    "secondary": "linear-gradient(135deg, #F093FB 0%, #F5576C 100%)",
                    "background": "linear-gradient(135deg, #FDFBFB 0%, #EBEDEE 100%)"
                },
                "fonts": {
                    "title": {"size": 38, "weight": "bold", "family": "Arial"},
                    "subtitle": {"size": 26, "weight": "semibold", "family": "Arial"},
                    "content": {"size": 20, "weight": "normal", "family": "Arial"},
                    "caption": {"size": 16, "weight": "light", "family": "Arial"}
                }
            },
            
            "education_pro": {
                "name": "Education Pro",
                "description": "Theme chuyên nghiệp cho giáo dục",
                "colors": {
                    "primary": "#2E86AB",      # Ocean Blue
                    "secondary": "#A23B72",    # Magenta
                    "accent": "#F18F01",       # Orange
                    "background": "#FAFAFA",   # Light Gray
                    "text": "#1A202C",         # Very Dark Gray
                    "success": "#38A169",      # Green
                    "warning": "#D69E2E",      # Yellow
                    "error": "#E53E3E"         # Red
                },
                "gradients": {
                    "primary": "linear-gradient(135deg, #2E86AB 0%, #A23B72 100%)",
                    "secondary": "linear-gradient(135deg, #F18F01 0%, #F39C12 100%)",
                    "background": "linear-gradient(135deg, #FAFAFA 0%, #F7FAFC 100%)"
                },
                "fonts": {
                    "title": {"size": 34, "weight": "bold", "family": "Calibri"},
                    "subtitle": {"size": 24, "weight": "semibold", "family": "Calibri"},
                    "content": {"size": 18, "weight": "normal", "family": "Calibri"},
                    "caption": {"size": 14, "weight": "light", "family": "Calibri"}
                }
            },
            
            "business_elegant": {
                "name": "Business Elegant",
                "description": "Theme thanh lịch cho doanh nghiệp", 
                "colors": {
                    "primary": "#1565C0",      # Deep Blue
                    "secondary": "#FF7043",    # Deep Orange
                    "accent": "#26A69A",       # Teal
                    "background": "#FFFFFF",   # Pure White
                    "text": "#263238",         # Blue Gray
                    "success": "#4CAF50",      # Green
                    "warning": "#FF9800",      # Orange  
                    "error": "#F44336"         # Red
                },
                "gradients": {
                    "primary": "linear-gradient(135deg, #1565C0 0%, #1976D2 100%)",
                    "secondary": "linear-gradient(135deg, #FF7043 0%, #FF5722 100%)",
                    "background": "linear-gradient(135deg, #FFFFFF 0%, #F5F5F5 100%)"
                },
                "fonts": {
                    "title": {"size": 36, "weight": "bold", "family": "Times New Roman"},
                    "subtitle": {"size": 26, "weight": "semibold", "family": "Times New Roman"},
                    "content": {"size": 20, "weight": "normal", "family": "Times New Roman"},
                    "caption": {"size": 16, "weight": "light", "family": "Times New Roman"}
                }
            },
            
            "creative_vibrant": {
                "name": "Creative Vibrant",
                "description": "Theme sáng tạo với màu sắc sinh động",
                "colors": {
                    "primary": "#E91E63",      # Pink
                    "secondary": "#9C27B0",    # Purple
                    "accent": "#00BCD4",       # Cyan
                    "background": "#FAFAFA",   # Light Gray
                    "text": "#212121",         # Dark Gray
                    "success": "#4CAF50",      # Green
                    "warning": "#FF9800",      # Orange
                    "error": "#F44336"         # Red
                },
                "gradients": {
                    "primary": "linear-gradient(135deg, #E91E63 0%, #9C27B0 100%)",
                    "secondary": "linear-gradient(135deg, #00BCD4 0%, #4CAF50 100%)",
                    "background": "linear-gradient(135deg, #FAFAFA 0%, #F0F0F0 100%)"
                },
                "fonts": {
                    "title": {"size": 40, "weight": "bold", "family": "Comic Sans MS"},
                    "subtitle": {"size": 28, "weight": "semibold", "family": "Comic Sans MS"},
                    "content": {"size": 22, "weight": "normal", "family": "Comic Sans MS"},
                    "caption": {"size": 18, "weight": "light", "family": "Comic Sans MS"}
                }
            }
        }
    
    def _init_icons(self) -> Dict[str, str]:
        """Khởi tạo bộ icon Unicode hiện đại"""
        return {
            # Education Icons
            "education": "🎓",
            "book": "📚", 
            "study": "📖",
            "learn": "🧠",
            "teacher": "👨‍🏫",
            "student": "👨‍🎓",
            "school": "🏫",
            "knowledge": "💡",
            
            # Technology Icons
            "python": "🐍",
            "code": "💻",
            "programming": "⌨️",
            "ai": "🤖",
            "data": "📊",
            "analysis": "📈",
            "algorithm": "⚙️", 
            "tech": "🔧",
            
            # Business Icons
            "business": "💼",
            "presentation": "📋",
            "meeting": "🤝",
            "strategy": "🎯",
            "growth": "📈",
            "success": "🏆",
            "team": "👥",
            "project": "📂",
            
            # Science Icons
            "biology": "🧬",
            "chemistry": "⚗️",
            "physics": "⚛️", 
            "math": "📐",
            "lab": "🔬",
            "experiment": "🧪",
            "research": "🔍",
            "discovery": "🌟",
            
            # UI Icons
            "arrow_right": "➡️",
            "arrow_down": "⬇️",
            "check": "✅",
            "cross": "❌",
            "star": "⭐",
            "heart": "❤️",
            "fire": "🔥",
            "rocket": "🚀",
            "sparkles": "✨",
            "magic": "🪄",
            
            # Status Icons
            "info": "ℹ️",
            "warning": "⚠️",
            "error": "🚫",
            "success": "🎉",
            "loading": "⏳",
            "complete": "✅"
        }
    
    def _init_layouts(self) -> Dict[str, Dict[str, Any]]:
        """Khởi tạo các layout template hiện đại"""
        return {
            "hero_layout": {
                "name": "Hero Layout",
                "description": "Layout hero cho slide đầu",
                "structure": {
                    "title_area": {"x": 0.5, "y": 2.0, "width": 9, "height": 2},
                    "subtitle_area": {"x": 0.5, "y": 4.5, "width": 9, "height": 1},
                    "image_area": {"x": 1, "y": 6, "width": 8, "height": 3},
                    "author_area": {"x": 7, "y": 9.5, "width": 2.5, "height": 0.5}
                }
            },
            
            "two_column_modern": {
                "name": "Two Column Modern",
                "description": "Layout hai cột hiện đại",
                "structure": {
                    "title_area": {"x": 0.5, "y": 0.3, "width": 9, "height": 1},
                    "left_column": {"x": 0.5, "y": 1.5, "width": 4.3, "height": 6},
                    "right_column": {"x": 5.2, "y": 1.5, "width": 4.3, "height": 6},
                    "footer_area": {"x": 0.5, "y": 8, "width": 9, "height": 0.5}
                }
            },
            
            "image_focus": {
                "name": "Image Focus",
                "description": "Layout tập trung vào ảnh",
                "structure": {
                    "title_area": {"x": 0.5, "y": 0.3, "width": 9, "height": 0.8},
                    "image_area": {"x": 0.5, "y": 1.3, "width": 6, "height": 4.5},
                    "content_area": {"x": 6.8, "y": 1.3, "width": 2.7, "height": 4.5},
                    "caption_area": {"x": 0.5, "y": 6, "width": 6, "height": 0.8}
                }
            },
            
            "content_rich": {
                "name": "Content Rich", 
                "description": "Layout nhiều nội dung",
                "structure": {
                    "title_area": {"x": 0.5, "y": 0.3, "width": 9, "height": 0.8},
                    "main_content": {"x": 0.5, "y": 1.3, "width": 9, "height": 5.5},
                    "sidebar": {"x": 7.5, "y": 1.3, "width": 2, "height": 5.5}
                }
            }
        }
    
    def get_theme(self, theme_name: str) -> Dict[str, Any]:
        """Lấy theme theo tên"""
        return self.themes.get(theme_name, self.themes["tech_gradient"])
    
    def get_icon(self, icon_name: str) -> str:
        """Lấy icon theo tên"""
        return self.icons.get(icon_name, "📄")
    
    def get_layout(self, layout_name: str) -> Dict[str, Any]:
        """Lấy layout theo tên"""
        return self.layouts.get(layout_name, self.layouts["two_column_modern"])
    
    def detect_theme_from_content(self, content: str) -> str:
        """Tự động detect theme phù hợp dựa trên nội dung"""
        content_lower = content.lower()
        
        # Python/Programming themes
        if any(word in content_lower for word in ['python', 'programming', 'code', 'lập trình']):
            return "python_modern"
        
        # Business themes  
        elif any(word in content_lower for word in ['business', 'doanh nghiệp', 'marketing', 'kinh doanh']):
            return "business_elegant"
            
        # Education themes
        elif any(word in content_lower for word in ['học', 'giáo dục', 'bài giảng', 'sinh viên']):
            return "education_pro"
            
        # Creative themes
        elif any(word in content_lower for word in ['sáng tạo', 'creative', 'design', 'nghệ thuật']):
            return "creative_vibrant"
            
        # Default modern theme
        else:
            return "tech_gradient"
    
    def get_subject_icon(self, subject: str) -> str:
        """Lấy icon phù hợp cho môn học"""
        subject_lower = subject.lower()
        
        if 'sinh học' in subject_lower or 'biology' in subject_lower:
            return self.get_icon("biology")
        elif 'vật lý' in subject_lower or 'physics' in subject_lower:
            return self.get_icon("physics")
        elif 'hóa học' in subject_lower or 'chemistry' in subject_lower:
            return self.get_icon("chemistry")
        elif 'toán' in subject_lower or 'math' in subject_lower:
            return self.get_icon("math")
        elif 'python' in subject_lower or 'lập trình' in subject_lower:
            return self.get_icon("python")
        elif 'marketing' in subject_lower or 'kinh doanh' in subject_lower:
            return self.get_icon("business")
        else:
            return self.get_icon("education")
    
    def create_color_palette(self, theme_name: str) -> Dict[str, RGBColor]:
        """Tạo color palette cho PowerPoint"""
        theme = self.get_theme(theme_name)
        colors = theme["colors"]
        
        palette = {}
        for color_name, hex_color in colors.items():
            # Convert hex to RGB
            hex_color = hex_color.lstrip('#')
            rgb = tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
            palette[color_name] = RGBColor(rgb[0], rgb[1], rgb[2])
            
        return palette
    
    def get_font_config(self, theme_name: str, font_type: str) -> Dict[str, Any]:
        """Lấy cấu hình font"""
        theme = self.get_theme(theme_name)
        return theme["fonts"].get(font_type, theme["fonts"]["content"])
    
    def list_available_themes(self) -> Dict[str, str]:
        """Liệt kê các theme có sẵn"""
        return {name: info["description"] for name, info in self.themes.items()}

# Khởi tạo hệ thống theme global
theme_system = ModernThemeSystem() 