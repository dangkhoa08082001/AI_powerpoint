# config.py - Updated với .env support
"""
Configuration file cho AI PowerPoint Generator
"""

import os
from typing import Dict, Any

# Load .env file
try:
    from dotenv import load_dotenv
    load_dotenv()
    print("✅ Loaded .env file")
except ImportError:
    print("⚠️ python-dotenv not installed, .env file will not be loaded")
except Exception as e:
    print(f"⚠️ Could not load .env file: {e}")

class Config:
    """Main configuration class"""
    
    # API Settings
    OPENAI_MODEL = "gpt-3.5-turbo"
    OPENAI_MAX_TOKENS = 3000
    OPENAI_TEMPERATURE = 0.7
    
    # Application Settings
    APP_TITLE = "🎓 AI PowerPoint Generator"
    APP_ICON = "🎓"
    DEFAULT_TEMPLATE = "education"
    
    # Streamlit Settings
    STREAMLIT_LAYOUT = "wide"
    STREAMLIT_SIDEBAR = "expanded"
    
    # PowerPoint Settings
    DEFAULT_SLIDE_COUNT = 8
    MAX_SLIDES_PER_PRESENTATION = 30
    MIN_SLIDES_PER_PRESENTATION = 3
    
    # Content Settings
    MAX_BULLET_POINTS = 6
    MIN_BULLET_POINTS = 2
    MAX_CONTENT_LENGTH = 200  # characters per bullet point
    
    # File Settings
    UPLOAD_FOLDER = "uploads"
    OUTPUT_FOLDER = "outputs"
    TEMP_FOLDER = "temp"
    MAX_FILE_SIZE = 10 * 1024 * 1024  # 10MB
    
    # Template Settings
    TEMPLATES = {
        "education": {
            "name": "Giáo Dục",
            "description": "Template cho bài giảng và giáo án",
            "primary_color": "#2E86AB",
            "secondary_color": "#A23B72", 
            "background_color": "#F18F01",
            "text_color": "#0F0F0F",
            "font_family": "Calibri",
            "font_sizes": {
                "title": 32,
                "subtitle": 24,
                "content": 18,
                "caption": 14
            }
        },
        "business": {
            "name": "Doanh Nghiệp",
            "description": "Template cho presentation doanh nghiệp",
            "primary_color": "#1565C0",
            "secondary_color": "#FFA726",
            "background_color": "#E3F2FD",
            "text_color": "#263238",
            "font_family": "Arial",
            "font_sizes": {
                "title": 36,
                "subtitle": 28,
                "content": 20,
                "caption": 16
            }
        },
        "training": {
            "name": "Đào Tạo",
            "description": "Template cho khóa học và workshop",
            "primary_color": "#7B1FA2",
            "secondary_color": "#FF7043",
            "background_color": "#F3E5F5",
            "text_color": "#424242",
            "font_family": "Segoe UI",
            "font_sizes": {
                "title": 34,
                "subtitle": 26,
                "content": 19,
                "caption": 15
            }
        }
    }
    
    # AI Prompts Templates
    SYSTEM_PROMPTS = {
        "education": """Bạn là một chuyên gia giáo dục với 20 năm kinh nghiệm trong việc thiết kế bài giảng và giáo án. 
        Bạn có khả năng tạo ra những bài giảng PowerPoint chất lượng cao, phù hợp với từng cấp học và môn học.
        Hãy tập trung vào:
        - Nội dung phù hợp với độ tuổi
        - Hoạt động tương tác
        - Ví dụ thực tế
        - Đánh giá học sinh""",
        
        "business": """Bạn là một chuyên gia tư vấn doanh nghiệp với kinh nghiệm sâu về thuyết trình và trình bày. 
        Bạn có thể tạo ra những presentation chuyên nghiệp cho môi trường công sở.
        Hãy tập trung vào:
        - Nội dung súc tích, rõ ràng
        - Dữ liệu và số liệu
        - Call-to-action
        - ROI và business impact""",
        
        "training": """Bạn là một chuyên gia đào tạo với khả năng thiết kế các khóa học và bài training hiệu quả.
        Bạn biết cách truyền đạt kiến thức một cách sinh động và dễ hiểu.
        Hãy tập trung vào:
        - Learning objectives rõ ràng
        - Hoạt động thực hành
        - Case studies
        - Assessment và feedback"""
    }
    
    # Error Messages
    ERROR_MESSAGES = {
        "api_key_missing": "❌ Vui lòng nhập OpenAI API key để sử dụng AI",
        "api_key_invalid": "❌ API key không hợp lệ. Vui lòng kiểm tra lại",
        "generation_failed": "❌ Không thể tạo presentation. Vui lòng thử lại",
        "save_failed": "❌ Không thể lưu file. Kiểm tra quyền ghi",
        "load_failed": "❌ Không thể tải file. Kiểm tra định dạng",
        "network_error": "❌ Lỗi kết nối mạng. Vui lòng kiểm tra internet",
        "quota_exceeded": "❌ Đã vượt quá giới hạn API. Vui lòng thử lại sau"
    }
    
    # Success Messages
    SUCCESS_MESSAGES = {
        "presentation_created": "✅ Đã tạo presentation thành công!",
        "slide_updated": "✅ Đã cập nhật slide!",
        "file_saved": "✅ Đã lưu file thành công!",
        "file_downloaded": "✅ File đã sẵn sàng để tải xuống!",
        "settings_updated": "✅ Đã cập nhật cài đặt!"
    }
    
    # Warning Messages  
    WARNING_MESSAGES = {
        "no_content": "⚠️ Chưa có nội dung để xử lý",
        "large_file": "⚠️ File có thể lớn, quá trình tạo sẽ mất thời gian",
        "many_slides": "⚠️ Presentation có nhiều slides, có thể ảnh hưởng hiệu suất",
        "beta_feature": "⚠️ Đây là tính năng beta, có thể chưa ổn định"
    }
    
    # UI Settings
    UI_SETTINGS = {
        "primary_color": "#667eea",
        "secondary_color": "#764ba2", 
        "success_color": "#28a745",
        "warning_color": "#ffc107",
        "error_color": "#dc3545",
        "info_color": "#17a2b8"
    }
    
    # Performance Settings
    PERFORMANCE = {
        "max_concurrent_requests": 5,
        "request_timeout": 30,  # seconds
        "cache_ttl": 3600,  # 1 hour
        "max_memory_usage": 500,  # MB
        "chunk_size": 1024,  # bytes
    }
    
    # Logging Settings
    LOGGING = {
        "level": "INFO",
        "format": "%(asctime)s - %(name)s - %(levelname)s - %(message)s",
        "file": "app.log",
        "max_size": 10 * 1024 * 1024,  # 10MB
        "backup_count": 5
    }
    
    # Security Settings
    SECURITY = {
        "allowed_file_types": [".txt", ".json", ".csv"],
        "max_input_length": 5000,
        "sanitize_input": True,
        "rate_limit": 60,  # requests per minute
    }

class DevelopmentConfig(Config):
    """Development environment configuration"""
    DEBUG = True
    OPENAI_TEMPERATURE = 0.8  # More creative for testing
    CACHE_TTL = 60  # Shorter cache for development
    LOG_LEVEL = "DEBUG"

class ProductionConfig(Config):
    """Production environment configuration"""
    DEBUG = False
    OPENAI_TEMPERATURE = 0.7  # More stable for production
    CACHE_TTL = 3600  # Longer cache for production
    LOG_LEVEL = "INFO"
    
    # Production security
    SECURITY = {
        **Config.SECURITY,
        "rate_limit": 30,  # Stricter rate limiting
        "max_input_length": 3000,  # Shorter input for safety
    }

class TestConfig(Config):
    """Test environment configuration"""
    DEBUG = True
    OPENAI_MODEL = "gpt-3.5-turbo"  # Cheaper model for testing
    OPENAI_MAX_TOKENS = 1000  # Reduced for testing
    CACHE_TTL = 0  # No cache for testing

# Environment-based config selection
def get_config() -> Config:
    """Get configuration based on environment"""
    env = os.getenv("ENVIRONMENT", "development").lower()
    
    if env == "production":
        return ProductionConfig()
    elif env == "test":
        return TestConfig()
    else:
        return DevelopmentConfig()

# Global config instance
config = get_config()

# Helper functions
def get_template_config(template_name: str) -> Dict[str, Any]:
    """Get template configuration by name"""
    return config.TEMPLATES.get(template_name, config.TEMPLATES["education"])

def get_system_prompt(template_type: str) -> str:
    """Get system prompt for template type"""
    return config.SYSTEM_PROMPTS.get(template_type, config.SYSTEM_PROMPTS["education"])

def validate_slide_count(count: int) -> bool:
    """Validate slide count within limits"""
    return config.MIN_SLIDES_PER_PRESENTATION <= count <= config.MAX_SLIDES_PER_PRESENTATION

def validate_content_length(content: str) -> bool:
    """Validate content length"""
    return len(content) <= config.MAX_CONTENT_LENGTH

def get_error_message(error_type: str) -> str:
    """Get error message by type"""
    return config.ERROR_MESSAGES.get(error_type, "❌ Đã có lỗi xảy ra")

def get_success_message(success_type: str) -> str:
    """Get success message by type"""
    return config.SUCCESS_MESSAGES.get(success_type, "✅ Thành công!")

def get_warning_message(warning_type: str) -> str:
    """Get warning message by type"""
    return config.WARNING_MESSAGES.get(warning_type, "⚠️ Cảnh báo!")

# Validation functions
def validate_api_key(api_key: str) -> bool:
    """Basic API key validation"""
    if not api_key:
        return False
    
    # OpenAI API keys start with 'sk-'
    if not api_key.startswith('sk-'):
        return False
    
    # Basic length check - Updated for new format
    if len(api_key) < 50:  # New OpenAI keys are longer
        return False
    
    return True

def get_api_key() -> str:
    """Get API key from environment or .env file"""
    api_key = os.getenv("OPENAI_API_KEY")
    if api_key:
        print(f"✅ Found API key: {api_key[:20]}...")
        return api_key
    else:
        print("❌ No API key found in environment")
        return ""

def sanitize_filename(filename: str) -> str:
    """Sanitize filename for safe file operations"""
    import re
    
    # Remove or replace unsafe characters
    safe_chars = re.compile(r'[^a-zA-Z0-9._\-\s]')
    filename = safe_chars.sub('_', filename)
    
    # Limit length
    max_length = 100
    if len(filename) > max_length:
        name, ext = os.path.splitext(filename)
        filename = name[:max_length-len(ext)] + ext
    
    return filename.strip()

def validate_user_input(user_input: str) -> tuple[bool, str]:
    """Validate user input for security"""
    if not user_input or not user_input.strip():
        return False, "Input rỗng"
    
    if len(user_input) > config.SECURITY["max_input_length"]:
        return False, f"Input quá dài (max {config.SECURITY['max_input_length']} ký tự)"
    
    # Basic XSS protection
    dangerous_patterns = ['<script', 'javascript:', 'onload=', 'onerror=']
    user_input_lower = user_input.lower()
    
    for pattern in dangerous_patterns:
        if pattern in user_input_lower:
            return False, "Input chứa nội dung không an toàn"
    
    return True, "Valid"

# Export main functions
__all__ = [
    'config',
    'get_config',
    'get_template_config', 
    'get_system_prompt',
    'validate_slide_count',
    'validate_content_length',
    'get_error_message',
    'get_success_message',
    'get_warning_message',
    'validate_api_key',
    'get_api_key',
    'sanitize_filename',
    'validate_user_input'
]

# Test API key on import
if __name__ == "__main__":
    api_key = get_api_key()
    if api_key:
        if validate_api_key(api_key):
            print("✅ API key is valid!")
        else:
            print("❌ API key format is invalid!")
    else:
        print("⚠️ No API key found!")