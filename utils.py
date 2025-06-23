# utils.py
"""
Utility functions cho AI PowerPoint Generator
"""

import json
import re
import os
import logging
import time                                 
import hashlib
from datetime import datetime, timedelta
from typing import Dict, List, Any, Optional, Tuple
from pathlib import Path
import streamlit as st
from functools import wraps
import traceback

# Setup logging
logger = logging.getLogger(__name__)

class Cache:
    """Simple in-memory cache for API responses"""
    
    def __init__(self, ttl: int = 3600):
        self.cache = {}
        self.timestamps = {}
        self.ttl = ttl  # Time to live in seconds
    
    def get(self, key: str) -> Optional[Any]:
        """Get value from cache"""
        if key in self.cache:
            if time.time() - self.timestamps[key] < self.ttl:
                return self.cache[key]
            else:
                # Expired, remove from cache
                del self.cache[key]
                del self.timestamps[key]
        return None
    
    def set(self, key: str, value: Any) -> None:
        """Set value in cache"""
        self.cache[key] = value
        self.timestamps[key] = time.time()
    
    def clear(self) -> None:
        """Clear all cache"""
        self.cache.clear()
        self.timestamps.clear()
    
    def size(self) -> int:
        """Get cache size"""
        return len(self.cache)

# Global cache instance
cache = Cache()

def timing_decorator(func):
    """Decorator to measure function execution time"""
    @wraps(func)
    def wrapper(*args, **kwargs):
        start_time = time.time()
        try:
            result = func(*args, **kwargs)
            end_time = time.time()
            logger.info(f"{func.__name__} executed in {end_time - start_time:.2f} seconds")
            return result
        except Exception as e:
            end_time = time.time()
            logger.error(f"{func.__name__} failed after {end_time - start_time:.2f} seconds: {str(e)}")
            raise
    return wrapper

def error_handler(func):
    """Decorator for error handling"""
    @wraps(func)
    def wrapper(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        except Exception as e:
            logger.error(f"Error in {func.__name__}: {str(e)}")
            logger.error(traceback.format_exc())
            # Re-raise the exception for the caller to handle
            raise
    return wrapper

def create_hash(text: str) -> str:
    """Create hash for caching purposes"""
    return hashlib.md5(text.encode()).hexdigest()

def clean_text(text: str) -> str:
    """Clean and normalize text"""
    if not text:
        return ""
    
    # Remove extra whitespace
    text = re.sub(r'\s+', ' ', text.strip())
    
    # Remove special characters that might cause issues
    text = re.sub(r'[^\w\s\-.,!?()áàảãạăắằẳẵặâấầẩẫậéèẻẽẹêếềểễệíìỉĩịóòỏõọôốồổỗộơớờởỡợúùủũụưứừửữựýỳỷỹỵđĐ]', '', text)
    
    return text

def extract_bullet_points(text: str) -> List[str]:
    """Extract bullet points from text"""
    if not text:
        return []
    
    # Split by common bullet markers
    lines = text.split('\n')
    bullet_points = []
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
        
        # Remove bullet markers
        line = re.sub(r'^[•\-\*\d+\.]\s*', '', line)
        
        if line:
            bullet_points.append(line)
    
    return bullet_points

def format_bullet_points(points: List[str], max_length: int = 200) -> List[str]:
    """Format bullet points with length limits"""
    formatted_points = []
    
    for point in points:
        point = clean_text(point)
        
        if len(point) > max_length:
            # Truncate at word boundary
            truncated = point[:max_length]
            last_space = truncated.rfind(' ')
            if last_space > max_length * 0.8:  # If we're not cutting too much
                point = truncated[:last_space] + "..."
            else:
                point = truncated + "..."
        
        formatted_points.append(point)
    
    return formatted_points

def validate_json_structure(data: Dict[str, Any], required_fields: List[str]) -> Tuple[bool, str]:
    """Validate JSON structure"""
    try:
        for field in required_fields:
            if field not in data:
                return False, f"Missing required field: {field}"
        
        return True, "Valid"
    
    except Exception as e:
        return False, f"Validation error: {str(e)}"

def safe_json_loads(json_string: str) -> Tuple[Optional[Dict], str]:
    """Safely parse JSON string"""
    try:
        data = json.loads(json_string)
        return data, "Success"
    except json.JSONDecodeError as e:
        return None, f"JSON decode error: {str(e)}"
    except Exception as e:
        return None, f"Unexpected error: {str(e)}"

def safe_json_dumps(data: Any, indent: int = 2) -> str:
    """Safely serialize to JSON"""
    try:
        return json.dumps(data, indent=indent, ensure_ascii=False)
    except Exception as e:
        logger.error(f"JSON serialization error: {str(e)}")
        return "{}"

def create_directories(directories: List[str]) -> None:
    """Create directories if they don't exist"""
    for directory in directories:
        Path(directory).mkdir(parents=True, exist_ok=True)

def get_file_size(file_path: str) -> int:
    """Get file size in bytes"""
    try:
        return os.path.getsize(file_path)
    except Exception:
        return 0

def format_file_size(size_bytes: int) -> str:
    """Format file size for display"""
    if size_bytes == 0:
        return "0 B"
    
    size_names = ["B", "KB", "MB", "GB"]
    i = 0
    
    while size_bytes >= 1024 and i < len(size_names) - 1:
        size_bytes /= 1024
        i += 1
    
    return f"{size_bytes:.1f} {size_names[i]}"

def generate_filename(base_name: str, extension: str = ".pptx", include_timestamp: bool = True) -> str:
    """Generate safe filename"""
    # Clean base name
    safe_name = re.sub(r'[^\w\s\-]', '', base_name)
    safe_name = re.sub(r'\s+', '_', safe_name)
    
    if include_timestamp:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"{safe_name}_{timestamp}{extension}"
    else:
        filename = f"{safe_name}{extension}"
    
    return filename

def truncate_text(text: str, max_length: int = 100, suffix: str = "...") -> str:
    """Truncate text with suffix"""
    if len(text) <= max_length:
        return text
    
    return text[:max_length - len(suffix)] + suffix

def extract_keywords(text: str, max_keywords: int = 5) -> List[str]:
    """Extract keywords from text (simple implementation)"""
    if not text:
        return []
    
    # Remove common Vietnamese stop words
    stop_words = {
        'là', 'và', 'của', 'có', 'được', 'trong', 'với', 'để', 'về', 'từ',
        'theo', 'như', 'sẽ', 'đã', 'cho', 'khi', 'mà', 'này', 'đó', 'những',
        'các', 'một', 'hai', 'ba', 'bốn', 'năm', 'sáu', 'bảy', 'tám', 'chín',
        'mười', 'the', 'and', 'or', 'but', 'in', 'on', 'at', 'to', 'for'
    }
    
    # Clean and split text
    words = re.findall(r'\w+', text.lower())
    
    # Filter out stop words and short words
    keywords = [word for word in words if len(word) > 3 and word not in stop_words]
    
    # Count frequency
    word_freq = {}
    for word in keywords:
        word_freq[word] = word_freq.get(word, 0) + 1
    
    # Sort by frequency and return top keywords
    sorted_words = sorted(word_freq.items(), key=lambda x: x[1], reverse=True)
    
    return [word for word, freq in sorted_words[:max_keywords]]

def detect_language(text: str) -> str:
    """Simple language detection for Vietnamese vs English"""
    if not text:
        return "unknown"
    
    # Vietnamese-specific characters
    vietnamese_chars = 'áàảãạăắằẳẵặâấầẩẫậéèẻẽẹêếềểễệíìỉĩịóòỏõọôốồổỗộơớờởỡợúùủũụưứừửữựýỳỷỹỵđĐ'
    
    vietnamese_count = sum(1 for char in text if char in vietnamese_chars)
    total_chars = len([char for char in text if char.isalpha()])
    
    if total_chars == 0:
        return "unknown"
    
    vietnamese_ratio = vietnamese_count / total_chars
    
    if vietnamese_ratio > 0.1:  # 10% Vietnamese characters indicates Vietnamese text
        return "vietnamese"
    else:
        return "english"

def format_duration(seconds: int) -> str:
    """Format duration in human readable format"""
    if seconds < 60:
        return f"{seconds} giây"
    elif seconds < 3600:
        minutes = seconds // 60
        remaining_seconds = seconds % 60
        if remaining_seconds == 0:
            return f"{minutes} phút"
        else:
            return f"{minutes} phút {remaining_seconds} giây"
    else:
        hours = seconds // 3600
        minutes = (seconds % 3600) // 60
        return f"{hours} giờ {minutes} phút"

def estimate_reading_time(text: str, words_per_minute: int = 200) -> int:
    """Estimate reading time in minutes"""
    if not text:
        return 0
    
    word_count = len(text.split())
    minutes = max(1, word_count // words_per_minute)
    
    return minutes

def create_progress_callback(total_steps: int, description: str = "Processing"):
    """Create progress callback for long operations"""
    
    def callback(current_step: int, step_description: str = ""):
        progress = current_step / total_steps
        
        if 'progress_bar' not in st.session_state:
            st.session_state.progress_bar = st.progress(0)
            st.session_state.progress_text = st.empty()
        
        st.session_state.progress_bar.progress(progress)
        
        if step_description:
            st.session_state.progress_text.text(f"{description}: {step_description}")
        else:
            st.session_state.progress_text.text(f"{description}: {current_step}/{total_steps}")
    
    return callback

def log_user_action(action: str, details: Dict[str, Any] = None):
    """Log user actions for analytics"""
    log_entry = {
        "timestamp": datetime.now().isoformat(),
        "action": action,
        "details": details or {},
        "session_id": st.session_state.get("session_id", "unknown")
    }
    
    logger.info(f"User action: {json.dumps(log_entry, ensure_ascii=False)}")

def handle_api_error(error: Exception) -> str:
    """Handle API errors and return user-friendly message"""
    error_str = str(error).lower()
    
    if "rate limit" in error_str or "quota" in error_str:
        return "🚫 Đã vượt quá giới hạn API. Vui lòng thử lại sau ít phút."
    elif "invalid api key" in error_str or "authentication" in error_str:
        return "🔑 API key không hợp lệ. Vui lòng kiểm tra lại API key."
    elif "timeout" in error_str:
        return "⏱️ Yêu cầu mất quá nhiều thời gian. Vui lòng thử lại."
    elif "network" in error_str or "connection" in error_str:
        return "🌐 Lỗi kết nối mạng. Vui lòng kiểm tra internet."
    else:
        return f"❌ Lỗi API: {str(error)}"

def backup_session_state(key: str = "backup"):
    """Backup current session state"""
    backup_data = {}
    
    for k, v in st.session_state.items():
        try:
            # Only backup serializable data
            json.dumps(v)
            backup_data[k] = v
        except (TypeError, ValueError):
            # Skip non-serializable objects
            continue
    
    st.session_state[f"_{key}_backup"] = backup_data
    st.session_state[f"_{key}_backup_time"] = datetime.now()

def restore_session_state(key: str = "backup") -> bool:
    """Restore session state from backup"""
    backup_key = f"_{key}_backup"
    
    if backup_key in st.session_state:
        backup_data = st.session_state[backup_key]
        
        for k, v in backup_data.items():
            st.session_state[k] = v
        
        return True
    
    return False

def get_memory_usage() -> Dict[str, float]:
    """Get current memory usage"""
    import psutil
    import os
    
    try:
        process = psutil.Process(os.getpid())
        memory_info = process.memory_info()
        
        return {
            "rss_mb": memory_info.rss / 1024 / 1024,  # Resident Set Size
            "vms_mb": memory_info.vms / 1024 / 1024,  # Virtual Memory Size
            "percent": process.memory_percent()
        }
    except ImportError:
        return {"error": "psutil not installed"}
    except Exception as e:
        return {"error": str(e)}

def cleanup_temp_files(temp_dir: str = "temp", max_age_hours: int = 24):
    """Clean up temporary files older than specified hours"""
    if not os.path.exists(temp_dir):
        return
    
    cutoff_time = datetime.now() - timedelta(hours=max_age_hours)
    
    for filename in os.listdir(temp_dir):
        file_path = os.path.join(temp_dir, filename)
        
        try:
            file_time = datetime.fromtimestamp(os.path.getmtime(file_path))
            
            if file_time < cutoff_time:
                os.remove(file_path)
                logger.info(f"Removed old temp file: {filename}")
        
        except Exception as e:
            logger.warning(f"Could not remove temp file {filename}: {str(e)}")

# Streamlit-specific utilities
def show_success_message(message: str, duration: int = 3):
    """Show success message that auto-disappears"""
    placeholder = st.empty()
    placeholder.success(message)
    time.sleep(duration)
    placeholder.empty()

def show_error_message(message: str, duration: int = 5):
    """Show error message that auto-disappears"""
    placeholder = st.empty()
    placeholder.error(message)
    time.sleep(duration)
    placeholder.empty()

def create_download_link(data: bytes, filename: str, mime_type: str, button_text: str = "Download"):
    """Create download link for binary data"""
    import base64
    
    b64 = base64.b64encode(data).decode()
    href = f'<a href="data:{mime_type};base64,{b64}" download="{filename}">{button_text}</a>'
    
    return href

def format_vietnamese_currency(amount: float) -> str:
    """Format currency in Vietnamese format"""
    if amount >= 1_000_000_000:
        return f"{amount/1_000_000_000:.1f} tỷ VNĐ"
    elif amount >= 1_000_000:
        return f"{amount/1_000_000:.1f} triệu VNĐ"
    elif amount >= 1_000:
        return f"{amount/1_000:.1f}K VNĐ"
    else:
        return f"{amount:,.0f} VNĐ"

# Export main functions
__all__ = [
    'Cache', 'cache',
    'timing_decorator', 'error_handler',
    'create_hash', 'clean_text', 'extract_bullet_points', 'format_bullet_points',
    'validate_json_structure', 'safe_json_loads', 'safe_json_dumps',
    'create_directories', 'get_file_size', 'format_file_size', 'generate_filename',
    'truncate_text', 'extract_keywords', 'detect_language', 'format_duration',
    'estimate_reading_time', 'create_progress_callback', 'log_user_action',
    'handle_api_error', 'backup_session_state', 'restore_session_state',
    'get_memory_usage', 'cleanup_temp_files', 'show_success_message',
    'show_error_message', 'create_download_link', 'format_vietnamese_currency'
]