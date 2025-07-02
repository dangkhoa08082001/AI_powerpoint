# dalle_generator.py
"""
Module để tạo ảnh sử dụng DALL-E API cho PowerPoint slides
"""

import openai
import requests
import os
import re
from typing import Dict, List, Optional, Any
import logging
from datetime import datetime
from PIL import Image
import io

logger = logging.getLogger(__name__)

class DALLEImageGenerator:
    """
    Class để tạo ảnh minh họa cho slides sử dụng DALL-E
    """
    
    def __init__(self, api_key: str):
        """
        Khởi tạo DALL-E Image Generator
        
        Args:
            api_key (str): OpenAI API key
        """
        openai.api_key = api_key
        self.dalle_size = "1024x1024"
        self.dalle_quality = "standard"
        self.images_dir = "dalle_images"
        
        # Tạo thư mục images nếu chưa có
        if not os.path.exists(self.images_dir):
            os.makedirs(self.images_dir)
    
    def generate_image_for_slide(self, slide_content: Dict[str, Any], topic: str) -> Optional[str]:
        """
        Tạo ảnh cho một slide cụ thể
        
        Args:
            slide_content (Dict): Nội dung slide
            topic (str): Chủ đề chính của presentation
            
        Returns:
            Optional[str]: Đường dẫn đến file ảnh
        """
        try:
            slide_title = slide_content.get("title", "")
            slide_type = slide_content.get("type", "")
            
            # Chỉ tạo ảnh cho content slides
            if slide_type not in ["content", "two_column"] or not slide_title:
                return None
            
            # Tạo prompt cho ảnh
            prompt = self._create_image_prompt(slide_title, topic)
            
            if not prompt:
                return None
            
            # Generate ảnh với DALL-E
            return self._generate_dalle_image(prompt, slide_title)
            
        except Exception as e:
            logger.error(f"Error generating image for slide: {str(e)}")
            return None
    
    def _create_image_prompt(self, slide_title: str, topic: str) -> str:
        """
        Tạo prompt cho DALL-E dựa trên nội dung slide
        
        Args:
            slide_title (str): Tiêu đề slide
            topic (str): Chủ đề chính
            
        Returns:
            str: Prompt đã tối ưu
        """
        # Keywords cho các môn học
        subject_keywords = {
            "sinh học": ["cell", "DNA", "organism", "biology", "molecular"],
            "vật lý": ["physics", "energy", "light", "force", "wave"],
            "hóa học": ["chemistry", "molecule", "reaction", "atom", "compound"],
            "toán học": ["mathematics", "geometry", "algebra", "equation", "graph"],
            "lịch sử": ["history", "timeline", "culture", "ancient", "civilization"],
            "địa lý": ["geography", "map", "earth", "climate", "landscape"],
            "kinh tế": ["economics", "business", "market", "finance", "chart"],
            "marketing": ["marketing", "business", "digital", "strategy", "growth"]
        }
        
        # Xác định môn học
        subject = self._detect_subject(topic, slide_title)
        
        # Tạo base prompt
        base_prompt = ""
        
        if subject == "sinh học":
            if any(word in slide_title.lower() for word in ["tế bào", "cell"]):
                base_prompt = "detailed biological cell structure, nucleus, organelles, scientific illustration"
            elif any(word in slide_title.lower() for word in ["dna", "gen"]):
                base_prompt = "DNA double helix structure, genetic material, molecular biology"
            elif any(word in slide_title.lower() for word in ["protein", "enzyme"]):
                base_prompt = "protein structure diagram, biochemistry illustration"
            else:
                base_prompt = f"biology concept illustration for {slide_title}"
                
        elif subject == "vật lý":
            if any(word in slide_title.lower() for word in ["quang", "ánh sáng", "light"]):
                base_prompt = "light physics diagram, optical phenomenon, wave properties"
            elif any(word in slide_title.lower() for word in ["điện", "electric"]):
                base_prompt = "electrical circuit diagram, physics illustration"
            elif any(word in slide_title.lower() for word in ["năng lượng", "energy"]):
                base_prompt = "energy transformation diagram, physics concept"
            else:
                base_prompt = f"physics concept illustration for {slide_title}"
                
        elif subject == "hóa học":
            if any(word in slide_title.lower() for word in ["phản ứng", "reaction"]):
                base_prompt = "chemical reaction diagram, molecular interaction"
            elif any(word in slide_title.lower() for word in ["nguyên tử", "atom"]):
                base_prompt = "atomic structure diagram, chemistry illustration"
            else:
                base_prompt = f"chemistry concept illustration for {slide_title}"
                
        elif subject == "toán học":
            if any(word in slide_title.lower() for word in ["hình học", "geometry"]):
                base_prompt = "geometric shapes and theorems, mathematical illustration"
            elif any(word in slide_title.lower() for word in ["đồ thị", "graph"]):
                base_prompt = "mathematical graph and functions, coordinate system"
            else:
                base_prompt = f"mathematics concept illustration for {slide_title}"
                
        elif subject == "marketing":
            if any(word in slide_title.lower() for word in ["digital", "số"]):
                base_prompt = "digital marketing infographic, modern business illustration"
            elif any(word in slide_title.lower() for word in ["strategy", "chiến lược"]):
                base_prompt = "business strategy diagram, marketing concept"
            else:
                base_prompt = f"marketing concept illustration for {slide_title}"
                
        else:
            # Generic educational illustration
            base_prompt = f"educational concept illustration for {slide_title}"
        
        # Thêm style modifiers
        style_modifiers = [
            "professional illustration",
            "clean design", 
            "educational style",
            "no text",
            "no words", 
            "vector art style",
            "modern and clear"
        ]
        
        final_prompt = base_prompt + ", " + ", ".join(style_modifiers)
        
        return final_prompt
    
    def _detect_subject(self, topic: str, slide_title: str) -> str:
        """
        Phát hiện môn học từ topic và slide title
        
        Args:
            topic (str): Chủ đề chính
            slide_title (str): Tiêu đề slide
            
        Returns:
            str: Môn học được phát hiện
        """
        combined_text = (topic + " " + slide_title).lower()
        
        if any(word in combined_text for word in ["sinh học", "biology", "tế bào", "cell", "dna"]):
            return "sinh học"
        elif any(word in combined_text for word in ["vật lý", "physics", "quang", "điện", "năng lượng"]):
            return "vật lý"
        elif any(word in combined_text for word in ["hóa học", "chemistry", "phản ứng", "nguyên tử"]):
            return "hóa học"
        elif any(word in combined_text for word in ["toán", "math", "hình học", "đại số"]):
            return "toán học"
        elif any(word in combined_text for word in ["marketing", "kinh doanh", "business"]):
            return "marketing"
        elif any(word in combined_text for word in ["lịch sử", "history"]):
            return "lịch sử"
        else:
            return "general"
    
    def _generate_dalle_image(self, prompt: str, slide_title: str) -> Optional[str]:
        """
        Gọi DALL-E API để tạo ảnh
        
        Args:
            prompt (str): Prompt cho DALL-E
            slide_title (str): Tiêu đề slide
            
        Returns:
            Optional[str]: Đường dẫn file ảnh
        """
        try:
            logger.info(f"Generating DALL-E image with prompt: {prompt}")
            
            response = openai.Image.create(
                prompt=prompt,
                n=1,
                size=self.dalle_size
            )
            
            if response.data and len(response.data) > 0:
                image_url = response.data[0].url
                
                # Download và lưu ảnh
                return self._download_and_save_image(image_url, slide_title, prompt)
            
        except Exception as e:
            logger.error(f"DALL-E API error: {str(e)}")
            
        return None
    
    def _download_and_save_image(self, image_url: str, slide_title: str, prompt: str) -> Optional[str]:
        """
        Download ảnh từ URL và lưu vào local
        
        Args:
            image_url (str): URL của ảnh từ DALL-E
            slide_title (str): Tiêu đề slide
            prompt (str): Prompt đã sử dụng
            
        Returns:
            Optional[str]: Đường dẫn file đã lưu
        """
        try:
            response = requests.get(image_url, stream=True, timeout=30)
            response.raise_for_status()
            
            # Tạo tên file an toàn
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            safe_title = re.sub(r'[^\w\s-]', '', slide_title)[:30]
            safe_title = re.sub(r'[\s_-]+', '_', safe_title)
            
            if not safe_title:
                safe_title = "slide"
            
            filename = f"dalle_{safe_title}_{timestamp}.png"
            filepath = os.path.join(self.images_dir, filename)
            
            # Lưu ảnh
            with open(filepath, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)
            
            logger.info(f"Saved DALL-E image: {filepath}")
            return filepath
            
        except Exception as e:
            logger.error(f"Error downloading/saving image: {str(e)}")
            return None
    
    def generate_images_for_presentation(self, presentation_data: Dict[str, Any]) -> Dict[str, str]:
        """
        Tạo ảnh cho toàn bộ presentation
        
        Args:
            presentation_data (Dict): Dữ liệu presentation
            
        Returns:
            Dict[str, str]: Mapping slide_number -> image_path
        """
        images = {}
        topic = presentation_data.get("presentation_info", {}).get("title", "")
        
        for slide in presentation_data.get("slides", []):
            slide_num = slide.get("slide_number")
            
            # Bỏ qua slide title (slide 1)
            if slide_num <= 1:
                continue
                
            image_path = self.generate_image_for_slide(slide, topic)
            if image_path:
                images[str(slide_num)] = image_path
            
        return images

# Utility functions cho integration
def create_dalle_generator(api_key: str) -> DALLEImageGenerator:
    """
    Factory function để tạo DALL-E generator
    
    Args:
        api_key (str): OpenAI API key
        
    Returns:
        DALLEImageGenerator: Instance đã khởi tạo
    """
    return DALLEImageGenerator(api_key)

 