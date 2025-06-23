# ai_content_generator.py
"""
Module để generate nội dung PowerPoint sử dụng ChatGPT API
"""

import openai
import json
import re
from typing import Dict, List, Optional, Any
import logging
from datetime import datetime

# Setup logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class AIContentGenerator:
    """
    Class để generate nội dung presentation sử dụng ChatGPT
    """
    
    def __init__(self, api_key: str, model: str = "gpt-3.5-turbo"):
        """
        Khởi tạo AI Content Generator
        
        Args:
            api_key (str): OpenAI API key
            model (str): Model để sử dụng (gpt-3.5-turbo, gpt-4, etc.)
        """
        openai.api_key = api_key
        self.model = model
        self.max_tokens = 3000
        self.temperature = 0.7
        
        # System prompts cho different types
        self.system_prompts = {
            "education": """Bạn là một chuyên gia giáo dục với 20 năm kinh nghiệm trong việc thiết kế bài giảng và giáo án. 
            Bạn có khả năng tạo ra những bài giảng PowerPoint chất lượng cao, phù hợp với từng cấp học và môn học.""",
            
            "business": """Bạn là một chuyên gia tư vấn doanh nghiệp với kinh nghiệm sâu về thuyết trình và trình bày. 
            Bạn có thể tạo ra những presentation chuyên nghiệp cho môi trường công sở.""",
            
            "training": """Bạn là một chuyên gia đào tạo với khả năng thiết kế các khóa học và bài training hiệu quả.
            Bạn biết cách truyền đạt kiến thức một cách sinh động và dễ hiểu."""
        }
    
    def analyze_user_request(self, user_message: str) -> Dict[str, Any]:
        """
        Phân tích yêu cầu của user để xác định loại presentation cần tạo
        
        Args:
            user_message (str): Tin nhắn của user
            
        Returns:
            Dict: Thông tin đã phân tích
        """
        try:
            analysis_prompt = f"""
            Phân tích yêu cầu sau và trích xuất thông tin:
            "{user_message}"
            
            Hãy xác định:
            1. Loại presentation (education/business/training)
            2. Chủ đề chính
            3. Đối tượng (học sinh lớp mấy, nhân viên, etc.)
            4. Thời gian trình bày (nếu có)
            5. Yêu cầu đặc biệt (nếu có)
            
            Trả về JSON format:
            {{
                "type": "education|business|training",
                "subject": "môn học hoặc lĩnh vực",
                "topic": "chủ đề chính",
                "audience": "đối tượng",
                "duration": "thời gian (phút)",
                "special_requirements": ["yêu cầu 1", "yêu cầu 2"],
                "estimated_slides": "số slides ước tính"
            }}
            """
            
            response = openai.ChatCompletion.create(
                model=self.model,
                messages=[
                    {"role": "system", "content": "Bạn là chuyên gia phân tích yêu cầu presentation."},
                    {"role": "user", "content": analysis_prompt}
                ],
                max_tokens=800,
                temperature=0.3
            )
            
            # Extract JSON from response
            content = response.choices[0].message.content
            json_match = re.search(r'\{.*\}', content, re.DOTALL)
            
            if json_match:
                analysis_data = json.loads(json_match.group())
                logger.info(f"Successfully analyzed user request: {analysis_data['topic']}")
                return analysis_data
            else:
                # Fallback analysis
                return self._fallback_analysis(user_message)
                
        except Exception as e:
            logger.error(f"Error analyzing user request: {str(e)}")
            return self._fallback_analysis(user_message)
    
    def generate_presentation_outline(self, analysis_data: Dict[str, Any]) -> Dict[str, Any]:
        """
        Generate outline cho presentation dựa trên analysis
        
        Args:
            analysis_data (Dict): Dữ liệu đã phân tích từ user request
            
        Returns:
            Dict: Outline chi tiết cho presentation
        """
        try:
            presentation_type = analysis_data.get('type', 'education')
            system_prompt = self.system_prompts.get(presentation_type, self.system_prompts['education'])
            
            outline_prompt = f"""
            Tạo outline chi tiết cho bài presentation với thông tin sau:
            - Loại: {analysis_data.get('type', 'education')}
            - Môn học/Lĩnh vực: {analysis_data.get('subject', '')}
            - Chủ đề: {analysis_data.get('topic', '')}
            - Đối tượng: {analysis_data.get('audience', '')}
            - Thời gian: {analysis_data.get('duration', '45')} phút
            - Yêu cầu đặc biệt: {', '.join(analysis_data.get('special_requirements', []))}
            
            Tạo outline bao gồm:
            1. Slide tiêu đề
            2. Slide mục tiêu/giới thiệu
            3. Các slide nội dung chính (3-7 slides)
            4. Slide thực hành/ví dụ (nếu phù hợp)
            5. Slide tổng kết/kết luận
            
            Format JSON:
            {{
                "presentation_info": {{
                    "title": "Tiêu đề bài presentation",
                    "subtitle": "Phụ đề",
                    "author": "Tên giảng viên",
                    "template": "education|business|training",
                    "estimated_duration": "thời gian",
                    "total_slides": "số slides"
                }},
                "slides": [
                    {{
                        "slide_number": 1,
                        "type": "title|content|two_column|image|chart|table",
                        "title": "Tiêu đề slide",
                        "purpose": "Mục đích của slide này",
                        "content_outline": ["Điểm 1", "Điểm 2", "Điểm 3"]
                    }}
                ]
            }}
            
            Đảm bảo nội dung:
            - Phù hợp với đối tượng
            - Logic và có tính liên kết
            - Thời gian hợp lý cho từng slide
            - Bao gồm elements tương tác (nếu là giáo dục)
            """
            
            response = openai.ChatCompletion.create(
                model=self.model,
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": outline_prompt}
                ],
                max_tokens=self.max_tokens,
                temperature=self.temperature
            )
            
            # Extract and parse JSON
            content = response.choices[0].message.content
            json_match = re.search(r'\{.*\}', content, re.DOTALL)
            
            if json_match:
                outline_data = json.loads(json_match.group())
                logger.info(f"Generated outline with {len(outline_data.get('slides', []))} slides")
                return outline_data
            else:
                return self._fallback_outline(analysis_data)
                
        except Exception as e:
            logger.error(f"Error generating outline: {str(e)}")
            return self._fallback_outline(analysis_data)
    
    def generate_detailed_content(self, outline_data: Dict[str, Any]) -> Dict[str, Any]:
        """
        Generate nội dung chi tiết cho từng slide
        
        Args:
            outline_data (Dict): Outline đã tạo
            
        Returns:
            Dict: Presentation data hoàn chỉnh
        """
        try:
            presentation_info = outline_data.get('presentation_info', {})
            slides_outline = outline_data.get('slides', [])
            
            # Create base presentation structure
            presentation_data = {
                "title": presentation_info.get('title', 'Bài Giảng'),
                "subtitle": presentation_info.get('subtitle', ''),
                "author": presentation_info.get('author', 'AI Assistant'),
                "template": presentation_info.get('template', 'education'),
                "generated_at": datetime.now().isoformat(),
                "slides": []
            }
            
            # Generate content for each slide
            for slide_outline in slides_outline:
                if slide_outline.get('type') == 'title':
                    # Skip title slide as it's handled by presentation_data
                    continue
                
                detailed_slide = self._generate_slide_content(slide_outline, presentation_info)
                if detailed_slide:
                    presentation_data['slides'].append(detailed_slide)
            
            logger.info(f"Generated detailed content for {len(presentation_data['slides'])} slides")
            return presentation_data
            
        except Exception as e:
            logger.error(f"Error generating detailed content: {str(e)}")
            return self._create_fallback_presentation()
    
    def _generate_slide_content(self, slide_outline: Dict[str, Any], presentation_info: Dict[str, Any]) -> Optional[Dict[str, Any]]:
        """
        Generate nội dung chi tiết cho một slide
        
        Args:
            slide_outline (Dict): Outline của slide
            presentation_info (Dict): Thông tin presentation
            
        Returns:
            Dict: Slide data hoàn chỉnh
        """
        try:
            slide_type = slide_outline.get('type', 'content')
            slide_title = slide_outline.get('title', '')
            content_outline = slide_outline.get('content_outline', [])
            
            content_prompt = f"""
            Tạo nội dung chi tiết cho slide:
            - Tiêu đề: {slide_title}
            - Loại slide: {slide_type}
            - Outline: {', '.join(content_outline)}
            - Context: Đây là slide thuộc bài "{presentation_info.get('title', '')}"
            
            Yêu cầu:
            - Nội dung phải chi tiết và đầy đủ
            - Phù hợp với tiêu đề slide
            - Dễ hiểu và logic
            - Độ dài phù hợp cho slide presentation
            
            Trả về nội dung dưới dạng list các bullet points (3-6 points).
            Chỉ trả về content, không cần format JSON.
            """
            
            response = openai.ChatCompletion.create(
                model=self.model,
                messages=[
                    {"role": "system", "content": "Bạn là chuyên gia tạo nội dung presentation chất lượng cao."},
                    {"role": "user", "content": content_prompt}
                ],
                max_tokens=800,
                temperature=self.temperature
            )
            
            content_text = response.choices[0].message.content.strip()
            
            # Parse content into list
            content_list = []
            for line in content_text.split('\n'):
                line = line.strip()
                if line and not line.startswith('#'):
                    # Remove bullet markers if present
                    line = re.sub(r'^[•\-\*\d+\.]\s*', '', line)
                    if line:
                        content_list.append(line)
            
            # Create slide data based on type
            slide_data = {
                "type": slide_type,
                "title": slide_title
            }
            
            if slide_type == "two_column":
                # Split content into two columns
                mid_point = len(content_list) // 2
                slide_data["left_content"] = content_list[:mid_point]
                slide_data["right_content"] = content_list[mid_point:]
            else:
                slide_data["content"] = content_list
            
            return slide_data
            
        except Exception as e:
            logger.error(f"Error generating slide content: {str(e)}")
            return None
    
    def enhance_content_with_examples(self, presentation_data: Dict[str, Any], topic: str) -> Dict[str, Any]:
        """
        Enhance presentation với examples và case studies
        
        Args:
            presentation_data (Dict): Presentation data hiện tại
            topic (str): Chủ đề để tạo examples
            
        Returns:
            Dict: Enhanced presentation data
        """
        try:
            examples_prompt = f"""
            Tạo các ví dụ và case studies phù hợp cho chủ đề: {topic}
            
            Cần tạo:
            1. 2-3 ví dụ thực tế dễ hiểu
            2. 1 case study chi tiết (nếu phù hợp)
            3. Các hoạt động thực hành
            
            Format: Trả về list các ví dụ, mỗi ví dụ trên một dòng.
            """
            
            response = openai.ChatCompletion.create(
                model=self.model,
                messages=[
                    {"role": "system", "content": "Bạn là chuyên gia tạo ví dụ và case studies cho giảng dạy."},
                    {"role": "user", "content": examples_prompt}
                ],
                max_tokens=1000,
                temperature=0.8
            )
            
            examples_text = response.choices[0].message.content.strip()
            examples_list = [line.strip() for line in examples_text.split('\n') if line.strip()]
            
            # Add examples slide
            if examples_list:
                examples_slide = {
                    "type": "content",
                    "title": "Ví dụ thực tế",
                    "content": examples_list[:5]  # Limit to 5 examples
                }
                
                # Insert before the last slide (usually conclusion)
                if len(presentation_data['slides']) > 0:
                    presentation_data['slides'].insert(-1, examples_slide)
                else:
                    presentation_data['slides'].append(examples_slide)
            
            logger.info("Enhanced presentation with examples")
            return presentation_data
            
        except Exception as e:
            logger.error(f"Error enhancing with examples: {str(e)}")
            return presentation_data
    
    def create_presentation_from_chat(self, user_message: str, include_examples: bool = True) -> Dict[str, Any]:
        """
        Tạo presentation hoàn chỉnh từ chat message của user
        
        Args:
            user_message (str): Message của user
            include_examples (bool): Có thêm examples không
            
        Returns:
            Dict: Complete presentation data
        """
        try:
            logger.info(f"Creating presentation from chat: {user_message[:100]}...")
            
            # Step 1: Analyze user request
            analysis = self.analyze_user_request(user_message)
            
            # Step 2: Generate outline
            outline = self.generate_presentation_outline(analysis)
            
            # Step 3: Generate detailed content
            presentation_data = self.generate_detailed_content(outline)
            
            # Step 4: Add examples if requested
            if include_examples:
                topic = analysis.get('topic', '')
                if topic:
                    presentation_data = self.enhance_content_with_examples(presentation_data, topic)
            
            # Add metadata
            presentation_data['ai_analysis'] = analysis
            presentation_data['generation_info'] = {
                'model_used': self.model,
                'generated_at': datetime.now().isoformat(),
                'user_request': user_message
            }
            
            logger.info("Successfully created presentation from chat")
            return presentation_data
            
        except Exception as e:
            logger.error(f"Error creating presentation from chat: {str(e)}")
            return self._create_fallback_presentation()
    
    def suggest_improvements(self, presentation_data: Dict[str, Any], user_feedback: str) -> Dict[str, Any]:
        """
        Suggest improvements dựa trên feedback của user
        
        Args:
            presentation_data (Dict): Presentation data hiện tại
            user_feedback (str): Feedback từ user
            
        Returns:
            Dict: Suggestions for improvement
        """
        try:
            improvement_prompt = f"""
            Dựa trên presentation hiện tại và feedback của user, hãy đề xuất cải thiện:
            
            Presentation title: {presentation_data.get('title', '')}
            Số slides hiện tại: {len(presentation_data.get('slides', []))}
            
            User feedback: "{user_feedback}"
            
            Hãy đề xuất:
            1. Những thay đổi cần thiết
            2. Slides nào cần sửa
            3. Nội dung nào cần thêm/bớt
            4. Cải thiện cấu trúc presentation
            
            Format JSON:
            {{
                "suggestions": [
                    {{
                        "type": "modify|add|remove|reorder",
                        "target": "slide number hoặc 'overall'",
                        "description": "Mô tả thay đổi",
                        "reason": "Lý do"
                    }}
                ],
                "priority": "high|medium|low",
                "estimated_changes": "số lượng thay đổi"
            }}
            """
            
            response = openai.ChatCompletion.create(
                model=self.model,
                messages=[
                    {"role": "system", "content": "Bạn là chuyên gia cải thiện presentation."},
                    {"role": "user", "content": improvement_prompt}
                ],
                max_tokens=1000,
                temperature=0.5
            )
            
            content = response.choices[0].message.content
            json_match = re.search(r'\{.*\}', content, re.DOTALL)
            
            if json_match:
                suggestions = json.loads(json_match.group())
                logger.info("Generated improvement suggestions")
                return suggestions
            else:
                return {"suggestions": [], "priority": "low", "estimated_changes": "0"}
                
        except Exception as e:
            logger.error(f"Error suggesting improvements: {str(e)}")
            return {"suggestions": [], "priority": "low", "estimated_changes": "0"}
    
    # Private helper methods
    def _fallback_analysis(self, user_message: str) -> Dict[str, Any]:
        """Fallback analysis when AI fails"""
        return {
            "type": "education",
            "subject": "Chung",
            "topic": user_message[:50] + "..." if len(user_message) > 50 else user_message,
            "audience": "Học sinh",
            "duration": "45",
            "special_requirements": [],
            "estimated_slides": "5-7"
        }
    
    def _fallback_outline(self, analysis_data: Dict[str, Any]) -> Dict[str, Any]:
        """Fallback outline when AI fails"""
        return {
            "presentation_info": {
                "title": analysis_data.get('topic', 'Bài Giảng'),
                "subtitle": f"Môn: {analysis_data.get('subject', 'Chung')}",
                "author": "AI Assistant",
                "template": "education",
                "estimated_duration": "45 phút",
                "total_slides": "5"
            },
            "slides": [
                {
                    "slide_number": 1,
                    "type": "content",
                    "title": "Giới thiệu",
                    "purpose": "Giới thiệu chủ đề",
                    "content_outline": ["Khái niệm cơ bản", "Tầm quan trọng", "Mục tiêu học tập"]
                },
                {
                    "slide_number": 2,
                    "type": "content",
                    "title": "Nội dung chính",
                    "purpose": "Trình bày nội dung",
                    "content_outline": ["Điểm chính 1", "Điểm chính 2", "Điểm chính 3"]
                },
                {
                    "slide_number": 3,
                    "type": "content",
                    "title": "Kết luận",
                    "purpose": "Tóm tắt và kết luận",
                    "content_outline": ["Tóm tắt", "Kết luận", "Câu hỏi thảo luận"]
                }
            ]
        }
    
    def _create_fallback_presentation(self) -> Dict[str, Any]:
        """Create fallback presentation when all else fails"""
        return {
            "title": "Bài Giảng",
            "subtitle": "Được tạo bởi AI",
            "author": "AI Assistant",
            "template": "education",
            "generated_at": datetime.now().isoformat(),
            "slides": [
                {
                    "type": "content",
                    "title": "Giới thiệu",
                    "content": [
                        "Chào mừng đến với bài giảng",
                        "Hôm nay chúng ta sẽ học về...",
                        "Mục tiêu của bài học"
                    ]
                },
                {
                    "type": "content", 
                    "title": "Nội dung chính",
                    "content": [
                        "Điểm quan trọng thứ nhất",
                        "Điểm quan trọng thứ hai",
                        "Điểm quan trọng thứ ba"
                    ]
                },
                {
                    "type": "content",
                    "title": "Kết luận",
                    "content": [
                        "Tóm tắt những điều đã học",
                        "Ứng dụng thực tế",
                        "Câu hỏi thảo luận"
                    ]
                }
            ]
        }


# Utility functions
def test_ai_generator(api_key: str, test_message: str = "Tạo bài giảng về Toán lớp 10 về phương trình bậc 2"):
    """
    Test AI Content Generator
    
    Args:
        api_key (str): OpenAI API key
        test_message (str): Test message
    """
    try:
        generator = AIContentGenerator(api_key)
        
        print(f"Testing with message: {test_message}")
        print("=" * 50)
        
        # Test analysis
        analysis = generator.analyze_user_request(test_message)
        print("Analysis:")
        print(json.dumps(analysis, indent=2, ensure_ascii=False))
        print()
        
        # Test full generation
        presentation_data = generator.create_presentation_from_chat(test_message)
        print("Generated presentation:")
        print(f"Title: {presentation_data.get('title', '')}")
        print(f"Slides: {len(presentation_data.get('slides', []))}")
        
        for i, slide in enumerate(presentation_data.get('slides', [])[:3]):  # Show first 3 slides
            print(f"  Slide {i+1}: {slide.get('title', '')}")
            content = slide.get('content', [])
            if content:
                print(f"    Content: {len(content)} points")
        
        return presentation_data
        
    except Exception as e:
        print(f"Test failed: {str(e)}")
        return None


if __name__ == "__main__":
    # Test the module
    print("AI Content Generator Module")
    print("Add your OpenAI API key to test this module")
    
    # Uncomment and add your API key to test
    # api_key = "your-openai-api-key-here"
    # test_ai_generator(api_key)