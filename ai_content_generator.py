# ai_content_generator.py
"""
Module để generate nội dung PowerPoint sử dụng ChatGPT API
Enhanced version với khả năng tương tác, tạo hình ảnh và tùy chỉnh theme
"""

import openai
import json
import re
import requests
import os
from typing import Dict, List, Optional, Any, Tuple
import logging
from datetime import datetime
from PIL import Image
import io

# Import custom modules
from dalle_generator import DALLEImageGenerator
from theme_system import ModernThemeSystem

# Setup logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class EnhancedAIContentGenerator:
    """
    Enhanced AI Content Generator với khả năng tương tác và tạo hình ảnh
    """
    
    def __init__(self, api_key: str, model: str = "gpt-3.5-turbo"):
        """
        Khởi tạo Enhanced AI Content Generator
        
        Args:
            api_key (str): OpenAI API key
            model (str): Model để sử dụng (gpt-3.5-turbo, gpt-4, etc.)
        """
        openai.api_key = api_key
        self.model = model
        self.max_tokens = 3000
        self.temperature = 0.7
        
        # Initialize subsystems
        self.dalle_generator = DALLEImageGenerator(api_key)
        self.theme_system = ModernThemeSystem()
        
        # Conversation state
        self.conversation_history = []
        self.current_context = {}
        
        # Enhanced system prompts
        self.system_prompts = {
            "education": """Bạn là một chuyên gia giáo dục với 20 năm kinh nghiệm trong việc thiết kế bài giảng tương tác. 
            Bạn có khả năng tạo ra những bài giảng PowerPoint chất lượng cao, phù hợp với từng cấp học và môn học.
            Bạn luôn hỏi câu hỏi để hiểu rõ nhu cầu và tạo nội dung tối ưu.""",
            
            "business": """Bạn là một chuyên gia tư vấn doanh nghiệp với kinh nghiệm sâu về thuyết trình và trình bày. 
            Bạn có thể tạo ra những presentation chuyên nghiệp cho môi trường công sở với visual design hiện đại.""",
            
            "training": """Bạn là một chuyên gia đào tạo với khả năng thiết kế các khóa học và bài training hiệu quả.
            Bạn biết cách truyền đạt kiến thức một cách sinh động, dễ hiểu và có tính tương tác cao."""
        }
        
        # Question templates for interactive gathering
        self.question_templates = {
            "basic_info": [
                "Chủ đề cụ thể bạn muốn trình bày là gì?",
                "Đối tượng khán giả của bạn là ai? (học sinh lớp mấy, nhân viên, v.v.)",
                "Thời gian dự kiến trình bày bao lâu?",
                "Bạn có yêu cầu đặc biệt nào không?"
            ],
            "content_depth": [
                "Bạn muốn nội dung có độ sâu như thế nào? (cơ bản/trung bình/nâng cao)",
                "Có phần nào bạn muốn tập trung nhiều hơn?",
                "Bạn có muốn thêm ví dụ thực tế hay case study không?",
                "Có cần thêm phần thực hành hay bài tập không?"
            ],
            "visual_preferences": [
                "Bạn thích style nào? (chuyên nghiệp/sáng tạo/hiện đại)",
                "Màu sắc ưa thích? (xanh dương/tím/cam/tự động)",
                "Có muốn thêm hình ảnh minh họa không?",
                "Số lượng slide mong muốn?"
            ]
        }
    
    def start_interactive_session(self, initial_request: str) -> Dict[str, Any]:
        """
        Bắt đầu session tương tác với người dùng
        
        Args:
            initial_request (str): Yêu cầu ban đầu của user
            
        Returns:
            Dict: Response với câu hỏi và context
        """
        try:
            # Phân tích yêu cầu ban đầu
            initial_analysis = self._analyze_initial_request(initial_request)
            
            # Xác định câu hỏi cần hỏi
            missing_info = self._identify_missing_information(initial_analysis)
            
            # Tạo câu hỏi tương tác
            questions = self._generate_interactive_questions(missing_info)
            
            # Lưu context
            self.current_context = {
                "initial_request": initial_request,
                "analysis": initial_analysis,
                "missing_info": missing_info,
                "questions_asked": [],
                "answers_collected": {},
                "session_id": datetime.now().isoformat()
            }
            
            return {
                "type": "interactive_questions",
                "message": "Tôi sẽ hỏi một vài câu hỏi để tạo bài thuyết trình tốt nhất cho bạn:",
                "questions": questions,
                "analysis": initial_analysis,
                "progress": "Bước 1/3: Thu thập thông tin cơ bản"
            }
            
        except Exception as e:
            logger.error(f"Error starting interactive session: {str(e)}")
            return self._fallback_response(initial_request)
    
    def process_user_answers(self, answers: Dict[str, Any]) -> Dict[str, Any]:
        """
        Xử lý câu trả lời của user và quyết định bước tiếp theo
        
        Args:
            answers (Dict): Câu trả lời của user
            
        Returns:
            Dict: Response với bước tiếp theo
        """
        try:
            # Cập nhật context với answers
            self.current_context["answers_collected"].update(answers)
            
            # Phân tích độ đầy đủ thông tin
            completeness = self._assess_information_completeness()
            
            if completeness["ready_to_generate"]:
                # Đủ thông tin để tạo presentation
                return self._proceed_to_generation()
            else:
                # Cần hỏi thêm câu hỏi
                additional_questions = self._generate_follow_up_questions(completeness)
                return {
                    "type": "follow_up_questions",
                    "message": "Cảm ơn! Tôi cần thêm một vài thông tin nữa:",
                    "questions": additional_questions,
                    "progress": f"Bước 2/3: Hoàn thiện thông tin ({completeness['completion_percentage']}%)"
                }
                
        except Exception as e:
            logger.error(f"Error processing user answers: {str(e)}")
            return self._proceed_to_generation()  # Fallback to generation
    
    def generate_enhanced_presentation(self, context: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
        """
        Tạo presentation với enhanced features
        
        Args:
            context (Optional[Dict]): Context từ interactive session
            
        Returns:
            Dict: Complete presentation data với images và theme
        """
        try:
            if context is None:
                context = self.current_context
            
            logger.info("Generating enhanced presentation...")
            
            # Step 1: Tạo outline chi tiết
            outline = self._generate_enhanced_outline(context)
            
            # Step 2: Tạo nội dung chi tiết
            presentation_data = self._generate_detailed_content_enhanced(outline, context)
            
            # Step 3: Phân tích và tạo hình ảnh
            image_analysis = self._analyze_content_for_images(presentation_data)
            presentation_data["image_suggestions"] = image_analysis
            
            # Step 4: Tự động chọn theme phù hợp
            recommended_theme = self._auto_select_theme(presentation_data, context)
            presentation_data["recommended_theme"] = recommended_theme
            
            # Step 5: Thêm icons và visual elements
            presentation_data = self._enhance_with_visual_elements(presentation_data)
            
            # Step 6: Generate images nếu được yêu cầu
            if context.get("answers_collected", {}).get("include_images", True):
                image_paths = self._generate_images_for_slides(presentation_data)
                presentation_data["generated_images"] = image_paths
                
                # Step 7: Update slides với generated image paths
                logger.info(f"Generated {len(image_paths)} images: {list(image_paths.keys())}")
                for slide_index_str, image_path in image_paths.items():
                    slide_index = int(slide_index_str)
                    if slide_index < len(presentation_data.get("slides", [])):
                        presentation_data["slides"][slide_index]["generated_image_path"] = image_path
                        logger.info(f"Updated slide {slide_index} with image: {image_path}")
            
            # Add metadata
            presentation_data["generation_info"] = {
                "model_used": self.model,
                "generated_at": datetime.now().isoformat(),
                "interactive_session": True,
                "context_used": context
            }
            
            logger.info("Enhanced presentation generated successfully")
            return presentation_data
            
        except Exception as e:
            logger.error(f"Error generating enhanced presentation: {str(e)}")
            return self._create_fallback_presentation()
    
    def _analyze_initial_request(self, request: str) -> Dict[str, Any]:
        """Phân tích yêu cầu ban đầu để xác định thông tin cơ bản"""
        try:
            analysis_prompt = f"""
            Phân tích yêu cầu sau và trích xuất thông tin có sẵn:
            "{request}"
            
            Xác định những gì đã biết và những gì cần hỏi thêm:
            1. Chủ đề/môn học
            2. Cấp độ/đối tượng  
            3. Thời gian
            4. Loại presentation
            5. Yêu cầu đặc biệt
            
            Trả về JSON format:
            {{
                "identified_info": {{
                    "topic": "chủ đề đã xác định hoặc null",
                    "subject": "môn học hoặc lĩnh vực",
                    "audience": "đối tượng đã xác định hoặc null", 
                    "duration": "thời gian hoặc null",
                "type": "education|business|training",
                    "special_requirements": ["yêu cầu đặc biệt"]
                }},
                "confidence_level": "high|medium|low",
                "missing_critical_info": ["thông tin quan trọng còn thiếu"]
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
            
            content = response.choices[0].message.content
            json_match = re.search(r'\{.*\}', content, re.DOTALL)
            
            if json_match:
                return json.loads(json_match.group())
            else:
                return self._fallback_analysis(request)
                
        except Exception as e:
            logger.error(f"Error analyzing initial request: {str(e)}")
            return self._fallback_analysis(request)
    
    def _identify_missing_information(self, analysis: Dict[str, Any]) -> List[str]:
        """Xác định thông tin còn thiếu"""
        missing = []
        identified = analysis.get("identified_info", {})
        
        if not identified.get("topic"):
            missing.append("topic")
        if not identified.get("audience"):
            missing.append("audience")
        if not identified.get("duration"):
            missing.append("duration")
        if analysis.get("confidence_level") == "low":
            missing.extend(["content_depth", "visual_preferences"])
        
        return missing
    
    def _generate_interactive_questions(self, missing_info: List[str]) -> List[Dict[str, Any]]:
        """Tạo câu hỏi tương tác dựa trên thông tin còn thiếu"""
        questions = []
        
        for info_type in missing_info:
            if info_type in ["topic", "audience", "duration"]:
                questions.extend(self._get_basic_questions(info_type))
            elif info_type == "content_depth":
                questions.extend(self._get_content_questions())
            elif info_type == "visual_preferences":
                questions.extend(self._get_visual_questions())
        
        # Limit to 3-4 questions per round
        return questions[:4]
    
    def _get_basic_questions(self, info_type: str) -> List[Dict[str, Any]]:
        """Get basic information questions"""
        question_map = {
            "topic": {
                "question": "Chủ đề cụ thể bạn muốn trình bày là gì?",
                "type": "text",
                "key": "topic",
                "required": True
            },
            "audience": {
                "question": "Đối tượng khán giả của bạn là ai?",
                "type": "select",
                "key": "audience",
                "options": ["Học sinh tiểu học", "Học sinh THCS", "Học sinh THPT", "Sinh viên đại học", "Nhân viên công ty", "Quản lý/Lãnh đạo", "Khách hàng", "Khác"],
                "required": True
            },
            "duration": {
                "question": "Thời gian dự kiến trình bày?",
                "type": "select", 
                "key": "duration",
                "options": ["15 phút", "30 phút", "45 phút", "60 phút", "90 phút", "Khác"],
                "required": True
            }
        }
        
        return [question_map.get(info_type, {})]
    
    def _get_content_questions(self) -> List[Dict[str, Any]]:
        """Get content-related questions"""
        return [
            {
                "question": "Mức độ chi tiết nội dung?",
                "type": "select",
                "key": "content_depth",
                "options": ["Cơ bản - Giới thiệu tổng quan", "Trung bình - Có ví dụ và chi tiết", "Nâng cao - Phân tích sâu và case study"],
                "required": True
            },
            {
                "question": "Bạn có muốn thêm ví dụ thực tế và bài tập không?",
                "type": "boolean",
                "key": "include_examples",
                "required": False
            }
        ]
    
    def _get_visual_questions(self) -> List[Dict[str, Any]]:
        """Get visual preference questions"""
        return [
            {
                "question": "Style presentation ưa thích?", 
                "type": "select",
                "key": "presentation_style",
                "options": ["Chuyên nghiệp - Business", "Giáo dục - Thân thiện", "Hiện đại - Tech", "Sáng tạo - Artistic"],
                "required": True
            },
            {
                "question": "Có muốn AI tự động tạo hình ảnh minh họa không?",
                "type": "boolean", 
                "key": "include_images",
                "required": False
            }
        ]
    
    def _assess_information_completeness(self) -> Dict[str, Any]:
        """Đánh giá độ đầy đủ thông tin"""
        answers = self.current_context.get("answers_collected", {})
        
        required_fields = ["topic", "audience", "duration"]
        optional_fields = ["content_depth", "presentation_style", "include_examples", "include_images"]
        
        required_filled = sum(1 for field in required_fields if answers.get(field))
        optional_filled = sum(1 for field in optional_fields if field in answers)
        
        total_possible = len(required_fields) + len(optional_fields)
        total_filled = required_filled + optional_filled
        
        completion_percentage = int((total_filled / total_possible) * 100)
        ready_to_generate = required_filled == len(required_fields)
        
        return {
            "ready_to_generate": ready_to_generate,
            "completion_percentage": completion_percentage,
            "missing_required": [f for f in required_fields if not answers.get(f)],
            "missing_optional": [f for f in optional_fields if f not in answers]
        }
    
    def _generate_follow_up_questions(self, completeness: Dict[str, Any]) -> List[Dict[str, Any]]:
        """Tạo câu hỏi follow-up"""
        questions = []
        
        # Hỏi required fields còn thiếu
        for field in completeness["missing_required"]:
            questions.extend(self._get_basic_questions(field))
        
        # Hỏi một vài optional fields quan trọng
        important_optional = ["content_depth", "presentation_style"]
        for field in important_optional:
            if field in completeness["missing_optional"]:
                if field == "content_depth":
                    questions.extend(self._get_content_questions()[:1])
                elif field == "presentation_style":
                    questions.extend(self._get_visual_questions()[:1])
        
        return questions[:3]  # Limit follow-up questions
    
    def _proceed_to_generation(self) -> Dict[str, Any]:
        """Tiến hành tạo presentation"""
        return {
            "type": "generation_ready",
            "message": "Hoàn tất thu thập thông tin! Đang tạo presentation...",
            "progress": "Bước 3/3: Tạo presentation",
            "estimated_time": "30-60 giây"
        }
    
    def _generate_enhanced_outline(self, context: Dict[str, Any]) -> Dict[str, Any]:
        """Tạo outline nâng cao dựa trên context"""
        try:
            answers = context.get("answers_collected", {})
            initial_analysis = context.get("analysis", {})
            
            # Combine information
            topic = answers.get("topic") or initial_analysis.get("identified_info", {}).get("topic", "")
            audience = answers.get("audience") or initial_analysis.get("identified_info", {}).get("audience", "")
            duration = answers.get("duration") or initial_analysis.get("identified_info", {}).get("duration", "45 phút")
            content_depth = answers.get("content_depth", "Trung bình")
            include_examples = answers.get("include_examples", True)
            
            outline_prompt = f"""
            Tạo outline chi tiết cho presentation với thông tin:
            - Chủ đề: {topic}
            - Đối tượng: {audience}
            - Thời gian: {duration}
            - Mức độ: {content_depth}
            - Có ví dụ: {include_examples}
            
            Yêu cầu outline:
            1. Cấu trúc logic và có tính thuyết phục
            2. Phù hợp với thời gian và đối tượng
            3. Bao gồm slide mở đầu hấp dẫn
            4. Nội dung chính chia thành 3-5 phần
            5. Slide kết luận với call-to-action
            6. Xác định slide nào cần hình ảnh minh họa
            
            Format JSON:
            {{
                "presentation_info": {{
                    "title": "Tiêu đề bài presentation",
                    "subtitle": "Phụ đề hấp dẫn",
                    "author": "Được tạo bởi AI",
                    "template": "education|business|training",
                    "estimated_duration": "{duration}",
                    "total_slides": "số slides",
                    "target_audience": "{audience}",
                    "difficulty_level": "{content_depth}"
                }},
                "slides": [
                    {{
                        "slide_number": 1,
                        "type": "title|content|two_column|image_focus|conclusion",
                        "title": "Tiêu đề slide",
                        "purpose": "Mục đích của slide",
                        "content_outline": ["Điểm 1", "Điểm 2"],
                        "needs_image": true/false,
                        "image_concept": "Ý tưởng hình ảnh nếu cần",
                        "estimated_time": "thời gian ước tính (phút)"
                    }}
                ]
            }}
            """
            
            system_prompt = self._get_system_prompt_for_context(context)
            
            response = openai.ChatCompletion.create(
                model=self.model,
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": outline_prompt}
                ],
                max_tokens=self.max_tokens,
                temperature=self.temperature
            )
            
            content = response.choices[0].message.content
            json_match = re.search(r'\{.*\}', content, re.DOTALL)
            
            if json_match:
                outline_data = json.loads(json_match.group())
                logger.info(f"Generated enhanced outline with {len(outline_data.get('slides', []))} slides")
                return outline_data
            else:
                return self._fallback_outline_enhanced(context)
                
        except Exception as e:
            logger.error(f"Error generating enhanced outline: {str(e)}")
            return self._fallback_outline_enhanced(context)
    
    def _generate_detailed_content_enhanced(self, outline_data: Dict[str, Any], context: Dict[str, Any]) -> Dict[str, Any]:
        """Generate nội dung chi tiết enhanced"""
        try:
            presentation_info = outline_data.get('presentation_info', {})
            slides_outline = outline_data.get('slides', [])
            
            # Create enhanced presentation structure
            presentation_data = {
                "title": presentation_info.get('title', 'Bài Giảng'),
                "subtitle": presentation_info.get('subtitle', ''),
                "author": presentation_info.get('author', 'AI Assistant'),
                "template": presentation_info.get('template', 'education'),
                "generated_at": datetime.now().isoformat(),
                "target_audience": presentation_info.get('target_audience', ''),
                "difficulty_level": presentation_info.get('difficulty_level', ''),
                "estimated_duration": presentation_info.get('estimated_duration', ''),
                "slides": []
            }
            
            # Generate enhanced content for each slide
            for slide_outline in slides_outline:
                if slide_outline.get('type') == 'title':
                    continue
                
                detailed_slide = self._generate_enhanced_slide_content(slide_outline, presentation_info, context)
                if detailed_slide:
                    presentation_data['slides'].append(detailed_slide)
            
            logger.info(f"Generated enhanced content for {len(presentation_data['slides'])} slides")
            return presentation_data
            
        except Exception as e:
            logger.error(f"Error generating enhanced detailed content: {str(e)}")
            return self._create_fallback_presentation()
    
    def _generate_enhanced_slide_content(self, slide_outline: Dict[str, Any], presentation_info: Dict[str, Any], context: Dict[str, Any]) -> Optional[Dict[str, Any]]:
        """Generate enhanced slide content"""
        try:
            slide_type = slide_outline.get('type', 'content')
            slide_title = slide_outline.get('title', '')
            content_outline = slide_outline.get('content_outline', [])
            needs_image = slide_outline.get('needs_image', False)
            image_concept = slide_outline.get('image_concept', '')
            
            # Get audience and content depth for context
            answers = context.get("answers_collected", {})
            audience = answers.get("audience", "")
            content_depth = answers.get("content_depth", "Trung bình")
            
            content_prompt = f"""
            Tạo nội dung chi tiết và hấp dẫn cho slide:
            - Tiêu đề: {slide_title}
            - Loại slide: {slide_type}
            - Outline: {', '.join(content_outline)}
            - Đối tượng: {audience}
            - Mức độ: {content_depth}
            - Context: Slide thuộc bài "{presentation_info.get('title', '')}"
            
            Yêu cầu nội dung:
            - Chi tiết phù hợp với mức độ "{content_depth}"
            - Ngôn ngữ phù hợp với "{audience}"
            - Có tính thuyết phục và hấp dẫn
            - Cấu trúc rõ ràng, dễ đọc trên slide
            - 3-6 bullet points chính, mỗi point ngắn gọn nhưng đầy đủ ý
            
            Chỉ trả về nội dung dưới dạng bullet points, không cần format JSON.
            """
            
            response = openai.ChatCompletion.create(
                model=self.model,
                messages=[
                    {"role": "system", "content": "Bạn là chuyên gia tạo nội dung presentation chất lượng cao, hấp dẫn và phù hợp với đối tượng."},
                    {"role": "user", "content": content_prompt}
                ],
                max_tokens=1000,
                temperature=self.temperature
            )
            
            content_text = response.choices[0].message.content.strip()
            
            # Parse content into list
            content_list = self._parse_content_to_list(content_text)
            
            # Create enhanced slide data
            slide_data = {
                "type": slide_type,
                "title": slide_title,
                "needs_image": needs_image,
                "image_concept": image_concept,
                "estimated_time": slide_outline.get("estimated_time", "2-3 phút")
            }
            
            if slide_type == "two_column":
                mid_point = len(content_list) // 2
                slide_data["left_content"] = content_list[:mid_point]
                slide_data["right_content"] = content_list[mid_point:]
            else:
                slide_data["content"] = content_list
            
            return slide_data
            
        except Exception as e:
            logger.error(f"Error generating enhanced slide content: {str(e)}")
            return None
    
    def _analyze_content_for_images(self, presentation_data: Dict[str, Any]) -> Dict[str, Any]:
        """Phân tích nội dung để xác định hình ảnh cần thiết"""
        try:
            image_analysis = {
                "total_slides": len(presentation_data.get("slides", [])),
                "slides_needing_images": [],
                "image_concepts": {},
                "priority_slides": []
            }
            
            for i, slide in enumerate(presentation_data.get("slides", [])):
                slide_title = slide.get("title", "")
                slide_content = slide.get("content", [])
                slide_type = slide.get("type", "")
                
                # Phân tích nhu cầu hình ảnh
                image_priority = self._assess_image_priority(slide_title, slide_content, slide_type)
                
                if image_priority["needs_image"]:
                    image_analysis["slides_needing_images"].append(i)
                    image_analysis["image_concepts"][i] = image_priority["concept"]
                    
                    if image_priority["priority"] == "high":
                        image_analysis["priority_slides"].append(i)
            
            return image_analysis
            
        except Exception as e:
            logger.error(f"Error analyzing content for images: {str(e)}")
            return {"slides_needing_images": [], "image_concepts": {}}
    
    def _assess_image_priority(self, title: str, content: List[str], slide_type: str) -> Dict[str, Any]:
        """Đánh giá ưu tiên hình ảnh cho slide"""
        title_lower = title.lower()
        content_text = " ".join(content).lower()
        
        # High priority keywords
        high_priority_keywords = [
            "cấu trúc", "mô hình", "sơ đồ", "biểu đồ", "quy trình", 
            "chu trình", "hệ thống", "kiến trúc", "phương pháp",
            "structure", "model", "process", "system", "method"
        ]
        
        # Science/tech keywords that benefit from visuals
        visual_benefit_keywords = [
            "tế bào", "phân tử", "nguyên tử", "protein", "dna",
            "mạch điện", "sóng", "năng lượng", "phản ứng",
            "thuật toán", "code", "lập trình", "dữ liệu"
        ]
        
        needs_image = False
        priority = "low"
        concept = ""
        
        # Check for high priority
        if any(keyword in title_lower or keyword in content_text for keyword in high_priority_keywords):
            needs_image = True
            priority = "high"
            concept = f"Diagram or illustration for {title}"
        
        # Check for visual benefit
        elif any(keyword in title_lower or keyword in content_text for keyword in visual_benefit_keywords):
            needs_image = True
            priority = "medium"
            concept = f"Scientific or technical illustration for {title}"
        
        # Content-heavy slides benefit from images
        elif len(content) > 4 and slide_type == "content":
            needs_image = True
            priority = "medium"
            concept = f"Supporting illustration for {title}"
        
        return {
            "needs_image": needs_image,
            "priority": priority,
            "concept": concept
        }
    
    def _auto_select_theme(self, presentation_data: Dict[str, Any], context: Dict[str, Any]) -> Dict[str, Any]:
        """Tự động chọn theme phù hợp"""
        answers = context.get("answers_collected", {})
        presentation_style = answers.get("presentation_style", "")
        topic = presentation_data.get("title", "").lower()
        
        # Theme mapping based on style and content
        if "Business" in presentation_style:
            recommended_theme = "business_elegant"
        elif "Tech" in presentation_style or any(keyword in topic for keyword in ["lập trình", "programming", "ai", "data"]):
            recommended_theme = "tech_gradient"
        elif "Artistic" in presentation_style:
            recommended_theme = "creative_vibrant"
        else:
            recommended_theme = "education_pro"
        
        theme_info = self.theme_system.get_theme(recommended_theme)
        
        return {
            "theme_name": recommended_theme,
            "theme_info": theme_info,
            "auto_selected": True,
            "reason": f"Selected based on style: {presentation_style} and content analysis"
        }
    
    def _enhance_with_visual_elements(self, presentation_data: Dict[str, Any]) -> Dict[str, Any]:
        """Thêm icons và visual elements"""
        try:
            # Detect subject for appropriate icons
            topic = presentation_data.get("title", "").lower()
            subject_icon = self.theme_system.get_subject_icon(topic)
            
            # Add icons to slides based on content
            for slide in presentation_data.get("slides", []):
                slide_title = slide.get("title", "").lower()
                
                # Add appropriate icon based on slide content
                if any(keyword in slide_title for keyword in ["giới thiệu", "introduction"]):
                    slide["icon"] = "🎯"
                elif any(keyword in slide_title for keyword in ["kết luận", "conclusion", "tóm tắt"]):
                    slide["icon"] = "🏆"
                elif any(keyword in slide_title for keyword in ["ví dụ", "example", "thực tế"]):
                    slide["icon"] = "💡"
                elif any(keyword in slide_title for keyword in ["phương pháp", "method", "cách"]):
                    slide["icon"] = "⚙️"
                else:
                    slide["icon"] = subject_icon
            
            # Add presentation-level visual metadata
            presentation_data["visual_elements"] = {
                "primary_icon": subject_icon,
                "color_scheme": "auto-selected",
                "visual_style": "modern_professional"
            }
            
            return presentation_data
            
        except Exception as e:
            logger.error(f"Error enhancing with visual elements: {str(e)}")
            return presentation_data
    
    def _generate_images_for_slides(self, presentation_data: Dict[str, Any]) -> Dict[str, str]:
        """Tạo hình ảnh cho các slides cần thiết"""
        try:
            image_paths = {}
            image_analysis = presentation_data.get("image_suggestions", {})
            slides_needing_images = image_analysis.get("slides_needing_images", [])
            
            topic = presentation_data.get("title", "")
            
            for slide_index in slides_needing_images:
                if slide_index < len(presentation_data.get("slides", [])):
                    slide = presentation_data["slides"][slide_index]
                    
                    # Generate image for this slide
                    image_path = self.dalle_generator.generate_image_for_slide(slide, topic)
                    
                    if image_path:
                        image_paths[slide_index] = image_path
                        logger.info(f"Generated image for slide {slide_index + 1}")
            
            return image_paths
                
        except Exception as e:
            logger.error(f"Error generating images for slides: {str(e)}")
            return {}
    
    # Helper methods
    def _get_system_prompt_for_context(self, context: Dict[str, Any]) -> str:
        """Get appropriate system prompt based on context"""
        analysis = context.get("analysis", {})
        presentation_type = analysis.get("identified_info", {}).get("type", "education")
        return self.system_prompts.get(presentation_type, self.system_prompts["education"])
    
    def _parse_content_to_list(self, content_text: str) -> List[str]:
        """Parse content text to list of points"""
        content_list = []
        for line in content_text.split('\n'):
            line = line.strip()
            if line and not line.startswith('#'):
                # Remove bullet markers if present
                line = re.sub(r'^[•\-\*\d+\.]\s*', '', line)
                if line:
                    content_list.append(line)
        return content_list
    
    def _fallback_response(self, request: str) -> Dict[str, Any]:
        """Fallback response when interactive session fails"""
        return {
            "type": "direct_generation",
            "message": "Đang tạo presentation dựa trên yêu cầu của bạn...",
            "fallback": True
        }
    
    def _fallback_analysis(self, user_message: str) -> Dict[str, Any]:
        """Fallback analysis when AI fails"""
        return {
            "identified_info": {
                "topic": user_message[:50] + "..." if len(user_message) > 50 else user_message,
                "subject": "Chung", 
                "audience": None,
                "duration": None,
            "type": "education",
                "special_requirements": []
            },
            "confidence_level": "low",
            "missing_critical_info": ["audience", "duration", "content_depth"]
        }
    
    def _fallback_outline_enhanced(self, context: Dict[str, Any]) -> Dict[str, Any]:
        """Enhanced fallback outline"""
        answers = context.get("answers_collected", {})
        topic = answers.get("topic", "Bài Giảng")
        
        return {
            "presentation_info": {
                "title": topic,
                "subtitle": "Được tạo bởi AI Assistant",
                "author": "AI Assistant",
                "template": "education",
                "estimated_duration": answers.get("duration", "45 phút"),
                "total_slides": "5",
                "target_audience": answers.get("audience", "Học sinh"),
                "difficulty_level": answers.get("content_depth", "Trung bình")
            },
            "slides": [
                {
                    "slide_number": 1,
                    "type": "content",
                    "title": "Giới thiệu",
                    "purpose": "Mở đầu thu hút",
                    "content_outline": ["Tổng quan chủ đề", "Tầm quan trọng", "Mục tiêu"],
                    "needs_image": True,
                    "image_concept": "Introduction illustration",
                    "estimated_time": "3 phút"
                },
                {
                    "slide_number": 2,
                    "type": "content",
                    "title": "Nội dung chính",
                    "purpose": "Trình bày kiến thức cốt lõi",
                    "content_outline": ["Khái niệm cơ bản", "Nguyên lý quan trọng", "Ứng dụng"],
                    "needs_image": True,
                    "image_concept": "Main content illustration",
                    "estimated_time": "5 phút"
                },
                {
                    "slide_number": 3,
                    "type": "content",
                    "title": "Kết luận",
                    "purpose": "Tóm tắt và hành động",
                    "content_outline": ["Tóm tắt chính", "Bài học kinh nghiệm rút ra", "Bước tiếp theo"],
                    "needs_image": False,
                    "image_concept": "",
                    "estimated_time": "2 phút"
                }
            ]
        }
    
    def _create_fallback_presentation(self) -> Dict[str, Any]:
        """Create enhanced fallback presentation"""
        return {
            "title": "Bài Giảng",
            "subtitle": "Được tạo bởi AI Assistant",
            "author": "AI Assistant",
            "template": "education",
            "generated_at": datetime.now().isoformat(),
            "target_audience": "Học sinh",
            "difficulty_level": "Trung bình",
            "recommended_theme": {
                "theme_name": "education_pro",
                "auto_selected": True
            },
            "slides": [
                {
                    "type": "content",
                    "title": "Giới thiệu chủ đề",
                    "content": [
                        "Chào mừng đến với bài giảng hôm nay",
                        "Chúng ta sẽ khám phá những kiến thức thú vị",
                        "Mục tiêu: Hiểu rõ và ứng dụng được kiến thức"
                    ],
                    "icon": "🎯",
                    "needs_image": True
                },
                {
                    "type": "content", 
                    "title": "Nội dung cốt lõi",
                    "content": [
                        "Khái niệm và định nghĩa cơ bản",
                        "Nguyên lý và quy luật quan trọng",
                        "Các ví dụ minh họa cụ thể",
                        "Ứng dụng trong thực tế"
                    ],
                    "icon": "📚",
                    "needs_image": True
                },
                {
                    "type": "content",
                    "title": "Tổng kết và hành động",
                    "content": [
                        "Tóm tắt những điểm quan trọng",
                        "Bài học kinh nghiệm rút ra",
                        "Hướng phát triển tiếp theo",
                        "Câu hỏi thảo luận"
                    ],
                    "icon": "🏆",
                    "needs_image": False
                }
            ]
        }

# Backward compatibility - keep original class name as alias
AIContentGenerator = EnhancedAIContentGenerator

# Test functions
