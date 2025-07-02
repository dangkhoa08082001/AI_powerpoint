# main_app.py
"""
Main Streamlit application - Ứng dụng chính để tạo PowerPoint với AI
Enhanced version với interactive features, auto image generation và smart theming
"""

import streamlit as st
import json
from datetime import datetime
from io import BytesIO
import logging
import traceback

# Import custom modules
from powerpoint_generator import PowerPointGenerator
from ai_content_generator import EnhancedAIContentGenerator
from dalle_generator import DALLEImageGenerator
from theme_system import ModernThemeSystem
from powerpoint_editor_module import PowerPointEditorModule

# Setup logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Page config
st.set_page_config(
    page_title="🎓 Enhanced AI PowerPoint Generator",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS với enhanced styles
st.markdown("""
<style>
    .main-header {
        text-align: center;
        background: linear-gradient(90deg, #667eea 0%, #764ba2 100%);
        padding: 2rem;
        border-radius: 10px;
        color: white;
        margin-bottom: 2rem;
    }
    
    .chat-message {
        background-color: #f0f2f6;
        padding: 1rem;
        border-radius: 10px;
        margin: 1rem 0;
        border-left: 4px solid #667eea;
        color: #2c3e50 !important;
    }
    
    .chat-message strong {
        color: #667eea !important;
    }
    
    .ai-response {
        background-color: #e8f4fd;
        padding: 1rem;
        border-radius: 10px;
        margin: 1rem 0;
        border-left: 4px solid #1f77b4;
        color: #2c3e50 !important;
    }
    
    .ai-response strong {
        color: #1f77b4 !important;
    }
    
    .interactive-question {
        background-color: #fff3cd;
        padding: 1rem;
        border-radius: 10px;
        margin: 1rem 0;
        border-left: 4px solid #ffc107;
    }
    
    .slide-preview {
        border: 2px solid #ddd;
        border-radius: 8px;
        padding: 1rem;
        margin: 0.5rem 0;
        background-color: #fafafa;
    }
    
    .slide-title {
        color: #2E86AB;
        font-weight: bold;
        font-size: 1.2em;
        margin-bottom: 0.5rem;
    }
    
    .slide-content {
        margin-left: 1rem;
    }
    
    .progress-indicator {
        background-color: #e9ecef;
        border-radius: 10px;
        height: 20px;
        margin: 1rem 0;
    }
    
    .progress-bar {
        background: linear-gradient(90deg, #28a745 0%, #20c997 100%);
        height: 100%;
        border-radius: 10px;
        transition: width 0.3s ease;
    }
    
    .theme-preview {
        border: 2px solid #ddd;
        border-radius: 8px;
        padding: 10px;
        margin: 5px;
        text-align: center;
        cursor: pointer;
    }
    
    .theme-preview:hover {
        border-color: #007bff;
    }
    
    .success-box {
        background-color: #d4edda;
        border: 1px solid #c3e6cb;
        border-radius: 5px;
        padding: 1rem;
        margin: 1rem 0;
    }
    
    .warning-box {
        background-color: #fff3cd;
        border: 1px solid #ffeaa7;
        border-radius: 5px;
        padding: 1rem;
        margin: 1rem 0;
    }
    
    .error-box {
        background-color: #f8d7da;
        border: 1px solid #f5c6cb;
        border-radius: 5px;
        padding: 1rem;
        margin: 1rem 0;
    }
    
    .feature-card {
        background: white;
        border-radius: 10px;
        padding: 1.5rem;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

class EnhancedPowerPointApp:
    """Enhanced main application class với interactive features"""
    
    def __init__(self):
        self.init_session_state()
        self.theme_system = ModernThemeSystem()
        
    def init_session_state(self):
        """Initialize enhanced session state variables"""
        if 'conversation_history' not in st.session_state:
            st.session_state.conversation_history = []
        
        if 'current_presentation' not in st.session_state:
            st.session_state.current_presentation = None
        
        if 'ai_generator' not in st.session_state:
            st.session_state.ai_generator = None
        
        if 'pp_generator' not in st.session_state:
            st.session_state.pp_generator = PowerPointGenerator()
        
        if 'presentation_data' not in st.session_state:
            st.session_state.presentation_data = None
        
        if 'editing_mode' not in st.session_state:
            st.session_state.editing_mode = False
        
        if 'dalle_generator' not in st.session_state:
            st.session_state.dalle_generator = None
        
        if 'enable_dalle' not in st.session_state:
            st.session_state.enable_dalle = True
        
        if 'selected_theme' not in st.session_state:
            st.session_state.selected_theme = 'education_pro'
        
        # Enhanced session state for interactive features
        if 'interactive_session' not in st.session_state:
            st.session_state.interactive_session = None
        
        if 'current_questions' not in st.session_state:
            st.session_state.current_questions = []
        
        # Enhanced PowerPoint Editor
        if 'enhanced_editor' not in st.session_state:
            st.session_state.enhanced_editor = PowerPointEditorModule()
        
        if 'user_answers' not in st.session_state:
            st.session_state.user_answers = {}
        
        if 'generation_phase' not in st.session_state:
            st.session_state.generation_phase = 'initial'  # initial, questions, generation, complete
        
        if 'auto_theme_enabled' not in st.session_state:
            st.session_state.auto_theme_enabled = True
    
    def setup_sidebar(self):
        """Setup enhanced sidebar với các cài đặt mới"""
        with st.sidebar:
            st.header("⚙️ Cài đặt AI Enhanced")
            
            # OpenAI API Key
            api_key = st.text_input(
                "🔑 OpenAI API Key",
                type="password",
                help="Nhập API key để sử dụng ChatGPT và DALL-E"
            )
            
            if api_key:
                if st.session_state.ai_generator is None:
                    try:
                        with st.spinner("🔌 Đang kết nối AI..."):
                            st.session_state.ai_generator = EnhancedAIContentGenerator(api_key)
                            st.session_state.dalle_generator = DALLEImageGenerator(api_key)
                        st.success("✅ Đã kết nối AI Enhanced + DALL-E!")
                    except Exception as e:
                        st.error(f"❌ Lỗi kết nối AI: {str(e)}")
                        st.error("Kiểm tra lại API key hoặc thử lại sau")
                else:
                    st.success("✅ AI đã kết nối sẵn sàng!")
            else:
                st.warning("⚠️ Cần API key để sử dụng tính năng AI")
                st.session_state.ai_generator = None
            
            st.divider()
            
            # Enhanced AI settings
            st.subheader("🤖 Cài đặt AI Enhanced")
            model_choice = st.selectbox(
                "Model",
                ["gpt-3.5-turbo", "gpt-4"],
                help="Chọn model ChatGPT"
            )
            
            interactive_mode = st.checkbox(
                "🗣️ Chế độ tương tác",
                value=True,
                help="AI sẽ hỏi câu hỏi để hiểu rõ nhu cầu"
            )
            
            # Enhanced DALL-E settings
            st.subheader("🎨 Cài đặt DALL-E Enhanced")
            enable_dalle = st.checkbox(
                "🖼️ Tự động tạo ảnh minh họa", 
                value=True,
                help="AI sẽ phân tích và tạo ảnh phù hợp cho từng slide"
            )
            
            image_quality = st.selectbox(
                "Chất lượng ảnh",
                ["standard", "hd"],
                help="Chất lượng ảnh DALL-E"
            )
            
            # Enhanced Theme settings
            st.subheader("🎨 Hệ thống Theme Thông minh")
            auto_theme = st.checkbox(
                "🎯 Tự động chọn theme",
                value=True,
                help="AI sẽ tự động chọn theme phù hợp với nội dung"
            )
            
            if not auto_theme:
                available_themes = self.theme_system.list_available_themes()
                selected_theme = st.selectbox(
                    "Template Theme",
                    options=list(available_themes.keys()),
                    format_func=lambda x: f"{available_themes[x]}",
                    help="Chọn theme thủ công"
                )
                st.session_state.selected_theme = selected_theme
            
            st.session_state.auto_theme_enabled = auto_theme
            
            st.divider()
            
            # Enhanced Quick actions
            st.subheader("⚡ Thao tác nhanh")
            
            if st.button("🔄 Reset Session"):
                self.reset_interactive_session()
                st.success("Đã reset session!")
            
            if st.session_state.presentation_data:
                if st.button("📊 Xem thống kê"):
                    self.show_presentation_stats()
            
            # Progress indicator
            if st.session_state.generation_phase != 'initial':
                self.render_progress_indicator()
    
    def render_header(self):
        """Render enhanced header"""
        st.markdown("""
        <div class="main-header">
            <h1>🎓 Enhanced AI PowerPoint Generator</h1>
            <p>Tạo presentation thông minh với AI tương tác, hình ảnh tự động và theme thông minh</p>
        </div>
        """, unsafe_allow_html=True)
        
        # Feature highlights
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.markdown("""
            <div class="feature-card">
                <h4>🗣️ Tương tác thông minh</h4>
                <p>AI hỏi câu hỏi để hiểu rõ nhu cầu</p>
            </div>
            """, unsafe_allow_html=True)
        
        with col2:
            st.markdown("""
            <div class="feature-card">
                <h4>🎨 Tạo ảnh tự động</h4>
                <p>DALL-E tạo ảnh minh họa phù hợp</p>
            </div>
            """, unsafe_allow_html=True)
        
        with col3:
            st.markdown("""
            <div class="feature-card">
                <h4>🎯 Theme thông minh</h4>
                <p>Tự động chọn theme phù hợp</p>
            </div>
            """, unsafe_allow_html=True)
        
        with col4:
            st.markdown("""
            <div class="feature-card">
                <h4>📊 Phân tích nội dung</h4>
                <p>AI phân tích và tối ưu hóa slides</p>
            </div>
            """, unsafe_allow_html=True)
    
    def render_interactive_chat_interface(self):
        """Render enhanced interactive chat interface"""
        st.subheader("💬 Trò chuyện với AI Assistant")
        
        # Chat history
        if st.session_state.conversation_history:
            for message in st.session_state.conversation_history:
                if message["role"] == "user":
                    st.markdown(f"""
                    <div class="chat-message">
                        <strong>👤 Bạn:</strong> {message["content"]}
                    </div>
                    """, unsafe_allow_html=True)
                else:
                    st.markdown(f"""
                    <div class="ai-response">
                        <strong>🤖 AI:</strong> {message["content"]}
                    </div>
                    """, unsafe_allow_html=True)
        
        # Current interactive questions
        if st.session_state.current_questions:
            self.render_interactive_questions()
        
        # Generation phase indicators
        if st.session_state.generation_phase == 'generation':
            st.markdown("""
            <div class="ai-response">
                <strong>🤖 AI:</strong> Đang tạo presentation với tất cả tính năng nâng cao...
                <br>• Phân tích nội dung cho hình ảnh
                <br>• Tự động chọn theme phù hợp  
                <br>• Tạo icons và visual elements
                <br>• Generate hình ảnh DALL-E
            </div>
            """, unsafe_allow_html=True)
            
            # Auto generate presentation
            self.auto_generate_presentation()
        
        # Initialize user_input variable
        user_input = ""
        
        # User input
        if st.session_state.generation_phase in ['initial', 'complete']:
            user_input = st.text_area(
                "Nhập yêu cầu của bạn:",
                placeholder="VD: Tạo bài giảng về Sinh học lớp 10 về cấu trúc tế bào...",
                height=100
            )
            
            col1, col2, col3 = st.columns([2, 1, 1])
            
            with col1:
                if st.button("🚀 Bắt đầu tạo presentation", type="primary"):
                    if user_input and st.session_state.ai_generator:
                        self.start_interactive_generation(user_input)
                    elif user_input:
                        st.warning("⚠️ Cần API key! Hoặc dùng 'Tạo mẫu' bên dưới.")
                    else:
                        st.warning("Vui lòng nhập yêu cầu!")
            
            with col2:
                if st.button("📝 Tạo nhanh"):
                    if user_input and st.session_state.ai_generator:
                        self.quick_generate(user_input)
                    elif user_input:
                        st.warning("⚠️ Cần API key! Hoặc dùng 'Tạo mẫu' bên dưới.")
                    else:
                        st.warning("Vui lòng nhập yêu cầu!")
            
            with col3:
                if st.button("💡 Gợi ý"):
                    self.show_suggestions()
            
            # Emergency fallback - tạo presentation mẫu không cần API
            if user_input and not st.session_state.ai_generator:
                st.markdown("---")
                st.markdown("#### 🆘 Không có API key?")
                if st.button("🎯 Tạo presentation mẫu", help="Tạo mẫu dựa trên yêu cầu, không cần API"):
                    self.create_sample_presentation(user_input)
    
    def render_interactive_questions(self):
        """Render interactive questions interface"""
        st.markdown("""
        <div class="interactive-question">
            <h4>🤖 AI cần thêm thông tin:</h4>
        </div>
        """, unsafe_allow_html=True)
        
        current_answers = {}
        
        for i, question in enumerate(st.session_state.current_questions):
            question_text = question.get("question", "")
            question_type = question.get("type", "text")
            question_key = question.get("key", f"q_{i}")
            required = question.get("required", False)
            
            st.markdown(f"**{question_text}** {'*' if required else ''}")
            
            if question_type == "text":
                answer = st.text_input(
                    f"Câu trả lời {i+1}:",
                    key=f"answer_{question_key}",
                    label_visibility="collapsed"
                )
                if answer:
                    current_answers[question_key] = answer
            
            elif question_type == "select":
                options = question.get("options", [])
                answer = st.selectbox(
                    f"Chọn {i+1}:",
                    options=[""] + options,
                    key=f"answer_{question_key}",
                    label_visibility="collapsed"
                )
                if answer:
                    current_answers[question_key] = answer
            
            elif question_type == "boolean":
                answer = st.checkbox(
                    "Có",
                    key=f"answer_{question_key}"
                )
                current_answers[question_key] = answer
        
        # Submit answers
        col1, col2 = st.columns([1, 3])
        
        with col1:
            if st.button("✅ Gửi câu trả lời"):
                self.process_interactive_answers(current_answers)
        
        with col2:
            if st.button("⏭️ Bỏ qua và tạo ngay"):
                self.skip_questions_and_generate()
    
    def start_interactive_generation(self, user_input: str):
        """Start interactive generation process"""
        try:
            # Add to conversation history
            st.session_state.conversation_history.append({
                "role": "user",
                "content": user_input
            })
            
            # Check if AI generator is available
            if not st.session_state.ai_generator:
                st.error("❌ Cần nhập API key trước!")
                return
            
            # Add loading indicator
            with st.spinner("🤖 AI đang phân tích yêu cầu..."):
                # Start interactive session with timeout
                try:
                    response = st.session_state.ai_generator.start_interactive_session(user_input)
                except Exception as api_error:
                    st.error(f"❌ Lỗi API: {str(api_error)}")
                    st.info("🔄 Đang chuyển sang chế độ tạo nhanh...")
                    self.quick_generate(user_input)
                    return
            
            if response and response.get("type") == "interactive_questions":
                st.session_state.current_questions = response.get("questions", [])
                st.session_state.generation_phase = "questions"
                
                ai_message = response.get("message", "Tôi cần thêm một số thông tin để tạo presentation tốt nhất cho bạn:")
                st.session_state.conversation_history.append({
                    "role": "assistant",
                    "content": ai_message
                })
                
                st.rerun()
            else:
                # Fallback to direct generation
                st.info("🔄 Chuyển sang tạo nhanh...")
                self.quick_generate(user_input)
                
        except Exception as e:
            st.error(f"❌ Lỗi khi bắt đầu: {str(e)}")
            st.info("🔄 Thử tạo nhanh thay thế...")
            self.quick_generate(user_input)
    
    def process_interactive_answers(self, answers: dict):
        """Process user answers from interactive questions"""
        try:
            # Update user answers
            st.session_state.user_answers.update(answers)
            
            # Process with AI
            response = st.session_state.ai_generator.process_user_answers(answers)
            
            if response.get("type") == "follow_up_questions":
                # More questions needed
                st.session_state.current_questions = response.get("questions", [])
                
                ai_message = response.get("message", "")
                st.session_state.conversation_history.append({
                    "role": "assistant",
                    "content": ai_message
                })
                
                st.rerun()
                
            elif response.get("type") == "generation_ready":
                # Ready to generate
                st.session_state.current_questions = []
                st.session_state.generation_phase = "generation"
                
                ai_message = response.get("message", "")
                st.session_state.conversation_history.append({
                    "role": "assistant", 
                    "content": ai_message
                })
                
                st.rerun()
                
        except Exception as e:
            st.error(f"Lỗi khi xử lý câu trả lời: {str(e)}")
    
    def auto_generate_presentation(self):
        """Auto generate presentation with all enhanced features"""
        try:
            with st.spinner("🎨 Đang tạo presentation với AI..."):
                # Generate enhanced presentation
                presentation_data = st.session_state.ai_generator.generate_enhanced_presentation()
                
                if presentation_data:
                    st.session_state.presentation_data = presentation_data
                    st.session_state.generation_phase = "complete"
                    
                    # Add success message
                    st.session_state.conversation_history.append({
                        "role": "assistant",
                        "content": f"✅ Đã tạo xong presentation '{presentation_data.get('title', '')}' với {len(presentation_data.get('slides', []))} slides! Hãy chuyển sang tab Preview để xem kết quả."
                    })
                    
                    st.success("🎉 Presentation đã được tạo thành công! Chuyển sang tab 'Preview' để xem.")
                    st.rerun()
                else:
                    st.error("❌ Không thể tạo presentation")
                    
        except Exception as e:
            st.error(f"Lỗi khi tạo presentation: {str(e)}")
            st.session_state.generation_phase = "complete"
    
    def quick_generate(self, user_input: str):
        """Quick generation without interactive questions"""
        try:
            with st.spinner("⚡ Tạo nhanh presentation..."):
                # Use fallback context
                fallback_context = {
                    "answers_collected": {
                        "topic": user_input,
                        "audience": "Học sinh",
                        "duration": "45 phút",
                        "content_depth": "Trung bình",
                        "presentation_style": "Giáo dục - Thân thiện",
                        "include_examples": True,
                        "include_images": st.session_state.enable_dalle
                    }
                }
                
                presentation_data = st.session_state.ai_generator.generate_enhanced_presentation(fallback_context)
                
                if presentation_data:
                    st.session_state.presentation_data = presentation_data
                    st.session_state.generation_phase = "complete"
                    
                    st.success("⚡ Tạo nhanh thành công! Chuyển sang tab 'Preview' để xem.")
                    st.rerun()
                    
        except Exception as e:
            st.error(f"Lỗi tạo nhanh: {str(e)}")
    
    def skip_questions_and_generate(self):
        """Skip remaining questions and generate with current info"""
        st.session_state.current_questions = []
        st.session_state.generation_phase = "generation"
        st.rerun()
    
    def show_suggestions(self):
        """Show example suggestions"""
        suggestions = [
            "Tạo bài giảng Sinh học lớp 10 về cấu trúc tế bào",
            "Presentation về Marketing Digital cho doanh nghiệp",
            "Bài thuyết trình về Trí tuệ nhân tạo và Machine Learning",
            "Giáo án Vật lý về sóng ánh sáng cho học sinh THPT",
            "Training về Kỹ năng giao tiếp cho nhân viên"
        ]
        
        st.markdown("### 💡 Gợi ý:")
        for suggestion in suggestions:
            if st.button(f"📝 {suggestion}", key=f"suggest_{suggestion[:20]}"):
                st.session_state.conversation_history.append({
                    "role": "user",
                    "content": suggestion
                })
                self.start_interactive_generation(suggestion)
    
    def create_sample_presentation(self, user_input: str):
        """Create sample presentation without API"""
        try:
            with st.spinner("🎯 Đang tạo presentation mẫu..."):
                # Create basic presentation structure
                sample_data = {
                    "title": f"Presentation về {user_input[:50]}",
                    "subtitle": "Được tạo bởi Enhanced AI PowerPoint Generator",
                    "author": "AI Assistant",
                    "template": "education",
                    "target_audience": "Học sinh/Nhân viên",
                    "estimated_duration": "30-45 phút",
                    "difficulty_level": "Trung bình",
                    "recommended_theme": {
                        "theme_name": "education_pro",
                        "auto_selected": False,
                        "reason": "Default theme for sample"
                    },
                    "image_suggestions": {
                        "total_slides": 4,
                        "slides_needing_images": [0, 1],
                        "image_concepts": {},
                        "priority_slides": []
                    },
                    "visual_elements": {
                        "primary_icon": "📊",
                        "color_scheme": "professional",
                        "visual_style": "clean_modern"
                    },
                    "slides": [
                        {
                            "type": "content",
                            "title": "Giới thiệu chủ đề",
                            "content": [
                                "Tổng quan về chủ đề",
                                "Mục tiêu của presentation",
                                "Nội dung chính sẽ trình bày"
                            ],
                            "icon": "🎯",
                            "needs_image": False,
                            "estimated_time": "5 phút"
                        },
                        {
                            "type": "content", 
                            "title": "Nội dung chính",
                            "content": [
                                "Điểm chính thứ nhất",
                                "Điểm chính thứ hai", 
                                "Điểm chính thứ ba"
                            ],
                            "icon": "📋",
                            "needs_image": True,
                            "image_concept": "Relevant illustration",
                            "estimated_time": "20 phút"
                        },
                        {
                            "type": "content",
                            "title": "Ví dụ và ứng dụng",
                            "content": [
                                "Ví dụ thực tế",
                                "Ứng dụng trong thực tiễn",
                                "Case study minh họa"
                            ],
                            "icon": "💡",
                            "needs_image": True,
                            "image_concept": "Example illustration",
                            "estimated_time": "15 phút"
                        },
                        {
                            "type": "content",
                            "title": "Kết luận",
                            "content": [
                                "Tóm tắt các điểm chính",
                                "Kết luận và đánh giá",
                                "Câu hỏi thảo luận"
                            ],
                            "icon": "🏆",
                            "needs_image": False,
                            "estimated_time": "5 phút"
                        }
                    ],
                    "generation_info": {
                        "model_used": "template_based",
                        "generated_at": datetime.now().isoformat(),
                        "interactive_session": False,
                        "features_used": ["template_generation", "basic_structure"]
                    }
                }
                
                st.session_state.presentation_data = sample_data
                st.session_state.generation_phase = "complete"
                
                # Add to conversation
                st.session_state.conversation_history.append({
                    "role": "assistant",
                    "content": "✅ Đã tạo presentation mẫu thành công! Đây là cấu trúc cơ bản, bạn có thể tùy chỉnh trong tab Preview và Download."
                })
                
                st.success("🎯 Tạo presentation mẫu thành công! Chuyển sang tab 'Preview' để xem.")
                st.rerun()
                
        except Exception as e:
            st.error(f"❌ Lỗi tạo mẫu: {str(e)}")
    
    def render_enhanced_presentation_preview(self):
        """Render enhanced presentation preview"""
        if not st.session_state.presentation_data:
            return
        
        data = st.session_state.presentation_data
        
        st.subheader("📋 Preview Presentation")
        
        # Enhanced presentation info
        col1, col2 = st.columns([2, 1])
        
        with col1:
            st.markdown(f"**📌 Tiêu đề:** {data.get('title', '')}")
            st.markdown(f"**📝 Phụ đề:** {data.get('subtitle', '')}")
            st.markdown(f"**👥 Đối tượng:** {data.get('target_audience', '')}")
            st.markdown(f"**⏱️ Thời gian:** {data.get('estimated_duration', '')}")
            st.markdown(f"**📊 Độ khó:** {data.get('difficulty_level', '')}")
        
        with col2:
            # Theme info
            theme_info = data.get('recommended_theme', {})
            if theme_info:
                st.markdown(f"**🎨 Theme:** {theme_info.get('theme_name', '')}")
                if theme_info.get('auto_selected'):
                    st.success("🎯 Tự động chọn theme")
            
            # Image info  
            image_suggestions = data.get('image_suggestions', {})
            if image_suggestions:
                slides_with_images = len(image_suggestions.get('slides_needing_images', []))
                st.markdown(f"**🖼️ Slides có ảnh:** {slides_with_images}")
        
        # Enhanced slide previews
        st.markdown("### 📑 Slides Preview")
        
        slides = data.get('slides', [])
        for i, slide in enumerate(slides):
            with st.expander(f"Slide {i+1}: {slide.get('title', '')} {slide.get('icon', '')}"):
                col1, col2 = st.columns([3, 1])
                
                with col1:
                    content = slide.get('content', [])
                    if content:
                        for point in content:
                            st.markdown(f"• {point}")
                    
                    # Two column content
                    left_content = slide.get('left_content', [])
                    right_content = slide.get('right_content', [])
                    if left_content or right_content:
                        lcol, rcol = st.columns(2)
                        with lcol:
                            for point in left_content:
                                st.markdown(f"• {point}")
                        with rcol:
                            for point in right_content:
                                st.markdown(f"• {point}")
                
                with col2:
                    st.markdown(f"**Loại:** {slide.get('type', '')}")
                    st.markdown(f"**Icon:** {slide.get('icon', 'N/A')}")
                    
                    if slide.get('needs_image'):
                        st.success("🖼️ Có ảnh")
                        concept = slide.get('image_concept', '')
                        if concept:
                            st.caption(f"Ý tưởng: {concept}")
                    else:
                        st.info("📝 Chỉ text")
                    
                    time_est = slide.get('estimated_time', '')
                    if time_est:
                        st.caption(f"⏱️ {time_est}")
    
    def render_progress_indicator(self):
        """Render progress indicator"""
        phase_map = {
            'initial': 0,
            'questions': 50,
            'generation': 80,
            'complete': 100
        }
        
        progress = phase_map.get(st.session_state.generation_phase, 0)
        
        st.markdown(f"""
        <div class="progress-indicator">
            <div class="progress-bar" style="width: {progress}%"></div>
        </div>
        <p style="text-align: center; margin: 0.5rem 0;">
            Tiến độ: {progress}% - {st.session_state.generation_phase.title()}
        </p>
        """, unsafe_allow_html=True)
    
    def show_presentation_stats(self):
        """Show enhanced presentation statistics"""
        if not st.session_state.presentation_data:
            return
        
        data = st.session_state.presentation_data
        
        with st.expander("📊 Thống kê chi tiết"):
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                total_slides = len(data.get('slides', []))
                st.metric("Total Slides", total_slides)
            
            with col2:
                images_count = len(data.get('image_suggestions', {}).get('slides_needing_images', []))
                st.metric("Slides với ảnh", images_count)
            
            with col3:
                total_content = sum(len(slide.get('content', [])) for slide in data.get('slides', []))
                st.metric("Tổng bullet points", total_content)
            
            with col4:
                duration = data.get('estimated_duration', '0')
                st.metric("Thời gian", duration)
            
            # Generation info
            gen_info = data.get('generation_info', {})
            if gen_info:
                st.markdown("**Thông tin tạo:**")
                st.json(gen_info)
    
    def reset_interactive_session(self):
        """Reset interactive session"""
        st.session_state.conversation_history = []
        st.session_state.current_questions = []
        st.session_state.user_answers = {}
        st.session_state.generation_phase = 'initial'
        st.session_state.interactive_session = None
        
        if st.session_state.ai_generator:
            st.session_state.ai_generator.current_context = {}
    
    def render_download_section(self):
        """Render enhanced download section"""
        if not st.session_state.presentation_data:
            return
        
        st.subheader("📥 Tải xuống Presentation")
        
        # Quick Edit button
        if st.button("🎨 Edit with Enhanced Editor", type="secondary", use_container_width=True, key="download_section_edit"):
            try:
                enhanced_editor = st.session_state.enhanced_editor
                result = enhanced_editor.start_editing(st.session_state.presentation_data)
                
                if result:
                    st.success("✅ Enhanced Editor đã khởi động!")
                    st.rerun()
                else:
                    st.error("❌ Không thể khởi động Enhanced Editor")
                    
            except Exception as e:
                st.error(f"❌ Lỗi khởi động Enhanced Editor: {str(e)}")
        
        st.markdown("---")
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            if st.button("📊 Tạo PowerPoint", type="primary"):
                try:
                    with st.spinner("🎨 Đang tạo file PowerPoint..."):
                        # Apply recommended theme if auto-selected
                        theme_info = st.session_state.presentation_data.get('recommended_theme', {})
                        if theme_info.get('auto_selected'):
                            selected_theme = theme_info.get('theme_name', 'education_pro')
                        else:
                            selected_theme = st.session_state.selected_theme
                        
                        # Generate PowerPoint with enhanced features
                        success = st.session_state.pp_generator.create_from_structured_data(
                            st.session_state.presentation_data
                        )
                        
                        if success:
                            pptx_buffer = st.session_state.pp_generator.save_to_buffer()
                        else:
                            pptx_buffer = None
                        
                        if pptx_buffer:
                            filename = f"{st.session_state.presentation_data.get('title', 'presentation')}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pptx"
                            
                            st.download_button(
                                label="⬇️ Download PowerPoint",
                                data=pptx_buffer,
                                file_name=filename,
                                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                            )
                        else:
                            st.error("❌ Không thể tạo file PowerPoint")
                            
                except Exception as e:
                    st.error(f"Lỗi khi tạo PowerPoint: {str(e)}")
        
        with col2:
            if st.button("📄 Export JSON"):
                json_data = json.dumps(st.session_state.presentation_data, indent=2, ensure_ascii=False)
                filename = f"{st.session_state.presentation_data.get('title', 'presentation')}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
                
                st.download_button(
                    label="⬇️ Download JSON",
                    data=json_data,
                    file_name=filename,
                    mime="application/json"
                )
        
        with col3:
            if st.button("🖼️ Xem ảnh đã tạo"):
                self.show_generated_images()
    
    def show_generated_images(self):
        """Show generated images"""
        if not st.session_state.presentation_data:
            return
        
        generated_images = st.session_state.presentation_data.get('generated_images', {})
        
        if generated_images:
            st.subheader("🖼️ Hình ảnh đã tạo")
            
            for slide_index, image_path in generated_images.items():
                slide = st.session_state.presentation_data['slides'][slide_index]
                st.markdown(f"**Slide {slide_index + 1}: {slide.get('title', '')}**")
                
                try:
                    from PIL import Image
                    image = Image.open(image_path)
                    st.image(image, caption=f"Ảnh cho slide {slide_index + 1}", width=300)
                except Exception as e:
                    st.error(f"Không thể hiển thị ảnh: {str(e)}")
        else:
            st.info("Chưa có ảnh nào được tạo")
    
    def run(self):
        """Run the enhanced application"""
        # Check if Enhanced Editor is in edit mode FIRST
        enhanced_editor = st.session_state.enhanced_editor
        
        if enhanced_editor.is_in_edit_mode():
            # Show exit button
            if st.button("🔙 Quay lại AI Generator", type="secondary", key="main_exit_editor"):
                enhanced_editor.exit_edit_mode()
                st.rerun()
            
            # Render Enhanced Editor
            try:
                enhanced_editor.render_editor_interface()
                return  # Exit early if in edit mode
            except Exception as e:
                st.error(f"❌ Lỗi Enhanced Editor: {str(e)}")
                st.code(traceback.format_exc())
                enhanced_editor.exit_edit_mode()
        
        self.setup_sidebar()
        self.render_header()
        
        # Main tabs
        tab1, tab2, tab3, tab4, tab5 = st.tabs(["💬 Chat AI", "📋 Preview", "🎨 Customize", "🎨 Editor", "📥 Download"])
        
        with tab1:
            self.render_interactive_chat_interface()
        
        with tab2:
            self.render_enhanced_presentation_preview()
        
        with tab3:
            if st.session_state.presentation_data:
                st.subheader("🎨 Tùy chỉnh Presentation")
                
                # Theme customization
                st.markdown("#### Chọn Theme")
                available_themes = self.theme_system.list_available_themes()
                
                cols = st.columns(3)
                for i, (theme_key, theme_name) in enumerate(available_themes.items()):
                    with cols[i % 3]:
                        if st.button(f"🎨 {theme_name}", key=f"theme_{theme_key}"):
                            st.session_state.selected_theme = theme_key
                            st.success(f"Đã chọn theme: {theme_name}")
                
                # Content editing (placeholder for future enhancement)
                st.markdown("#### Chỉnh sửa nội dung")
                st.info("Tính năng chỉnh sửa nội dung sẽ được phát triển trong phiên bản tiếp theo")
            else:
                st.info("Chưa có presentation để tùy chỉnh")
        
        with tab4:
            # Enhanced PowerPoint Editor Tab
            if st.session_state.presentation_data:
                st.subheader("🎨 Enhanced PowerPoint Editor")
                st.markdown("### Chỉnh sửa presentation với giao diện như PowerPoint thật!")
                
                # Editor info
                col1, col2 = st.columns([2, 1])
                
                with col1:
                    st.markdown("""
                    **✨ Tính năng Enhanced Editor:**
                    - 🌈 **5 Theme system đẹp** với gradient chuyên nghiệp
                    - 🎨 **Fabric.js editor** drag & drop như PowerPoint
                    - ⌨️ **Keyboard shortcuts** (Ctrl+C/V, Delete)
                    - 📝 **Text editing** với font, color, size
                    - 🔷 **Shapes & Images** với hiệu ứng đẹp
                    - 💾 **Export PPTX** trực tiếp
                    """)
                
                with col2:
                    # Quick stats
                    data = st.session_state.presentation_data
                    st.metric("📄 Slides", len(data.get('slides', [])))
                    st.metric("🎨 Theme", data.get('theme_hint', 'Default'))
                
                # Launch Enhanced Editor button
                st.markdown("---")
                
                col1, col2, col3 = st.columns([1, 2, 1])
                with col2:
                    if st.button("🚀 Launch Enhanced Editor", type="primary", use_container_width=True, key="editor_tab_launch"):
                        with st.spinner("🎨 Đang khởi động Enhanced Editor..."):
                            try:
                                enhanced_editor = st.session_state.enhanced_editor
                                result = enhanced_editor.start_editing(st.session_state.presentation_data)
                                
                                if result:
                                    st.success("✅ Enhanced Editor đã khởi động!")
                                    st.rerun()
                                else:
                                    st.error("❌ Không thể khởi động Enhanced Editor")
                                    
                            except Exception as e:
                                st.error(f"❌ Lỗi khởi động Enhanced Editor: {str(e)}")
                                st.code(traceback.format_exc())
                
                # Quick preview
                st.markdown("---")
                with st.expander("👀 Preview Presentation Data", expanded=False):
                    st.json(st.session_state.presentation_data)
                    
            else:
                st.info("🤖 Vui lòng tạo presentation với AI trước khi sử dụng Enhanced Editor")
                st.markdown("""
                **Hướng dẫn:**
                1. Vào tab **💬 Chat AI** 
                2. Tạo presentation với AI
                3. Quay lại tab **🎨 Editor** này
                4. Click **🚀 Launch Enhanced Editor**
                """)
        
        with tab5:
            self.render_download_section()

def main():
    """Main function để chạy ứng dụng"""
    try:
        app = EnhancedPowerPointApp()
        app.run()
    except Exception as e:
        st.error(f"Lỗi ứng dụng: {str(e)}")
        st.error(traceback.format_exc())

if __name__ == "__main__":
    main()