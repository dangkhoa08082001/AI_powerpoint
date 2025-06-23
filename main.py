# main_app.py
"""
Main Streamlit application - Ứng dụng chính để tạo PowerPoint với AI
"""

import streamlit as st
import json
from datetime import datetime
from io import BytesIO
import logging
import traceback

# Import custom modules
from powerpoint_generator import PowerPointGenerator
from ai_content_generator import AIContentGenerator

# Setup logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Page config
st.set_page_config(
    page_title="🎓 AI PowerPoint Generator",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS
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
    }
    
    .ai-response {
        background-color: #e8f4fd;
        padding: 1rem;
        border-radius: 10px;
        margin: 1rem 0;
        border-left: 4px solid #1f77b4;
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
</style>
""", unsafe_allow_html=True)

class PowerPointApp:
    """Main application class"""
    
    def __init__(self):
        self.init_session_state()
        
    def init_session_state(self):
        """Initialize session state variables"""
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
    
    def setup_sidebar(self):
        """Setup sidebar với các cài đặt"""
        with st.sidebar:
            st.header("⚙️ Cài đặt")
            
            # OpenAI API Key
            api_key = st.text_input(
                "🔑 OpenAI API Key",
                type="password",
                help="Nhập API key để sử dụng ChatGPT"
            )
            
            if api_key:
                if st.session_state.ai_generator is None or st.session_state.ai_generator.model != "gpt-3.5-turbo":
                    try:
                        st.session_state.ai_generator = AIContentGenerator(api_key)
                        st.success("✅ Đã kết nối AI thành công!")
                    except Exception as e:
                        st.error(f"❌ Lỗi kết nối AI: {str(e)}")
            else:
                st.warning("⚠️ Cần API key để sử dụng AI")
            
            st.divider()
            
            # Model settings
            st.subheader("🤖 Cài đặt AI")
            model_choice = st.selectbox(
                "Model",
                ["gpt-3.5-turbo", "gpt-4"],
                help="Chọn model ChatGPT"
            )
            
            include_examples = st.checkbox(
                "Thêm ví dụ thực tế",
                value=True,
                help="AI sẽ tự động thêm các ví dụ và case studies"
            )
            
            st.divider()
            
            # Quick actions
            st.subheader("⚡ Thao tác nhanh")
            
            if st.button("🗑️ Xóa lịch sử chat", use_container_width=True):
                st.session_state.conversation_history = []
                st.rerun()
            
            if st.button("🔄 Reset tất cả", use_container_width=True):
                for key in ['conversation_history', 'current_presentation', 'presentation_data', 'editing_mode']:
                    if key in st.session_state:
                        del st.session_state[key]
                st.rerun()
            
            # Display info
            if st.session_state.presentation_data:
                st.divider()
                st.subheader("📊 Thông tin presentation")
                data = st.session_state.presentation_data
                st.write(f"**Tiêu đề:** {data.get('title', 'N/A')}")
                st.write(f"**Slides:** {len(data.get('slides', []))}")
                st.write(f"**Template:** {data.get('template', 'N/A')}")
    
    def render_header(self):
        """Render main header"""
        st.markdown("""
        <div class="main-header">
            <h1>🎓 AI PowerPoint Generator</h1>
            <p>Tạo bài giảng PowerPoint chuyên nghiệp với AI trong vài giây</p>
        </div>
        """, unsafe_allow_html=True)
    
    def render_chat_interface(self):
        """Render chat interface"""
        st.subheader("💬 Chat với AI để tạo PowerPoint")
        
        # Display conversation history
        if st.session_state.conversation_history:
            st.write("**Lịch sử hội thoại:**")
            for i, msg in enumerate(st.session_state.conversation_history):
                if msg['role'] == 'user':
                    st.markdown(f"""
                    <div class="chat-message">
                        <strong>👤 Bạn:</strong> {msg['content']}
                    </div>
                    """, unsafe_allow_html=True)
                else:
                    st.markdown(f"""
                    <div class="ai-response">
                        <strong>🤖 AI:</strong> {msg['content']}
                    </div>
                    """, unsafe_allow_html=True)
        
        # Chat input
        user_input = st.text_area(
            "Nhập yêu cầu của bạn:",
            placeholder="Ví dụ: Tạo giúp tôi một bài giảng PowerPoint về Toán lớp 10 chủ đề phương trình bậc 2",
            height=100
        )
        
        col1, col2 = st.columns([3, 1])
        
        with col1:
            send_button = st.button("🚀 Gửi yêu cầu", type="primary", use_container_width=True)
        
        with col2:
            example_button = st.button("💡 Ví dụ", use_container_width=True)
        
        # Handle example button
        if example_button:
            example_requests = [
                "Tạo bài giảng PowerPoint về Lịch sử lớp 9 chủ đề Cách mạng tháng 8",
                "Làm presentation về Marketing Digital cho nhân viên công ty",
                "Tạo bài giảng Sinh học lớp 12 về di truyền học",
                "Làm slide thuyết trình về Machine Learning cơ bản"
            ]
            
            st.write("**Ví dụ các yêu cầu:**")
            for example in example_requests:
                if st.button(f"📝 {example}", key=f"example_{example[:20]}"):
                    st.session_state.example_request = example
                    st.rerun()
        
        # Handle example selection
        if 'example_request' in st.session_state:
            user_input = st.session_state.example_request
            del st.session_state.example_request
            send_button = True
        
        # Handle send button
        if send_button and user_input.strip():
            if st.session_state.ai_generator is None:
                st.error("❌ Vui lòng nhập OpenAI API key ở sidebar!")
                return
            
            # Add user message to history
            st.session_state.conversation_history.append({
                'role': 'user',
                'content': user_input,
                'timestamp': datetime.now()
            })
            
            # Process AI request
            self.process_ai_request(user_input)
    
    def process_ai_request(self, user_input: str):
        """Process AI request and generate presentation"""
        try:
            with st.spinner("🤖 AI đang xử lý yêu cầu của bạn..."):
                # Generate presentation data using AI
                presentation_data = st.session_state.ai_generator.create_presentation_from_chat(
                    user_input, 
                    include_examples=True
                )
                
                # Store presentation data
                st.session_state.presentation_data = presentation_data
                
                # Add AI response to history
                ai_response = f"Đã tạo thành công bài giảng '{presentation_data.get('title', 'Không có tiêu đề')}' với {len(presentation_data.get('slides', []))} slides."
                
                st.session_state.conversation_history.append({
                    'role': 'ai',
                    'content': ai_response,
                    'timestamp': datetime.now()
                })
                
                # Auto switch to preview mode
                st.session_state.editing_mode = True
                
                st.success("✅ Đã tạo thành công! Kiểm tra kết quả bên dưới.")
                st.rerun()
                
        except Exception as e:
            error_msg = f"Lỗi khi tạo presentation: {str(e)}"
            logger.error(f"AI processing error: {str(e)}")
            logger.error(traceback.format_exc())
            
            st.session_state.conversation_history.append({
                'role': 'ai',
                'content': error_msg,
                'timestamp': datetime.now()
            })
            
            st.error(error_msg)
    
    def render_presentation_preview(self):
        """Render presentation preview and editing interface"""
        if not st.session_state.presentation_data:
            return
        
        data = st.session_state.presentation_data
        
        st.subheader("📋 Xem trước và chỉnh sửa")
        
        # Presentation info
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.metric("📝 Tiêu đề", data.get('title', 'N/A'))
        
        with col2:
            st.metric("📊 Số slides", len(data.get('slides', [])))
        
        with col3:
            st.metric("🎨 Template", data.get('template', 'education'))
        
        # Edit presentation info
        with st.expander("✏️ Chỉnh sửa thông tin presentation", expanded=False):
            new_title = st.text_input("Tiêu đề:", value=data.get('title', ''))
            new_subtitle = st.text_input("Phụ đề:", value=data.get('subtitle', ''))
            new_author = st.text_input("Tác giả:", value=data.get('author', ''))
            new_template = st.selectbox("Template:", ['education', 'business', 'training'], 
                                      index=['education', 'business', 'training'].index(data.get('template', 'education')))
            
            if st.button("💾 Cập nhật thông tin"):
                st.session_state.presentation_data.update({
                    'title': new_title,
                    'subtitle': new_subtitle,
                    'author': new_author,
                    'template': new_template
                })
                st.success("✅ Đã cập nhật thông tin!")
                st.rerun()
        
        # Tabs for different views
        tab1, tab2, tab3 = st.tabs(["👀 Xem trước", "✏️ Chỉnh sửa", "🤖 AI Editor"])
        
        with tab1:
            self.render_slides_preview()
        
        with tab2:
            self.render_slides_editor()
        
        with tab3:
            self.render_ai_editor()
        
        # Download section
        st.divider()
        self.render_download_section()
    
    def render_slides_preview(self):
        """Render slides preview"""
        data = st.session_state.presentation_data
        slides = data.get('slides', [])
        
        if not slides:
            st.warning("Chưa có slides nào.")
            return
        
        st.write(f"**Xem trước {len(slides)} slides:**")
        
        for i, slide in enumerate(slides):
            with st.container():
                st.markdown(f"""
                <div class="slide-preview">
                    <div class="slide-title">Slide {i + 1}: {slide.get('title', 'Không có tiêu đề')}</div>
                    <div class="slide-content">
                """, unsafe_allow_html=True)
                
                content = slide.get('content', [])
                if isinstance(content, list):
                    for item in content:
                        st.write(f"• {item}")
                else:
                    st.write(content)
                
                # Handle special slide types
                if slide.get('type') == 'two_column':
                    col1, col2 = st.columns(2)
                    with col1:
                        st.write("**Cột trái:**")
                        for item in slide.get('left_content', []):
                            st.write(f"• {item}")
                    with col2:
                        st.write("**Cột phải:**")
                        for item in slide.get('right_content', []):
                            st.write(f"• {item}")
                
                st.markdown("</div></div>", unsafe_allow_html=True)
    
    def render_slides_editor(self):
        """Render slides editor"""
        data = st.session_state.presentation_data
        slides = data.get('slides', [])
        
        if not slides:
            st.warning("Chưa có slides nào để chỉnh sửa.")
            return
        
        st.write("**Chỉnh sửa từng slide:**")
        
        for i, slide in enumerate(slides):
            with st.expander(f"📝 Slide {i + 1}: {slide.get('title', 'Không có tiêu đề')}", expanded=False):
                # Edit title
                new_title = st.text_input(f"Tiêu đề slide {i + 1}:", 
                                        value=slide.get('title', ''), 
                                        key=f"slide_title_{i}")
                
                # Edit content
                if slide.get('type') == 'two_column':
                    col1, col2 = st.columns(2)
                    with col1:
                        st.write("**Cột trái:**")
                        left_content = '\n'.join(slide.get('left_content', []))
                        new_left = st.text_area(f"Nội dung cột trái slide {i + 1}:", 
                                              value=left_content, 
                                              key=f"slide_left_{i}")
                    with col2:
                        st.write("**Cột phải:**")
                        right_content = '\n'.join(slide.get('right_content', []))
                        new_right = st.text_area(f"Nội dung cột phải slide {i + 1}:", 
                                               value=right_content, 
                                               key=f"slide_right_{i}")
                else:
                    content = slide.get('content', [])
                    if isinstance(content, list):
                        content_text = '\n'.join(content)
                    else:
                        content_text = str(content)
                    
                    new_content = st.text_area(f"Nội dung slide {i + 1}:", 
                                             value=content_text, 
                                             height=150,
                                             key=f"slide_content_{i}")
                
                # Action buttons
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    if st.button(f"💾 Lưu slide {i + 1}", key=f"save_slide_{i}"):
                        # Update slide data
                        st.session_state.presentation_data['slides'][i]['title'] = new_title
                        
                        if slide.get('type') == 'two_column':
                            st.session_state.presentation_data['slides'][i]['left_content'] = new_left.split('\n')
                            st.session_state.presentation_data['slides'][i]['right_content'] = new_right.split('\n')
                        else:
                            st.session_state.presentation_data['slides'][i]['content'] = new_content.split('\n')
                        
                        st.success(f"✅ Đã lưu slide {i + 1}!")
                
                with col2:
                    if st.button(f"🗑️ Xóa slide {i + 1}", key=f"delete_slide_{i}"):
                        if len(slides) > 1:
                            del st.session_state.presentation_data['slides'][i]
                            st.success(f"✅ Đã xóa slide {i + 1}!")
                            st.rerun()
                        else:
                            st.error("❌ Không thể xóa slide cuối cùng!")
                
                with col3:
                    if st.button(f"📄 Nhân bản slide {i + 1}", key=f"duplicate_slide_{i}"):
                        import copy
                        new_slide = copy.deepcopy(slide)
                        new_slide['title'] += " (Bản sao)"
                        st.session_state.presentation_data['slides'].insert(i + 1, new_slide)
                        st.success(f"✅ Đã nhân bản slide {i + 1}!")
                        st.rerun()
        
        # Add new slide
        st.divider()
        with st.expander("➕ Thêm slide mới"):
            new_slide_title = st.text_input("Tiêu đề slide mới:")
            new_slide_type = st.selectbox("Loại slide:", ['content', 'two_column', 'image', 'chart', 'table'])
            new_slide_content = st.text_area("Nội dung slide mới:")
            
            if st.button("➕ Thêm slide"):
                new_slide = {
                    'type': new_slide_type,
                    'title': new_slide_title,
                    'content': new_slide_content.split('\n') if new_slide_content else []
                }
                st.session_state.presentation_data['slides'].append(new_slide)
                st.success("✅ Đã thêm slide mới!")
                st.rerun()
    
    def render_ai_editor(self):
        """Render AI-powered editor"""
        if st.session_state.ai_generator is None:
            st.warning("⚠️ Cần API key để sử dụng AI Editor")
            return
        
        st.write("**Chỉnh sửa với sự hỗ trợ của AI:**")
        
        # AI improvement suggestions
        user_feedback = st.text_area(
            "Nhập feedback để AI đề xuất cải thiện:",
            placeholder="Ví dụ: Thêm nhiều ví dụ thực tế hơn, slides quá dài, cần thêm hình ảnh..."
        )
        
        if st.button("🤖 Lấy đề xuất từ AI") and user_feedback:
            with st.spinner("AI đang phân tích và đề xuất cải thiện..."):
                try:
                    suggestions = st.session_state.ai_generator.suggest_improvements(
                        st.session_state.presentation_data, 
                        user_feedback
                    )
                    
                    st.write("**Đề xuất cải thiện từ AI:**")
                    
                    for suggestion in suggestions.get('suggestions', []):
                        st.markdown(f"""
                        <div class="ai-response">
                            <strong>🎯 {suggestion.get('type', 'modify').title()}:</strong> {suggestion.get('description', '')}<br>
                            <em>Lý do: {suggestion.get('reason', '')}</em>
                        </div>
                        """, unsafe_allow_html=True)
                    
                    priority = suggestions.get('priority', 'medium')
                    if priority == 'high':
                        st.error(f"⚠️ Ưu tiên cao: {suggestions.get('estimated_changes', '0')} thay đổi cần thiết")
                    elif priority == 'medium':
                        st.warning(f"📋 Ưu tiên trung bình: {suggestions.get('estimated_changes', '0')} thay đổi đề xuất")
                    else:
                        st.info(f"✅ Ưu tiên thấp: {suggestions.get('estimated_changes', '0')} thay đổi nhỏ")
                
                except Exception as e:
                    st.error(f"❌ Lỗi khi lấy đề xuất từ AI: {str(e)}")
        
        # Quick AI actions
        st.divider()
        st.write("**Thao tác nhanh với AI:**")
        
        col1, col2 = st.columns(2)
        
        with col1:
            if st.button("✨ Thêm ví dụ thực tế", use_container_width=True):
                # Get current topic from presentation
                title = st.session_state.presentation_data.get('title', '')
                if title:
                    with st.spinner("AI đang tạo ví dụ..."):
                        try:
                            enhanced_data = st.session_state.ai_generator.enhance_content_with_examples(
                                st.session_state.presentation_data, 
                                title
                            )
                            st.session_state.presentation_data = enhanced_data
                            st.success("✅ Đã thêm ví dụ thực tế!")
                            st.rerun()
                        except Exception as e:
                            st.error(f"❌ Lỗi khi thêm ví dụ: {str(e)}")
        
        with col2:
            if st.button("🔄 AI tối ưu nội dung", use_container_width=True):
                st.info("💡 Tính năng đang phát triển...")
    
    def render_download_section(self):
        """Render download section"""
        st.subheader("📥 Tải xuống PowerPoint")
        
        if not st.session_state.presentation_data:
            st.warning("Chưa có presentation để tải xuống.")
            return
        
        col1, col2 = st.columns([2, 1])
        
        with col1:
            download_button = st.button("🎯 Tạo và tải xuống PowerPoint", type="primary", use_container_width=True)
        
        with col2:
            # Save JSON button
            json_data = json.dumps(st.session_state.presentation_data, indent=2, ensure_ascii=False)
            st.download_button(
                label="💾 Tải JSON",
                data=json_data,
                file_name=f"presentation_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json",
                mime="application/json",
                use_container_width=True
            )
        
        if download_button:
            with st.spinner("🔄 Đang tạo file PowerPoint..."):
                try:
                    # Create PowerPoint using generator
                    success = st.session_state.pp_generator.create_from_structured_data(
                        st.session_state.presentation_data
                    )
                    
                    if success:
                        # Get PowerPoint buffer
                        pptx_buffer = st.session_state.pp_generator.save_to_buffer()
                        
                        if pptx_buffer:
                            # Generate filename
                            title = st.session_state.presentation_data.get('title', 'presentation')
                            safe_title = "".join(c for c in title if c.isalnum() or c in (' ', '-', '_')).rstrip()
                            filename = f"{safe_title}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pptx"
                            
                            # Download button
                            st.download_button(
                                label="📥 Tải xuống PowerPoint",
                                data=pptx_buffer,
                                file_name=filename,
                                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                                use_container_width=True
                            )
                            
                            st.success("✅ PowerPoint đã sẵn sàng để tải xuống!")
                            
                            # Show stats
                            st.info(f"📊 Đã tạo {st.session_state.pp_generator.get_slide_count()} slides")
                        else:
                            st.error("❌ Lỗi khi tạo file PowerPoint")
                    else:
                        st.error("❌ Không thể tạo PowerPoint từ dữ liệu hiện tại")
                
                except Exception as e:
                    logger.error(f"Download error: {str(e)}")
                    logger.error(traceback.format_exc())
                    st.error(f"❌ Lỗi khi tạo PowerPoint: {str(e)}")
    
    def run(self):
        """Run the main application"""
        try:
            # Setup sidebar
            self.setup_sidebar()
            
            # Render header
            self.render_header()
            
            # Main content
            if not st.session_state.presentation_data:
                # Show chat interface when no presentation
                self.render_chat_interface()
                
                # Show example gallery
                st.divider()
                st.subheader("💡 Ví dụ sử dụng")
                
                col1, col2 = st.columns(2)
                
                with col1:
                    st.markdown("""
                    **🎓 Giáo dục:**
                    - "Tạo bài giảng Toán lớp 10 về phương trình bậc 2"
                    - "Làm slide Lịch sử về Cách mạng tháng 8"
                    - "Bài giảng Sinh học về di truyền học"
                    """)
                
                with col2:
                    st.markdown("""
                    **💼 Doanh nghiệp:**
                    - "Presentation về Marketing Digital"
                    - "Slide training nhân viên mới"
                    - "Thuyết trình kế hoạch kinh doanh Q4"
                    """)
            
            else:
                # Show presentation preview and editing
                self.render_presentation_preview()
                
                # Option to start new presentation
                st.divider()
                if st.button("🆕 Tạo presentation mới"):
                    for key in ['presentation_data', 'editing_mode']:
                        if key in st.session_state:
                            del st.session_state[key]
                    st.rerun()
        
        except Exception as e:
            logger.error(f"Application error: {str(e)}")
            logger.error(traceback.format_exc())
            st.error(f"❌ Lỗi ứng dụng: {str(e)}")
            
            if st.button("🔄 Restart ứng dụng"):
                st.session_state.clear()
                st.rerun()


def main():
    """Main function"""
    app = PowerPointApp()
    app.run()


if __name__ == "__main__":
    main()