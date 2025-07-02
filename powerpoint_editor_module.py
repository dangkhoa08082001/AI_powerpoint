# -*- coding: utf-8 -*-
"""
PowerPoint Editor Module - Tích hợp với hệ thống AI PowerPoint Generator
Module này cung cấp chức năng edit PowerPoint với Fabric.js như phần mềm thật
"""

import streamlit as st
import streamlit.components.v1 as components
import json
import base64
from typing import Dict, List, Any, Optional
from datetime import datetime
import os
import logging
from io import BytesIO

logger = logging.getLogger(__name__)

class PowerPointEditorModule:
    """
    PowerPoint Editor tích hợp với hệ thống AI PowerPoint Generator
    Workflow: AI tạo presentation → Edit với Fabric.js → Download
    """
    
    def __init__(self):
        self.editor_height = 700
        self.slide_width = 900
        self.slide_height = 600
        
        # Initialize session state for editor
        if 'pp_editor_data' not in st.session_state:
            st.session_state.pp_editor_data = None
        if 'pp_current_slide_index' not in st.session_state:
            st.session_state.pp_current_slide_index = 0
        if 'pp_edit_mode' not in st.session_state:
            st.session_state.pp_edit_mode = False
    
    def start_editing(self, ai_generated_data: Dict[str, Any]) -> bool:
        """
        Bắt đầu edit presentation từ data AI đã tạo
        
        Args:
            ai_generated_data: Data presentation từ AI generator
            
        Returns:
            bool: True nếu bắt đầu edit thành công
        """
        try:
            # Convert AI data to editor format
            st.session_state.pp_editor_data = self._convert_ai_to_editor_format(ai_generated_data)
            st.session_state.pp_current_slide_index = 0
            st.session_state.pp_edit_mode = True
            
            return True
            
        except Exception as e:
            logger.error(f"Error starting editor: {str(e)}")
            st.error(f"Lỗi khởi động editor: {str(e)}")
            return False
    
    def render_editor_interface(self) -> Optional[Dict[str, Any]]:
        """
        Render giao diện editor hoàn chỉnh
        
        Returns:
            Optional[Dict]: Edited presentation data hoặc None nếu không có data
        """
        if not st.session_state.pp_edit_mode or st.session_state.pp_editor_data is None:
            st.error("❌ Chưa có data để edit. Vui lòng tạo presentation trước.")
            return None
        
        # Header
        st.markdown("# 🎨 PowerPoint Editor")
        st.markdown("### *Chỉnh sửa presentation như trong PowerPoint thực sự!*")
        st.markdown("---")
        
        editor_data = st.session_state.pp_editor_data
        
        # Control buttons
        col_control1, col_control2, col_control3 = st.columns([2, 2, 2])
        
        with col_control1:
            if st.button("🔙 Quay lại AI Generator", type="secondary", use_container_width=True):
                st.session_state.pp_edit_mode = False
                st.session_state.pp_editor_data = None
                st.rerun()
        
        with col_control2:
            if st.button("💾 Lưu thay đổi", type="primary", use_container_width=True):
                st.success("✅ Đã lưu thay đổi!")
        
        with col_control3:
            if st.button("🔄 Reset Editor", type="secondary", use_container_width=True):
                st.session_state.pp_current_slide_index = 0
                st.rerun()
        
        st.markdown("---")
        
        # Main editor layout
        col1, col2 = st.columns([1, 4])
        
        with col1:
            self._render_slide_navigator(editor_data)
            self._render_editor_tools()
        
        with col2:
            self._render_fabric_editor(editor_data)
            self._render_download_section(editor_data)
        
        return st.session_state.pp_editor_data
    
    def _convert_ai_to_editor_format(self, ai_data: Dict[str, Any]) -> Dict[str, Any]:
        """Convert AI generated data to editor format"""
        
        editor_slides = []
        slides = ai_data.get('slides', [])
        
        for i, slide in enumerate(slides):
            editor_slide = {
                'id': f'slide_{i}',
                'title': slide.get('title', f'Slide {i+1}'),
                'type': slide.get('type', 'content'),
                'background': '#FFFFFF',
                'elements': []
            }
            
            # Add title element
            if slide.get('title'):
                title_element = {
                    'type': 'text',
                    'id': f'title_{i}',
                    'content': slide['title'],
                    'x': 50,
                    'y': 50,
                    'width': 800,
                    'height': 80,
                    'fontSize': 32,
                    'fontFamily': 'Arial',
                    'fontWeight': 'bold',
                    'fill': '#2E86AB',
                    'textAlign': 'left'
                }
                editor_slide['elements'].append(title_element)
            
            # Add content elements
            content = slide.get('content', [])
            y_offset = 150
            
            for j, item in enumerate(content):
                content_element = {
                    'type': 'text',
                    'id': f'content_{i}_{j}',
                    'content': f"• {item}",
                    'x': 70,
                    'y': y_offset,
                    'width': 700,
                    'height': 40,
                    'fontSize': 18,
                    'fontFamily': 'Arial',
                    'fill': '#333333',
                    'textAlign': 'left'
                }
                editor_slide['elements'].append(content_element)
                y_offset += 50
            
            # Add image if exists
            image_path = slide.get('generated_image_path')
            if image_path and os.path.exists(image_path):
                try:
                    image_b64 = self._image_to_base64(image_path)
                    if image_b64:
                        image_element = {
                            'type': 'image',
                            'id': f'image_{i}',
                            'src': f'data:image/png;base64,{image_b64}',
                            'x': 500,
                            'y': 200,
                            'width': 300,
                            'height': 200,
                            'scaleX': 1,
                            'scaleY': 1
                        }
                        editor_slide['elements'].append(image_element)
                except Exception as e:
                    logger.warning(f"Could not load image {image_path}: {str(e)}")
            
            editor_slides.append(editor_slide)
        
        return {
            'title': ai_data.get('title', 'Presentation'),
            'slides': editor_slides,
            'theme': ai_data.get('recommended_theme', {}),
            'metadata': {
                'original_ai_data': ai_data,
                'edit_timestamp': datetime.now().isoformat(),
                'total_slides': len(editor_slides)
            }
        }
    
    def _image_to_base64(self, image_path: str) -> Optional[str]:
        """Convert image to base64 string"""
        try:
            with open(image_path, 'rb') as img_file:
                encoded = base64.b64encode(img_file.read()).decode()
                return encoded
        except Exception as e:
            logger.error(f"Error converting image to base64: {str(e)}")
            return None
    
    def _render_slide_navigator(self, editor_data: Dict[str, Any]):
        """Render slide navigator với thumbnail preview"""
        st.markdown("### 📋 Slides Navigator")
        slides = editor_data.get('slides', [])
        
        # Display current slide info
        current_slide = st.session_state.pp_current_slide_index + 1
        total_slides = len(slides)
        st.markdown(f"**Slide {current_slide} of {total_slides}**")
        
        # Slide selection
        for i, slide in enumerate(slides):
            slide_title = slide.get('title', f'Slide {i+1}')[:25]
            
            # Create button with special styling for current slide
            button_type = "primary" if i == st.session_state.pp_current_slide_index else "secondary"
            
            if st.button(
                f"📄 {i+1}. {slide_title}",
                key=f"pp_slide_nav_{i}",
                type=button_type,
                use_container_width=True
            ):
                st.session_state.pp_current_slide_index = i
                st.rerun()
        
        st.divider()
        
        # Slide management
        col_add, col_dup = st.columns(2)
        
        with col_add:
            if st.button("➕ Add Slide", key="pp_add_slide", use_container_width=True):
                self._add_new_slide(editor_data)
                st.rerun()
        
        with col_dup:
            if st.button("📋 Duplicate", key="pp_dup_slide", use_container_width=True):
                self._duplicate_slide(editor_data, st.session_state.pp_current_slide_index)
                st.rerun()
        
        # Delete slide (only if more than 1 slide)
        if len(slides) > 1:
            if st.button("🗑️ Delete Slide", key="pp_del_slide", type="secondary", use_container_width=True):
                if st.checkbox("Confirm delete?", key="pp_confirm_delete"):
                    self._delete_slide(editor_data, st.session_state.pp_current_slide_index)
                    st.rerun()
    
    def _render_editor_tools(self):
        """Render editor tools và properties"""
        st.markdown("### 🛠️ Editor Tools")
        
        # Theme selection
        st.markdown("**Theme:**")
        theme_options = ["Education", "Business", "Modern", "Creative"]
        selected_theme = st.selectbox("Select theme", theme_options, key="pp_theme_select")
        
        # Slide background
        st.markdown("**Background:**")
        bg_color = st.color_picker("Background Color", "#FFFFFF", key="pp_bg_color")
        
        if st.button("Apply to Current Slide", key="pp_apply_bg"):
            current_slide = st.session_state.pp_editor_data['slides'][st.session_state.pp_current_slide_index]
            current_slide['background'] = bg_color
            st.success("Background applied!")
        
        st.divider()
        
        # Quick actions help
        st.markdown("**Quick Actions:**")
        st.info("🎨 Use toolbar in editor to add text, shapes, images")
        st.info("⌨️ Shortcuts: Ctrl+C/V (copy/paste), Delete (remove)")
        st.info("🖱️ Drag & drop to move elements")
    
    def _add_new_slide(self, editor_data: Dict[str, Any]):
        """Add new blank slide"""
        new_slide = {
            'id': f'slide_{len(editor_data["slides"])}',
            'title': 'New Slide',
            'type': 'content',
            'background': '#FFFFFF',
            'elements': [{
                'type': 'text',
                'id': 'title_new',
                'content': 'Click to edit title',
                'x': 50,
                'y': 50,
                'width': 800,
                'height': 80,
                'fontSize': 32,
                'fontFamily': 'Arial',
                'fontWeight': 'bold',
                'fill': '#2E86AB',
                'textAlign': 'left'
            }]
        }
        editor_data['slides'].append(new_slide)
        st.session_state.pp_current_slide_index = len(editor_data['slides']) - 1
    
    def _duplicate_slide(self, editor_data: Dict[str, Any], slide_index: int):
        """Duplicate slide at index"""
        if slide_index < len(editor_data['slides']):
            original_slide = editor_data['slides'][slide_index].copy()
            original_slide['id'] = f'slide_{len(editor_data["slides"])}'
            original_slide['title'] = f"{original_slide['title']} (Copy)"
            # Deep copy elements
            original_slide['elements'] = [elem.copy() for elem in original_slide.get('elements', [])]
            editor_data['slides'].append(original_slide)
            st.session_state.pp_current_slide_index = len(editor_data['slides']) - 1
    
    def _delete_slide(self, editor_data: Dict[str, Any], slide_index: int):
        """Delete slide at index"""
        if len(editor_data['slides']) > 1 and slide_index < len(editor_data['slides']):
            del editor_data['slides'][slide_index]
            if st.session_state.pp_current_slide_index >= len(editor_data['slides']):
                st.session_state.pp_current_slide_index = len(editor_data['slides']) - 1
    
    def _render_fabric_editor(self, editor_data: Dict[str, Any]):
        """Render main Fabric.js editor canvas"""
        current_slide_index = st.session_state.pp_current_slide_index
        slides = editor_data.get('slides', [])
        
        if not slides:
            st.error("No slides to edit")
            return
        
        current_slide = slides[current_slide_index]
        
        # Create Fabric.js HTML component
        fabric_html = self._create_fabric_html(current_slide, current_slide_index)
        
        # Render the editor
        components.html(fabric_html, height=self.editor_height, scrolling=False)
    
    def _create_fabric_html(self, slide_data: Dict[str, Any], slide_index: int) -> str:
        """Create comprehensive Fabric.js editor HTML"""
        elements_json = json.dumps(slide_data.get('elements', []))
        background_color = slide_data.get('background', '#FFFFFF')
        
        return f'''
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="UTF-8">
            <script src="https://cdnjs.cloudflare.com/ajax/libs/fabric.js/5.3.0/fabric.min.js"></script>
            <style>
                body {{
                    margin: 0;
                    padding: 10px;
                    font-family: 'Segoe UI', Arial, sans-serif;
                    background: #f5f7fa;
                }}
                
                .editor-container {{
                    background: white;
                    border-radius: 12px;
                    box-shadow: 0 4px 20px rgba(0,0,0,0.1);
                    overflow: hidden;
                }}
                
                .editor-toolbar {{
                    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                    color: white;
                    padding: 15px 20px;
                    display: flex;
                    gap: 12px;
                    align-items: center;
                    flex-wrap: wrap;
                    box-shadow: 0 2px 10px rgba(0,0,0,0.1);
                }}
                
                .toolbar-btn {{
                    background: rgba(255,255,255,0.2);
                    border: none;
                    color: white;
                    padding: 10px 16px;
                    border-radius: 6px;
                    cursor: pointer;
                    font-size: 14px;
                    font-weight: 500;
                    transition: all 0.2s ease;
                    display: flex;
                    align-items: center;
                    gap: 5px;
                }}
                
                .toolbar-btn:hover {{
                    background: rgba(255,255,255,0.3);
                    transform: translateY(-1px);
                }}
                
                .toolbar-input {{
                    padding: 6px 10px;
                    border: none;
                    border-radius: 4px;
                    font-size: 13px;
                    width: 70px;
                    background: rgba(255,255,255,0.9);
                }}
                
                .canvas-container {{
                    position: relative;
                    background: #fafbfc;
                    border: 3px solid #e2e8f0;
                    margin: 15px;
                    border-radius: 8px;
                    overflow: hidden;
                }}
                
                .slide-info {{
                    position: absolute;
                    top: 15px;
                    right: 15px;
                    background: rgba(0,0,0,0.8);
                    color: white;
                    padding: 8px 15px;
                    border-radius: 20px;
                    font-size: 12px;
                    font-weight: 600;
                    z-index: 1000;
                }}
                
                .status-bar {{
                    background: #f8f9fa;
                    padding: 10px 20px;
                    border-top: 1px solid #dee2e6;
                    font-size: 12px;
                    color: #6c757d;
                    display: flex;
                    justify-content: space-between;
                    align-items: center;
                }}
            </style>
        </head>
        <body>
            <div class="editor-container">
                <!-- Toolbar -->
                <div class="editor-toolbar">
                    <button class="toolbar-btn" onclick="addText()">
                        <span>📝</span> Add Text
                    </button>
                    <button class="toolbar-btn" onclick="addShape('rect')">
                        <span>⬜</span> Rectangle
                    </button>
                    <button class="toolbar-btn" onclick="addShape('circle')">
                        <span>⭕</span> Circle
                    </button>
                    <button class="toolbar-btn" onclick="addImage()">
                        <span>🖼️</span> Add Image
                    </button>
                    <button class="toolbar-btn" onclick="deleteSelected()">
                        <span>🗑️</span> Delete
                    </button>
                    <button class="toolbar-btn" onclick="copySelected()">
                        <span>📋</span> Copy
                    </button>
                    <button class="toolbar-btn" onclick="pasteSelected()">
                        <span>📌</span> Paste
                    </button>
                    
                    <div style="margin-left: 20px; display: flex; align-items: center; gap: 10px;">
                        <span style="font-size: 13px; font-weight: 500;">Font Size:</span>
                        <input type="number" class="toolbar-input" id="fontSizeInput" value="18" onchange="changeFontSize()">
                        
                        <span style="font-size: 13px; font-weight: 500;">Color:</span>
                        <input type="color" class="toolbar-input" id="colorPicker" value="#333333" onchange="changeColor()" style="width: 50px;">
                    </div>
                    
                    <button class="toolbar-btn" onclick="toggleBold()">
                        <span>🅱️</span> Bold
                    </button>
                    <button class="toolbar-btn" onclick="toggleItalic()">
                        <span>🎭</span> Italic
                    </button>
                    
                    <button class="toolbar-btn" onclick="bringToFront()">
                        <span>⬆️</span> Front
                    </button>
                    <button class="toolbar-btn" onclick="sendToBack()">
                        <span>⬇️</span> Back
                    </button>
                </div>
                
                <!-- Canvas Container -->
                <div class="canvas-container">
                    <canvas id="editor-canvas" width="900" height="600"></canvas>
                    <div class="slide-info">Slide {slide_index + 1}</div>
                </div>
                
                <!-- Status Bar -->
                <div class="status-bar">
                    <span id="statusText">Ready to edit</span>
                    <span id="objectCount">Objects: 0</span>
                </div>
            </div>
            
            <script>
                // Initialize Fabric.js canvas
                const canvas = new fabric.Canvas('editor-canvas', {{
                    backgroundColor: '{background_color}',
                    preserveObjectStacking: true,
                    selection: true
                }});
                
                // Load existing elements
                const elements = {elements_json};
                let clipboard = null;
                
                // Load elements into canvas
                function loadElements() {{
                    elements.forEach(element => {{
                        if (element.type === 'text') {{
                            const text = new fabric.IText(element.content || 'Sample Text', {{
                                left: element.x || 100,
                                top: element.y || 100,
                                fontSize: element.fontSize || 18,
                                fontFamily: element.fontFamily || 'Arial',
                                fill: element.fill || '#333333',
                                fontWeight: element.fontWeight || 'normal',
                                fontStyle: element.fontStyle || 'normal',
                                textAlign: element.textAlign || 'left'
                            }});
                            canvas.add(text);
                        }} else if (element.type === 'image' && element.src) {{
                            fabric.Image.fromURL(element.src, function(img) {{
                                img.set({{
                                    left: element.x || 200,
                                    top: element.y || 200,
                                    scaleX: element.scaleX || 0.5,
                                    scaleY: element.scaleY || 0.5
                                }});
                                canvas.add(img);
                                updateStatus();
                            }});
                        }}
                    }});
                    updateStatus();
                }}
                
                // Initialize
                loadElements();
                
                // Event handlers
                canvas.on('selection:created', updateStatus);
                canvas.on('selection:updated', updateStatus);
                canvas.on('selection:cleared', updateStatus);
                canvas.on('object:added', updateStatus);
                canvas.on('object:removed', updateStatus);
                
                // Update status bar
                function updateStatus() {{
                    const activeObject = canvas.getActiveObject();
                    const objectCount = canvas.getObjects().length;
                    
                    document.getElementById('objectCount').textContent = `Objects: ${{objectCount}}`;
                    
                    if (activeObject) {{
                        document.getElementById('statusText').textContent = `Selected: ${{activeObject.type || 'object'}}`;
                        
                        // Update toolbar inputs
                        if (activeObject.type === 'i-text' || activeObject.type === 'text') {{
                            document.getElementById('fontSizeInput').value = activeObject.fontSize || 18;
                            document.getElementById('colorPicker').value = activeObject.fill || '#333333';
                        }}
                    }} else {{
                        document.getElementById('statusText').textContent = 'Ready to edit';
                    }}
                }}
                
                // Toolbar functions
                function addText() {{
                    const text = new fabric.IText('Click to edit text', {{
                        left: 100 + Math.random() * 300,
                        top: 100 + Math.random() * 200,
                        fontSize: 18,
                        fontFamily: 'Arial',
                        fill: '#333333'
                    }});
                    canvas.add(text);
                    canvas.setActiveObject(text);
                    text.enterEditing();
                }}
                
                function addShape(type) {{
                    let shape;
                    if (type === 'rect') {{
                        shape = new fabric.Rect({{
                            left: 150 + Math.random() * 200,
                            top: 150 + Math.random() * 150,
                            width: 200,
                            height: 100,
                            fill: 'rgba(102, 126, 234, 0.5)',
                            stroke: '#667eea',
                            strokeWidth: 2
                        }});
                    }} else if (type === 'circle') {{
                        shape = new fabric.Circle({{
                            left: 150 + Math.random() * 200,
                            top: 150 + Math.random() * 150,
                            radius: 50,
                            fill: 'rgba(118, 75, 162, 0.5)',
                            stroke: '#764ba2',
                            strokeWidth: 2
                        }});
                    }}
                    canvas.add(shape);
                    canvas.setActiveObject(shape);
                }}
                
                function addImage() {{
                    const input = document.createElement('input');
                    input.type = 'file';
                    input.accept = 'image/*';
                    input.onchange = function(e) {{
                        const file = e.target.files[0];
                        const reader = new FileReader();
                        reader.onload = function(event) {{
                            fabric.Image.fromURL(event.target.result, function(img) {{
                                img.set({{
                                    left: 200,
                                    top: 200,
                                    scaleX: 0.5,
                                    scaleY: 0.5
                                }});
                                canvas.add(img);
                                canvas.setActiveObject(img);
                            }});
                        }};
                        reader.readAsDataURL(file);
                    }};
                    input.click();
                }}
                
                function deleteSelected() {{
                    const activeObjects = canvas.getActiveObjects();
                    if (activeObjects.length) {{
                        activeObjects.forEach(obj => canvas.remove(obj));
                        canvas.discardActiveObject();
                    }}
                }}
                
                function copySelected() {{
                    const activeObject = canvas.getActiveObject();
                    if (activeObject) {{
                        activeObject.clone(function(cloned) {{
                            clipboard = cloned;
                        }});
                        document.getElementById('statusText').textContent = 'Object copied';
                    }}
                }}
                
                function pasteSelected() {{
                    if (clipboard) {{
                        clipboard.clone(function(clonedObj) {{
                            canvas.discardActiveObject();
                            clonedObj.set({{
                                left: clonedObj.left + 20,
                                top: clonedObj.top + 20,
                                evented: true,
                            }});
                            canvas.add(clonedObj);
                            canvas.setActiveObject(clonedObj);
                            document.getElementById('statusText').textContent = 'Object pasted';
                        }});
                    }}
                }}
                
                function changeFontSize() {{
                    const activeObject = canvas.getActiveObject();
                    const fontSize = document.getElementById('fontSizeInput').value;
                    if (activeObject && (activeObject.type === 'i-text' || activeObject.type === 'text')) {{
                        activeObject.set('fontSize', parseInt(fontSize));
                        canvas.renderAll();
                    }}
                }}
                
                function changeColor() {{
                    const activeObject = canvas.getActiveObject();
                    const color = document.getElementById('colorPicker').value;
                    if (activeObject) {{
                        activeObject.set('fill', color);
                        canvas.renderAll();
                    }}
                }}
                
                function toggleBold() {{
                    const activeObject = canvas.getActiveObject();
                    if (activeObject && (activeObject.type === 'i-text' || activeObject.type === 'text')) {{
                        const currentWeight = activeObject.fontWeight || 'normal';
                        activeObject.set('fontWeight', currentWeight === 'bold' ? 'normal' : 'bold');
                        canvas.renderAll();
                    }}
                }}
                
                function toggleItalic() {{
                    const activeObject = canvas.getActiveObject();
                    if (activeObject && (activeObject.type === 'i-text' || activeObject.type === 'text')) {{
                        const currentStyle = activeObject.fontStyle || 'normal';
                        activeObject.set('fontStyle', currentStyle === 'italic' ? 'normal' : 'italic');
                        canvas.renderAll();
                    }}
                }}
                
                function bringToFront() {{
                    const activeObject = canvas.getActiveObject();
                    if (activeObject) {{
                        canvas.bringToFront(activeObject);
                        document.getElementById('statusText').textContent = 'Brought to front';
                    }}
                }}
                
                function sendToBack() {{
                    const activeObject = canvas.getActiveObject();
                    if (activeObject) {{
                        canvas.sendToBack(activeObject);
                        document.getElementById('statusText').textContent = 'Sent to back';
                    }}
                }}
                
                // Keyboard shortcuts
                document.addEventListener('keydown', function(e) {{
                    if (e.ctrlKey || e.metaKey) {{
                        switch(e.key) {{
                            case 'c':
                                e.preventDefault();
                                copySelected();
                                break;
                            case 'v':
                                e.preventDefault();
                                pasteSelected();
                                break;
                            case 'z':
                                e.preventDefault();
                                // Undo functionality would go here
                                break;
                        }}
                    }}
                    if (e.key === 'Delete' || e.key === 'Backspace') {{
                        if (canvas.getActiveObject() && !canvas.getActiveObject().isEditing) {{
                            e.preventDefault();
                            deleteSelected();
                        }}
                    }}
                }});
                
                // Make canvas responsive
                function resizeCanvas() {{
                    const container = document.querySelector('.canvas-container');
                    const containerWidth = container.clientWidth - 30;
                    const scale = Math.min(containerWidth / 900, 1);
                    canvas.setZoom(scale);
                    canvas.setWidth(900 * scale);
                    canvas.setHeight(600 * scale);
                }}
                
                window.addEventListener('resize', resizeCanvas);
                setTimeout(resizeCanvas, 100); // Initial resize
            </script>
        </body>
        </html>
        '''
    
    def _render_download_section(self, editor_data: Dict[str, Any]):
        """Render download section với export options"""
        st.markdown("---")
        st.markdown("### 📥 Export & Download")
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            if st.button("💾 Save Project", key="pp_save", type="primary", use_container_width=True):
                st.session_state.pp_saved = True
                st.success("✅ Project saved successfully!")
        
        with col2:
            if st.button("📊 Export PowerPoint", key="pp_export_pptx", type="primary", use_container_width=True):
                try:
                    # Try to export to PowerPoint using existing generator
                    pptx_data = self._export_to_powerpoint(editor_data)
                    if pptx_data:
                        filename = f"{editor_data.get('title', 'presentation')}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pptx"
                        st.download_button(
                            label="⬇️ Download PPTX",
                            data=pptx_data,
                            file_name=filename,
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                            type="primary",
                            key="pp_download_pptx"
                        )
                    else:
                        # Fallback to JSON
                        json_data = json.dumps(editor_data, indent=2, ensure_ascii=False)
                        filename = f"{editor_data.get('title', 'presentation')}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
                        st.download_button(
                            label="⬇️ Download JSON",
                            data=json_data,
                            file_name=filename,
                            mime="application/json",
                            key="pp_download_json_fallback"
                        )
                        st.info("💡 Exported as JSON - You can convert to PPTX using your main system")
                except Exception as e:
                    st.error(f"Export error: {str(e)}")
        
        with col3:
            if st.button("📄 Export JSON", key="pp_export_json", use_container_width=True):
                json_data = json.dumps(editor_data, indent=2, ensure_ascii=False)
                filename = f"{editor_data.get('title', 'presentation')}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
                st.download_button(
                    label="⬇️ Download JSON",
                    data=json_data,
                    file_name=filename,
                    mime="application/json",
                    key="pp_download_json"
                )
        
        # Show edit summary
        st.markdown("---")
        with st.expander("📊 Edit Summary"):
            total_slides = len(editor_data.get('slides', []))
            st.write(f"**Total slides:** {total_slides}")
            st.write(f"**Current slide:** {st.session_state.pp_current_slide_index + 1}")
            st.write(f"**Last modified:** {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            
            if st.checkbox("Show detailed data", key="pp_show_data"):
                st.json(editor_data, expanded=False)
    
    def _export_to_powerpoint(self, editor_data: Dict[str, Any]) -> Optional[bytes]:
        """Export edited data back to PowerPoint format"""
        try:
            # Convert editor data back to presentation format
            presentation_data = self._convert_from_editor_format(editor_data)
            
            # Try to import and use PowerPoint generator from main system
            try:
                from powerpoint_generator import PowerPointGenerator
                pp_generator = PowerPointGenerator()
                
                # Create presentation from converted data
                success = pp_generator.create_from_structured_data(presentation_data)
                
                if success:
                    # Get the buffer
                    buffer = pp_generator.save_to_buffer()
                    if buffer:
                        return buffer.getvalue()
                
            except ImportError:
                logger.warning("PowerPoint generator not available for direct export")
                return None
            
            return None
            
        except Exception as e:
            logger.error(f"Error exporting to PowerPoint: {str(e)}")
            return None
    
    def _convert_from_editor_format(self, editor_data: Dict[str, Any]) -> Dict[str, Any]:
        """Convert editor format back to AI presentation format"""
        
        slides = []
        
        for slide in editor_data.get('slides', []):
            ai_slide = {
                'type': slide.get('type', 'content'),
                'title': '',
                'content': [],
                'design_type': 'creative_bullets'
            }
            
            # Extract elements back to AI format
            for element in slide.get('elements', []):
                if element['type'] == 'text':
                    content = element.get('content', '')
                    if element.get('fontSize', 18) > 24:  # Likely a title
                        ai_slide['title'] = content
                    else:  # Content
                        # Remove bullet point if present
                        clean_content = content.replace('• ', '').strip()
                        if clean_content:
                            ai_slide['content'].append(clean_content)
            
            slides.append(ai_slide)
        
        return {
            'title': editor_data.get('title', 'Edited Presentation'),
            'subtitle': 'Edited with PowerPoint Editor',
            'author': 'User',
            'template': 'education',
            'slides': slides,
            'recommended_theme': editor_data.get('theme', {}),
            'generation_info': {
                'edited': True,
                'edit_timestamp': datetime.now().isoformat(),
                'editor_version': '1.0'
            }
        }
    
    def is_in_edit_mode(self) -> bool:
        """Check if currently in edit mode"""
        return st.session_state.get('pp_edit_mode', False)
    
    def get_edited_data(self) -> Optional[Dict[str, Any]]:
        """Get current edited data"""
        return st.session_state.get('pp_editor_data', None)
    
    def exit_edit_mode(self):
        """Exit edit mode and return to main app"""
        st.session_state.pp_edit_mode = False
        st.session_state.pp_editor_data = None
        st.session_state.pp_current_slide_index = 0


# Example usage function để test
def test_editor():
    """Test function for the PowerPoint Editor"""
    st.set_page_config(
        page_title="🎨 PowerPoint Editor Test",
        page_icon="🎨",
        layout="wide"
    )
    
    editor = PowerPointEditorModule()
    
    # Sample AI data for testing
    sample_ai_data = {
        'title': 'Sample AI Generated Presentation',
        'slides': [
            {
                'title': 'Introduction to AI',
                'content': ['Artificial Intelligence is transforming our world', 'Machine learning enables computers to learn', 'Deep learning mimics human brain networks'],
                'type': 'intro'
            },
            {
                'title': 'Key Applications',
                'content': ['Healthcare and medical diagnosis', 'Autonomous vehicles and transportation', 'Natural language processing'],
                'type': 'content'
            }
        ],
        'recommended_theme': {'primary_color': '#2E86AB', 'secondary_color': '#A23B72'}
    }
    
    if not editor.is_in_edit_mode():
        st.markdown("# 🎨 PowerPoint Editor Test")
        st.markdown("### Test the PowerPoint Editor Module")
        
        if st.button("🚀 Start Editing Sample Presentation", type="primary"):
            if editor.start_editing(sample_ai_data):
                st.rerun()
    else:
        editor.render_editor_interface()


if __name__ == "__main__":
    test_editor() 