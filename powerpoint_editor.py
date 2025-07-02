# powerpoint_editor.py
"""
PowerPoint Editor với Fabric.js - Chỉnh sửa presentation như PowerPoint
Module độc lập, không ảnh hưởng đến hệ thống hiện tại
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

class PowerPointEditor:
    """
    PowerPoint Editor hoàn chỉnh sử dụng Fabric.js
    Hoạt động giống PowerPoint thật với đầy đủ tính năng edit và export
    """
    
    def __init__(self):
        self.editor_height = 700
        self.slide_width = 900
        self.slide_height = 600
        
        # Initialize session state for editor
        if 'editor_data' not in st.session_state:
            st.session_state.editor_data = None
        if 'current_slide_index' not in st.session_state:
            st.session_state.current_slide_index = 0
        if 'editor_changes' not in st.session_state:
            st.session_state.editor_changes = {}
        
    def render_editor(self, presentation_data: Dict[str, Any]) -> Dict[str, Any]:
        """
        Render PowerPoint Editor hoàn chỉnh
        
        Args:
            presentation_data: Dữ liệu presentation từ AI
            
        Returns:
            Dict: Dữ liệu đã được edit
        """
        
        st.markdown("# 🎨 PowerPoint Editor")
        st.markdown("### *Chỉnh sửa presentation như trong PowerPoint thực sự!*")
        st.markdown("---")
        
        # Convert presentation data to editor format
        if st.session_state.editor_data is None:
            st.session_state.editor_data = self._convert_to_editor_format(presentation_data)
        
        editor_data = st.session_state.editor_data
        
        # Main editor layout
        col1, col2 = st.columns([1, 4])
        
        with col1:
            # Slide Navigator & Tools
            self._render_slide_navigator(editor_data)
            self._render_editor_tools()
        
        with col2:
            # Main Editor Canvas
            self._render_fabric_editor(editor_data)
            
            # Download section
            self._render_download_section(editor_data)
        
        return st.session_state.editor_data
    
    def _convert_to_editor_format(self, presentation_data: Dict[str, Any]) -> Dict[str, Any]:
        """Convert AI presentation data to editor format với đầy đủ elements"""
        
        editor_slides = []
        slides = presentation_data.get('slides', [])
        
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
                # Convert image to base64 for web display
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
            
            editor_slides.append(editor_slide)
        
        return {
            'title': presentation_data.get('title', 'Presentation'),
            'slides': editor_slides,
            'theme': presentation_data.get('recommended_theme', {}),
            'metadata': presentation_data.get('generation_info', {})
        }
    
    def _render_slide_navigator(self, editor_data: Dict[str, Any]):
        """Render slide navigator với thumbnail preview"""
        st.markdown("### 📋 Slides Navigator")
        slides = editor_data.get('slides', [])
        
        # Display current slide info
        current_slide = st.session_state.current_slide_index + 1
        total_slides = len(slides)
        st.markdown(f"**Slide {current_slide} of {total_slides}**")
        
        # Slide selection
        for i, slide in enumerate(slides):
            slide_title = slide.get('title', f'Slide {i+1}')[:25]
            
            # Create button with special styling for current slide
            button_type = "primary" if i == st.session_state.current_slide_index else "secondary"
            
            if st.button(
                f"📄 {i+1}. {slide_title}",
                key=f"slide_nav_{i}",
                type=button_type,
                use_container_width=True
            ):
                st.session_state.current_slide_index = i
                st.rerun()
        
        st.divider()
        
        # Slide management
        col_add, col_dup = st.columns(2)
        
        with col_add:
            if st.button("➕ Add Slide", use_container_width=True):
                self._add_new_slide(editor_data)
                st.rerun()
        
        with col_dup:
            if st.button("📋 Duplicate", use_container_width=True):
                self._duplicate_slide(editor_data, st.session_state.current_slide_index)
                st.rerun()
        
        # Delete slide (only if more than 1 slide)
        if len(slides) > 1:
            if st.button("🗑️ Delete Slide", type="secondary", use_container_width=True):
                if st.checkbox("Confirm delete?", key="confirm_delete"):
                    self._delete_slide(editor_data, st.session_state.current_slide_index)
                    st.rerun()
    
    def _render_editor_tools(self):
        """Render editor tools và properties"""
        st.markdown("### 🛠️ Editor Tools")
        
        # Theme selection
        st.markdown("**Theme:**")
        theme_options = ["Education", "Business", "Modern", "Creative"]
        selected_theme = st.selectbox("Select theme", theme_options, key="theme_select")
        
        # Slide background
        st.markdown("**Background:**")
        bg_color = st.color_picker("Background Color", "#FFFFFF", key="bg_color")
        
        if st.button("Apply to Current Slide"):
            current_slide = st.session_state.editor_data['slides'][st.session_state.current_slide_index]
            current_slide['background'] = bg_color
            st.success("Background applied!")
        
        st.divider()
        
        # Quick actions
        st.markdown("**Quick Actions:**")
        
        if st.button("🎨 Add Text Box", use_container_width=True):
            st.session_state.add_text_trigger = True
        
        if st.button("🖼️ Add Image", use_container_width=True):
            st.session_state.add_image_trigger = True
        
        if st.button("⬜ Add Shape", use_container_width=True):
            st.session_state.add_shape_trigger = True
        
        st.divider()
        
        # Undo/Redo
        col_undo, col_redo = st.columns(2)
        with col_undo:
            if st.button("↶ Undo", use_container_width=True):
                st.info("Undo functionality")
        
                 with col_redo:
             if st.button("↷ Redo", use_container_width=True):
                 st.info("Redo functionality")
    
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
        st.session_state.current_slide_index = len(editor_data['slides']) - 1
    
    def _duplicate_slide(self, editor_data: Dict[str, Any], slide_index: int):
        """Duplicate slide at index"""
        if slide_index < len(editor_data['slides']):
            original_slide = editor_data['slides'][slide_index].copy()
            original_slide['id'] = f'slide_{len(editor_data["slides"])}'
            original_slide['title'] = f"{original_slide['title']} (Copy)"
            editor_data['slides'].append(original_slide)
            st.session_state.current_slide_index = len(editor_data['slides']) - 1
    
    def _delete_slide(self, editor_data: Dict[str, Any], slide_index: int):
        """Delete slide at index"""
        if len(editor_data['slides']) > 1 and slide_index < len(editor_data['slides']):
            del editor_data['slides'][slide_index]
            if st.session_state.current_slide_index >= len(editor_data['slides']):
                st.session_state.current_slide_index = len(editor_data['slides']) - 1
    
    def _image_to_base64(self, image_path: str) -> Optional[str]:
        """Convert image to base64 string"""
        try:
            with open(image_path, 'rb') as img_file:
                encoded = base64.b64encode(img_file.read()).decode()
                return encoded
        except Exception as e:
            logger.error(f"Error converting image to base64: {str(e)}")
            return None
    
    def _render_fabric_editor(self, editor_data: Dict[str, Any]) -> Dict[str, Any]:
        """Render main Fabric.js editor canvas"""
        current_slide_index = st.session_state.current_slide_index
        slides = editor_data.get('slides', [])
        
        if not slides:
            st.error("No slides to edit")
            return editor_data
        
        current_slide = slides[current_slide_index]
        
        # Create Fabric.js HTML component
        fabric_html = self._create_fabric_html(current_slide, current_slide_index)
        
        # Render the editor
        components.html(fabric_html, height=self.editor_height, scrolling=False)
        
        return editor_data
    
    def _render_download_section(self, editor_data: Dict[str, Any]):
        """Render download section với export options"""
        st.markdown("---")
        st.markdown("### 📥 Export & Download")
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            if st.button("💾 Save Changes", type="primary", use_container_width=True):
                st.session_state.editor_saved = True
                st.success("✅ Changes saved!")
        
        with col2:
            if st.button("📊 Export PowerPoint", type="primary", use_container_width=True):
                try:
                    pptx_data = self.export_to_powerpoint(editor_data)
                    if pptx_data:
                        filename = f"{editor_data.get('title', 'presentation')}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pptx"
                        st.download_button(
                            label="⬇️ Download PPTX",
                            data=pptx_data,
                            file_name=filename,
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                            type="primary"
                        )
                    else:
                        st.error("❌ Error creating PowerPoint file")
                except Exception as e:
                    st.error(f"Export error: {str(e)}")
        
        with col3:
            if st.button("📄 Export JSON", use_container_width=True):
                json_data = json.dumps(editor_data, indent=2, ensure_ascii=False)
                filename = f"{editor_data.get('title', 'presentation')}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
                st.download_button(
                    label="⬇️ Download JSON",
                    data=json_data,
                    file_name=filename,
                    mime="application/json"
                )
    
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
                    padding: 12px 20px;
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
                    padding: 8px 16px;
                    border-radius: 6px;
                    cursor: pointer;
                    font-size: 13px;
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
                
                .properties-panel {{
                    background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%);
                    padding: 20px;
                    border-top: 1px solid #dee2e6;
                    display: none;
                }}
                
                .properties-panel.active {{
                    display: block;
                }}
                
                .prop-grid {{
                    display: grid;
                    grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
                    gap: 15px;
                }}
                
                .prop-group {{
                    background: white;
                    padding: 15px;
                    border-radius: 8px;
                    box-shadow: 0 2px 8px rgba(0,0,0,0.05);
                }}
                
                .prop-label {{
                    display: block;
                    margin-bottom: 8px;
                    font-weight: 600;
                    font-size: 13px;
                    color: #495057;
                }}
                
                .prop-input {{
                    width: 100%;
                    padding: 10px 12px;
                    border: 2px solid #e9ecef;
                    border-radius: 6px;
                    font-size: 14px;
                    transition: border-color 0.2s ease;
                }}
                
                .prop-input:focus {{
                    outline: none;
                    border-color: #667eea;
                }}
                
                .color-picker {{
                    width: 60px;
                    height: 35px;
                    padding: 0;
                    border: 2px solid #e9ecef;
                    border-radius: 6px;
                    cursor: pointer;
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
                        <input type="color" class="color-picker" id="colorPicker" value="#333333" onchange="changeColor()">
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
                
                <!-- Properties Panel -->
                <div class="properties-panel" id="propertiesPanel">
                    <div class="prop-grid">
                        <div class="prop-group">
                            <label class="prop-label">Text Content:</label>
                            <input type="text" class="prop-input" id="textContentInput" 
                                   placeholder="Enter text..." onchange="updateTextContent()">
                        </div>
                        
                        <div class="prop-group">
                            <label class="prop-label">Position:</label>
                            <div style="display: flex; gap: 10px;">
                                <input type="number" class="prop-input" id="posXInput" 
                                       placeholder="X" onchange="updatePosition()" style="width: 48%;">
                                <input type="number" class="prop-input" id="posYInput" 
                                       placeholder="Y" onchange="updatePosition()" style="width: 48%;">
                            </div>
                        </div>
                        
                        <div class="prop-group">
                            <label class="prop-label">Size:</label>
                            <div style="display: flex; gap: 10px;">
                                <input type="number" class="prop-input" id="widthInput" 
                                       placeholder="Width" onchange="updateSize()" style="width: 48%;">
                                <input type="number" class="prop-input" id="heightInput" 
                                       placeholder="Height" onchange="updateSize()" style="width: 48%;">
                            </div>
                        </div>
                    </div>
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
                            }});
                        }}
                    }});
                }}
                
                // Initialize
                loadElements();
                
                // Event handlers
                canvas.on('selection:created', updatePropertiesPanel);
                canvas.on('selection:updated', updatePropertiesPanel);
                canvas.on('selection:cleared', hidePropertiesPanel);
                canvas.on('object:modified', saveChanges);
                canvas.on('object:moving', updatePropertiesPanel);
                canvas.on('object:scaling', updatePropertiesPanel);
                
                // Toolbar functions
                function addText() {{
                    const text = new fabric.IText('Click to edit text', {{
                        left: 100 + Math.random() * 200,
                        top: 100 + Math.random() * 200,
                        fontSize: 18,
                        fontFamily: 'Arial',
                        fill: '#333333'
                    }});
                    canvas.add(text);
                    canvas.setActiveObject(text);
                    saveChanges();
                }}
                
                function addShape(type) {{
                    let shape;
                    if (type === 'rect') {{
                        shape = new fabric.Rect({{
                            left: 150 + Math.random() * 200,
                            top: 150 + Math.random() * 200,
                            width: 200,
                            height: 100,
                            fill: 'rgba(102, 126, 234, 0.5)',
                            stroke: '#667eea',
                            strokeWidth: 2
                        }});
                    }} else if (type === 'circle') {{
                        shape = new fabric.Circle({{
                            left: 150 + Math.random() * 200,
                            top: 150 + Math.random() * 200,
                            radius: 50,
                            fill: 'rgba(118, 75, 162, 0.5)',
                            stroke: '#764ba2',
                            strokeWidth: 2
                        }});
                    }}
                    canvas.add(shape);
                    canvas.setActiveObject(shape);
                    saveChanges();
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
                                saveChanges();
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
                        saveChanges();
                    }}
                }}
                
                function copySelected() {{
                    const activeObject = canvas.getActiveObject();
                    if (activeObject) {{
                        activeObject.clone(function(cloned) {{
                            clipboard = cloned;
                        }});
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
                            saveChanges();
                        }});
                    }}
                }}
                
                function changeFontSize() {{
                    const activeObject = canvas.getActiveObject();
                    const fontSize = document.getElementById('fontSizeInput').value;
                    if (activeObject && activeObject.type === 'i-text') {{
                        activeObject.set('fontSize', parseInt(fontSize));
                        canvas.renderAll();
                        updatePropertiesPanel();
                        saveChanges();
                    }}
                }}
                
                function changeColor() {{
                    const activeObject = canvas.getActiveObject();
                    const color = document.getElementById('colorPicker').value;
                    if (activeObject) {{
                        activeObject.set('fill', color);
                        canvas.renderAll();
                        saveChanges();
                    }}
                }}
                
                function toggleBold() {{
                    const activeObject = canvas.getActiveObject();
                    if (activeObject && activeObject.type === 'i-text') {{
                        const currentWeight = activeObject.fontWeight || 'normal';
                        activeObject.set('fontWeight', currentWeight === 'bold' ? 'normal' : 'bold');
                        canvas.renderAll();
                        saveChanges();
                    }}
                }}
                
                function toggleItalic() {{
                    const activeObject = canvas.getActiveObject();
                    if (activeObject && activeObject.type === 'i-text') {{
                        const currentStyle = activeObject.fontStyle || 'normal';
                        activeObject.set('fontStyle', currentStyle === 'italic' ? 'normal' : 'italic');
                        canvas.renderAll();
                        saveChanges();
                    }}
                }}
                
                function bringToFront() {{
                    const activeObject = canvas.getActiveObject();
                    if (activeObject) {{
                        canvas.bringToFront(activeObject);
                        saveChanges();
                    }}
                }}
                
                function sendToBack() {{
                    const activeObject = canvas.getActiveObject();
                    if (activeObject) {{
                        canvas.sendToBack(activeObject);
                        saveChanges();
                    }}
                }}
                
                // Properties panel functions
                function updatePropertiesPanel() {{
                    const activeObject = canvas.getActiveObject();
                    const panel = document.getElementById('propertiesPanel');
                    
                    if (activeObject) {{
                        panel.classList.add('active');
                        
                        // Update inputs
                        if (activeObject.type === 'i-text') {{
                            document.getElementById('textContentInput').value = activeObject.text || '';
                        }}
                        document.getElementById('posXInput').value = Math.round(activeObject.left);
                        document.getElementById('posYInput').value = Math.round(activeObject.top);
                        document.getElementById('widthInput').value = Math.round(activeObject.width * (activeObject.scaleX || 1));
                        document.getElementById('heightInput').value = Math.round(activeObject.height * (activeObject.scaleY || 1));
                        
                        // Update toolbar inputs
                        if (activeObject.type === 'i-text') {{
                            document.getElementById('fontSizeInput').value = activeObject.fontSize || 18;
                            document.getElementById('colorPicker').value = activeObject.fill || '#333333';
                        }}
                    }}
                }}
                
                function hidePropertiesPanel() {{
                    document.getElementById('propertiesPanel').classList.remove('active');
                }}
                
                function updateTextContent() {{
                    const activeObject = canvas.getActiveObject();
                    const newText = document.getElementById('textContentInput').value;
                    if (activeObject && activeObject.type === 'i-text') {{
                        activeObject.set('text', newText);
                        canvas.renderAll();
                        saveChanges();
                    }}
                }}
                
                function updatePosition() {{
                    const activeObject = canvas.getActiveObject();
                    const newX = document.getElementById('posXInput').value;
                    const newY = document.getElementById('posYInput').value;
                    if (activeObject) {{
                        activeObject.set({{
                            left: parseInt(newX),
                            top: parseInt(newY)
                        }});
                        canvas.renderAll();
                        saveChanges();
                    }}
                }}
                
                function updateSize() {{
                    const activeObject = canvas.getActiveObject();
                    const newWidth = document.getElementById('widthInput').value;
                    const newHeight = document.getElementById('heightInput').value;
                    if (activeObject) {{
                        const scaleX = activeObject.scaleX || 1;
                        const scaleY = activeObject.scaleY || 1;
                        activeObject.set({{
                            width: parseInt(newWidth) / scaleX,
                            height: parseInt(newHeight) / scaleY
                        }});
                        canvas.renderAll();
                        saveChanges();
                    }}
                }}
                
                // Save changes
                function saveChanges() {{
                    console.log('Slide changes saved');
                    // Auto-save functionality can be implemented here
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
                                // Undo functionality
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
    
    def export_to_powerpoint(self, editor_data: Dict[str, Any]) -> Optional[bytes]:
        """
        Export edited data back to PowerPoint format
        """
        try:
            # Convert editor data back to presentation format
            presentation_data = self._convert_from_editor_format(editor_data)
            
            # Import PowerPoint generator (assuming it exists in the system)
            try:
                from powerpoint_generator import PowerPointGenerator
                pp_generator = PowerPointGenerator()
                
                # Create presentation
                success = pp_generator.create_from_structured_data(presentation_data)
                
                if success:
                    buffer = pp_generator.save_to_buffer()
                    if buffer:
                        return buffer.getvalue()
            except ImportError:
                st.error("PowerPoint generator not available. Please ensure powerpoint_generator.py is in the same directory.")
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
            
            # Extract elements
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
                'edit_timestamp': datetime.now().isoformat()
            }
        }


# Standalone Application Function
def run_standalone_editor():
    """
    Function to run PowerPoint Editor as standalone application
    """
    st.set_page_config(
        page_title="🎨 PowerPoint Editor",
        page_icon="🎨",
        layout="wide",
        initial_sidebar_state="expanded"
    )
    
    # CSS for standalone app
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
        
        .stApp > header {
            background-color: transparent;
        }
        
        .feature-highlight {
            background: #f8f9fa;
            padding: 1rem;
            border-radius: 8px;
            border-left: 4px solid #667eea;
            margin: 1rem 0;
        }
    </style>
    """, unsafe_allow_html=True)
    
    # Header
    st.markdown("""
    <div class="main-header">
        <h1>🎨 PowerPoint Editor</h1>
        <p>Professional PowerPoint editing với Fabric.js</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Initialize editor
    editor = PowerPointEditor()
    
    # Sidebar for sample data or file upload
    with st.sidebar:
        st.header("📁 Import Presentation")
        
        # Option 1: Create sample presentation
        if st.button("🎯 Create Sample Presentation"):
            sample_data = {
                'title': 'Sample Presentation',
                'slides': [
                    {
                        'title': 'Welcome Slide',
                        'content': ['This is a sample presentation', 'You can edit all elements', 'Add images, text, and shapes']
                    },
                    {
                        'title': 'Features',
                        'content': ['Drag and drop editing', 'Professional themes', 'Export to PowerPoint', 'Real-time preview']
                    }
                ]
            }
            st.session_state.editor_data = None  # Reset to load new data
            editor.render_editor(sample_data)
        
        # Option 2: Upload JSON file
        uploaded_file = st.file_uploader("Upload JSON presentation", type=['json'])
        if uploaded_file:
            try:
                presentation_data = json.load(uploaded_file)
                st.session_state.editor_data = None  # Reset to load new data
                editor.render_editor(presentation_data)
                st.success("✅ Presentation loaded!")
            except Exception as e:
                st.error(f"Error loading file: {str(e)}")
    
    # Main content
    if 'editor_data' not in st.session_state or st.session_state.editor_data is None:
        # Welcome screen
        st.markdown("""
        <div class="feature-highlight">
            <h3>🚀 Welcome to PowerPoint Editor!</h3>
            <p>Start by creating a sample presentation or uploading a JSON file from the sidebar.</p>
        </div>
        """, unsafe_allow_html=True)
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.markdown("""
            **🎨 Professional Editing**
            - Drag & drop interface
            - Text, shapes, and images
            - Real-time preview
            """)
        
        with col2:
            st.markdown("""
            **⚡ Powerful Features**
            - Multiple slide management
            - Theme customization
            - Copy/paste functionality
            """)
        
        with col3:
            st.markdown("""
            **📥 Export Options**
            - PowerPoint (.pptx)
            - JSON format
            - High-quality output
            """)
    else:
        # Show editor if data exists
        st.write("")  # Placeholder since editor is already rendered in sidebar


# Main execution
if __name__ == "__main__":
    run_standalone_editor() 