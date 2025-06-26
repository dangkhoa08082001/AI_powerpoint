# 🎓 AI PowerPoint Generator

🤖 **Automatically generate PowerPoint presentations using AI with DALL-E image generation**

[![Streamlit App](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://your-app-url.streamlit.app)
[![Python](https://img.shields.io/badge/python-3.7+-blue.svg)](https://python.org)
[![OpenAI](https://img.shields.io/badge/powered%20by-OpenAI-green.svg)](https://openai.com)

## ✨ Features

### 🤖 AI Content Generation
- **ChatGPT Integration**: Smart automatic lecture content generation
- **DALL-E Images**: Automatic illustration generation for each slide
- **Content Analysis**: AI analyzes topics and creates appropriate images

### 🖼️ Smart Image Processing
- **Two-column layout**: Image on left, content on right
- **Auto-resize**: Automatic image size adjustment
- **Fallback placeholder**: Display placeholder when image fails

### 📝 Content Processing
- **Smart truncation**: Intelligent truncation by words/sentences
- **Bullet points**: Automatic creation of 4-5 key points per slide
- **Template system**: Support for multiple templates (education, business)

## 🚀 Quick Start

### 1. Installation
```bash
pip install -r requirements.txt
```

### 2. Run Application
```bash
streamlit run main.py
```

### 3. Usage
1. Enter your OpenAI API Key in the sidebar
2. Type a request like: "Create a lecture with images about Cell Biology"
3. Wait for AI to generate presentation with images
4. Download the PowerPoint file

## 📁 Project Structure

```
AI_powerpoint/
├── main.py                    # 🎯 Main Streamlit application
├── ai_content_generator.py    # 🤖 AI content generation + DALL-E
├── powerpoint_generator.py    # 📊 PowerPoint file creation
├── theme_system.py           # 🎨 Theme and styling system
├── requirements.txt          # 📦 Dependencies
└── README.md                # 📖 Documentation
```

## 🔧 Environment Setup

Create a `.env` file in the root directory:
```
OPENAI_API_KEY=your_openai_api_key_here
```

## 🎨 Template Types

### Education (Default)
- Primary color: Blue (#2E86AB)
- Font size: Title 32pt, Content 18pt
- Best for: Lectures, educational content

### Business
- Primary color: Navy (#1565C0)
- Font size: Title 36pt, Content 20pt
- Best for: Business presentations

## 📚 Supported Subjects

- **Biology**: Biological diagrams, cells, DNA
- **Physics**: Physics diagrams, optics
- **Chemistry**: Molecular structures, reactions
- **Mathematics**: Geometry, algebra
- **Marketing**: Business infographics
- **History**: Historical illustrations, culture

## 🐛 Troubleshooting

### Common Issues

1. **"❌ API key required"**
   - Enter OpenAI API key in sidebar
   - Ensure API key is valid

2. **"❌ DALL-E error"**
   - Check OpenAI credits
   - Try again with different topic

3. **"Images not displaying"**
   - Normal! Will show placeholder
   - PowerPoint file still created successfully

## 📊 Technical Specifications

- **Python**: 3.7+
- **Streamlit**: Web interface
- **OpenAI**: GPT-3.5 + DALL-E
- **python-pptx**: PowerPoint generation

## 🌐 Deploy on Streamlit Cloud

1. Fork this repository
2. Connect to [Streamlit Cloud](https://streamlit.io/cloud)
3. Add your `OPENAI_API_KEY` to secrets
4. Deploy with one click!

## 🤝 Contributing

1. Fork the project
2. Create your feature branch: `git checkout -b feature/amazing-feature`
3. Commit your changes: `git commit -m 'Add amazing feature'`
4. Push to the branch: `git push origin feature/amazing-feature`
5. Open a Pull Request

## 📝 License

This project is open source and available under the [MIT License](LICENSE).

## 🙏 Acknowledgments

- **OpenAI** for GPT and DALL-E APIs
- **Streamlit** for the amazing web framework
- **python-pptx** for PowerPoint generation

---

**Made with ❤️ using OpenAI GPT + DALL-E** 