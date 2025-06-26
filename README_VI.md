# 🎓 AI PowerPoint Generator - Đã Tối Ưu

## 📋 Mô tả

AI PowerPoint Generator là ứng dụng tự động tạo bài thuyết trình PowerPoint sử dụng trí tuệ nhân tạo với khả năng tạo ảnh DALL-E. Chương trình đã được tối ưu hóa với interface hoàn toàn bằng tiếng Việt.

## ✨ Tính năng chính

### 🤖 Tạo nội dung AI
- **ChatGPT Integration**: Tự động tạo nội dung bài giảng thông minh
- **DALL-E Images**: Tạo ảnh minh họa tự động cho từng slide
- **Phân tích nội dung**: AI phân tích chủ đề và tạo ảnh phù hợp

### 🖼️ Xử lý ảnh thông minh
- **Layout hai cột**: Ảnh bên trái, nội dung bên phải
- **Auto-resize**: Tự động điều chỉnh kích thước ảnh
- **Fallback placeholder**: Hiển thị placeholder khi ảnh lỗi

### 📝 Xử lý nội dung
- **Smart truncation**: Cắt ngắn thông minh theo từ/câu
- **Bullet points**: Tự động tạo 4-5 điểm chính cho mỗi slide
- **Template system**: Hỗ trợ nhiều template (education, business)

## 🚀 Cài đặt và sử dụng

### 1. Cài đặt dependencies
```bash
pip install -r requirements.txt
```

### 2. Chạy ứng dụng
```bash
streamlit run main.py
```

### 3. Sử dụng
1. Nhập OpenAI API Key vào sidebar
2. Gõ yêu cầu như: "Tạo bài giảng có hình ảnh về Sinh học tế bào"
3. Chờ AI tạo presentation với ảnh
4. Tải xuống file PowerPoint

## 📁 Cấu trúc Project (Đã tối ưu)

```
duan_pp/
├── main.py                    # 🎯 Ứng dụng Streamlit chính
├── ai_content_generator.py    # 🤖 AI tạo nội dung + DALL-E
├── powerpoint_generator.py    # 📊 Tạo file PowerPoint
├── requirements.txt           # 📦 Dependencies
├── .env                       # 🔑 API keys (tạo file này)
├── .gitignore                 # 🚫 Git ignore
├── dalle_config.json          # ⚙️ Cấu hình DALL-E
├── dalle_images/              # 🖼️ Ảnh DALL-E đã tạo
├── images/                    # 📸 Ảnh curated
├── DEPLOYMENT_CHECKLIST.md    # 📋 Checklist deploy
└── README_VI.md              # 📖 Hướng dẫn tiếng Việt
```

## 🔧 File .env cần tạo

Tạo file `.env` trong thư mục gốc:
```
OPENAI_API_KEY=your_openai_api_key_here
```

## 🎨 Các loại template

### Education (Mặc định)
- Màu chủ đạo: Xanh dương (#2E86AB)  
- Font size: Title 32pt, Content 18pt
- Phù hợp: Bài giảng, giáo dục

### Business
- Màu chủ đạo: Xanh navy (#1565C0)
- Font size: Title 36pt, Content 20pt  
- Phù hợp: Thuyết trình doanh nghiệp

## 📚 Hướng dẫn sử dụng chi tiết

### Tạo bài giảng với ảnh
```
Ví dụ input: "Tạo bài giảng có hình ảnh về Sinh học tế bào"

Kết quả:
- 6 slides tự động
- 4 slides có ảnh DALL-E
- Nội dung được phân tích thông minh
- Layout hai cột cho slides ảnh
```

### Các môn học được hỗ trợ
- **Sinh học**: Sơ đồ sinh học, tế bào, DNA
- **Vật lý**: Diagram vật lý, quang học
- **Hóa học**: Cấu trúc phân tử, phản ứng
- **Toán học**: Hình học, đại số
- **Marketing**: Infographics doanh nghiệp
- **Lịch sử**: Minh họa lịch sử, văn hóa

## 🔧 Tính năng đã fix

### ✅ Layout Fix
- Ảnh không còn tràn ra ngoài slide
- Layout hai cột ổn định
- Tự động resize ảnh khi cần

### ✅ Content Fix  
- Không còn cắt nội dung vô nghĩa như "...trong..."
- Smart truncation theo từ và câu
- Đảm bảo nội dung hoàn chỉnh

### ✅ Image Fix
- Placeholder thông minh khi ảnh lỗi
- Vẫn hiển thị nội dung đầy đủ
- Xử lý lỗi ảnh graceful

## 🐛 Troubleshooting

### Lỗi thường gặp

1. **"❌ Cần API key"**
   - Nhập OpenAI API key vào sidebar
   - Đảm bảo API key hợp lệ

2. **"❌ Lỗi DALL-E"**  
   - Kiểm tra credit OpenAI
   - Thử lại với chủ đề khác

3. **"Ảnh không hiển thị"**
   - Bình thường! Sẽ hiển thị placeholder
   - File PowerPoint vẫn được tạo thành công

### Performance tips
- Chủ đề ngắn gọn cho kết quả tốt hơn
- Restart app nếu gặp lỗi memory
- Xóa ảnh cũ trong `dalle_images/` để tiết kiệm dung lượng

## 📊 Thông số kỹ thuật

- **Python**: 3.7+
- **Streamlit**: Web interface
- **OpenAI**: GPT-3.5 + DALL-E
- **python-pptx**: Tạo PowerPoint
- **File size**: ~50MB (đã tối ưu từ 100MB+)

## 🆕 Version History

### v2.0 (Latest) - Đã tối ưu
- ✅ Chuyển toàn bộ interface sang tiếng Việt
- ✅ Fix layout slides ảnh
- ✅ Smart content truncation  
- ✅ Tối ưu cấu trúc project (giảm 50% file)
- ✅ Cải thiện xử lý lỗi

### v1.0 - Version gốc
- Tính năng cơ bản tạo PowerPoint
- DALL-E integration
- Multiple templates

## 👨‍💻 Phát triển

Để contribute hoặc customize:

1. Fork project
2. Tạo branch mới: `git checkout -b feature/ten-tinh-nang`
3. Code và test
4. Commit: `git commit -m "Thêm tính năng XYZ"`
5. Push: `git push origin feature/ten-tinh-nang`
6. Tạo Pull Request

## 📞 Hỗ trợ

Nếu gặp vấn đề:
1. Kiểm tra log trong terminal
2. Restart ứng dụng bằng nút "🔄 Khởi động lại"
3. Xóa cache: `Remove-Item __pycache__ -Recurse -Force`

---

**Made with ❤️ using OpenAI GPT + DALL-E** 