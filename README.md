
# 🎤 AI-Driven Presentation Generator 📊  

![Python](https://img.shields.io/badge/Python-3.x-blue?style=flat&logo=python)  
![FastAPI](https://img.shields.io/badge/FastAPI-0.95.2-green?style=flat&logo=fastapi)  
![OpenAI API](https://img.shields.io/badge/OpenAI-API-orange?style=flat&logo=openai)  
![License](https://img.shields.io/github/license/your-repo-name/AI-Presentation-Generator)  

## 📝 Overview  
This project leverages **OpenAI's GPT-4** and **Python-pptx** to generate PowerPoint presentations dynamically based on user-defined topics. Users can select from multiple templates, and the app supports speech recognition for hands-free input. Built with **FastAPI**, the tool makes presentations quickly and efficiently.

---

## 🚀 Features  
- **AI-Powered Content Generation**: Generate slides with detailed outlines and bullet points.  
- **Custom Templates**: Choose between templates such as Minimalist, Corporate, Creative, Innovative, and Aster.  
- **Voice Input Support**: Use Google Speech Recognition for hands-free topic selection.  
- **REST API with FastAPI**: Easily integrate the presentation generator with other applications.  
- **Image Handling**: Include Q&A and Thank You images on specific slides for enhanced presentation aesthetics.

---

## 🛠️ Installation  

1. **Clone the repository**:  
   ```bash
   git clone https://github.com/AlishbahGhazanfar/AI-Presentation-Generator.git
   cd AI-Presentation-Generator
   ```

2. **Install the dependencies**:  
   ```bash
   pip install -r requirements.txt
   ```

## 🖥️ Usage  

1. **Run the FastAPI server**:  
   ```bash
   uvicorn main:app --reload
   ```

2. **Generate a Presentation**:  
   Send a **POST** request to:  
   ```
   http://localhost:8000/generate_presentation/
   ```
   **Parameters**:
   - `topic` (str): Topic of the presentation  
   - `slide_count` (int): Number of slides (between 5 and 18)  
   - `template_choice` (str): Template name (e.g., Minimalist, Corporate)  

3. **Access the root endpoint**:  
   ```
   http://localhost:8000/
   ```
   It returns a welcome message confirming the server is running.

---

## 📂 Project Structure  
```
📦 AI-Presentation-Generator  
├── main.py                  # Main FastAPI application  
├── requirements.txt         # Python dependencies  
├── templates/               # PPTX template files  
├── images/                  # Image assets (Q&A, Thank You)  
└── README.md                # Documentation  
```

---

## 🔑 API Endpoints  

- **POST /generate_presentation/**:  
  Generate a PowerPoint presentation with the provided topic and template.  

- **GET /**:  
  Root endpoint returning a welcome message.


---

## 💬 Contributing  
Contributions are welcome! Feel free to open issues or submit pull requests to improve this project.

---

