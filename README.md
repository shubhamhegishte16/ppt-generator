# 📊 PPT Final Editor

A Flask-based web application that automates the creation of project presentation slides. By filling out a simple form and uploading relevant images, users can generate a professional PowerPoint presentation based on a predefined template.

## 🚀 Features

- **Automated Text Replacement**: Automatically replaces placeholders like `{{PROJECT_TITLE}}`, `{{STUDENT_1}}`, etc., in the PowerPoint template.
- **Image Insertion**: Specifically positions uploaded guide photos and project screenshots within the slides.
- **Dynamic Styling**: Automatically adjusts font sizes for key fields (e.g., Program Name).
- **Word Limit Enforcement**: Limits the project description to 100 words to ensure it fits perfectly on the slide.
- **Downloadable Output**: Generates and serves the edited `.pptx` file directly for download.

## 🛠️ Tech Stack

- **Backend**: Python, Flask
- **PPT Manipulation**: `python-pptx`
- **Image Processing**: `Pillow`
- **Web Server**: `gunicorn` (for deployment)
- **Frontend**: HTML/CSS (Jinja2 Templates)

## ⚙️ Installation & Setup

1.  **Clone the repository**:
    ```bash
    git clone <repository-url>
    cd ppt-final-editor
    ```

2.  **Create a virtual environment** (optional but recommended):
    ```bash
    python -m venv venv
    source venv/bin/activate  # On Windows: venv\Scripts\activate
    ```

3.  **Install dependencies**:
    ```bash
    pip install -r requirements.txt
    ```

4.  **Run the application**:
    ```bash
    python app.py
    ```
    The app will be available at `http://127.0.0.1:5000/`.

## 📖 Usage

1.  Open the web interface.
2.  Fill in the project details (Program Name, Title, Domain, Guide, etc.).
3.  Add student names and their respective application roles.
4.  Upload a guide photo and 1 to 3 project screenshots.
5.  Click **Generate** to receive your customized PowerPoint presentation.

## ☁️ Deployment

The project includes a `Procfile` and is ready for deployment on platforms like Heroku.

```text
web: gunicorn app:app
```

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details (if applicable).
