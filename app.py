from flask import Flask, render_template, request, send_file
from werkzeug.utils import secure_filename
import os
import tempfile
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Inches, Pt
from PIL import Image

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = tempfile.mkdtemp()
app.config['TEMPLATE_FILE'] = 'shubh.pptx'

ALLOWED_EXTENSIONS = {'png', 'jpg', 'jpeg'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def resize_image(path, w, h):
    img = Image.open(path)
    img = img.resize((int(w * 96), int(h * 96)), Image.Resampling.LANCZOS)
    return img

# -------- SAFE TEXT REPLACEMENT FUNCTION --------
def replace_text_in_shape(shape, old_text, new_text):
    if not shape.has_text_frame:
        return

    text_frame = shape.text_frame
    for paragraph in text_frame.paragraphs:
        if old_text in paragraph.text:
            paragraph.text = paragraph.text.replace(old_text, new_text)
            for run in paragraph.runs:
                run.font.size = Pt(8)
                run.font.name = "Segoe UI"

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/generate', methods=['POST'])
def generate_ppt():
    try:
        # -------- FORM DATA --------
        project_title = request.form['projectTitle']
        project_domain = request.form['projectDomain']
        guide_name = request.form['guideName']
        applications = request.form['applications']
        brief_idea = request.form['briefIdea']
        student1 = request.form['student1']
        student2 = request.form['student2']
        student3 = request.form.get('student3', '')

        green_image = request.files.get('greenAreaPhoto')
        screenshots = request.files.getlist('projectScreenshots')

        # -------- VALIDATION --------
        if not green_image or not allowed_file(green_image.filename):
            return "Green area image is required.", 400

        if len(screenshots) != 2:
            return "Please upload exactly 2 screenshots.", 400

        for s in screenshots:
            if not allowed_file(s.filename):
                return "Invalid screenshot file type.", 400

        # -------- LOAD TEMPLATE --------
        prs = Presentation(app.config['TEMPLATE_FILE'])
        slide = prs.slides[0]

        # -------- TEXT REPLACEMENT (UPDATED & SAFE) --------
        for shape in slide.shapes:
            if shape.has_text_frame:
                replace_text_in_shape(shape, "{{PROJECT_TITLE}}", project_title)
                replace_text_in_shape(shape, "{{PROJECT_DOMAIN}}", project_domain)
                replace_text_in_shape(shape, "{{GUIDE_NAME}}", guide_name)
                replace_text_in_shape(shape, "{{APPLICATIONS}}", applications)
                replace_text_in_shape(shape, "{{PROJECT_DESCRIPTION}}", brief_idea)
                replace_text_in_shape(shape, "{{STUDENT_1}}", student1)
                replace_text_in_shape(shape, "{{STUDENT_2}}", student2)
                replace_text_in_shape(
                    shape,
                    "{{STUDENT_3}}",
                    student3 if student3 else ""
                )

        # -------- COLLECT IMAGE PLACEHOLDERS --------
        green_placeholder = None
        screenshot_placeholders = []

        for shape in slide.shapes:
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                if not green_placeholder:
                    green_placeholder = shape
                else:
                    screenshot_placeholders.append(shape)

        if len(screenshot_placeholders) < 2:
            return "Template must contain 2 screenshot placeholders.", 500

        # -------- INSERT GREEN IMAGE --------
        green_path = os.path.join(
            app.config['UPLOAD_FOLDER'],
            secure_filename(green_image.filename)
        )
        green_image.save(green_path)

        green_img = resize_image(green_path, 2.44, 1.11)
        green_resized = green_path.replace('.', '_green.')
        green_img.save(green_resized)

        slide.shapes._spTree.remove(green_placeholder._element)
        slide.shapes.add_picture(
            green_resized,
            green_placeholder.left,
            green_placeholder.top,
            green_placeholder.width,
            green_placeholder.height
        )

        # -------- INSERT 2 SCREENSHOTS --------
        for i in range(2):
            file = screenshots[i]
            placeholder = screenshot_placeholders[i]

            img_path = os.path.join(
                app.config['UPLOAD_FOLDER'],
                secure_filename(file.filename)
            )
            file.save(img_path)

            resized = resize_image(img_path, 1.59, 0.89)
            resized_path = img_path.replace('.', '_shot.')
            resized.save(resized_path)

            slide.shapes._spTree.remove(placeholder._element)
            slide.shapes.add_picture(
                resized_path,
                placeholder.left,
                placeholder.top,
                placeholder.width,
                placeholder.height
            )

        # -------- SAVE PPT --------
        output = os.path.join(app.config['UPLOAD_FOLDER'], 'output.pptx')
        prs.save(output)

        return send_file(
            output,
            as_attachment=True,
            download_name='project_presentation.pptx'
        )

    except Exception as e:
        return f"Error generating PPT: {e}", 500

if __name__ == '__main__':
    app.run()
