import os
import pandas as pd
from openpyxl import load_workbook
from flask_sqlalchemy import SQLAlchemy
from werkzeug.utils import secure_filename
from openpyxl.drawing.image import Image as OpenpyxlImage
from flask import Flask, request, render_template, flash, redirect, url_for


app = Flask(__name__)

app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///questions.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['UPLOAD_FOLDER'] = 'static/uploads/'
app.config['SECRET_KEY'] = 'afsarkhan'

db = SQLAlchemy(app)

class Question(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    previous_year = db.Column(db.String(50))
    level = db.Column(db.String(50))
    question_text = db.Column(db.String(255))
    question_image = db.Column(db.String(255), nullable=True)
    option1_text = db.Column(db.String(255), nullable=True)
    option1_image = db.Column(db.String(255), nullable=True)
    option2_text = db.Column(db.String(255), nullable=True)
    option2_image = db.Column(db.String(255), nullable=True)
    option3_text = db.Column(db.String(255), nullable=True)
    option3_image = db.Column(db.String(255), nullable=True)
    option4_text = db.Column(db.String(255), nullable=True)
    option4_image = db.Column(db.String(255), nullable=True)
    explanation = db.Column(db.Text)
    answer = db.Column(db.String(50))

with app.app_context():
    db.create_all()




@app.route('/upload', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        file = request.files['file']
        if file and file.filename.endswith('.xlsx'):
            filename = secure_filename(file.filename)
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(file_path)

            import_from_excel(file_path)

            flash('Data has been imported successfully', 'success')
            return redirect(url_for('upload_file'))

    return render_template('upload.html')




def save_image_from_excel(img, folder, image_name):
    if img:
        image_path = os.path.join(app.config['UPLOAD_FOLDER'], folder)
        if not os.path.exists(image_path):
            os.makedirs(image_path)

        image_file_path = os.path.join(image_path, f"{image_name}.png")

        with open(image_file_path, 'wb') as f:
            f.write(img._data()) 

        return image_file_path
    return None


def import_from_excel(file_path):
    wb = load_workbook(filename=file_path)
    sheet = wb.active

    for i, row in enumerate(sheet.iter_rows(min_row=2, max_row=sheet.max_row, values_only=False), start=1):
        previous_year = row[0].value
        level = row[1].value
        question_text = row[2].value
        explanation = row[7].value
        answer = row[8].value

        question_image = None
        if sheet._images:
            for image in sheet._images:
                if image.anchor._from.row == i and image.anchor._from.col == 2:
                    question_image = image

        option1_image, option2_image, option3_image, option4_image = None, None, None, None
        for image in sheet._images:
            if image.anchor._from.row == i:
                if image.anchor._from.col == 3:
                    option1_image = image
                elif image.anchor._from.col == 4: 
                    option2_image = image
                elif image.anchor._from.col == 5: 
                    option3_image = image
                elif image.anchor._from.col == 6:
                    option4_image = image

        question_image_path = save_image_from_excel(question_image, 'questions', f'question_{i}')
        option1_image_path = save_image_from_excel(option1_image, 'options', f'option1_{i}')
        option2_image_path = save_image_from_excel(option2_image, 'options', f'option2_{i}')
        option3_image_path = save_image_from_excel(option3_image, 'options', f'option3_{i}')
        option4_image_path = save_image_from_excel(option4_image, 'options', f'option4_{i}')

        question = Question(
            previous_year=previous_year,
            level=level,
            question_text=question_text if isinstance(question_text, str) else None,
            question_image=question_image_path,
            option1_text=row[3].value if isinstance(row[3].value, str) else None,
            option1_image=option1_image_path,
            option2_text=row[4].value if isinstance(row[4].value, str) else None,
            option2_image=option2_image_path,
            option3_text=row[5].value if isinstance(row[5].value, str) else None,
            option3_image=option3_image_path,
            option4_text=row[6].value if isinstance(row[6].value, str) else None,
            option4_image=option4_image_path,
            explanation=explanation,
            answer=answer
        )
        db.session.add(question)

    db.session.commit()




@app.route('/questions')
def display():
    questions = Question.query.all()
    for question in questions:
        if question.question_image:
            print(question.question_image)
    return render_template('questions.html', questions = questions)


if __name__ == "__main__":
    app.run(debug=True)