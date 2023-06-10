from docx import Document
from unidecode import unidecode
import json


def convert_docx_to_json(docx_path):
    doc = Document(docx_path)
    questions = []
    question = {}
    question_number = 0
    current_topic = ""
    current_year = ""
    current_explanation = ""
    has_image = False
    images = []

    for paragraph in doc.paragraphs:
        line = unidecode(paragraph.text.strip())

        if line.lower() == "image: yes":
            has_image = True
        elif line.lower() == "image: no":
            has_image = False

        if has_image and line.startswith("Image URL:"):
            image_url = line.split("Image URL:")[1].strip()
            images.append(image_url)

        if line.startswith(str(question_number + 1) + "."):
            if question:
                question["topic"] = current_topic
                question["year"] = current_year
                question["explanation"] = current_explanation
                question["has_image"] = has_image
                if has_image:
                    question["images"] = images
                questions.append(question)

            question_number += 1
            question = {
                "id": question_number,
                "question": line[len(str(question_number)) + 1:].strip(),
                "a": "",
                "b": "",
                "c": "",
                "d": "",
                "answer": "",
                "topic": "",
                "year": "",
                "explanation": ""
            }
            current_topic = ""
            current_year = ""
            current_explanation = ""
            has_image = False
            images = []

        elif line.startswith("a."):
            question["a"] = line[3:].strip()

        elif line.startswith("b."):
            question["b"] = line[3:].strip()

        elif line.startswith("c."):
            question["c"] = line[3:].strip()

        elif line.startswith("d."):
            question["d"] = line[3:].strip()

        elif line.startswith("topic."):
            current_topic = line.split("topic.")[1].strip()

        elif line.startswith("year."):
            current_year = line.split("year.")[1].strip()

        elif line.startswith("explanation."):
            current_explanation = line.split("explanation.")[1].strip()

        elif line.startswith("answer."):
            answer_letter = line.split("answer.")[1].strip().lower()
            question["answer"] = answer_letter

    if question:
        question["topic"] = current_topic
        question["year"] = current_year
        question["explanation"] = current_explanation
        question["has_image"] = has_image
        if has_image:
            question["images"] = images
        questions.append(question)

    json_data = json.dumps(questions, indent=4)
    return json_data


# Provide the path to the input Word document (.docx)
docx_path = "obj-test.docx"

# Convert the Word document to JSON
json_data = convert_docx_to_json(docx_path)

# Save the JSON data to a file
output_file = "output.json"
with open(output_file, "w") as f:
    f.write(json_data)
