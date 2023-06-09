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

    for paragraph in doc.paragraphs:
        line = unidecode(paragraph.text.strip())

        if line.startswith(str(question_number + 1) + "."):
            if question:
                question["topic"] = current_topic
                question["year"] = current_year
                question["explanation"] = current_explanation
                questions.append(question)

            question_number += 1
            question = {
                "id": question_number,
                "question": line[len(str(question_number)) + 1:].strip(),
                "option_a": "",
                "option_b": "",
                "option_c": "",
                "option_d": "",
                "answer": "",
                "topic": "",
                "year": "",
                "explanation": ""
            }
            current_topic = ""
            current_year = ""
            current_explanation = ""

        elif line.startswith("a."):
            question["option_a"] = line[3:].strip()

        elif line.startswith("b."):
            question["option_b"] = line[3:].strip()

        elif line.startswith("c."):
            question["option_c"] = line[3:].strip()

        elif line.startswith("d."):
            question["option_d"] = line[3:].strip()

        elif line.startswith("topic."):
            current_topic = line.split("topic.")[1].strip()

        elif line.startswith("year."):
            current_year = line.split("year.")[1].strip()

        elif line.startswith("explanation."):
            current_explanation = line.split("explanation.")[1].strip()

        elif line.startswith("answer."):
            answer_letter = line.split("answer.")[1].strip().lower()
            answer_mapping = {
                "a": "option_a",
                "b": "option_b",
                "c": "option_c",
                "d": "option_d"
            }
            for option_key, option_value in answer_mapping.items():
                if answer_letter == option_key:
                    question["answer"] = option_value

    if question:
        question["topic"] = current_topic
        question["year"] = current_year
        question["explanation"] = current_explanation
        questions.append(question)

    json_data = json.dumps(questions, indent=4)
    return json_data


# Provide the path to the input Word document (.docx)
docx_path = "test.docx"

# Convert the Word document to JSON
json_data = convert_docx_to_json(docx_path)

# Save the JSON data to a file
output_file = "output.json"
with open(output_file, "w") as f:
    f.write(json_data)
