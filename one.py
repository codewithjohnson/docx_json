import json
from docx import Document

def convert_docx_to_json(file_path):
    doc = Document(file_path)
    data = []
    question_number = 1
    current_question = None
    option_map = {
        "A.": "option_a",
        "B.": "option_b",
        "C.": "option_c",
        "D.": "option_d"
    }

    for paragraph in doc.paragraphs:
        if question_number <= 50:
            if paragraph.text.startswith(str(question_number) + "."):
                if current_question is not None:
                    data.append(current_question)

                current_question = {
                    "id": question_number,
                    "question": paragraph.text.split(".", 1)[1].strip(),
                    "option_a": "",
                    "option_b": "",
                    "option_c": "",
                    "option_d": "",
                    "answer": "",
                    "topic": "",
                    "repeated": "",
                    "year": ""
                }
                question_number += 1
            elif current_question is not None:
                for option_key, option_value in option_map.items():
                    if paragraph.text.startswith(option_key):
                        current_question[option_value] = paragraph.text.split(".", 1)[1].strip()

    if current_question is not None:
        data.append(current_question)

    json_data = json.dumps(data, indent=4)
    return json_data

# Usage
docx_file = "crs_2015.docx"
json_data = convert_docx_to_json(docx_file)
print(json_data)
