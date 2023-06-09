from docx import Document
from unidecode import unidecode
import json

def convert_docx_to_json(docx_path):
    doc = Document(docx_path)
    questions = []
    question_number = 1

    for paragraph in doc.paragraphs:
        line = unidecode(paragraph.text.strip())

        if line.startswith(str(question_number) + "."):
            question = {
                "id": question_number,
                "question": line[len(str(question_number))+1:].strip(),
                "option_a": "",
                "option_b": "",
                "option_c": "",
                "option_d": "",
                "answer": "",
                "topic": "",
                "repeated": "",
                "year": ""
            }
            questions.append(question)
            question_number += 1

        elif line.startswith("a."):
            questions[-1]["option_a"] = line[3:].strip()

        elif line.startswith("b."):
            questions[-1]["option_b"] = line[3:].strip()

        elif line.startswith("c."):
            questions[-1]["option_c"] = line[3:].strip()

        elif line.startswith("d."):
            questions[-1]["option_d"] = line[3:].strip()

        elif line.startswith("answer."):
            answer_letter = line.split("answer.")[1].strip().lower()
            answer_mapping = {
                "a": "option_a",
                "b": "option_b",
                "c": "option_c",
                "d": "option_d"
            }
            questions[-1]["answer"] = answer_mapping.get(answer_letter, "")

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
