import pandas as pd
from docx import Document
from typing import List, Tuple, Dict
import os
import datetime
import time
from tqdm import tqdm


def get_data(data_path: str) -> Tuple[int, pd.DataFrame, List[str]]:
    df = pd.read_excel(data_path, engine='openpyxl')
    return len(df), df, df.columns.tolist()


def fill_invitations(template_path: str, data, key_format: str) -> Document:
    doc = Document(template_path)
    for paragraph in doc.paragraphs:
        for key, value in data.items():
            if " " not in str(key_format):
                key = key.replace(' ', '_')
            key_format = key_format.replace("_", "")
            key = key_format.replace("Keyword", key)
            opening_bracket = key_format[0]
            closing_bracket = key_format[-1]

            if key in paragraph.text:
                temp = ""
                for run in paragraph.runs:
                    # print(f"Run text:", run.text)
                    if run.text == opening_bracket or temp != "":
                        # print("Temp open:", temp)
                        temp += run.text
                        run.clear()

                    if opening_bracket in run.text and closing_bracket not in run.text:
                        # print("Temp open 2:", temp)
                        temp_2 = run.text.split(opening_bracket)
                        # print("Split list:", temp_2)
                        run.text = temp_2[0]
                        # print("Curr run: ", run.text)
                        temp += (opening_bracket + temp_2[1])

                    if opening_bracket in run.text and closing_bracket in run.text:
                        run.text = run.text.replace(f"{key}", str(value))

                    if closing_bracket in temp:
                        # print("Temp closing:", temp)
                        run.text = temp + run.text
                        # print("Addition:", run.text)
                        run.text = run.text.replace(f"{key}", str(value))
                        # print("End:", run.text)
                        temp = ""
                    # print()
                    # print(paragraph.text)
                    # print("-" * 50)

    # Process tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:  # Iterate through paragraphs inside the cell
                    # Iterate through runs inside the paragraph
                    for key, value in data.items():
                        if " " not in str(key_format):
                            key = key.replace(' ', '_')
                        key_format = key_format.replace("_", "")
                        key = key_format.replace("Keyword", key)
                        opening_bracket = key_format[0]
                        closing_bracket = key_format[-1]

                        temp = ""
                        for run in paragraph.runs:
                            if opening_bracket in run.text and closing_bracket in run.text:
                                run.text = run.text.replace(f"{key}", str(value))
                                continue

                            if run.text == opening_bracket or temp != "":
                                temp += run.text
                                run.clear()

                            if opening_bracket in run.text and closing_bracket not in run.text:
                                temp_2 = run.text.split(opening_bracket)
                                run.text = temp_2[0]
                                temp += (opening_bracket + temp_2[1])

                            if closing_bracket in temp:
                                run.text = temp + run.text
                                run.text = run.text.replace(f"{key}", str(value))
                                temp = ""
    return doc


def merge_documents(documents: List[Document]) -> Document:
    merged_doc = Document()
    for i, doc in enumerate(documents):
        # Append each element of the current document to the merged document
        if i != len(documents) - 1:
            doc.add_page_break()

        for element in doc.element.body:
            merged_doc.element.body.append(element)

    return merged_doc



def main():
    pass

if __name__ == "__main__":
    main()
