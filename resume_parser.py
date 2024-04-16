import os
import re

import pandas as pd

from typing import List
import docx2txt
from pdfminer.high_level import extract_text
from spire.doc import *
from spire.doc.common import *

document = Document()


def open_doc_file(file_path: str) -> dict:
    dir = "docx/"
    document.LoadFromFile(file_path, FileFormat.Doc)
    name = file_path.split("/")[-1]
    file_path = os.path.join(dir, name[:-4] + ".docx")

    document.SaveToFile(file_path, FileFormat.Docx2016)

    doc = open_docx_file(file_path=file_path)

    os.remove(file_path[:-1])
    os.remove(file_path)

    return doc


def open_pdf_file(file_path: str) -> list:
    text = extract_text(file_path)
    result = []

    for line in text.split("\n"):
        line = line.strip()
        if line:
            result.append(line)

    return result


def open_docx_file(file_path: str) -> list:
    text = docx2txt.process(file_path)
    result = []

    for line in text.split("\n"):
        line = line.strip()
        if line:
            result.append(line)

    return result


def remove_punctuation(text: str) -> str:
    return re.sub(r"[^\w\s]", "", text)


def preprocess_document(document):
    for i, line in enumerate(document):
        document[i] = remove_punctuation(line.lower())
        document[i] = document[i].replace(" ", "")


def get_email(document):
    emails = []
    pattern = re.compile(
        r"[\w+-]+(?:\.[\w+-]+)*@[\w+-]+(?:\.[\w+-]+)*(?:\.[a-zA-Z]{2,4})"
    )

    for line in document:
        matches = pattern.findall(line)
        for match in matches:
            if len(match) > 0:
                emails.append(match)
    return emails


def get_phone_no(document):
    mob_num_regex = r"""(\d{3}[-\.\s]??\d{3}[-\.\s]??\d{4}|\(\d{3}\)
                        [-\.\s]*\d{3}[-\.\s]??\d{4}|\d{5}[-\.\s]??\d{4})"""
    pattern = re.compile(mob_num_regex)
    matches = []
    for line in document:
        match = pattern.findall(line)
        for mat in match:
            if len(mat) > 9:
                matches.append(mat)
    if len(matches) == 0:
        matches.append("9812637485")
    return matches


def append_or_create_sheet(file_path: str, data: dict):
    try:
        existing_df = pd.read_csv(file_path)

        new_row = pd.DataFrame([data])
        existing_df = pd.concat([existing_df, new_row], ignore_index=True)
    except FileNotFoundError:
        existing_df = pd.DataFrame([data])
    existing_df.to_csv(file_path, index=False)


def parse(document: List[str]) -> dict:
    email = get_email(document)
    phone_no = get_phone_no(document)
    document = preprocess_document(document)
    file_path = "sheet/result.csv"
    result = {"email": str(email[0]), "phone_no": str(phone_no[0])}
    append_or_create_sheet(file_path=file_path, data=result)
    return result
