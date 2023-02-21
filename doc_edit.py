import json
from docx import Document
from docx.opc.exceptions import PackageNotFoundError


class DocEdit:
    def open_document(self, dir: str) -> Document:
        """
        Open the document .docx to be edited.

        :param dir: document's directory
        :return .docx file
        """
        try:
            document = Document(f'{dir}.docx')
        except (PackageNotFoundError, AttributeError) as error:
            return print(f'{error}')

        return document

    def create_form(self, save_dir: str, keys: list):
        """
        Create a .json form to be used later to update a document
        with the function update_document()

        :param save_dir: file's save directory
        :param *args: keys of the form
        """
        form = {}

        for i in keys:
            form[f'{i}_'] = ''

        try:
            with open(save_dir, 'w') as fp:
                json.dump(form, fp)

        except (FileNotFoundError, PermissionError, TypeError) as error:
            return print(f'{error}')

    def upload_form(self, dir: str):
        """
        Upload a .json form with already saved keys that
        can be used as a parameter in update_document() function.

        :param dir: form's directory
        :return dict
        """
        try:
            with open(dir, 'r') as fp:
                data = json.load(fp)

                return data

        except (FileNotFoundError, IndexError) as error:
            return print(f'{error}')

    def update_document(self, document, save_dir: str, form=None, **kwargs):
        """
        It generates a .docx document with the form keys filled in. Or keys
        and values can be passed as parameters in dictionary format.

        :param document: .docx file
        :param save_dir: file's save directory
        :param form: ready form
        :param **kwargs: Key and values parameters
        """
        try:
            if form is not None:
                fields = form
            else:
                fields = {}
                fields.update(kwargs)
                fields = fields['kwargs']

            for paragraph in document.paragraphs:
                for field in fields:
                    value = fields[field]
                    paragraph.text = paragraph.text.replace(
                        field, str(value)
                    )

            document.save(f'{save_dir}.docx')

        except (FileNotFoundError, KeyError, AttributeError) as error:
            return print(f'{error}')


if __name__ == '__main__':
    # Init the program
    doc_edit = DocEdit()
    document = doc_edit.open_document("<document's directory>")

    # Create a form
    doc_edit.create_form(
        "<form's save directory>",
        keys=[
            'name', 'age', 'email'  # form's fields
        ]
    )

    # Upload created form
    form = doc_edit.upload_form(
        "<form's directory>",
    )

    # Updating the uploaded form
    for key, value in form.items():
        value = str(input(f'{key}: '))
        form[key] = value

    # Using the uploaded form to fill the document fields
    doc_edit.update_document(
        document,
        "<edited document's directory>",
        form
    )
